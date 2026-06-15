# -*- coding: utf-8 -*-
"""
Bowman Thermometry — serial read-out, calibration check, live logging and Excel export.
Intented to work with the Linduino and the four-wire readout of Bowman-Sensors.
This is a updated version including UI and direct plotting that was modified with Claude Opus 4.8 to simplify the structure and increase readability.
Run:
    python bowman_thermometry.py

Dependencies: numpy, pandas, matplotlib, pyserial, xlsxwriter
    (scipy is only required if the calibration ``.npy`` files contain scipy spline
    objects, which are unpickled by ``numpy.load``).

@author: Benjamin Kahlert
"""

import os
import sys
import time
import queue
import datetime
import threading
import configparser
from dataclasses import dataclass, field
from functools import partial

import numpy as np
import pandas as pd
import serial
from serial.tools import list_ports
import tkinter as tk
from tkinter import messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


# ===================== CONSTANTS =====================
SERIAL_BAUD = 115200
SENSOR_S1 = "Sensor4"           # device identifier mapped to sensor 1
SENSOR_S2 = "Sensor9"           # device identifier mapped to sensor 2
LINE_DELIMITER = "_"            # device line layout: <idx>_<ms>_<sensor>_<value>
INVALID_READING = 2 ** 21 - 1  # 2097151: sensor out-of-range / error sentinel
HEADING_FONT = ("TkDefaultFont", 16)


def style_window(window, title, on_close=None):
    """Apply the shared look (title, raise, briefly keep on top) to a Tk window."""
    window.title(title)
    window.lift()
    window.attributes('-topmost', True)
    window.after(100, lambda: window.attributes('-topmost', False))
    if on_close is not None:
        window.protocol("WM_DELETE_WINDOW", on_close)


# ===================== STATE =====================
@dataclass
class MeasurementState:
    """Calibration inputs and results shared across the measurement steps."""
    cal1: object = None                                          # spline ln(R)->T, sensor 1
    cal2: object = None                                          # spline ln(R)->T, sensor 2
    cal_paths: list = field(default_factory=lambda: [None, None])
    sensor_names: list = field(default_factory=lambda: [None, None])
    R_cal: list = field(default_factory=lambda: [None, None])   # resistance at calibration
    T_cal: object = field(default_factory=lambda: [None, None])  # uncorrected T at calibration
    T_ref: object = field(default_factory=lambda: [None, None])  # reference T at calibration
    cal_delta: object = field(default_factory=lambda: [0, 0])   # T_ref - T_cal, added to transform


# ===================== CONFIG =====================
class ConfigManager:
    """Load and persist ``config.ini`` located next to this script."""

    def __init__(self):
        self.config_path = os.path.join(
            os.path.dirname(os.path.realpath(__file__)), 'config.ini')
        self.load_config()

    def load_config(self):
        self.current_config = configparser.ConfigParser()
        self.current_config.read(self.config_path)

    def save_config(self):
        with open(self.config_path, 'w') as configfile:
            self.current_config.write(configfile)


# ===================== MAIN APP =====================
class ThermometryUI:
    def __init__(self):
        self._ser = None
        self.configuration = ConfigManager()
        self.state = MeasurementState()

        # Single hidden root; every window below is a Toplevel sharing it.
        self.root = tk.Tk()
        self.root.withdraw()
        try:
            InitThermometry(self).run()         # select sensors + (optional) calibration files
            LinduinoHandler(self).run()         # connect to the serial device
            if self.state.cal1 is not None or self.state.cal2 is not None:
                CheckCalWindow(self).run()       # verify an existing calibration
            self.logger = ThermometryLoggerApp(self)
            self.logger.run()                   # live logging until stopped
            self.save_and_format()              # export to Excel
        finally:
            self.root.destroy()

    def set_serial(self, ser):
        self._ser = ser

    def get_serial(self):
        return self._ser

    def save_and_format(self, path=None):
        self.create_dataframe()
        self.create_merged_dataframe()
        self.create_overview_frame()

        date = self.logger.start_time.strftime("%Y.%m.%d")

        main_path = path or self.configuration.current_config['paths']['output_path']
        save_path = os.path.join(main_path, date)
        os.makedirs(save_path, exist_ok=True)

        i = 1
        while os.path.exists(os.path.join(save_path, f"Measurement {i}.xlsx")):
            i += 1

        file_path = os.path.join(save_path, f"Measurement {i}.xlsx")

        try:
            with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                self.df.to_excel(writer, sheet_name="Data", index=False)
                self.df2.to_excel(writer, sheet_name="Merged Data", index=False)
                self.df_info.to_excel(writer, sheet_name="Info", index=False)

            print(f"Data saved to: {file_path}")

        except Exception as e:
            print("Saving of data failed:")
            print(e)
            raise

    def create_dataframe(self):
        """Create a DataFrame with structured serial + system time data."""

        def pad_entries(entries, max_len, width=3):
            if len(entries) == 0:
                return [(np.nan,) * width] * max_len
            pad_len = max_len - len(entries)
            pad_element = (np.nan,) * len(entries[0])
            return list(entries) + [pad_element] * pad_len

        def transform_entries(entries, transform):
            # entries are (sys_time, device_time, raw value); the transform applies to the raw value only
            if not entries:
                return []
            if transform is None:
                return [(e[0], np.nan) for e in entries]
            values = np.asarray([e[2] for e in entries], dtype=float)
            transformed = np.asarray(transform(values), dtype=float)
            return [(e[0], t) for e, t in zip(entries, transformed)]

        # --- Extract ---
        s1_entries = self.logger.serial_logger_1.get_entries()
        s2_entries = self.logger.serial_logger_2.get_entries()
        user_entries = self.logger.user_logger.get_entries()

        s1_entries_t = transform_entries(s1_entries, self.logger.transform1)
        s2_entries_t = transform_entries(s2_entries, self.logger.transform2)

        max_len = max(len(s1_entries), len(s2_entries), len(user_entries))

        s1_entries = pad_entries(s1_entries, max_len)
        s1_entries_t = pad_entries(s1_entries_t, max_len, width=2)
        s2_entries = pad_entries(s2_entries, max_len)
        s2_entries_t = pad_entries(s2_entries_t, max_len, width=2)

        # --- User ---
        user_times, user_vals = zip(*user_entries) if user_entries else ([], [])
        user_times = list(user_times) + [np.nan] * (max_len - len(user_entries))
        user_vals = list(user_vals) + [np.nan] * (max_len - len(user_entries))

        # --- DataFrame ---
        df = pd.DataFrame({
            "sys time":       [e[0] for e in s1_entries],
            "time_S1":        [e[1] for e in s1_entries],
            "S1_raw":         [e[2] for e in s1_entries],
            "S1_transformed": [e[1] for e in s1_entries_t],

            "sys time 2":     [e[0] for e in s2_entries],
            "time_S2":        [e[1] for e in s2_entries],
            "S2_raw":         [e[2] for e in s2_entries],
            "S2_transformed": [e[1] for e in s2_entries_t],

            "user_time":  user_times,
            "user_value": user_vals,
        })

        # --- Relative time ---
        all_serial_times = pd.concat([df["time_S1"], df["time_S2"]]).dropna()
        if not all_serial_times.empty:
            t0 = all_serial_times.min()
            df["time_S1"] = (df["time_S1"] - t0) / 1000
            df["time_S2"] = (df["time_S2"] - t0) / 1000

        all_sys_times = pd.concat([df["sys time"], df["sys time 2"], df["user_time"]]).dropna()
        if not all_sys_times.empty:
            t0_sys = all_sys_times.min()
            df["sys time"] -= t0_sys
            df["sys time 2"] -= t0_sys
            df["user_time"] = df["user_time"] - t0_sys

        self.df = df

    def create_merged_dataframe(self):
        """Build a DataFrame with S1, S2, and user values on a single time axis."""
        # ---------------- Create individual series ----------------
        s1 = pd.Series(data=self.df["S1_raw"].values, index=self.df["time_S1"]).dropna()
        s2 = pd.Series(data=self.df["S2_raw"].values, index=self.df["time_S2"]).dropna()
        user = pd.Series(data=self.df["user_value"].values, index=self.df["user_time"]).dropna()

        # Synchronise serial device time and system time:
        if len(user) > 0:
            user_0 = np.nanmin(user.index.values)
            d1 = (self.df["sys time"] - user_0).abs()
            d2 = (self.df["sys time 2"] - user_0).abs()
            offset = None
            if d1.notna().any() and (not d2.notna().any() or d1.min() <= d2.min()):
                idx = d1.idxmin()
                offset = self.df.loc[idx, "time_S1"] - self.df.loc[idx, "sys time"]
            elif d2.notna().any():
                idx = d2.idxmin()
                offset = self.df.loc[idx, "time_S2"] - self.df.loc[idx, "sys time 2"]

            if offset is not None and not np.isnan(offset):
                user.index = user.index + offset

        # ---------------- Merge all timestamps ----------------
        all_times = np.unique(np.concatenate([s1.index.values, s2.index.values, user.index.values]))
        all_times = all_times[~np.isnan(all_times)]

        # ---------------- Reindex / align ----------------
        self.df2 = pd.DataFrame({
            "time": all_times,
            "S1_raw": s1.reindex(all_times).values,
            "S2_raw": s2.reindex(all_times).values,
            "user_value": user.reindex(all_times).values,
        })

    def create_overview_frame(self):
        """Build the 'Info' sheet summarising the run and the calibration."""
        def reformat_timedelta(td):
            total_seconds = int(td.total_seconds())
            hours, rem = divmod(total_seconds, 3600)
            minutes, seconds = divmod(rem, 60)
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

        categories1 = ['Date', 'Local time at M start', 'Local time at M end',
                       'M-duration (hh:mm:ss)', 'M-duration (Linduino time)/s', 'Info']
        values1 = [
            self.logger.start_time.strftime("%Y.%m.%d"),
            self.logger.start_time.strftime("%H:%M:%S"),
            self.logger.end_time.strftime("%H:%M:%S"),
            reformat_timedelta(self.logger.end_time - self.logger.start_time),
            self.df2['time'].max() - self.df2['time'].min(),
            'The first system time values can be misleading as the buffer is quickly '
            'read-out. If a calibration offset was determined it is automatically '
            'applied in the transform',
        ]

        categories2 = ['Identifier', 'calibration path', 'Calibration R', 'Calibration T', 'Reference T']
        valuesS1 = [self.state.sensor_names[0], self.state.cal_paths[0],
                    self.state.R_cal[0], self.state.T_cal[0], self.state.T_ref[0]]
        valuesS2 = [self.state.sensor_names[1], self.state.cal_paths[1],
                    self.state.R_cal[1], self.state.T_cal[1], self.state.T_ref[1]]

        info_frame1 = pd.DataFrame()
        info_frame1['General information'] = categories1
        info_frame1['Values'] = values1
        info_frame1[''] = None

        info_frame2 = pd.DataFrame()
        info_frame2['Measurement information'] = categories2
        info_frame2['Sensor 1'] = valuesS1
        info_frame2['Sensor 2'] = valuesS2
        self.df_info = info_frame1.join(info_frame2, how='outer')


# ===================== INIT WINDOW =====================
class InitThermometry:
    def __init__(self, app):
        self.app = app
        self.configuration = app.configuration
        self._init_window()
        self._init_variables()
        self._init_fields()
        self._populate_window()

    def run(self):
        self.app.root.wait_window(self.main_window)

    def _init_window(self):
        self.main_window = tk.Toplevel(self.app.root)
        style_window(self.main_window, 'Bowman Thermometry', self.exit_all)

    def _init_variables(self):
        prev = self.configuration.current_config['prev_settings']
        self.name_var_1 = tk.StringVar(value=prev['name1'])
        self.name_var_2 = tk.StringVar(value=prev['name2'])
        self.dir_var_1 = tk.StringVar(value=prev['cal1'])
        self.dir_var_2 = tk.StringVar(value=prev['cal2'])
        # run without a calibration file to collect data for a new calibration
        self.no_cal_var = tk.BooleanVar(value=False)

    def _init_fields(self):
        self.title_label = tk.Label(self.main_window, text='Bowman-Thermometry Selection Pane', font=HEADING_FONT)
        self.s1_label = tk.Label(self.main_window, text='S1')
        self.s2_label = tk.Label(self.main_window, text='S2')
        self.name1_label = tk.Label(self.main_window, text='Designation:')
        self.name2_label = tk.Label(self.main_window, text='Designation:')
        self.cal1_label = tk.Label(self.main_window, text='Calibration:')
        self.cal2_label = tk.Label(self.main_window, text='Calibration:')

        self.name1_entry = tk.Entry(self.main_window, textvariable=self.name_var_1)
        self.name2_entry = tk.Entry(self.main_window, textvariable=self.name_var_2)

        self.dir_entry_1 = tk.Entry(self.main_window, textvariable=self.dir_var_1, width=40)
        self.dir_entry_1.bind("<Button-1>", partial(self.open_directory, variable=self.dir_var_1))

        self.dir_entry_2 = tk.Entry(self.main_window, textvariable=self.dir_var_2, width=40)
        self.dir_entry_2.bind("<Button-1>", partial(self.open_directory, variable=self.dir_var_2))

        self.no_cal_check = tk.Checkbutton(
            self.main_window, text='No calibration file (collect calibration data)',
            variable=self.no_cal_var, command=self._toggle_cal_fields)

        self.submit_btn = tk.Button(self.main_window, text="Submit", command=self._start)
        self.exit_btn = tk.Button(self.main_window, text="Exit", command=self.exit_all)

    def _populate_window(self):
        self.title_label.grid(row=0, column=1, columnspan=2, pady=15)
        self.s1_label.grid(row=1, column=1, pady=15)
        self.s2_label.grid(row=1, column=2, pady=15)
        self.name1_label.grid(row=2, column=1)
        self.name2_label.grid(row=2, column=2)
        self.name1_entry.grid(row=3, column=1)
        self.name2_entry.grid(row=3, column=2)
        self.cal1_label.grid(row=4, column=1)
        self.cal2_label.grid(row=4, column=2)
        self.dir_entry_1.grid(row=5, column=1)
        self.dir_entry_2.grid(row=5, column=2)
        self.no_cal_check.grid(row=6, column=1, columnspan=2, pady=(8, 0))
        self.submit_btn.grid(row=7, column=1)
        self.exit_btn.grid(row=7, column=2)

    def open_directory(self, event=None, variable=None):
        if getattr(self, "_dialog_open", False):
            return "break"

        self._dialog_open = True
        try:
            path = filedialog.askopenfilename(
                initialdir=self.configuration.current_config['paths']['calibration_path'],
                filetypes=[("NumPy files", "*.npy *.npz"), ("All files", "*.*")]
            )
            if path:
                variable.set(path)
        finally:
            self._dialog_open = False

        return "break"

    def _toggle_cal_fields(self):
        field_state = "disabled" if self.no_cal_var.get() else "normal"
        self.dir_entry_1.config(state=field_state)
        self.dir_entry_2.config(state=field_state)

    def _start(self):
        no_cal = self.no_cal_var.get()
        try:
            if no_cal:
                cal1 = cal2 = None
            else:
                cal1 = np.load(self.dir_var_1.get(), allow_pickle=True).item()
                cal2 = np.load(self.dir_var_2.get(), allow_pickle=True).item()
            prev = self.configuration.current_config['prev_settings']
            if not no_cal:
                prev['cal1'] = self.dir_var_1.get()
                prev['cal2'] = self.dir_var_2.get()
            prev['name1'] = self.name_var_1.get()
            prev['name2'] = self.name_var_2.get()
            self.configuration.save_config()
        except Exception as e:
            messagebox.showerror("Load error", str(e))
            return

        self.app.state.cal1 = cal1
        self.app.state.cal2 = cal2
        self.app.state.cal_paths = [None, None] if no_cal else [self.dir_var_1.get(), self.dir_var_2.get()]
        self.app.state.sensor_names = [self.name_var_1.get(), self.name_var_2.get()]
        self.main_window.destroy()

    def exit_all(self):
        self.main_window.destroy()
        sys.exit()


# ===================== SERIAL HANDLER =====================
class LinduinoHandler:
    def __init__(self, app):
        self.app = app
        self.ser = None
        self._init_window()

    def run(self):
        self.app.root.wait_window(self.main_window)

    def _init_window(self):
        self.main_window = tk.Toplevel(self.app.root)
        style_window(self.main_window, 'Connecting', self.exit_all)
        tk.Label(self.main_window, text="Connecting to Device...").pack(padx=10, pady=10)
        # Probe after the window is shown so the GUI stays responsive
        self.main_window.after(100, self._try_connect)

    def _try_connect(self):
        for p in list_ports.comports():
            try:
                ser = serial.Serial(p.device, SERIAL_BAUD, timeout=1)
                time.sleep(1)

                line = ser.readline().decode(errors='ignore').strip()
                print("Probe:", p.device, line)

                if "Sensor" in line:
                    self.ser = ser
                    self.app.set_serial(ser)
                    tk.Label(self.main_window, text=f"Connected to {p.device}").pack(pady=5)
                    self.main_window.after(3000, self.main_window.destroy)
                    return  # Connected successfully
                else:
                    ser.close()

            except serial.SerialException:
                pass

        # If no device found, show retry dialog
        self._show_retry_dialog()

    def _show_retry_dialog(self):
        self.dialog = tk.Toplevel(self.main_window)
        self.dialog.title("Device not found")
        self.dialog.lift()
        self.dialog.attributes('-topmost', True)
        self.dialog.transient(self.main_window)  # Keep on top of main window
        self.dialog.grab_set()                   # Block interaction with main window
        self.dialog.focus_force()

        tk.Label(self.dialog, text="No valid serial device found.\nCheck cable & power.").pack(padx=10, pady=10)
        tk.Button(self.dialog, text="Retry", command=self.retry).pack(side="left", padx=10, pady=10)
        tk.Button(self.dialog, text="Exit", command=self.exit_all).pack(side="right", padx=10, pady=10)

        # Wait here until dialog is closed
        self.main_window.wait_window(self.dialog)

    def retry(self):
        self.dialog.destroy()
        # Retry after short delay to allow GUI update
        self.main_window.after(100, self._try_connect)

    def exit_all(self):
        ser = self.app.get_serial()
        if ser and ser.is_open:
            ser.close()
        self.main_window.destroy()
        sys.exit()


# ===================== CAL WINDOW =====================
class CheckCalWindow:
    def __init__(self, app):
        self.running = True
        self.app = app
        self.ser = app.get_serial()
        self._init_window()
        self._init_vars()
        self._init_objects()
        self._populate_window()
        self.read_ser()

    def run(self):
        self.app.root.wait_window(self.main_window)

    def _init_window(self):
        self.main_window = tk.Toplevel(self.app.root)
        style_window(self.main_window, 'Bowman Thermometry', self._skip)

    def safe_float(self, val):
        """Convert to float or return np.nan if empty/invalid."""
        try:
            return float(val)
        except (ValueError, TypeError):
            return np.nan

    def _init_vars(self):
        self.vcmd = (self.main_window.register(self.validate_number), "%P")
        self._buffer = bytearray()

        self.R1_var = tk.StringVar()
        self.R2_var = tk.StringVar()
        self.T1_var = tk.StringVar()
        self.T2_var = tk.StringVar()
        self.ref_T1_var = tk.StringVar()
        self.ref_T2_var = tk.StringVar()

    def _init_objects(self):
        self.title_label = tk.Label(self.main_window, text='Calibration check', font=HEADING_FONT)
        self.s1_label = tk.Label(self.main_window, text='S1')
        self.s2_label = tk.Label(self.main_window, text='S2')
        self.r_label = tk.Label(self.main_window, text='R')
        self.t_label = tk.Label(self.main_window, text='T[°C]')
        self.tref_label = tk.Label(self.main_window, text='T-ref[°C]')

        self.R1_entry = tk.Entry(self.main_window, textvariable=self.R1_var)
        self.R2_entry = tk.Entry(self.main_window, textvariable=self.R2_var)
        self.T1_entry = tk.Entry(self.main_window, textvariable=self.T1_var)
        self.T2_entry = tk.Entry(self.main_window, textvariable=self.T2_var)
        self.T1_ref_entry = tk.Entry(self.main_window, textvariable=self.ref_T1_var, validate="key", validatecommand=self.vcmd)
        self.T2_ref_entry = tk.Entry(self.main_window, textvariable=self.ref_T2_var, validate="key", validatecommand=self.vcmd)

        self.skip_btn = tk.Button(self.main_window, text="Skip", command=self._skip)
        self.save_btn = tk.Button(self.main_window, text="Save", command=self._save)

    def _populate_window(self):
        self.title_label.grid(row=0, column=1, columnspan=2, pady=15)
        self.s1_label.grid(row=1, column=1)
        self.s2_label.grid(row=1, column=2)
        self.r_label.grid(row=2, column=0)
        self.R1_entry.grid(row=2, column=1)
        self.R2_entry.grid(row=2, column=2)
        self.t_label.grid(row=3, column=0)
        self.T1_entry.grid(row=3, column=1)
        self.T2_entry.grid(row=3, column=2)
        self.tref_label.grid(row=4, column=0)
        self.T1_ref_entry.grid(row=4, column=1)
        self.T2_ref_entry.grid(row=4, column=2)
        self.skip_btn.grid(row=5, column=2, pady=10)
        self.save_btn.grid(row=5, column=3)

    def _save(self):
        self.running = False
        state = self.app.state
        state.R_cal = [self.safe_float(self.R1_var.get()), self.safe_float(self.R2_var.get())]
        state.T_ref = np.array([self.safe_float(self.ref_T1_var.get()), self.safe_float(self.ref_T2_var.get())])
        state.T_cal = np.array([self.safe_float(self.T1_var.get()), self.safe_float(self.T2_var.get())])
        # A missing reference or sensor reading must not poison the transform with NaN
        state.cal_delta = np.nan_to_num(state.T_ref - state.T_cal)
        self.main_window.destroy()

    def _skip(self):
        self.running = False
        state = self.app.state
        state.R_cal = [None, None]
        state.T_ref = [None, None]
        state.T_cal = [None, None]
        state.cal_delta = [0, 0]
        self.main_window.destroy()

    def read_ser(self):
        """Non-blocking poll: drain whatever is waiting and update both sensors."""
        try:
            if self.ser and self.ser.is_open:
                waiting = self.ser.in_waiting
                if waiting:
                    self._buffer.extend(self.ser.read(waiting))
                    while b'\n' in self._buffer:
                        line_bytes, _, rest = self._buffer.partition(b'\n')
                        self._buffer = bytearray(rest)
                        self._handle_line(line_bytes.decode(errors='ignore').strip())
        except Exception as e:
            print("Serial read error:", e)

        if self.running:
            self.main_window.after(200, self.read_ser)

    def _handle_line(self, line):
        parts = line.split(LINE_DELIMITER)
        if len(parts) < 4:
            return
        try:
            value = float(parts[3])
        except ValueError:
            return
        if value == INVALID_READING:
            value = np.nan
        if parts[2] == SENSOR_S1:
            self._show_reading(value, self.R1_var, self.T1_var, self.app.state.cal1)
        elif parts[2] == SENSOR_S2:
            self._show_reading(value, self.R2_var, self.T2_var, self.app.state.cal2)

    def _show_reading(self, value, r_var, t_var, cal):
        if np.isnan(value):
            r_var.set("nan")
            t_var.set("")
            return
        r_var.set(round(value, 2))
        try:
            t_var.set(round(float(cal(np.log(value))), 2))
        except Exception:
            t_var.set("")

    def validate_number(self, new_value):
        """Allow only a number between 0 and 100 (decimals allowed) in the ref-T fields."""
        if new_value in ("", "."):  # allow empty field and a leading decimal point
            return True
        try:
            return 0 <= float(new_value) <= 100
        except ValueError:
            return False


# ===================== LOGGER =====================
class NumericLogger:
    """Simple numeric logger for timestamped values."""
    def __init__(self):
        self.entries = []  # list of (timestamp, value) or (timestamp, device_ts, value)

    def add(self, value, timestamp=None):
        if timestamp is None:
            timestamp = time.time()
        if isinstance(value, (list, tuple)) and len(value) == 2:
            self.entries.append((timestamp, value[0], value[1]))
        else:
            self.entries.append((timestamp, value))

    def get_entries(self):
        # Return a copy: the serial worker thread may append while the GUI reads.
        return list(self.entries)

    def remove_last(self):
        if self.entries:
            return self.entries.pop()
        return None


class ThermometryLoggerApp:
    def __init__(self, app, display_interval=2.0):
        self.app = app
        self.ser = app.get_serial()
        self.display_interval = display_interval
        # Last display timestamps for the two serial listboxes
        self.last_display_time_1 = 0
        self.last_display_time_2 = 0

        self._init_window()
        self._init_variables()
        self._init_fields()
        self._populate_window()
        self._init_loggers()
        self._init_plot()
        self.main_window.after(500, self._update_plot)

    def run(self):
        if self.ser:
            self.serial_thread = threading.Thread(target=self._serial_worker, daemon=True)
            self.serial_thread.start()
        self._poll_queues()
        self.app.root.wait_window(self.main_window)

    def _init_loggers(self):
        self.start_time = datetime.datetime.now()
        self.user_logger = NumericLogger()
        self.serial_logger_1 = NumericLogger()
        self.serial_logger_2 = NumericLogger()

        self.queue1 = queue.Queue()
        self.queue2 = queue.Queue()

        self.stop_event = threading.Event()
        self._redraw_failed = False
        self._build_transforms()
        # Without any calibration, temperature display is unavailable.
        if self.transform1 is None and self.transform2 is None:
            self.radio_T.config(state="disabled")
            if self.display_mode.get() == "T":
                self.display_mode.set("R")

    def _build_transforms(self):
        """Raw-resistance -> temperature. None for a sensor that has no calibration."""
        state = self.app.state
        self.transform1 = (lambda e: state.cal1(np.log(e)) + state.cal_delta[0]) if state.cal1 is not None else None
        self.transform2 = (lambda e: state.cal2(np.log(e)) + state.cal_delta[1]) if state.cal2 is not None else None

    # ---------------- UI -----------------
    def _init_window(self):
        self.main_window = tk.Toplevel(self.app.root)
        # Closing the window must end the measurement cleanly, otherwise no data is saved.
        style_window(self.main_window, 'Bowman Thermometry', self._stop)

    def _init_variables(self):
        self.user_var = tk.StringVar()
        self.interval_var = tk.StringVar(value=str(self.display_interval))
        # Mutually exclusive sensor display mode: 'R' raw resistance, 'T' temperature, 'lnR' log resistance
        self.display_mode = tk.StringVar(value="R")
        # Plot the user input on its own (secondary) y-axis instead of sharing the sensor axis
        self.user_axis_var = tk.BooleanVar(value=True)

    def _init_fields(self):
        # Labels
        self.title_label = tk.Label(self.main_window, text='Thermometry Logger', font=('Arial', 16))
        self.interval_label = tk.Label(self.main_window, text='Display interval (s):')
        self.hint_label = tk.Label(self.main_window, text='Add number, "-" to remove last, empty=last+1')
        self.user_col_label = tk.Label(self.main_window, text='USER')
        self.s1_col_label = tk.Label(self.main_window, text='S1')
        self.s2_col_label = tk.Label(self.main_window, text='S2')

        # Entry and buttons
        self.user_entry = tk.Entry(self.main_window, textvariable=self.user_var)
        self.user_entry.bind("<Return>", lambda e: self.main_window.after(1, self._add_user_entry))
        self.add_btn = tk.Button(self.main_window, text="Add", command=self._add_user_entry)
        self.add_btn.configure(takefocus=False)
        self.stop_btn = tk.Button(self.main_window, text="Stop", command=self._stop)
        self.interval_entry = tk.Entry(self.main_window, textvariable=self.interval_var, width=5)

        # Listboxes (column 0 user, column 1 sensor 1, column 2 sensor 2)
        self.user_listbox = tk.Listbox(self.main_window, width=60)
        self.s1_listbox = tk.Listbox(self.main_window, width=60)
        self.s2_listbox = tk.Listbox(self.main_window, width=60)

        # Sensor display mode (mutually exclusive) + user-axis toggle
        self.mode_frame = tk.LabelFrame(self.main_window, text='Sensor display')
        self.radio_R = tk.Radiobutton(self.mode_frame, text='R', value='R',
                                      variable=self.display_mode, command=self._redraw)
        self.radio_T = tk.Radiobutton(self.mode_frame, text='T', value='T',
                                      variable=self.display_mode, command=self._redraw)
        self.radio_lnR = tk.Radiobutton(self.mode_frame, text='ln(R)', value='lnR',
                                        variable=self.display_mode, command=self._redraw)
        self.user_axis_check = tk.Checkbutton(self.main_window, text='User on own axis',
                                              variable=self.user_axis_var, command=self._redraw)

    def _populate_window(self):
        self.title_label.grid(row=0, column=0, columnspan=3, pady=10, sticky="w")
        self.interval_label.grid(row=1, column=0, sticky="w")
        self.interval_entry.grid(row=1, column=1, sticky="w")
        self.hint_label.grid(row=2, column=0, pady=(5, 10), sticky="w")
        self.user_entry.grid(row=3, column=0, sticky="w")
        self.add_btn.grid(row=3, column=1, padx=5, sticky="w")
        self.stop_btn.grid(row=3, column=2, padx=5, sticky="w")

        # Column headers aligned with the listboxes below
        self.user_col_label.grid(row=4, column=0, sticky="w")
        self.s1_col_label.grid(row=4, column=1, sticky="w")
        self.s2_col_label.grid(row=4, column=2, sticky="w")

        self.radio_R.pack(side="left")
        self.radio_T.pack(side="left")
        self.radio_lnR.pack(side="left")
        self.mode_frame.grid(row=5, column=0, columnspan=2, sticky="w", pady=(5, 0))
        self.user_axis_check.grid(row=5, column=2, sticky="w")

        # ---------- Listboxes ----------
        self.user_listbox.grid(row=6, column=0, sticky="nsew")
        self.s1_listbox.grid(row=6, column=1, sticky="nsew")
        self.s2_listbox.grid(row=6, column=2, sticky="nsew")

        # ---------------- Row/Column Weight ----------------
        for row in range(8):  # rows 0-7
            self.main_window.grid_rowconfigure(row, weight=0)
        self.main_window.grid_rowconfigure(6, weight=1)  # listboxes expand
        for col in range(3):
            self.main_window.grid_columnconfigure(col, weight=1)

    def _init_plot(self):
        self.user_color = "tab:green"
        self.fig = Figure(figsize=(7, 3), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax_user = self.ax.twinx()  # secondary y-axis dedicated to the user input

        self.line_s1, = self.ax.plot([], [], '*', linestyle='-', label="S1")
        self.line_s2, = self.ax.plot([], [], '*', linestyle='-', label="S2")
        # One user line per axis; only the active one carries data (see _redraw).
        self.line_user, = self.ax.plot([], [], 'D', color=self.user_color, label="USER")
        self.line_user_sec, = self.ax_user.plot([], [], 'D', color=self.user_color, label="USER")

        self.ax.set_xlabel("Time [s]", labelpad=6)
        self.ax.set_ylabel("Resistance [Ω]", labelpad=8)
        self.ax_user.set_ylabel("User value", labelpad=8, color=self.user_color)
        self.ax_user.tick_params(axis='y', colors=self.user_color)
        # Legend handles come from both axes; line_user is the representative USER entry.
        self.ax.legend([self.line_s1, self.line_s2, self.line_user],
                       ["S1", "S2", "USER"], loc="upper left")
        self.ax.grid(True)
        self.fig.tight_layout()

        self.plot_frame = tk.Frame(self.main_window)
        self.plot_frame.grid(row=7, column=0, columnspan=3, sticky="nsew")
        self.main_window.grid_rowconfigure(7, weight=2)  # make canvas row expandable

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def _update_plot(self):
        # Periodic loop: draw, then reschedule. Callbacks use _redraw to avoid stacking timers.
        if self.stop_event.is_set():
            return
        self._redraw()
        self.main_window.after(500, self._update_plot)

    def _set_user_axis_visible(self, visible):
        """Show/hide the secondary (user) y-axis without removing the axis."""
        self.ax_user.get_yaxis().set_visible(visible)
        self.ax_user.spines["right"].set_visible(visible)

    def _redraw(self):
        if not hasattr(self, "ax"):  # plot not built yet (callback fired too early)
            return
        try:
            now = time.time()

            u_entries = self.user_logger.get_entries()
            s1_entries = self.serial_logger_1.get_entries()
            s2_entries = self.serial_logger_2.get_entries()

            ux, uy = self._downsample_by_interval(u_entries, self.display_interval)
            s1x, s1y = self._downsample_by_interval(s1_entries, self.display_interval)
            s2x, s2y = self._downsample_by_interval(s2_entries, self.display_interval)

            # Make x relative to the first timestamp (nicer axis)
            if ux:
                t0 = min(ux)
            elif s1x:
                t0 = min(s1x)
            elif s2x:
                t0 = min(s2x)
            else:
                t0 = now

            ux = [t - t0 for t in ux]
            s1x = [t - t0 for t in s1x]
            s2x = [t - t0 for t in s2x]

            # --- Sensor traces: single mutually exclusive display mode ---
            s1y = np.asarray(s1y, dtype=float)
            s2y = np.asarray(s2y, dtype=float)
            mode = self.display_mode.get()
            if mode == "T":
                if self.transform1 is not None:
                    self.line_s1.set_data(s1x, self.transform1(s1y))
                else:
                    self.line_s1.set_data([], [])
                if self.transform2 is not None:
                    self.line_s2.set_data(s2x, self.transform2(s2y))
                else:
                    self.line_s2.set_data([], [])
                self.ax.set_ylabel("Temperature [°C]", labelpad=8)
            elif mode == "lnR":
                self.line_s1.set_data(s1x, np.log(s1y))
                self.line_s2.set_data(s2x, np.log(s2y))
                self.ax.set_ylabel("ln(Resistance)", labelpad=8)
            else:  # "R"
                self.line_s1.set_data(s1x, s1y)
                self.line_s2.set_data(s2x, s2y)
                self.ax.set_ylabel("Resistance [Ω]", labelpad=8)

            # --- User markers: own axis or shared with the sensor axis ---
            if self.user_axis_var.get():
                self.line_user.set_data([], [])
                self.line_user_sec.set_data(ux, uy)
                self._set_user_axis_visible(True)
            else:
                self.line_user.set_data(ux, uy)
                self.line_user_sec.set_data([], [])
                self._set_user_axis_visible(False)

            for ax in (self.ax, self.ax_user):
                ax.relim()
                ax.autoscale_view()

            self.canvas.draw_idle()
            self._redraw_failed = False

        except Exception as e:
            if not getattr(self, "_redraw_failed", False):  # log once, don't flood at 2 Hz
                print("Plot update error:", e)
                self._redraw_failed = True

    def _downsample_by_interval(self, entries, interval):
        """
        entries: list of (timestamp, value) OR (timestamp, idx, value)
        returns: (xs, ys) downsampled to 1 per interval bin
        """
        xs, ys = [], []
        last_bin = None

        for item in entries:
            if len(item) == 2:
                t, v = item
            else:
                t, _, v = item

            if not isinstance(v, (int, float)):
                continue

            bin_id = int(t // interval)
            if bin_id != last_bin:
                xs.append(t)
                ys.append(v)
                last_bin = bin_id

        return xs, ys

    # ---------------- User Input -----------------
    def _add_user_entry(self, event=None):
        value_str = self.user_var.get().strip()
        self.user_var.set("")

        # Delete last
        if value_str == "-":
            removed = self.user_logger.remove_last()
            if removed:
                for j in range(self.user_listbox.size() - 1, -1, -1):
                    if self.user_listbox.get(j).startswith("[USER]"):
                        self.user_listbox.delete(j)
                        break
            return

        # Empty -> last + 1
        if value_str == "":
            if self.user_logger.entries:
                _, last_val = self.user_logger.entries[-1][:2]
                value = last_val + 1
            else:
                messagebox.showwarning("No previous value", "No previous value to increment")
                return
        else:
            try:
                value = float(value_str)
            except ValueError:
                messagebox.showwarning("Invalid input", "Enter a number, '-' to delete, or leave blank to increment last value")
                return

        self.user_logger.add(value)
        self.user_listbox.insert(tk.END, f"[USER] {value}")

    # ---------------- Serial Worker with buffer -----------------
    def _serial_worker(self):
        buffer = bytearray()
        while not self.stop_event.is_set():
            try:
                if self.ser and self.ser.is_open:
                    n = self.ser.in_waiting
                    if n:
                        buffer.extend(self.ser.read(n))
                        while b'\n' in buffer:
                            line_bytes, _, buffer = buffer.partition(b'\n')
                            line = line_bytes.decode(errors="ignore").strip()
                            parts = line.split(LINE_DELIMITER)
                            if len(parts) < 4:
                                if line:
                                    print(f"Non-conforming line: {line}")
                                continue
                            try:
                                device_ts = float(parts[1])
                                val = float(parts[3])
                            except ValueError:
                                print(f"Non-conforming line: {line}")
                                continue
                            sensor_id = parts[2]
                            ts = time.time()
                            if val == INVALID_READING:
                                val = np.nan

                            if sensor_id == SENSOR_S1:
                                self.serial_logger_1.add((int(device_ts), val), ts)
                                self.queue1.put((val, ts))
                            elif sensor_id == SENSOR_S2:
                                self.serial_logger_2.add((int(device_ts), val), ts)
                                self.queue2.put((val, ts))
                            else:
                                print(f"Non-conforming line: {line}")
                    else:
                        time.sleep(0.01)
                else:
                    time.sleep(0.1)
            except Exception:
                time.sleep(0.05)

    def _get_display_interval(self):
        try:
            return max(float(self.interval_var.get()), 0.1)
        except (ValueError, tk.TclError):
            return self.display_interval

    # ---------------- Poll Queues -----------------
    def _poll_queues(self):
        if self.stop_event.is_set():
            return
        self.display_interval = self._get_display_interval()

        while not self.queue1.empty():
            value, ts = self.queue1.get()
            if ts - self.last_display_time_1 >= self.display_interval:
                self.s1_listbox.insert(tk.END, f"[SERIAL] {value}")
                self.last_display_time_1 = ts

        while not self.queue2.empty():
            value, ts = self.queue2.get()
            if ts - self.last_display_time_2 >= self.display_interval:
                self.s2_listbox.insert(tk.END, f"[SERIAL] {value}")
                self.last_display_time_2 = ts

        self.main_window.after(100, self._poll_queues)

    # ---------------- Stop -----------------
    def _stop(self):
        if self.stop_event.is_set():
            return
        self.stop_event.set()
        self.end_time = datetime.datetime.now()
        if self.ser and self.ser.is_open:
            self.ser.close()
        messagebox.showinfo("Stopped", "Logging stopped")

        print("User entries:")
        for t, v in self.user_logger.get_entries():
            print(f"{time.strftime('%H:%M:%S', time.localtime(t))} {v}")

        print("\nSerial S1 entries:")
        for t, idx, v in self.serial_logger_1.get_entries():
            print(f"{time.strftime('%H:%M:%S', time.localtime(t))} idx={idx} val={v}")

        print("\nSerial S2 entries:")
        for t, idx, v in self.serial_logger_2.get_entries():
            print(f"{time.strftime('%H:%M:%S', time.localtime(t))} idx={idx} val={v}")
        self.main_window.destroy()


# ===================== RUN =====================
def main():
    ThermometryUI()


if __name__ == "__main__":
    main()