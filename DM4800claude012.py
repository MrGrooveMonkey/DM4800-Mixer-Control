#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DM4800 Mixer Controller v0.4.0
Dependencies: pip install PyQt5 mido python-rtmidi
"""

import os
import sys
import csv
import json
import datetime
from dataclasses import dataclass
from typing import Dict, Tuple, Optional, List
from collections import deque
from PyQt5 import QtCore, QtGui, QtWidgets
import mido

# Constants
CSV_DEFAULT = "DM4800_CC_Values.csv"
SCRIBBLE_CSV = "scribblestrip.csv"
MONO_STEREO_CSV = "DM4800MonoStereo.csv"
SETTINGS_FILE = "dm4800_settings.json"
MIDI_LOG_FILE = "DM4800_MIDI_Log.txt"

CHANNEL_DB_POINTS = [(127, "+10 dB"), (114, "+5 dB"), (102, "0 dB"), (89, "-5 dB"), (77, "-10 dB"), (58, "-20 dB"), (0, "-∞")]
MASTER_DB_POINTS = [(127, "0 dB"), (114, "-5 dB"), (102, "-10 dB"), (77, "-20 dB"), (0, "-∞")]

BASE_FADER_HEIGHT = 320
BASE_STRIP_WIDTH = 92
BASE_STEREO_FADER_WIDTH_MULTIPLIER = 1.5
BASE_FONT = 11
BASE_SMALL_FONT = 10
BASE_BUTTON_H = 28
BASE_PAN_DIAL = 58

PAN_CENTER_CC = 64
FADER_ZERO_DB_CC = 102
MASTER_ZERO_DB_CC = 127
PAN_CENTER_COLOR = "#00ff00"

MAX_SCRIBBLE_LENGTH = 24
MAX_SCRIBBLE_CHARS_PER_ROW = 12
DEFAULT_SCRIBBLE_TEXT = "-"
MAX_MONO_STEREO_LENGTH = 8
DEFAULT_MONO_STEREO_TEXT = "Mono"
MAX_MIDI_MESSAGES = 1000

@dataclass
class MidiMapping:
    section: str
    number: int
    type: str
    midi_channel: int
    cc_number: int

class SettingsManager:
    def __init__(self, settings_file: str = SETTINGS_FILE):
        self.settings_file = settings_file
        self.settings = self.load_settings()

    def load_settings(self) -> dict:
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading settings: {e}")
        return {'midi_in_port': '', 'midi_out_port': '', 'horizontal_zoom': 100, 'vertical_zoom': 100, 'remember_ports': True}

    def save_settings(self):
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def get(self, key: str, default=None):
        return self.settings.get(key, default)

    def set(self, key: str, value):
        self.settings[key] = value

class MidiMonitorWindow(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("MIDI Monitor")
        self.resize(600, 400)
        self.setWindowFlags(QtCore.Qt.Window)
        self.messages = deque(maxlen=MAX_MIDI_MESSAGES)
        self.logging_enabled = False
        self.setup_ui()
        
    def setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        controls_layout = QtWidgets.QHBoxLayout()
        
        self.clear_btn = QtWidgets.QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_messages)
        controls_layout.addWidget(self.clear_btn)
        
        self.log_checkbox = QtWidgets.QCheckBox("Log?")
        self.log_checkbox.toggled.connect(self.toggle_logging)
        controls_layout.addWidget(self.log_checkbox)
        
        controls_layout.addStretch()
        layout.addLayout(controls_layout)
        
        self.text_area = QtWidgets.QTextEdit()
        self.text_area.setReadOnly(True)
        self.text_area.setFont(QtGui.QFont("Courier", 9))
        layout.addWidget(self.text_area)
        
    def add_message(self, direction: str, channel: int, cc: int, value: int):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        message = f"{timestamp} {direction:3} Ch:{channel:2d} CC:{cc:3d} Val:{value:3d}"
        self.messages.append(message)
        self.update_display()
        if self.logging_enabled:
            self.log_message(message)
    
    def update_display(self):
        self.text_area.clear()
        self.text_area.append("\n".join(self.messages))
        scrollbar = self.text_area.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def clear_messages(self):
        self.messages.clear()
        self.text_area.clear()
    
    def toggle_logging(self, enabled: bool):
        self.logging_enabled = enabled
        if enabled:
            self.log_message(f"=== MIDI Logging Started at {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    
    def log_message(self, message: str):
        try:
            with open(MIDI_LOG_FILE, 'a') as f:
                f.write(message + '\n')
        except Exception as e:
            print(f"Error writing to log file: {e}")

class MidiPortSelectionDialog(QtWidgets.QDialog):
    def __init__(self, title: str, ports: List[str], current_port: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.selected_port = current_port
        
        layout = QtWidgets.QVBoxLayout(self)
        
        current_display = current_port if current_port else "None"
        current_label = QtWidgets.QLabel(f"Currently Selected Port: {current_display}")
        current_label.setStyleSheet("font-weight: bold; padding: 5px; background-color: #f0f0f0; border: 1px solid #ccc;")
        layout.addWidget(current_label)
        
        label = QtWidgets.QLabel(f"Select {title.lower()}:")
        layout.addWidget(label)
        
        self.port_list = QtWidgets.QListWidget()
        for port in ports:
            item = QtWidgets.QListWidgetItem(port)
            if port == current_port:
                item.setSelected(True)
                self.port_list.setCurrentItem(item)
            self.port_list.addItem(item)
        
        self.port_list.itemDoubleClicked.connect(self.accept)
        layout.addWidget(self.port_list)
        
        button_layout = QtWidgets.QHBoxLayout()
        self.ok_btn = QtWidgets.QPushButton("OK")
        self.ok_btn.clicked.connect(self.accept)
        button_layout.addWidget(self.ok_btn)
        
        self.cancel_btn = QtWidgets.QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        self.resize(300, 250)
    
    def get_selected_port(self) -> str:
        current_item = self.port_list.currentItem()
        return current_item.text() if current_item else ""

class MonoStereoManager:
    def __init__(self, csv_file: str = MONO_STEREO_CSV):
        self.csv_file = csv_file
        self.mono_stereo_data: Dict[str, str] = {}
        self.stereo_pairs: Dict[str, str] = {}
        self.load_mono_stereo_data()

    def _generate_default_data(self) -> Dict[str, str]:
        data = {}
        for i in range(1, 65):
            data[f"Channel {i}"] = DEFAULT_MONO_STEREO_TEXT
        for i in range(1, 25):
            data[f"Bus {i}"] = DEFAULT_MONO_STEREO_TEXT
        for i in range(1, 13):
            data[f"Aux {i}"] = DEFAULT_MONO_STEREO_TEXT
        return data

    def load_mono_stereo_data(self):
        if not os.path.exists(self.csv_file):
            self.mono_stereo_data = self._generate_default_data()
            self.save_mono_stereo_data()
            return

        try:
            with open(self.csv_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                self.mono_stereo_data = {}
                for row in reader:
                    key = row.get('Channel_Bus_Aux', '').strip()
                    value = row.get('Mono_Stereo', DEFAULT_MONO_STEREO_TEXT).strip()
                    if key:
                        self.mono_stereo_data[key] = value

            default_data = self._generate_default_data()
            for key in default_data:
                if key not in self.mono_stereo_data:
                    self.mono_stereo_data[key] = DEFAULT_MONO_STEREO_TEXT

        except Exception as e:
            print(f"Error loading mono/stereo data: {e}")
            self.mono_stereo_data = self._generate_default_data()
            self.save_mono_stereo_data()

        self._detect_stereo_pairs()

    def save_mono_stereo_data(self):
        try:
            with open(self.csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Channel_Bus_Aux', 'Mono_Stereo'])
                for i in range(1, 65):
                    key = f"Channel {i}"
                    value = self.mono_stereo_data.get(key, DEFAULT_MONO_STEREO_TEXT)
                    writer.writerow([key, value])
                for i in range(1, 25):
                    key = f"Bus {i}"
                    value = self.mono_stereo_data.get(key, DEFAULT_MONO_STEREO_TEXT)
                    writer.writerow([key, value])
                for i in range(1, 13):
                    key = f"Aux {i}"
                    value = self.mono_stereo_data.get(key, DEFAULT_MONO_STEREO_TEXT)
                    writer.writerow([key, value])
        except Exception as e:
            print(f"Error saving mono/stereo data: {e}")

    def _detect_stereo_pairs(self):
        self.stereo_pairs = {}
        for section in ["Channel", "Bus", "Aux"]:
            max_num = {"Channel": 64, "Bus": 24, "Aux": 12}[section]
            for i in range(1, max_num + 1, 2):
                odd_key = f"{section} {i}"
                even_key = f"{section} {i + 1}"
                if (self.mono_stereo_data.get(odd_key, "").lower() == "stereo" and i + 1 <= max_num):
                    self.stereo_pairs[odd_key] = even_key
                    self.stereo_pairs[even_key] = odd_key

    def is_stereo_pair(self, key: str) -> bool:
        return key in self.stereo_pairs

    def get_stereo_partner(self, key: str) -> Optional[str]:
        return self.stereo_pairs.get(key)

    def is_stereo_left(self, key: str) -> bool:
        if key not in self.stereo_pairs:
            return False
        try:
            parts = key.split()
            number = int(parts[1])
            return number % 2 == 1
        except:
            return False

    def is_stereo_right(self, key: str) -> bool:
        if key not in self.stereo_pairs:
            return False
        try:
            parts = key.split()
            number = int(parts[1])
            return number % 2 == 0
        except:
            return False

class ScribbleStripManager:
    def __init__(self, csv_file: str = SCRIBBLE_CSV):
        self.csv_file = csv_file
        self.scribble_data: Dict[str, str] = {}
        self.load_scribble_data()

    def _generate_default_data(self) -> Dict[str, str]:
        data = {}
        for i in range(1, 65):
            data[f"Channel {i}"] = DEFAULT_SCRIBBLE_TEXT
        for i in range(1, 25):
            data[f"Bus {i}"] = DEFAULT_SCRIBBLE_TEXT
        for i in range(1, 13):
            data[f"Aux {i}"] = DEFAULT_SCRIBBLE_TEXT
        return data

    def load_scribble_data(self):
        if not os.path.exists(self.csv_file):
            self.scribble_data = self._generate_default_data()
            self.save_scribble_data()
            return

        try:
            with open(self.csv_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                self.scribble_data = {}
                for row in reader:
                    key = row.get('Channel_Bus_Aux', '').strip()
                    value = row.get('Value', DEFAULT_SCRIBBLE_TEXT).strip()
                    if key:
                        self.scribble_data[key] = value

            default_data = self._generate_default_data()
            for key in default_data:
                if key not in self.scribble_data:
                    self.scribble_data[key] = DEFAULT_SCRIBBLE_TEXT

        except Exception as e:
            print(f"Error loading scribble data: {e}")
            self.scribble_data = self._generate_default_data()
            self.save_scribble_data()

    def save_scribble_data(self):
        try:
            with open(self.csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Channel_Bus_Aux', 'Value'])
                for i in range(1, 65):
                    key = f"Channel {i}"
                    value = self.scribble_data.get(key, DEFAULT_SCRIBBLE_TEXT)
                    writer.writerow([key, value])
                for i in range(1, 25):
                    key = f"Bus {i}"
                    value = self.scribble_data.get(key, DEFAULT_SCRIBBLE_TEXT)
                    writer.writerow([key, value])
                for i in range(1, 13):
                    key = f"Aux {i}"
                    value = self.scribble_data.get(key, DEFAULT_SCRIBBLE_TEXT)
                    writer.writerow([key, value])
        except Exception as e:
            print(f"Error saving scribble data: {e}")

    def get_scribble_text(self, key: str) -> str:
        return self.scribble_data.get(key, DEFAULT_SCRIBBLE_TEXT)

    def set_scribble_text(self, key: str, value: str):
        value = value[:MAX_SCRIBBLE_LENGTH] if len(value) > MAX_SCRIBBLE_LENGTH else value
        self.scribble_data[key] = value
        self.save_scribble_data()

class MidiManager(QtCore.QObject):
    midi_message_received = QtCore.pyqtSignal(int, int, int)
    midi_message_sent = QtCore.pyqtSignal(int, int, int)

    def __init__(self):
        super().__init__()
        self._inport: Optional[mido.ports.BaseInput] = None
        self._outport: Optional[mido.ports.BaseOutput] = None
        self.current_in_port = ""
        self.current_out_port = ""

    def list_input_ports(self) -> List[str]:
        return mido.get_input_names()

    def list_output_ports(self) -> List[str]:
        return mido.get_output_names()

    def open_input(self, input_name: str) -> bool:
        if self._inport:
            self._inport.close()
            self._inport = None
        
        if input_name:
            try:
                self._inport = mido.open_input(input_name, callback=self._mido_callback)
                self.current_in_port = input_name
                print(f"Opened MIDI input: {input_name}")
                return True
            except Exception as e:
                print(f"Error opening MIDI input {input_name}: {e}")
                self.current_in_port = ""
                return False
        return False

    def open_output(self, output_name: str) -> bool:
        if self._outport:
            self._outport.close()
            self._outport = None
        
        if output_name:
            try:
                self._outport = mido.open_output(output_name)
                self.current_out_port = output_name
                print(f"Opened MIDI output: {output_name}")
                return True
            except Exception as e:
                print(f"Error opening MIDI output {output_name}: {e}")
                self.current_out_port = ""
                return False
        return False

    def close(self):
        if self._inport:
            self._inport.close()
            self._inport = None
            self.current_in_port = ""
        if self._outport:
            self._outport.close()
            self._outport = None
            self.current_out_port = ""

    def _mido_callback(self, msg: mido.Message):
        if msg.type == 'control_change':
            self.midi_message_received.emit(msg.channel + 1, msg.control, msg.value)

    def send_cc(self, channel_1: int, cc: int, value: int):
        if not self._outport:
            return
        ch0 = max(0, min(15, channel_1 - 1))
        val = max(0, min(127, int(value)))
        self._outport.send(mido.Message('control_change', channel=ch0, control=int(cc), value=val))
        self.midi_message_sent.emit(channel_1, cc, val)

    def parse_section_and_number(s: str) -> Tuple[str, int]:
        raw = s.strip().lower()
        if raw.startswith("channel"):
            return ("channel", int(raw.split()[-1]))
        if raw.startswith("bus") or raw.startswith("buss"):
            return ("bus", int(raw.split()[-1]))
        if raw.startswith("aux"):
            return ("aux", int(raw.split()[-1]))
        if "master" in raw:
            return ("master", 1)
        parts = raw.split()
        try:
            return (parts[0], int(parts[-1]))
        except Exception:
            return (raw, 1)

    def load_midi_mappings(csv_path: str) -> Dict[Tuple[str,int,str], MidiMapping]:
        mappings: Dict[Tuple[str,int,str], MidiMapping] = {}
        if not os.path.isfile(csv_path):
            raise FileNotFoundError(f"CSV not found at {csv_path}")
        with open(csv_path, newline='') as f:
            reader = csv.DictReader(f)
            for row in reader:
                section_raw = row.get('Channel_Bus_Aux', '')
                type_raw = row.get('Type', '').strip().lower()
                midi_ch = int(row.get('MIDI Channel', '1'))
                cc_num = int(row.get('CC_Number', '0'))
                section, number = parse_section_and_number(section_raw)
                key = (section, number, type_raw)
                mappings[key] = MidiMapping(section, number, type_raw, midi_ch, cc_num)
        return mappings

    class ScribbleStripTextEdit(QtWidgets.QTextEdit):
        textCommitted = QtCore.pyqtSignal(str)

        def __init__(self, parent=None):
            super().__init__(parent)
            self.setMaximumHeight(60)
            self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
            self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
            self.setLineWrapMode(QtWidgets.QTextEdit.WidgetWidth)
            self.setStyleSheet("""
                QTextEdit {
                    background-color: #333333; color: white; border: 1px solid #666666;
                    border-radius: 3px; padding: 2px; font-size: 10px; font-family: Arial;
                }
                QTextEdit:focus { border: 2px solid #0078d4; background-color: #444444; }
            """)
            self.textChanged.connect(self._on_text_changed)

        def _on_text_changed(self):
            text = self.toPlainText()
            if len(text) > MAX_SCRIBBLE_LENGTH:
                cursor = self.textCursor()
                cursor.deletePreviousChar()

        def _format_text_for_display(self, text: str) -> str:
            if len(text) <= MAX_SCRIBBLE_CHARS_PER_ROW:
                return text
            else:
                row1 = text[:MAX_SCRIBBLE_CHARS_PER_ROW]
                row2 = text[MAX_SCRIBBLE_CHARS_PER_ROW:MAX_SCRIBBLE_LENGTH]
                return f"{row1}\n{row2}"

        def setText(self, text: str):
            formatted = self._format_text_for_display(text)
            super().setPlainText(formatted)

        def getText(self) -> str:
            return self.toPlainText().replace('\n', '')

        def keyPressEvent(self, event: QtGui.QKeyEvent):
            if event.key() in [QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter]:
                self.clearFocus()
            else:
                super().keyPressEvent(event)

        def focusOutEvent(self, event: QtGui.QFocusEvent):
            super().focusOutEvent(event)
            self.textCommitted.emit(self.getText())

    class ScaleWidget(QtWidgets.QWidget):
        def __init__(self, is_master=False, parent=None):
            super().__init__(parent)
            self.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
            self.font_scale = 1.0
            self.is_master = is_master

        def set_scale(self, s: float):
            self.font_scale = s
            self.update()

        def paintEvent(self, event: QtGui.QPaintEvent):
            painter = QtGui.QPainter(self)
            rect = self.rect().adjusted(6, 6, -6, -6)
            painter.setRenderHint(QtGui.QPainter.Antialiasing, True)

            painter.setPen(QtGui.QPen(QtCore.Qt.gray, 1))
            painter.drawLine(rect.right()-2, rect.top(), rect.right()-2, rect.bottom())

            def y_for_cc(cc: int) -> int:
                t = cc / 127.0
                return int(rect.top() + (1.0 - t) * rect.height())

            painter.setPen(QtGui.QPen(QtCore.Qt.black, 1))
            font = painter.font()
            font.setPointSizeF(max(7.0, BASE_SMALL_FONT * self.font_scale))
            painter.setFont(font)

            scale_points = MASTER_DB_POINTS if self.is_master else CHANNEL_DB_POINTS

            for cc, label in scale_points:
                y = y_for_cc(cc)
                painter.drawLine(rect.right()-6, y, rect.right()-2, y)
                text_rect = QtCore.QRect(rect.left(), y-8, rect.width()-8, 16)
                painter.drawText(text_rect, QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter, label)

    class CustomFaderSlider(QtWidgets.QSlider):
        doubleClicked = QtCore.pyqtSignal()

        def __init__(self, is_master=False, parent=None):
            super().__init__(QtCore.Qt.Vertical, parent)
            self.setRange(0, 127)
            self.setInvertedAppearance(False)
            self.setTickPosition(QtWidgets.QSlider.NoTicks)
            self.is_master = is_master
            self.current_value = MASTER_ZERO_DB_CC if is_master else FADER_ZERO_DB_CC
            self.setStyleSheet(self._get_fader_style())

        def _interpolate_color(self, start_rgb: tuple, end_rgb: tuple, factor: float) -> str:
            factor = max(0.0, min(1.0, factor))
            r = int(start_rgb[0] + (end_rgb[0] - start_rgb[0]) * factor)
            g = int(start_rgb[1] + (end_rgb[1] - start_rgb[1]) * factor)
            b = int(start_rgb[2] + (end_rgb[2] - start_rgb[2]) * factor)
            return f"rgb({r}, {g}, {b})"

        def _get_fader_color(self, value: int) -> str:
            zero_db_value = MASTER_ZERO_DB_CC if self.is_master else FADER_ZERO_DB_CC
        
            if value == zero_db_value:
                return "#00ff00"
            elif value > zero_db_value:
                factor = (value - zero_db_value) / (127 - zero_db_value)
                start_color = (237, 190, 0)
                end_color = (255, 0, 0)
                return self._interpolate_color(start_color, end_color, factor)
            else:
                factor = (zero_db_value - value) / zero_db_value
                start_color = (0, 186, 0)
                end_color = (0, 60, 0)
                return self._interpolate_color(start_color, end_color, factor)

        def _get_fader_style(self):
            color = self._get_fader_color(self.current_value)
            return f"""
                QSlider:groove:vertical {{
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #333333, stop:0.5 #555555, stop:1 #333333);
                    width: 8px; border: 1px solid #222222; border-radius: 2px;
                }}
                QSlider:handle:vertical {{
                    background: {color}; border: 1px solid #444444; height: 20px; width: 24px; margin: 0px -8px; border-radius: 3px;
                }}
                QSlider:handle:vertical:pressed {{
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #00aa00, stop:0.5 #00dd00, stop:1 #00aa00);
                }}
            """

        def update_color(self, value: int):
            self.current_value = value
            self.setStyleSheet(self._get_fader_style())

        def paintEvent(self, event):
            super().paintEvent(event)
            painter = QtGui.QPainter(self)
            painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
            opt = QtWidgets.QStyleOptionSlider()
            self.initStyleOption(opt)
            handle = self.style().subControlRect(QtWidgets.QStyle.CC_Slider, opt, QtWidgets.QStyle.SC_SliderHandle, self)
            painter.setPen(QtGui.QPen(QtCore.Qt.black, 2))
            center_y = handle.center().y()
            painter.drawLine(handle.left() + 2, center_y, handle.right() - 2, center_y)

        def mouseDoubleClickEvent(self, e: QtGui.QMouseEvent):
            if e.button() == QtCore.Qt.LeftButton:
                self.doubleClicked.emit()

    class FaderWithScale(QtWidgets.QWidget):
        valueChanged = QtCore.pyqtSignal(int)

        def __init__(self, initial=FADER_ZERO_DB_CC, h_scale=1.0, v_scale=1.0, is_master=False, is_stereo=False, parent=None):
            super().__init__(parent)
            self.h_scale_factor = h_scale
            self.v_scale_factor = v_scale
            self.is_master = is_master
            self.is_stereo = is_stereo
            self.slider = CustomFaderSlider(is_master=is_master)
            self.slider.setValue(initial)

            self.scale = ScaleWidget(is_master=is_master)
            self.scale.setMinimumWidth(46)

            layout = QtWidgets.QHBoxLayout(self)
            layout.setContentsMargins(2,2,2,2)
            layout.setSpacing(6)
            layout.addWidget(self.scale)
            layout.addWidget(self.slider)

            self.slider.valueChanged.connect(self._on_value_changed)
            self.slider.doubleClicked.connect(self.reset_to_zero_db)

            self.apply_scale(self.h_scale_factor, self.v_scale_factor)
            self.slider.update_color(initial)

        def _on_value_changed(self, value):
            self.slider.update_color(value)
            self.valueChanged.emit(value)

        def reset_to_zero_db(self):
            zero_db_value = MASTER_ZERO_DB_CC if self.is_master else FADER_ZERO_DB_CC
            self.slider.setValue(zero_db_value)

        def setValue(self, v: int):
            self.slider.blockSignals(True)
            self.slider.setValue(int(max(0, min(127, v))))
            self.slider.blockSignals(False)
            self.slider.update_color(v)

        def value(self) -> int:
            return self.slider.value()

        def apply_scale(self, h_scale: float, v_scale: float):
            self.h_scale_factor = h_scale
            self.v_scale_factor = v_scale
            fader_h = int(BASE_FADER_HEIGHT * v_scale)
            self.slider.setFixedHeight(fader_h)
        
            if self.is_stereo:
                fader_w = int(40 * BASE_STEREO_FADER_WIDTH_MULTIPLIER * h_scale)
                self.slider.setMinimumWidth(fader_w)
        
            self.scale.set_scale(min(h_scale, v_scale))

    class PanDial(QtWidgets.QDial):
        doubleClicked = QtCore.pyqtSignal()

        def __init__(self, parent=None):
            super().__init__(parent)
            self.setStyleSheet(self._get_pan_style())

        def _get_pan_style(self):
            return """
                QDial {
                    background-color: #444444;
                    border: 2px solid #666666;
                    border-radius: 29px;
                }
            """

        def update_color(self, value: int):
            if value == PAN_CENTER_CC:
                self.setStyleSheet(f"""
                    QDial {{
                        background-color: {PAN_CENTER_COLOR};
                        border: 2px solid #00cc00;
                        border-radius: 29px;
                    }}
                """)
            else:
                self.setStyleSheet(self._get_pan_style())

        def mouseDoubleClickEvent(self, e: QtGui.QMouseEvent):
            if e.button() == QtCore.Qt.LeftButton:
                self.doubleClicked.emit()
            super().mouseDoubleClickEvent(e)

    class MuteButton(QtWidgets.QPushButton):
        toggledCC = QtCore.pyqtSignal(int)

        def __init__(self, parent=None):
            super().__init__("MUTE", parent)
            self.setCheckable(True)
            self.update_style()
            self.toggled.connect(self._on_toggled)

        def _on_toggled(self, on: bool):
            self.update_style()
            self.toggledCC.emit(127 if on else 0)

        def update_style(self):
            if self.isChecked():
                self.setStyleSheet("QPushButton { background-color: rgb(249,7,7); color: white; font-weight: bold; }")
            else:
                self.setStyleSheet("QPushButton { background-color: #444; color: black; }")

        def set_from_cc(self, value: int):
            on = (value >= 64)
            self.blockSignals(True)
            self.setChecked(on)
            self.blockSignals(False)
            self.update_style()

    class ChannelStrip(QtWidgets.QFrame):
        scribbleTextChanged = QtCore.pyqtSignal(str, str)

        def __init__(self, title: str, scribble_key: str = "", has_pan: bool=True, has_mute: bool=True, has_scribble: bool=True, h_scale: float=1.0, v_scale: float=1.0, is_master: bool=False, mono_stereo_manager=None, is_stereo_pair: bool=False, stereo_partner_num: int=0, parent=None):
            super().__init__(parent)
            self.h_scale_factor = h_scale
            self.v_scale_factor = v_scale
            self.is_master = is_master
            self.is_stereo_pair = is_stereo_pair
            self.stereo_partner_num = stereo_partner_num
            self.scribble_key = scribble_key
            self.mono_stereo_manager = mono_stereo_manager
            self.stereo_partner_strip = None
            self.setFrameShape(QtWidgets.QFrame.StyledPanel)

            display_title = title
            title_color = "black"
        
            if is_stereo_pair and stereo_partner_num > 0:
                channel_num = int(scribble_key.split()[-1]) if scribble_key else 0
                display_title = f"Ch {channel_num} - {stereo_partner_num}"
                title_color = "rgb(20,20,225)"
            elif scribble_key and "Channel" in scribble_key:
                channel_num = int(scribble_key.split()[-1])
                display_title = f"CH {channel_num}"

            self.v = QtWidgets.QVBoxLayout(self)
            self.v.setContentsMargins(6,6,6,6)
            self.v.setSpacing(6)
            self.setMinimumWidth(int(BASE_STRIP_WIDTH * self.h_scale_factor))

            self.title_label = QtWidgets.QLabel(display_title)
            self.title_label.setAlignment(QtCore.Qt.AlignCenter)
            self.title_label.setStyleSheet(f"font-weight: bold; color: {title_color};")
            self.v.addWidget(self.title_label)

            if has_pan:
                self.pan = PanDial()
                self.pan.setRange(0,127)
                self.pan.setValue(PAN_CENTER_CC)
                self.pan.setNotchesVisible(True)
                self.pan.doubleClicked.connect(self._pan_reset)
                self.pan.valueChanged.connect(self._on_pan_changed)
                self.v.addWidget(self.pan, 0, QtCore.Qt.AlignHCenter)

                self.pan_value = QtWidgets.QLabel("C")
                self.pan_value.setAlignment(QtCore.Qt.AlignCenter)
                self.v.addWidget(self.pan_value)
                self.pan.update_color(PAN_CENTER_CC)
            else:
                self.pan = None
                self.pan_value = None

            if has_mute:
                self.mute = MuteButton()
                self.v.addWidget(self.mute)
            else:
                self.mute = None

            initial_val = MASTER_ZERO_DB_CC if is_master else FADER_ZERO_DB_CC
            self.fader = FaderWithScale(initial=initial_val, h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=is_master, is_stereo=is_stereo_pair)
            self.v.addWidget(self.fader, 1)

            if has_scribble and not is_master:
                self.scribble = ScribbleStripTextEdit()
                self.scribble.textCommitted.connect(self._on_scribble_changed)
                self.v.addWidget(self.scribble)
            else:
                self.scribble = None

            if is_master:
                self.v.addStretch(1)

            self.apply_scale(self.h_scale_factor, self.v_scale_factor)

        def set_stereo_partner(self, partner_strip):
            self.stereo_partner_strip = partner_strip

        def mirror_from_partner(self, control_type: str, value: int):
            if control_type == "fader":
                self.fader.setValue(value)
            elif control_type == "pan" and self.pan:
                self.pan.blockSignals(True)
                self.pan.setValue(value)
                self.pan.blockSignals(False)
                self._on_pan_changed(value)
            elif control_type == "mute" and self.mute:
                self.mute.set_from_cc(value)

        def sync_to_partner(self, control_type: str, value: int):
            if self.stereo_partner_strip:
                self.stereo_partner_strip.mirror_from_partner(control_type, value)

        def set_scribble_text(self, text: str):
            if self.scribble:
                self.scribble.blockSignals(True)
                self.scribble.setText(text)
                self.scribble.blockSignals(False)

        def _on_scribble_changed(self, new_text: str):
            if self.scribble_key:
                self.scribbleTextChanged.emit(self.scribble_key, new_text)

        def _on_pan_changed(self, v: int):
            if self.pan_value:
                if v == 64:
                    self.pan_value.setText("C")
                elif v < 64:
                    left_val = max(v - 64, -63)
                    self.pan_value.setText(str(left_val))
                else:
                    self.pan_value.setText(str(v))

            if self.pan:
                self.pan.update_color(v)

        def _pan_reset(self):
            if self.pan:
                self.pan.setValue(PAN_CENTER_CC)

        def apply_scale(self, h_scale: float, v_scale: float):
            self.h_scale_factor = h_scale
            self.v_scale_factor = v_scale

            self.setMinimumWidth(int(BASE_STRIP_WIDTH * h_scale))
            self.setMaximumWidth(int(BASE_STRIP_WIDTH * h_scale) + 40)

            title_f = int(max(9, BASE_FONT * h_scale))
            small_f = int(max(8, BASE_SMALL_FONT * h_scale))
        
            current_color = "rgb(20,20,225)" if self.is_stereo_pair else "black"
            self.title_label.setStyleSheet(f"font-size: {title_f}px; font-weight: bold; color: {current_color};")
        
            if self.pan_value:
                self.pan_value.setStyleSheet(f"font-size: {small_f}px;")

            if self.mute:
                self.mute.setFixedHeight(int(BASE_BUTTON_H * v_scale))

            if self.pan:
                size = int(BASE_PAN_DIAL * h_scale)
                self.pan.setFixedSize(size, size)

            if self.scribble:
                scribble_h = int(max(40, 60 * v_scale))
                self.scribble.setMaximumHeight(scribble_h)

            self.fader.apply_scale(h_scale, v_scale)
class MixerWindow(QtWidgets.QMainWindow):
    def __init__(self, csv_path: str):
        super().__init__()
        self.setWindowTitle("DM4800 Mixer Controller (v0.4.0)")
        self.resize(1700, 1000)

        self.setStyleSheet("""
            QMainWindow { background-color: rgb(115,115,116); color: black; }
            QWidget { background-color: rgb(115,115,116); color: black; }
            QScrollArea { background-color: rgb(115,115,116); color: black; }
            QLabel { color: black; }
            QPushButton { color: black; }
            QComboBox { color: black; }
        """)

        self.settings = SettingsManager()
        self.mappings = load_midi_mappings(csv_path)
        self.scribble_manager = ScribbleStripManager()
        self.mono_stereo_manager = MonoStereoManager()
        self.midi = MidiManager()
        self.midi.midi_message_received.connect(self.on_midi_cc)
        
        self.midi_monitor = MidiMonitorWindow()
        self.midi.midi_message_received.connect(lambda ch, cc, val: self.midi_monitor.add_message("IN ", ch, cc, val))
        self.midi.midi_message_sent.connect(lambda ch, cc, val: self.midi_monitor.add_message("OUT", ch, cc, val))

        self.h_scale_factor = self.settings.get('horizontal_zoom', 100) / 100.0
        self.v_scale_factor = self.settings.get('vertical_zoom', 100) / 100.0

        self._create_menus()

        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        self.main_v = QtWidgets.QVBoxLayout(central)
        self.main_v.setContentsMargins(8,8,8,8)
        self.main_v.setSpacing(8)

        self._build_zoom_row()

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        self.main_v.addWidget(scroll, 1)

        container = QtWidgets.QWidget()
        self.rows_v = QtWidgets.QVBoxLayout(container)
        self.rows_v.setContentsMargins(4,4,4,4)
        self.rows_v.setSpacing(12)
        scroll.setWidget(container)

        self.channel_widgets: Dict[Tuple[str,int], ChannelStrip] = {}
        self.bus_widgets: Dict[int, ChannelStrip] = {}
        self.aux_widgets: Dict[int, ChannelStrip] = {}
        self.master_widget: Optional[ChannelStrip] = None
        self.stereo_pairs: Dict[str, ChannelStrip] = {}

        self._build_rows()
        self._setup_stereo_pairs()
        self._load_scribble_strips()
        self._install_zoom_shortcuts()

        if self.settings.get('remember_ports', True):
            self._auto_open_midi_ports()

    def closeEvent(self, event):
        self.settings.save_settings()
        self.midi.close()
        self.midi_monitor.close()
        super().closeEvent(event)

    def _create_menus(self):
        menubar = self.menuBar()
        midi_menu = menubar.addMenu("MIDI")
        
        self.midi_in_action = QtWidgets.QAction("Select MIDI In", self)
        self.midi_in_action.triggered.connect(self._select_midi_in)
        midi_menu.addAction(self.midi_in_action)
        
        self.midi_out_action = QtWidgets.QAction("Select MIDI Out", self)
        self.midi_out_action.triggered.connect(self._select_midi_out)
        midi_menu.addAction(self.midi_out_action)
        
        midi_menu.addSeparator()
        
        self.remember_ports_action = QtWidgets.QAction("Remember port selection", self)
        self.remember_ports_action.setCheckable(True)
        self.remember_ports_action.setChecked(self.settings.get('remember_ports', True))
        self.remember_ports_action.toggled.connect(self._toggle_remember_ports)
        midi_menu.addAction(self.remember_ports_action)
        
        midi_menu.addSeparator()
        
        self.refresh_ports_action = QtWidgets.QAction("Refresh MIDI ports", self)
        self.refresh_ports_action.triggered.connect(self._refresh_midi_ports)
        midi_menu.addAction(self.refresh_ports_action)
        
        midi_menu.addSeparator()
        
        self.midi_monitor_action = QtWidgets.QAction("MIDI Monitor", self)
        self.midi_monitor_action.triggered.connect(self._show_midi_monitor)
        midi_menu.addAction(self.midi_monitor_action)
        
        about_menu = menubar.addMenu("About")
        self.about_action = QtWidgets.QAction("Version/Author/License", self)
        self.about_action.triggered.connect(self._show_about)
        about_menu.addAction(self.about_action)

    def _select_midi_in(self):
        ports = self.midi.list_input_ports()
        if not ports:
            QtWidgets.QMessageBox.information(self, "No MIDI Inputs", "No MIDI input ports available.")
            return
        
        dialog = MidiPortSelectionDialog("MIDI Input Port", ports, self.midi.current_in_port, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            selected_port = dialog.get_selected_port()
            if selected_port:
                success = self.midi.open_input(selected_port)
                if success and self.settings.get('remember_ports', True):
                    self.settings.set('midi_in_port', selected_port)
                    self.settings.save_settings()

    def _select_midi_out(self):
        ports = self.midi.list_output_ports()
        if not ports:
            QtWidgets.QMessageBox.information(self, "No MIDI Outputs", "No MIDI output ports available.")
            return
        
        dialog = MidiPortSelectionDialog("MIDI Output Port", ports, self.midi.current_out_port, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            selected_port = dialog.get_selected_port()
            if selected_port:
                success = self.midi.open_output(selected_port)
                if success and self.settings.get('remember_ports', True):
                    self.settings.set('midi_out_port', selected_port)
                    self.settings.save_settings()

    def _toggle_remember_ports(self, checked: bool):
        self.settings.set('remember_ports', checked)
        self.settings.save_settings()

    def _refresh_midi_ports(self):
        saved_in = self.settings.get('midi_in_port', '')
        saved_out = self.settings.get('midi_out_port', '')
        
        available_in = self.midi.list_input_ports()
        available_out = self.midi.list_output_ports()
        
        missing_ports = []
        if saved_in and saved_in not in available_in:
            missing_ports.append(f"MIDI Input: {saved_in}")
        if saved_out and saved_out not in available_out:
            missing_ports.append(f"MIDI Output: {saved_out}")
        
        if missing_ports:
            message = "The following previously selected ports are no longer available:\n\n"
            message += "\n".join(missing_ports)
            QtWidgets.QMessageBox.warning(self, "MIDI Ports Missing", message)
        else:
            QtWidgets.QMessageBox.information(self, "MIDI Ports", "MIDI ports refreshed successfully.")

    def _show_midi_monitor(self):
        self.midi_monitor.show()
        self.midi_monitor.raise_()
        self.midi_monitor.activateWindow()

    def _show_about(self):
        about_text = """Version: v0.4.0
Author: Your Name
License: MIT License"""
        QtWidgets.QMessageBox.about(self, "About DM4800 Mixer Controller", about_text)

    def _auto_open_midi_ports(self):
        saved_in = self.settings.get('midi_in_port', '')
        saved_out = self.settings.get('midi_out_port', '')

        if saved_in:
            available_in = self.midi.list_input_ports()
            if saved_in in available_in:
                self.midi.open_input(saved_in)
                print(f"Auto-opened MIDI input: {saved_in}")

        if saved_out:
            available_out = self.midi.list_output_ports()
            if saved_out in available_out:
                self.midi.open_output(saved_out)
                print(f"Auto-opened MIDI output: {saved_out}")

    def _build_zoom_row(self):
        zoom_h = QtWidgets.QHBoxLayout()

        zoom_h.addWidget(QtWidgets.QLabel("H-Zoom"))
        self.h_zoom_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.h_zoom_slider.setRange(50, 125)
        self.h_zoom_slider.setValue(int(self.h_scale_factor * 100))
        self.h_zoom_slider.valueChanged.connect(self._on_h_zoom_changed)
        self.h_zoom_value_lbl = QtWidgets.QLabel(f"{int(self.h_scale_factor * 100)}%")
        zoom_h.addWidget(self.h_zoom_slider, 1)
        zoom_h.addWidget(self.h_zoom_value_lbl)

        zoom_h.addSpacing(20)

        zoom_h.addWidget(QtWidgets.QLabel("V-Zoom"))
        self.v_zoom_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.v_zoom_slider.setRange(50, 125)
        self.v_zoom_slider.setValue(int(self.v_scale_factor * 100))
        self.v_zoom_slider.valueChanged.connect(self._on_v_zoom_changed)
        self.v_zoom_value_lbl = QtWidgets.QLabel(f"{int(self.v_scale_factor * 100)}%")
        zoom_h.addWidget(self.v_zoom_slider, 1)
        zoom_h.addWidget(self.v_zoom_value_lbl)

        self.main_v.addLayout(zoom_h)

    def _install_zoom_shortcuts(self):
        zoom_in = QtWidgets.QShortcut(QtGui.QKeySequence.ZoomIn, self)
        zoom_out = QtWidgets.QShortcut(QtGui.QKeySequence.ZoomOut, self)
        zoom_reset = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+0"), self)
        zoom_reset_mac = QtWidgets.QShortcut(QtGui.QKeySequence("Meta+0"), self)

        zoom_in.activated.connect(lambda: self.h_zoom_slider.setValue(min(125, self.h_zoom_slider.value()+10)))
        zoom_out.activated.connect(lambda: self.h_zoom_slider.setValue(max(50, self.h_zoom_slider.value()-10)))
        zoom_reset.activated.connect(lambda: (self.h_zoom_slider.setValue(100), self.v_zoom_slider.setValue(100)))
        zoom_reset_mac.activated.connect(lambda: (self.h_zoom_slider.setValue(100), self.v_zoom_slider.setValue(100)))

    def _build_rows(self):
        def add_channel_row(start, end, include_pan=True, include_master=False):
            row = QtWidgets.QWidget()
            h = QtWidgets.QHBoxLayout(row)
            h.setContentsMargins(2,2,2,2)
            h.setSpacing(6)
            
            strips_added = 0
            ch = start
            
            while ch <= end and strips_added < 24:
                scribble_key = f"Channel {ch}"
                
                if self.mono_stereo_manager.is_stereo_pair(scribble_key):
                    if self.mono_stereo_manager.is_stereo_left(scribble_key):
                        partner_key = self.mono_stereo_manager.get_stereo_partner(scribble_key)
                        partner_num = int(partner_key.split()[1]) if partner_key else 0
                        
                        strip = ChannelStrip(f"Ch {ch}", scribble_key=scribble_key, has_pan=include_pan, 
                                           has_mute=True, has_scribble=True,
                                           h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, 
                                           is_master=False, mono_stereo_manager=self.mono_stereo_manager,
                                           is_stereo_pair=True, stereo_partner_num=partner_num)
                        h.addWidget(strip)
                        self.channel_widgets[("channel", ch)] = strip
                        self._wire_strip_controls("channel", ch, strip)
                        strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                        
                        strips_added += 2
                        ch += 2
                    else:
                        ch += 1
                else:
                    strip = ChannelStrip(f"Ch {ch}", scribble_key=scribble_key, has_pan=include_pan, 
                                       has_mute=True, has_scribble=True,
                                       h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, 
                                       is_master=False, mono_stereo_manager=self.mono_stereo_manager,
                                       is_stereo_pair=False)
                    h.addWidget(strip)
                    self.channel_widgets[("channel", ch)] = strip
                    self._wire_strip_controls("channel", ch, strip)
                    strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                    
                    strips_added += 1
                    ch += 1
            
            if include_master:
                h.addSpacing(12)
                ms = ChannelStrip("Master", has_pan=False, has_mute=False, has_scribble=False,
                                h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=True)
                h.addWidget(ms)
                self.master_widget = ms
                self._wire_strip_controls("master", 1, ms)
            
            self.rows_v.addWidget(row)
            return ch

        next_ch = add_channel_row(1, 64, include_pan=True, include_master=True)
        
        if next_ch <= 64:
            next_ch = add_channel_row(next_ch, 64, include_pan=True, include_master=False)
        
        if next_ch <= 64:
            add_channel_row(next_ch, 64, include_pan=True, include_master=False)

        bus_row = QtWidgets.QWidget()
        hbus = QtWidgets.QHBoxLayout(bus_row)
        hbus.setContentsMargins(2,2,2,2)
        hbus.setSpacing(6)
        for b in range(1,25):
            scribble_key = f"Bus {b}"
            strip = ChannelStrip(f"Bus {b}", scribble_key=scribble_key, has_pan=False, has_mute=True, has_scribble=True,
                               h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=False, mono_stereo_manager=self.mono_stereo_manager)
            hbus.addWidget(strip)
            self.bus_widgets[b] = strip
            self._wire_strip_controls("bus", b, strip)
            strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
        self.rows_v.addWidget(bus_row)

        aux_row = QtWidgets.QWidget()
        haux = QtWidgets.QHBoxLayout(aux_row)
        haux.setContentsMargins(2,2,2,2)
        haux.setSpacing(6)
        for a in range(1,13):
            scribble_key = f"Aux {a}"
            strip = ChannelStrip(f"Aux {a}", scribble_key=scribble_key, has_pan=False, has_mute=True, has_scribble=True,
                               h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=False, mono_stereo_manager=self.mono_stereo_manager)
            haux.addWidget(strip)
            self.aux_widgets[a] = strip
            self._wire_strip_controls("aux", a, strip)
            strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
        self.rows_v.addWidget(aux_row)

        self.rows_v.addStretch(1)

    def _wire_strip_controls(self, section: str, number: int, strip: ChannelStrip):
        def map_for(t: str) -> Optional[MidiMapping]:
            return self.mappings.get((section, number, t))

        key = f"{section.title()} {number}"
        
        m = map_for('fader')
        if m:
            strip.fader.valueChanged.connect(lambda v, m=m: self.midi.send_cc(m.midi_channel, m.cc_number, v))

        if strip.mute:
            m = map_for('mute')
            if m:
                strip.mute.toggledCC.connect(lambda v, m=m: self.midi.send_cc(m.midi_channel, m.cc_number, v))

        if strip.pan:
            m = map_for('pan')
            if m:
                strip.pan.valueChanged.connect(lambda v, m=m: self.midi.send_cc(m.midi_channel, m.cc_number, v))

    def _setup_stereo_pairs(self):
        for number, strip in self.bus_widgets.items():
            key = f"Bus {number}"
            partner_key = self.mono_stereo_manager.get_stereo_partner(key)
            if partner_key:
                partner_num = int(partner_key.split()[1])
                partner_strip = self.bus_widgets.get(partner_num)
                if partner_strip:
                    strip.stereo_partner_strip = partner_strip

        for number, strip in self.aux_widgets.items():
            key = f"Aux {number}"
            partner_key = self.mono_stereo_manager.get_stereo_partner(key)
            if partner_key:
                partner_num = int(partner_key.split()[1])
                partner_strip = self.aux_widgets.get(partner_num)
                if partner_strip:
                    strip.stereo_partner_strip = partner_strip

    def _load_scribble_strips(self):
        for (section, number), strip in self.channel_widgets.items():
            if section == "channel":
                key = f"Channel {number}"
                text = self.scribble_manager.get_scribble_text(key)
                strip.set_scribble_text(text)

        for number, strip in self.bus_widgets.items():
            key = f"Bus {number}"
            text = self.scribble_manager.get_scribble_text(key)
            strip.set_scribble_text(text)

        for number, strip in self.aux_widgets.items():
            key = f"Aux {number}"
            text = self.scribble_manager.get_scribble_text(key)
            strip.set_scribble_text(text)

    def _on_scribble_text_changed(self, key: str, new_text: str):
        self.scribble_manager.set_scribble_text(key, new_text)

    def _on_h_zoom_changed(self, val: int):
        self.h_scale_factor = val / 100.0
        self.h_zoom_value_lbl.setText(f"{val}%")
        self.settings.set('horizontal_zoom', val)
        self._apply_all_scales()

    def _on_v_zoom_changed(self, val: int):
        self.v_scale_factor = val / 100.0
        self.v_zoom_value_lbl.setText(f"{val}%")
        self.settings.set('vertical_zoom', val)
        self._apply_all_scales()

    def _apply_all_scales(self):
        for key, strip in self.channel_widgets.items():
            strip.apply_scale(self.h_scale_factor, self.v_scale_factor)
        for strip in self.bus_widgets.values():
            strip.apply_scale(self.h_scale_factor, self.v_scale_factor)
        if self.master_widget:
            self.master_widget.apply_scale(self.h_scale_factor, self.v_scale_factor)
        for strip in self.aux_widgets.values():
            strip.apply_scale(self.h_scale_factor, self.v_scale_factor)

    def on_midi_cc(self, channel_1: int, cc: int, value: int):
        if not hasattr(self, "_rev_index"):
            self._rev_index = {}
            for key, mm in self.mappings.items():
                self._rev_index[(mm.midi_channel, mm.cc_number)] = key

        key = self._rev_index.get((channel_1, cc))
        if not key:
            return
        section, number, typ = key

        if section == "channel":
            strip = self.channel_widgets.get((section, number))
            if not strip:
                channel_key = f"Channel {number}"
                if self.mono_stereo_manager.is_stereo_right(channel_key):
                    left_channel_key = self.mono_stereo_manager.get_stereo_partner(channel_key)
                    if left_channel_key:
                        left_number = int(left_channel_key.split()[1])
                        strip = self.channel_widgets.get((section, left_number))
                
                if not strip:
                    return
        elif section == "bus":
            strip = self.bus_widgets.get(number)
        elif section == "aux":
            strip = self.aux_widgets.get(number)
        elif section == "master":
            strip = self.master_widget
        else:
            strip = None
            
        if not strip:
            return

        if typ == "fader":
            strip.fader.setValue(value)
        elif typ == "mute" and strip.mute:
            strip.mute.set_from_cc(value)
        elif typ == "pan" and strip.pan:
            strip.pan.blockSignals(True)
            strip.pan.setValue(value)
            strip.pan.blockSignals(False)
            strip._on_pan_changed(value)

    def _get_strip_by_key(self, key: str) -> Optional[ChannelStrip]:
        try:
            parts = key.split()
            section = parts[0].lower()
            number = int(parts[1])

            if section == "channel":
                return self.channel_widgets.get(("channel", number))
            elif section == "bus":
                return self.bus_widgets.get(number)
            elif section == "aux":
                return self.aux_widgets.get(number)
            elif section == "master":
                return self.master_widget
        except:
            pass
        return None

def locate_csv() -> Optional[str]:
    here = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(here, CSV_DEFAULT)
    if os.path.isfile(candidate):
        return candidate
    candidate = os.path.join(os.getcwd(), CSV_DEFAULT)
    if os.path.isfile(candidate):
        return candidate
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    path, _ = QtWidgets.QFileDialog.getOpenFileName(None, "Locate DM4800_CC_Values.csv", os.getcwd(), "CSV Files (*.csv)")
    return path or None

def main():
    app = QtWidgets.QApplication(sys.argv)
    csv_path = locate_csv()
    if not csv_path:
        QtWidgets.QMessageBox.critical(None, "Error", "Cannot proceed without DM4800_CC_Values.csv")
        return
    win = MixerWindow(csv_path)
    win.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()                
