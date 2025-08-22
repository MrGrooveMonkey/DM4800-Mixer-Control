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
        self.slider
