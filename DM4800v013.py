#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DM4800 Mixer Controller v0.4.1 - Enhanced Stereo Pair Display
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
        return {
            'midi_in_port': '', 
            'midi_out_port': '', 
            'horizontal_zoom': 100, 
            'vertical_zoom': 100, 
            'remember_ports': True,
            'stereo_display_mode': 'linked_pair',  # 'wide_fader' or 'linked_pair'
            'show_stereo_indicators': True
        }

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

    def connect_input(self, port_name: str):
        if self._inport:
            self._inport.close()
            self._inport = None
        self.current_in_port = ""
        
        if port_name:
            try:
                self._inport = mido.open_input(port_name, callback=self._on_message)
                self.current_in_port = port_name
            except Exception as e:
                print(f"Failed to open input port '{port_name}': {e}")

    def connect_output(self, port_name: str):
        if self._outport:
            self._outport.close()
            self._outport = None
        self.current_out_port = ""
        
        if port_name:
            try:
                self._outport = mido.open_output(port_name)
                self.current_out_port = port_name
            except Exception as e:
                print(f"Failed to open output port '{port_name}': {e}")

    def _on_message(self, msg: mido.Message):
        if msg.type == 'control_change':
            self.midi_message_received.emit(msg.channel, msg.control, msg.value)

    def send_cc(self, channel: int, cc: int, value: int):
        if self._outport and 0 <= channel <= 15 and 0 <= cc <= 127 and 0 <= value <= 127:
            msg = mido.Message('control_change', channel=channel, control=cc, value=value)
            try:
                self._outport.send(msg)
                self.midi_message_sent.emit(channel, cc, value)
            except Exception as e:
                print(f"Failed to send MIDI: {e}")

    def is_connected(self) -> bool:
        return (self._inport is not None) and (self._outport is not None)
        
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

class FaderScale(QtWidgets.QWidget):
    def __init__(self, is_master: bool = False, parent=None):
        super().__init__(parent)
        self.is_master = is_master
        self.db_points = MASTER_DB_POINTS if is_master else CHANNEL_DB_POINTS
        self.setMinimumWidth(24)
        
    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        
        rect = self.rect()
        painter.fillRect(rect, QtGui.QColor(80, 80, 80))
        
        painter.setPen(QtGui.QPen(QtGui.QColor(200, 200, 200), 1))
        font = QtGui.QFont("Arial", 7)
        painter.setFont(font)
        
        for cc_val, db_str in self.db_points:
            y = int(rect.height() - (cc_val / 127.0) * rect.height())
            painter.drawLine(0, y, 8, y)
            painter.drawText(10, y + 3, db_str)

class FaderWithScale(QtWidgets.QWidget):
    valueChanged = QtCore.pyqtSignal(int)

    def __init__(self, initial: int = FADER_ZERO_DB_CC, h_scale: float = 1.0, v_scale: float = 1.0, 
                 is_master: bool = False, is_stereo: bool = False, parent=None):
        super().__init__(parent)
        self.h_scale_factor = h_scale
        self.v_scale_factor = v_scale
        self.is_master = is_master
        self.is_stereo = is_stereo
        
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        self.scale = FaderScale(is_master=is_master)
        layout.addWidget(self.scale)

        self.fader = QtWidgets.QSlider(QtCore.Qt.Vertical)
        self.fader.setRange(0, 127)
        self.fader.setValue(initial)
        self.fader.valueChanged.connect(self.valueChanged.emit)
        layout.addWidget(self.fader)

        self.apply_scale(h_scale, v_scale)

    def apply_scale(self, h_scale: float, v_scale: float):
        self.h_scale_factor = h_scale
        self.v_scale_factor = v_scale
        
        width = int(BASE_STRIP_WIDTH * h_scale)
        if self.is_stereo:
            width = int(width * 2.0)  # 2x width for wide stereo faders
        
        height = int(BASE_FADER_HEIGHT * v_scale)
        
        self.setMinimumWidth(width)
        self.setMaximumWidth(width)
        self.setMinimumHeight(height)
        self.setMaximumHeight(height)

    def setValue(self, value: int):
        self.fader.setValue(value)

    def value(self) -> int:
        return self.fader.value()

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

    def __init__(self, title: str, scribble_key: str = "", has_pan: bool=True, has_mute: bool=True, 
                 has_scribble: bool=True, h_scale: float=1.0, v_scale: float=1.0, is_master: bool=False, 
                 mono_stereo_manager=None, is_stereo_pair: bool=False, stereo_partner_num: int=0, parent=None):
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
            if scribble_key and "Channel" in scribble_key:
                channel_num = int(scribble_key.split()[-1])
                display_title = f"Ch {channel_num} - {stereo_partner_num}"
            elif scribble_key and "Bus" in scribble_key:
                bus_num = int(scribble_key.split()[-1])
                display_title = f"Bus {bus_num} - {stereo_partner_num}"
            elif scribble_key and "Aux" in scribble_key:
                aux_num = int(scribble_key.split()[-1])
                display_title = f"Aux {aux_num} - {stereo_partner_num}"
            title_color = "rgb(20,20,225)"
        elif scribble_key and "Channel" in scribble_key:
            channel_num = int(scribble_key.split()[-1])
            display_title = f"CH {channel_num}"
        elif scribble_key and "Bus" in scribble_key:
            bus_num = int(scribble_key.split()[-1])
            display_title = f"BUS {bus_num}"
        elif scribble_key and "Aux" in scribble_key:
            aux_num = int(scribble_key.split()[-1])
            display_title = f"AUX {aux_num}"

        self.v = QtWidgets.QVBoxLayout(self)
        self.v.setContentsMargins(6,6,6,6)
        self.v.setSpacing(6)
        self.setMinimumWidth(int(BASE_STRIP_WIDTH * self.h_scale_factor))

        # Title with stereo indicator
        title_layout = QtWidgets.QHBoxLayout()
        self.title_label = QtWidgets.QLabel(display_title)
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.title_label.setStyleSheet(f"font-weight: bold; color: {title_color};")
        title_layout.addWidget(self.title_label)
        
        # Add stereo indicator badge
        if is_stereo_pair:
            self.stereo_badge = QtWidgets.QLabel("ST")
            self.stereo_badge.setStyleSheet("""
                QLabel {
                    background-color: rgb(20,20,225);
                    color: white;
                    border-radius: 8px;
                    padding: 1px 4px;
                    font-size: 8px;
                    font-weight: bold;
                }
            """)
            self.stereo_badge.setFixedSize(20, 16)
            self.stereo_badge.setAlignment(QtCore.Qt.AlignCenter)
            title_layout.addWidget(self.stereo_badge)
        
        self.v.addLayout(title_layout)

        if has_pan:
            self.pan = PanDial()
            self.pan.setRange(0,127)
            self.pan.setValue(PAN_CENTER_CC)
            self.pan.setNotchesVisible(True)
            self.pan.doubleClicked.connect(self._pan_reset)
            self.pan.valueChanged.connect(self._on_pan_changed)
            self.v.addWidget(self.pan, 0, QtCore.Qt.AlignHCenter)

            # Use "Bal" for stereo pairs, "Pan" for mono
            self.pan_value = QtWidgets.QLabel("C")
            self.pan_value.setAlignment(QtCore.Qt.AlignCenter)
            self.v.addWidget(self.pan_value)
            
            # Add pan/balance label
            pan_label_text = "Bal" if is_stereo_pair else "Pan"
            self.pan_label = QtWidgets.QLabel(pan_label_text)
            self.pan_label.setAlignment(QtCore.Qt.AlignCenter)
            self.pan_label.setStyleSheet("font-size: 8px; color: gray;")
            self.v.addWidget(self.pan_label)
            
            self.pan.update_color(PAN_CENTER_CC)
        else:
            self.pan = None
            self.pan_value = None
            self.pan_label = None

        if has_mute:
            self.mute = MuteButton()
            self.v.addWidget(self.mute)
        else:
            self.mute = None

        initial_val = MASTER_ZERO_DB_CC if is_master else FADER_ZERO_DB_CC
        self.fader = FaderWithScale(initial=initial_val, h_scale=self.h_scale_factor, 
                                  v_scale=self.v_scale_factor, is_master=is_master, is_stereo=False)
        self.v.addWidget(self.fader, 1)

        if has_scribble:
            self.scribble = QtWidgets.QLineEdit()
            self.scribble.setMaxLength(MAX_SCRIBBLE_LENGTH)
            self.scribble.setAlignment(QtCore.Qt.AlignCenter)
            self.scribble.setText(DEFAULT_SCRIBBLE_TEXT)
            self.scribble.textChanged.connect(self._on_scribble_changed)
            self.v.addWidget(self.scribble)
        else:
            self.scribble = None

        self.apply_scale(h_scale, v_scale)

    def _pan_reset(self):
        self.pan.setValue(PAN_CENTER_CC)

    def _on_pan_changed(self, val: int):
        if val == PAN_CENTER_CC:
            self.pan_value.setText("C")
        elif val < PAN_CENTER_CC:
            self.pan_value.setText(f"L{PAN_CENTER_CC - val}")
        else:
            self.pan_value.setText(f"R{val - PAN_CENTER_CC}")
        self.pan.update_color(val)

    def _on_scribble_changed(self):
        if self.scribble_key:
            self.scribbleTextChanged.emit(self.scribble_key, self.scribble.text())

    def get_fader_value(self) -> int:
        return self.fader.value()

    def set_fader_value(self, val: int):
        self.fader.setValue(val)

    def get_pan_value(self) -> int:
        return self.pan.value() if self.pan else PAN_CENTER_CC

    def set_pan_value(self, val: int):
        if self.pan:
            self.pan.setValue(val)

    def get_mute_value(self) -> bool:
        return self.mute.isChecked() if self.mute else False

    def set_mute_value(self, on: bool):
        if self.mute:
            self.mute.setChecked(on)

    def set_scribble_text(self, text: str):
        if self.scribble:
            self.scribble.blockSignals(True)
            self.scribble.setText(text)
            self.scribble.blockSignals(False)

    def apply_scale(self, h_scale: float, v_scale: float):
        self.h_scale_factor = h_scale
        self.v_scale_factor = v_scale
        
        width = int(BASE_STRIP_WIDTH * h_scale)
        self.setMinimumWidth(width)
        self.setMaximumWidth(width)
        
        title_f = int(max(8, BASE_FONT * h_scale))
        small_f = int(max(8, BASE_SMALL_FONT * h_scale))
        
        current_color = "rgb(20,20,225)" if self.is_stereo_pair else "black"
        self.title_label.setStyleSheet(f"font-size: {title_f}px; font-weight: bold; color: {current_color};")
        
        if self.pan_value:
            self.pan_value.setStyleSheet(f"font-size: {small_f}px;")

        if self.pan_label:
            self.pan_label.setStyleSheet(f"font-size: {int(small_f*0.8)}px; color: gray;")

        if self.mute:
            self.mute.setFixedHeight(int(BASE_BUTTON_H * v_scale))

        if self.pan:
            size = int(BASE_PAN_DIAL * h_scale)
            self.pan.setFixedSize(size, size)

        if self.scribble:
            scribble_h = int(max(40, 60 * v_scale))
            self.scribble.setMaximumHeight(scribble_h)

        self.fader.apply_scale(h_scale, v_scale)
        
class WideStereoStrip(QtWidgets.QFrame):
    """Wide stereo channel strip that combines two channels into a single wide control strip"""
    scribbleTextChanged = QtCore.pyqtSignal(str, str)

    def __init__(self, title: str, left_scribble_key: str = "", right_scribble_key: str = "", 
                 section_type: str = "channel", left_num: int = 1, right_num: int = 2,
                 has_pan: bool=True, has_mute: bool=True, has_scribble: bool=True, 
                 h_scale: float=1.0, v_scale: float=1.0, mono_stereo_manager=None, parent=None):
        super().__init__(parent)
        self.h_scale_factor = h_scale
        self.v_scale_factor = v_scale
        self.left_scribble_key = left_scribble_key
        self.right_scribble_key = right_scribble_key
        self.section_type = section_type
        self.left_num = left_num
        self.right_num = right_num
        self.mono_stereo_manager = mono_stereo_manager
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        
        # Enhanced visual styling for stereo pairs
        self.setStyleSheet("""
            WideStereoStrip {
                border: 2px solid rgb(20,20,225);
                border-radius: 4px;
                background-color: rgba(20,20,225,10);
            }
        """)

        # Create title with stereo indicators
        if section_type == "channel":
            display_title = f"CH {left_num}-{right_num}"
        elif section_type == "bus":
            display_title = f"BUS {left_num}-{right_num}"
        else:  # aux
            display_title = f"AUX {left_num}-{right_num}"

        self.v = QtWidgets.QVBoxLayout(self)
        self.v.setContentsMargins(6,6,6,6)
        self.v.setSpacing(6)
        
        # Double width for stereo
        self.setMinimumWidth(int(BASE_STRIP_WIDTH * 2.0 * self.h_scale_factor))

        # Title with L-R indicators
        title_layout = QtWidgets.QHBoxLayout()
        
        self.left_indicator = QtWidgets.QLabel("L")
        self.left_indicator.setStyleSheet("""
            QLabel {
                background-color: rgb(20,20,225);
                color: white;
                border-radius: 8px;
                padding: 1px 4px;
                font-size: 8px;
                font-weight: bold;
            }
        """)
        self.left_indicator.setFixedSize(16, 16)
        self.left_indicator.setAlignment(QtCore.Qt.AlignCenter)
        title_layout.addWidget(self.left_indicator)
        
        self.title_label = QtWidgets.QLabel(display_title)
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.title_label.setStyleSheet("font-weight: bold; color: rgb(20,20,225);")
        title_layout.addWidget(self.title_label, 1)
        
        self.right_indicator = QtWidgets.QLabel("R")
        self.right_indicator.setStyleSheet("""
            QLabel {
                background-color: rgb(20,20,225);
                color: white;
                border-radius: 8px;
                padding: 1px 4px;
                font-size: 8px;
                font-weight: bold;
            }
        """)
        self.right_indicator.setFixedSize(16, 16)
        self.right_indicator.setAlignment(QtCore.Qt.AlignCenter)
        title_layout.addWidget(self.right_indicator)
        
        self.v.addLayout(title_layout)

        if has_pan:
            self.pan = PanDial()
            self.pan.setRange(0,127)
            self.pan.setValue(PAN_CENTER_CC)
            self.pan.setNotchesVisible(True)
            self.pan.doubleClicked.connect(self._pan_reset)
            self.pan.valueChanged.connect(self._on_pan_changed)
            self.v.addWidget(self.pan, 0, QtCore.Qt.AlignHCenter)

            # Balance control for stereo
            self.pan_value = QtWidgets.QLabel("C")
            self.pan_value.setAlignment(QtCore.Qt.AlignCenter)
            self.v.addWidget(self.pan_value)
            
            # Add "Bal" label
            self.pan_label = QtWidgets.QLabel("Bal")
            self.pan_label.setAlignment(QtCore.Qt.AlignCenter)
            self.pan_label.setStyleSheet("font-size: 8px; color: rgb(20,20,225);")
            self.v.addWidget(self.pan_label)
            
            self.pan.update_color(PAN_CENTER_CC)
        else:
            self.pan = None
            self.pan_value = None
            self.pan_label = None

        if has_mute:
            self.mute = MuteButton()
            self.v.addWidget(self.mute)
        else:
            self.mute = None

        # Wide fader (2x width)
        self.fader = FaderWithScale(initial=FADER_ZERO_DB_CC, h_scale=self.h_scale_factor, 
                                  v_scale=self.v_scale_factor, is_master=False, is_stereo=True)
        self.v.addWidget(self.fader, 1)

        if has_scribble:
            self.scribble = QtWidgets.QLineEdit()
            self.scribble.setMaxLength(MAX_SCRIBBLE_LENGTH)
            self.scribble.setAlignment(QtCore.Qt.AlignCenter)
            self.scribble.setText(DEFAULT_SCRIBBLE_TEXT)
            self.scribble.textChanged.connect(self._on_scribble_changed)
            self.v.addWidget(self.scribble)
        else:
            self.scribble = None

        self.apply_scale(h_scale, v_scale)

    def _pan_reset(self):
        self.pan.setValue(PAN_CENTER_CC)

    def _on_pan_changed(self, val: int):
        if val == PAN_CENTER_CC:
            self.pan_value.setText("C")
        elif val < PAN_CENTER_CC:
            self.pan_value.setText(f"L{PAN_CENTER_CC - val}")
        else:
            self.pan_value.setText(f"R{val - PAN_CENTER_CC}")
        self.pan.update_color(val)

    def _on_scribble_changed(self):
        # Update both left and right scribble keys with the same text
        text = self.scribble.text()
        if self.left_scribble_key:
            self.scribbleTextChanged.emit(self.left_scribble_key, text)
        if self.right_scribble_key:
            self.scribbleTextChanged.emit(self.right_scribble_key, text)

    def get_fader_value(self) -> int:
        return self.fader.value()

    def set_fader_value(self, val: int):
        self.fader.setValue(val)

    def get_pan_value(self) -> int:
        return self.pan.value() if self.pan else PAN_CENTER_CC

    def set_pan_value(self, val: int):
        if self.pan:
            self.pan.setValue(val)

    def get_mute_value(self) -> bool:
        return self.mute.isChecked() if self.mute else False

    def set_mute_value(self, on: bool):
        if self.mute:
            self.mute.setChecked(on)

    def set_scribble_text(self, text: str):
        if self.scribble:
            self.scribble.blockSignals(True)
            self.scribble.setText(text)
            self.scribble.blockSignals(False)

    def apply_scale(self, h_scale: float, v_scale: float):
        self.h_scale_factor = h_scale
        self.v_scale_factor = v_scale
        
        # Double width for wide stereo strips
        width = int(BASE_STRIP_WIDTH * 2.0 * h_scale)
        self.setMinimumWidth(width)
        self.setMaximumWidth(width)
        
        title_f = int(max(8, BASE_FONT * h_scale))
        small_f = int(max(8, BASE_SMALL_FONT * h_scale))
        
        self.title_label.setStyleSheet(f"font-size: {title_f}px; font-weight: bold; color: rgb(20,20,225);")
        
        if self.pan_value:
            self.pan_value.setStyleSheet(f"font-size: {small_f}px;")

        if self.pan_label:
            self.pan_label.setStyleSheet(f"font-size: {int(small_f*0.8)}px; color: rgb(20,20,225);")

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
        self.setWindowTitle("DM4800 Mixer Controller (v0.4.1)")
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

        # Widget storage
        self.channel_widgets = {}
        self.bus_widgets = {}
        self.aux_widgets = {}
        self.master_widget = None

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
        self.rows_v.setSpacing(8)
        
        self._build_rows()
        
        container.setLayout(self.rows_v)
        scroll.setWidget(container)

        self._load_all_scribble_texts()
        self._build_shortcuts()

    def _create_menus(self):
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("&File")
        
        midi_monitor_action = file_menu.addAction("MIDI Monitor")
        midi_monitor_action.triggered.connect(self.midi_monitor.show)

        file_menu.addSeparator()
        
        exit_action = file_menu.addAction("Exit")
        exit_action.triggered.connect(self.close)

        # MIDI menu
        midi_menu = menubar.addMenu("&MIDI")
        
        select_input_action = midi_menu.addAction("Select Input Port...")
        select_input_action.triggered.connect(self._select_midi_input)
        
        select_output_action = midi_menu.addAction("Select Output Port...")
        select_output_action.triggered.connect(self._select_midi_output)

        midi_menu.addSeparator()
        
        disconnect_action = midi_menu.addAction("Disconnect All")
        disconnect_action.triggered.connect(self._disconnect_midi)

        # Settings menu
        settings_menu = menubar.addMenu("&Settings")
        
        # Stereo display toggle
        self.stereo_display_action = QtWidgets.QAction("Display Stereo Pairs as Single Channel Strips", self)
        self.stereo_display_action.setCheckable(True)
        self.stereo_display_action.setChecked(self.settings.get('stereo_display_mode') == 'wide_fader')
        self.stereo_display_action.triggered.connect(self._toggle_stereo_display_mode)
        settings_menu.addAction(self.stereo_display_action)

    def _build_zoom_row(self):
        zoom_row = QtWidgets.QWidget()
        h_zoom = QtWidgets.QHBoxLayout(zoom_row)
        h_zoom.setContentsMargins(0,0,0,0)
        h_zoom.setSpacing(12)

        h_zoom.addWidget(QtWidgets.QLabel("H-Zoom:"))
        self.h_zoom_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.h_zoom_slider.setRange(50, 125)
        self.h_zoom_slider.setValue(int(self.h_scale_factor * 100))
        self.h_zoom_slider.valueChanged.connect(self._on_h_zoom_changed)
        h_zoom.addWidget(self.h_zoom_slider)

        self.h_zoom_value_lbl = QtWidgets.QLabel(f"{int(self.h_scale_factor * 100)}%")
        self.h_zoom_value_lbl.setMinimumWidth(40)
        h_zoom.addWidget(self.h_zoom_value_lbl)

        h_zoom.addSpacing(20)

        h_zoom.addWidget(QtWidgets.QLabel("V-Zoom:"))
        self.v_zoom_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.v_zoom_slider.setRange(50, 125)
        self.v_zoom_slider.setValue(int(self.v_scale_factor * 100))
        self.v_zoom_slider.valueChanged.connect(self._on_v_zoom_changed)
        h_zoom.addWidget(self.v_zoom_slider)

        self.v_zoom_value_lbl = QtWidgets.QLabel(f"{int(self.v_scale_factor * 100)}%")
        self.v_zoom_value_lbl.setMinimumWidth(40)
        h_zoom.addWidget(self.v_zoom_value_lbl)

        # Add stereo display checkbox
        h_zoom.addSpacing(20)
        self.stereo_checkbox = QtWidgets.QCheckBox("Display Stereo Pairs as Single Channel Strips")
        self.stereo_checkbox.setChecked(self.settings.get('stereo_display_mode') == 'wide_fader')
        self.stereo_checkbox.toggled.connect(self._on_stereo_checkbox_changed)
        h_zoom.addWidget(self.stereo_checkbox)

        h_zoom.addStretch()
        self.main_v.addWidget(zoom_row)
        
    def _toggle_stereo_display_mode(self):
        """Toggle stereo display mode via menu"""
        current_mode = self.settings.get('stereo_display_mode', 'linked_pair')
        new_mode = 'wide_fader' if current_mode == 'linked_pair' else 'linked_pair'
        self.settings.set('stereo_display_mode', new_mode)
        self.settings.save_settings()
        
        # Update checkbox to match
        self.stereo_checkbox.blockSignals(True)
        self.stereo_checkbox.setChecked(new_mode == 'wide_fader')
        self.stereo_checkbox.blockSignals(False)
        
        self._rebuild_interface()

    def _on_stereo_checkbox_changed(self, checked: bool):
        """Handle stereo display checkbox change"""
        new_mode = 'wide_fader' if checked else 'linked_pair'
        self.settings.set('stereo_display_mode', new_mode)
        self.settings.save_settings()
        
        # Update menu action to match
        self.stereo_display_action.blockSignals(True)
        self.stereo_display_action.setChecked(checked)
        self.stereo_display_action.blockSignals(False)
        
        self._rebuild_interface()

    def _rebuild_interface(self):
        """Rebuild the entire interface when stereo display mode changes"""
        # Clear existing widgets
        for i in reversed(range(self.rows_v.count())):
            child = self.rows_v.itemAt(i).widget()
            if child:
                child.setParent(None)
        
        # Clear widget dictionaries
        self.channel_widgets.clear()
        self.bus_widgets.clear()
        self.aux_widgets.clear()
        self.master_widget = None
        
        # Rebuild
        self._build_rows()
        self._load_all_scribble_texts()

    def _build_rows(self):
        use_wide_stereo = self.settings.get('stereo_display_mode') == 'wide_fader'
        
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
                        
                        if use_wide_stereo:
                            # Create wide stereo strip
                            strip = WideStereoStrip(
                                f"Ch {ch}-{partner_num}", 
                                left_scribble_key=scribble_key,
                                right_scribble_key=partner_key,
                                section_type="channel",
                                left_num=ch, 
                                right_num=partner_num,
                                has_pan=include_pan, 
                                has_mute=True, 
                                has_scribble=True,
                                h_scale=self.h_scale_factor, 
                                v_scale=self.v_scale_factor,
                                mono_stereo_manager=self.mono_stereo_manager
                            )
                        else:
                            # Create regular stereo pair with visual indicators
                            strip = ChannelStrip(
                                f"Ch {ch}", 
                                scribble_key=scribble_key, 
                                has_pan=include_pan, 
                                has_mute=True, 
                                has_scribble=True,
                                h_scale=self.h_scale_factor, 
                                v_scale=self.v_scale_factor, 
                                is_master=False, 
                                mono_stereo_manager=self.mono_stereo_manager,
                                is_stereo_pair=True, 
                                stereo_partner_num=partner_num
                            )
                        
                        h.addWidget(strip)
                        self.channel_widgets[("channel", ch)] = strip
                        self._wire_strip_controls("channel", ch, strip)
                        strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                        
                        strips_added += 2
                        ch += 2
                    else:
                        ch += 1
                else:
                    strip = ChannelStrip(
                        f"Ch {ch}", 
                        scribble_key=scribble_key, 
                        has_pan=include_pan, 
                        has_mute=True, 
                        has_scribble=True,
                        h_scale=self.h_scale_factor, 
                        v_scale=self.v_scale_factor, 
                        is_master=False, 
                        mono_stereo_manager=self.mono_stereo_manager,
                        is_stereo_pair=False
                    )
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

        # Add channel rows
        next_ch = add_channel_row(1, 64, include_pan=True, include_master=True)
        if next_ch <= 64:
            next_ch = add_channel_row(next_ch, 64, include_pan=True, include_master=False)
        if next_ch <= 64:
            add_channel_row(next_ch, 64, include_pan=True, include_master=False)
            
# Add bus row
        bus_row = QtWidgets.QWidget()
        hbus = QtWidgets.QHBoxLayout(bus_row)
        hbus.setContentsMargins(2,2,2,2)
        hbus.setSpacing(6)
        
        strips_added = 0
        b = 1
        
        while b <= 24 and strips_added < 24:
            scribble_key = f"Bus {b}"
            
            if self.mono_stereo_manager.is_stereo_pair(scribble_key):
                if self.mono_stereo_manager.is_stereo_left(scribble_key):
                    partner_key = self.mono_stereo_manager.get_stereo_partner(scribble_key)
                    partner_num = int(partner_key.split()[1]) if partner_key else 0
                    
                    if use_wide_stereo:
                        # Create wide stereo bus strip
                        strip = WideStereoStrip(
                            f"Bus {b}-{partner_num}", 
                            left_scribble_key=scribble_key,
                            right_scribble_key=partner_key,
                            section_type="bus",
                            left_num=b, 
                            right_num=partner_num,
                            has_pan=False, 
                            has_mute=True, 
                            has_scribble=True,
                            h_scale=self.h_scale_factor, 
                            v_scale=self.v_scale_factor,
                            mono_stereo_manager=self.mono_stereo_manager
                        )
                    else:
                        # Create regular stereo bus pair
                        strip = ChannelStrip(
                            f"Bus {b}", 
                            scribble_key=scribble_key, 
                            has_pan=False, 
                            has_mute=True, 
                            has_scribble=True,
                            h_scale=self.h_scale_factor, 
                            v_scale=self.v_scale_factor, 
                            is_master=False, 
                            mono_stereo_manager=self.mono_stereo_manager,
                            is_stereo_pair=True, 
                            stereo_partner_num=partner_num
                        )
                    
                    hbus.addWidget(strip)
                    self.bus_widgets[b] = strip
                    self._wire_strip_controls("bus", b, strip)
                    strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                    
                    strips_added += 2
                    b += 2
                else:
                    b += 1
            else:
                strip = ChannelStrip(f"Bus {b}", scribble_key=scribble_key, has_pan=False, has_mute=True, has_scribble=True,
                                   h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=False, mono_stereo_manager=self.mono_stereo_manager)
                hbus.addWidget(strip)
                self.bus_widgets[b] = strip
                self._wire_strip_controls("bus", b, strip)
                strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                
                strips_added += 1
                b += 1
        
        self.rows_v.addWidget(bus_row)

        # Add aux row
        aux_row = QtWidgets.QWidget()
        haux = QtWidgets.QHBoxLayout(aux_row)
        haux.setContentsMargins(2,2,2,2)
        haux.setSpacing(6)
        
        strips_added = 0
        a = 1
        
        while a <= 12 and strips_added < 12:
            scribble_key = f"Aux {a}"
        
            if self.mono_stereo_manager.is_stereo_pair(scribble_key):
                if self.mono_stereo_manager.is_stereo_left(scribble_key):
                    partner_key = self.mono_stereo_manager.get_stereo_partner(scribble_key)
                    partner_num = int(partner_key.split()[1]) if partner_key else 0
                    
                    if use_wide_stereo:
                        # Create wide stereo aux strip
                        strip = WideStereoStrip(
                            f"Aux {a}-{partner_num}", 
                            left_scribble_key=scribble_key,
                            right_scribble_key=partner_key,
                            section_type="aux",
                            left_num=a, 
                            right_num=partner_num,
                            has_pan=False, 
                            has_mute=True, 
                            has_scribble=True,
                            h_scale=self.h_scale_factor, 
                            v_scale=self.v_scale_factor,
                            mono_stereo_manager=self.mono_stereo_manager
                        )
                    else:
                        # Create regular stereo aux pair
                        strip = ChannelStrip(
                            f"Aux {a}", 
                            scribble_key=scribble_key, 
                            has_pan=False, 
                            has_mute=True, 
                            has_scribble=True,
                            h_scale=self.h_scale_factor, 
                            v_scale=self.v_scale_factor, 
                            is_master=False, 
                            mono_stereo_manager=self.mono_stereo_manager,
                            is_stereo_pair=True, 
                            stereo_partner_num=partner_num
                        )
                    
                    haux.addWidget(strip)
                    self.aux_widgets[a] = strip
                    self._wire_strip_controls("aux", a, strip)
                    strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                    
                    strips_added += 2
                    a += 2
                else:
                    a += 1
            else:
                strip = ChannelStrip(f"Aux {a}", scribble_key=scribble_key, has_pan=False, has_mute=True, has_scribble=True,
                                   h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=False, mono_stereo_manager=self.mono_stereo_manager)
                haux.addWidget(strip)
                self.aux_widgets[a] = strip
                self._wire_strip_controls("aux", a, strip)
                strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                
                strips_added += 1
                a += 1
        
        self.rows_v.addWidget(aux_row)
        
    def _wire_strip_controls(self, section: str, number: int, strip):
        # Wire fader
        fader_key = (section, number, "fader")
        if fader_key in self.mappings:
            mm = self.mappings[fader_key]
            strip.fader.valueChanged.connect(lambda val: self._send_midi_cc(mm.midi_channel, mm.cc_number, val))

        # Wire pan (only for channels)
        if hasattr(strip, 'pan') and strip.pan:
            pan_key = (section, number, "pan")
            if pan_key in self.mappings:
                mm = self.mappings[pan_key]
                strip.pan.valueChanged.connect(lambda val: self._send_midi_cc(mm.midi_channel, mm.cc_number, val))

        # Wire mute
        if hasattr(strip, 'mute') and strip.mute:
            mute_key = (section, number, "mute")
            if mute_key in self.mappings:
                mm = self.mappings[mute_key]
                strip.mute.toggledCC.connect(lambda val: self._send_midi_cc(mm.midi_channel, mm.cc_number, val))

        # Wire stereo partner for wide strips
        if isinstance(strip, WideStereoStrip):
            # Wire partner fader
            partner_fader_key = (section, strip.right_num, "fader")
            if partner_fader_key in self.mappings:
                mm = self.mappings[partner_fader_key]
                strip.fader.valueChanged.connect(lambda val: self._send_midi_cc(mm.midi_channel, mm.cc_number, val))

            # Wire partner pan (only for channels)
            if hasattr(strip, 'pan') and strip.pan:
                partner_pan_key = (section, strip.right_num, "pan")
                if partner_pan_key in self.mappings:
                    mm = self.mappings[partner_pan_key]
                    strip.pan.valueChanged.connect(lambda val: self._send_midi_cc(mm.midi_channel, mm.cc_number, val))

            # Wire partner mute
            if hasattr(strip, 'mute') and strip.mute:
                partner_mute_key = (section, strip.right_num, "mute")
                if partner_mute_key in self.mappings:
                    mm = self.mappings[partner_mute_key]
                    strip.mute.toggledCC.connect(lambda val: self._send_midi_cc(mm.midi_channel, mm.cc_number, val))

    def _send_midi_cc(self, channel_1: int, cc: int, value: int):
        self.midi.send_cc(channel_1 - 1, cc, value)

    def _select_midi_input(self):
        ports = self.midi.list_input_ports()
        if not ports:
            QtWidgets.QMessageBox.information(self, "No Input Ports", "No MIDI input ports are available.")
            return
        
        dialog = MidiPortSelectionDialog("Select MIDI Input Port", ports, self.midi.current_in_port, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            selected_port = dialog.get_selected_port()
            if selected_port:
                self.midi.connect_input(selected_port)
                if self.settings.get('remember_ports', True):
                    self.settings.set('midi_in_port', selected_port)
                    self.settings.save_settings()

    def _select_midi_output(self):
        ports = self.midi.list_output_ports()
        if not ports:
            QtWidgets.QMessageBox.information(self, "No Output Ports", "No MIDI output ports are available.")
            return
        
        dialog = MidiPortSelectionDialog("Select MIDI Output Port", ports, self.midi.current_out_port, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            selected_port = dialog.get_selected_port()
            if selected_port:
                self.midi.connect_output(selected_port)
                if self.settings.get('remember_ports', True):
                    self.settings.set('midi_out_port', selected_port)
                    self.settings.save_settings()

    def _disconnect_midi(self):
        self.midi.connect_input("")
        self.midi.connect_output("")

    def _build_shortcuts(self):
        zoom_in = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl++"), self)
        zoom_out = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+-"), self)
        zoom_reset = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+0"), self)
        zoom_reset_mac = QtWidgets.QShortcut(QtGui.QKeySequence("Meta+0"), self)

        zoom_in.activated.connect(lambda: self.h_zoom_slider.setValue(min(125, self.h_zoom_slider.value()+10)))
        zoom_out.activated.connect(lambda: self.h_zoom_slider.setValue(max(50, self.h_zoom_slider.value()-10)))
        zoom_reset.activated.connect(lambda: (self.h_zoom_slider.setValue(100), self.v_zoom_slider.setValue(100)))
        zoom_reset_mac.activated.connect(lambda: (self.h_zoom_slider.setValue(100), self.v_zoom_slider.setValue(100)))

    def _load_all_scribble_texts(self):
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
                # Handle stereo right channel messages
                channel_key = f"Channel {number}"
                if self.mono_stereo_manager.is_stereo_right(channel_key):
                    left_channel_key = self.mono_stereo_manager.get_stereo_partner(channel_key)
                    if left_channel_key:
                        left_number = int(left_channel_key.split()[1])
                        strip = self.channel_widgets.get((section, left_number))
                        
            if strip:
                if typ == "fader":
                    strip.set_fader_value(value)
                elif typ == "pan":
                    strip.set_pan_value(value)
                elif typ == "mute":
                    strip.set_mute_value(value >= 64)

        elif section == "bus":
            strip = self.bus_widgets.get(number)
            if not strip:
                # Handle stereo right bus messages
                bus_key = f"Bus {number}"
                if self.mono_stereo_manager.is_stereo_right(bus_key):
                    left_bus_key = self.mono_stereo_manager.get_stereo_partner(bus_key)
                    if left_bus_key:
                        left_number = int(left_bus_key.split()[1])
                        strip = self.bus_widgets.get(left_number)
                        
            if strip:
                if typ == "fader":
                    strip.set_fader_value(value)
                elif typ == "mute":
                    strip.set_mute_value(value >= 64)

        elif section == "aux":
            strip = self.aux_widgets.get(number)
            if not strip:
                # Handle stereo right aux messages
                aux_key = f"Aux {number}"
                if self.mono_stereo_manager.is_stereo_right(aux_key):
                    left_aux_key = self.mono_stereo_manager.get_stereo_partner(aux_key)
                    if left_aux_key:
                        left_number = int(left_aux_key.split()[1])
                        strip = self.aux_widgets.get(left_number)
                        
            if strip:
                if typ == "fader":
                    strip.set_fader_value(value)
                elif typ == "mute":
                    strip.set_mute_value(value >= 64)

        elif section == "master" and self.master_widget:
            if typ == "fader":
                self.master_widget.set_fader_value(value)

    def closeEvent(self, event):
        self.settings.save_settings()
        self.midi.connect_input("")
        self.midi.connect_output("")
        event.accept()

def main():
    app = QtWidgets.QApplication(sys.argv)
    
    csv_path = CSV_DEFAULT
    if len(sys.argv) > 1:
        csv_path = sys.argv[1]
    
    if not os.path.exists(csv_path):
        QtWidgets.QMessageBox.critical(
            None, "File Not Found", 
            f"The MIDI mapping file '{csv_path}' was not found.\n"
            f"Please ensure the file exists in the current directory."
        )
        return 1

    window = MixerWindow(csv_path)
    
    # Auto-connect to saved MIDI ports if remember_ports is enabled
    if window.settings.get('remember_ports', True):
        saved_in = window.settings.get('midi_in_port', '')
        saved_out = window.settings.get('midi_out_port', '')
        
        available_in = window.midi.list_input_ports()
        available_out = window.midi.list_output_ports()
        
        if saved_in in available_in:
            window.midi.connect_input(saved_in)
        if saved_out in available_out:
            window.midi.connect_output(saved_out)

    window.show()
    return app.exec_()

if __name__ == "__main__":
    sys.exit(main())                                                                            