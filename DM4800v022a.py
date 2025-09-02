#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DM4800 Mixer Controller v022a
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

BASE_FADER_HEIGHT = 280  # Reduced from 390
BASE_STRIP_WIDTH = 92
BASE_STEREO_FADER_WIDTH_MULTIPLIER = 1.5
BASE_FONT = 11
BASE_SMALL_FONT = 10
BASE_BUTTON_H = 28
BASE_PAN_DIAL = 58

PAN_CENTER_CC = 64
FADER_ZERO_DB_CC = 102
MASTER_ZERO_DB_CC = 127
MASTER_UI_MAX_CC = 140  # Allow fader to go above 0dB in UI for cap visibility
PAN_CENTER_COLOR = "#00ff00"  # Default value - actual color managed by ColorThemeManager

MAX_SCRIBBLE_LENGTH = 24
MAX_SCRIBBLE_CHARS_PER_ROW = 14
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
            'vertical_zoom': 100, 
            'remember_ports': True,
            'stereo_display_mode': 'linked_pair',
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

class ColorThemeManager(QtCore.QObject):
    """Central color management system for the DM4800 application"""
    
    # Signals emitted when colors change
    backgroundColorChanged = QtCore.pyqtSignal(str)  # New background color
    faderZeroDbColorChanged = QtCore.pyqtSignal(str)  # New fader 0dB color
    panCenterColorChanged = QtCore.pyqtSignal(str)    # New pan center color
    panOffCenterColorChanged = QtCore.pyqtSignal(str)  # New pan off-center color
    panLeftColorChanged = QtCore.pyqtSignal(str)      # New pan left color
    panRightColorChanged = QtCore.pyqtSignal(str)     # New pan right color
    panUseSeparateLRColorsChanged = QtCore.pyqtSignal(bool)  # Toggle for L/R colors
    faderGradientEnabledChanged = QtCore.pyqtSignal(bool)  # Gradient on/off
    masterStripBackgroundColorChanged = QtCore.pyqtSignal(str)  # Main fader strip background
    colorsChanged = QtCore.pyqtSignal()  # General signal for any color change
    undoAvailableChanged = QtCore.pyqtSignal(bool)  # Undo availability changed
    
    def __init__(self, settings_manager: 'SettingsManager', parent=None):
        super().__init__(parent)
        self.settings = settings_manager
        
        # Undo system - store previous state for single-level undo
        self._previous_colors = None
        self._undo_available = False
        
        # Default colors - matching current app colors
        self._default_colors = {
            'background_color': '#737374',  # Current rgb(115,115,116)
            'fader_zero_db_color': '#00ff00',  # Current green at 0dB
            'pan_center_color': '#00ff00',     # Current PAN_CENTER_COLOR
            'pan_off_center_color': '#444444',  # Pan knob when not at center (single color mode)
            'pan_left_color': '#ff6600',      # Pan knob when turned left
            'pan_right_color': '#0088ff',     # Pan knob when turned right
            'pan_use_separate_lr_colors': False,  # Toggle for separate L/R colors
            'enable_fader_gradient': True,     # Current behavior has gradients
            'master_strip_background_color': '#737374'  # Main fader strip background
        }
        
        # Initialize settings with defaults if they don't exist
        self._init_color_settings()
    
    def _init_color_settings(self):
        """Initialize color settings with defaults if they don't exist"""
        for key, default_value in self._default_colors.items():
            if self.settings.get(key) is None:
                self.settings.set(key, default_value)
    
    # Background Color Methods
    def get_background_color(self) -> str:
        """Get current background color"""
        return self.settings.get('background_color', self._default_colors['background_color'])
    
    def set_background_color(self, color: str):
        """Set background color and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('background_color', color)
            self.settings.save_settings()
            self.backgroundColorChanged.emit(color)
            self.colorsChanged.emit()

    def reset_background_color(self):
        """Reset background color to default"""
        self.set_background_color(self._default_colors['background_color'])
    
    # Fader 0dB Color Methods  
    def get_fader_zero_db_color(self) -> str:
        """Get current fader 0dB color"""
        return self.settings.get('fader_zero_db_color', self._default_colors['fader_zero_db_color'])
    
    def set_fader_zero_db_color(self, color: str):
        """Set fader 0dB color and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('fader_zero_db_color', color)
            self.settings.save_settings()
            self.faderZeroDbColorChanged.emit(color)
            self.colorsChanged.emit()

    def reset_fader_zero_db_color(self):
        """Reset fader 0dB color to default"""
        self.set_fader_zero_db_color(self._default_colors['fader_zero_db_color'])
    
    # Pan Center Color Methods
    def get_pan_center_color(self) -> str:
        """Get current pan center color"""
        return self.settings.get('pan_center_color', self._default_colors['pan_center_color'])
    
    def set_pan_center_color(self, color: str):
        """Set pan center color and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('pan_center_color', color)
            self.settings.save_settings()
            self.panCenterColorChanged.emit(color)
            self.colorsChanged.emit()

    def reset_pan_center_color(self):
        """Reset pan center color to default"""
        self.set_pan_center_color(self._default_colors['pan_center_color'])
    
    # Pan Off-Center Color Methods
    def get_pan_off_center_color(self) -> str:
        """Get current pan off-center color"""
        return self.settings.get('pan_off_center_color', self._default_colors['pan_off_center_color'])
    
    def set_pan_off_center_color(self, color: str):
        """Set pan off center color and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('pan_off_center_color', color)
            self.settings.save_settings()
            self.panOffCenterColorChanged.emit(color)
            self.colorsChanged.emit()

    def reset_pan_off_center_color(self):
        """Reset pan off-center color to default"""
        self.set_pan_off_center_color(self._default_colors['pan_off_center_color'])
    
    # Pan Left Color Methods
    def get_pan_left_color(self) -> str:
        """Get current pan left color"""
        return self.settings.get('pan_left_color', self._default_colors['pan_left_color'])
    
    def set_pan_left_color(self, color: str):
        """Set pan left color and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('pan_left_color', color)
            self.settings.save_settings()
            self.panLeftColorChanged.emit(color)
            self.colorsChanged.emit()

    def reset_pan_left_color(self):
        """Reset pan left color to default"""
        self.set_pan_left_color(self._default_colors['pan_left_color'])
    
    # Pan Right Color Methods
    def get_pan_right_color(self) -> str:
        """Get current pan right color"""
        return self.settings.get('pan_right_color', self._default_colors['pan_right_color'])

    def set_pan_right_color(self, color: str):
        """Set pan right color and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('pan_right_color', color)
            self.settings.save_settings()
            self.panRightColorChanged.emit(color)
            self.colorsChanged.emit()

    def reset_pan_right_color(self):
        """Reset pan right color to default"""
        self.set_pan_right_color(self._default_colors['pan_right_color'])
    
    # Pan Use Separate L/R Colors Methods
    def get_pan_use_separate_lr_colors(self) -> bool:
        """Get current pan use separate L/R colors setting"""
        return self.settings.get('pan_use_separate_lr_colors', self._default_colors['pan_use_separate_lr_colors'])
    
    def set_pan_use_separate_lr_colors(self, color: str):
        """Set pan use separate lr colors and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('pan_use_separate_lr_colors', color)
            self.settings.save_settings()
            self.panUseSeparateLRColorsChanged.emit(color)
            self.colorsChanged.emit()

    def reset_pan_use_separate_lr_colors(self):
        """Reset pan use separate L/R colors to default"""
        self.set_pan_use_separate_lr_colors(self._default_colors['pan_use_separate_lr_colors'])
    
    # Fader Gradient Methods
    def get_fader_gradient_enabled(self) -> bool:
        """Get current fader gradient setting"""
        return self.settings.get('enable_fader_gradient', self._default_colors['enable_fader_gradient'])
    
    def set_fader_gradient_enabled(self, enabled: bool):
        """Set fader gradient enabled and emit change signal"""
        self.settings.set('enable_fader_gradient', enabled)
        self.settings.save_settings()
        self.faderGradientEnabledChanged.emit(enabled)
        self.colorsChanged.emit()

    def set_fader_gradient_enabled(self, color: str):
        """Set fader gradient enabled and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('enable_fader_gradient', color)
            self.settings.save_settings()
            self.faderGradientEnabledChanged.emit(color)
            self.colorsChanged.emit()

    def reset_fader_gradient_enabled(self):
        """Reset fader gradient setting to default"""
        self.set_fader_gradient_enabled(self._default_colors['enable_fader_gradient'])
    
    # Master Strip Background Color Methods
    def get_master_strip_background_color(self) -> str:
        """Get current master strip background color"""
        return self.settings.get('master_strip_background_color', self._default_colors['master_strip_background_color'])
    
    def set_master_strip_background_color(self, color: str):
        """Set master strip background color and emit change signal"""
        if self._validate_color(color):
            # Save current state for undo before making change
            self._save_current_state()
            
            self.settings.set('master_strip_background_color', color)
            self.settings.save_settings()
            self.masterStripBackgroundColorChanged.emit(color)
            self.colorsChanged.emit()

    def reset_master_strip_background_color(self):
        """Reset master strip background color to default"""
        self.set_master_strip_background_color(self._default_colors['master_strip_background_color'])
    
    # Utility Methods
    def _validate_color(self, color: str) -> bool:
        """Validate color string format"""
        if not color:
            return False
        
        # Remove whitespace
        color = color.strip()
        
        # Check for hex format (#RRGGBB or #RGB)
        if color.startswith('#'):
            hex_part = color[1:]
            if len(hex_part) == 6 or len(hex_part) == 3:
                try:
                    int(hex_part, 16)
                    return True
                except ValueError:
                    return False
        
        # Check for rgb() format
        if color.startswith('rgb(') and color.endswith(')'):
            try:
                rgb_values = color[4:-1].split(',')
                if len(rgb_values) == 3:
                    for val in rgb_values:
                        num = int(val.strip())
                        if not (0 <= num <= 255):
                            return False
                    return True
            except (ValueError, IndexError):
                return False
        
        return False
    
    def hex_to_rgb(self, hex_color: str) -> tuple:
        """Convert hex color to RGB tuple"""
        if not hex_color.startswith('#'):
            return (0, 0, 0)
        
        hex_color = hex_color[1:]
        if len(hex_color) == 3:
            hex_color = ''.join([c*2 for c in hex_color])
        
        try:
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        except (ValueError, IndexError):
            return (0, 0, 0)
    
    def rgb_to_hex(self, r: int, g: int, b: int) -> str:
        """Convert RGB values to hex color string"""
        return f"#{r:02x}{g:02x}{b:02x}"
    
    def reset_all_colors(self):
        """Reset all colors to defaults"""
        # Save current state before resetting everything
        self._save_current_state()
        
        self.reset_background_color()
        self.reset_fader_zero_db_color()
        self.reset_pan_center_color()
        self.reset_pan_off_center_color()
        self.reset_pan_left_color()
        self.reset_pan_right_color()
        self.reset_pan_use_separate_lr_colors()
        self.reset_fader_gradient_enabled()
        self.reset_master_strip_background_color()

    # Undo System Methods
    def _save_current_state(self):
        """Save current color state for undo functionality"""
        self._previous_colors = {
            'background_color': self.get_background_color(),
            'fader_zero_db_color': self.get_fader_zero_db_color(),
            'pan_center_color': self.get_pan_center_color(),
            'pan_off_center_color': self.get_pan_off_center_color(),
            'pan_left_color': self.get_pan_left_color(),
            'pan_right_color': self.get_pan_right_color(),
            'pan_use_separate_lr_colors': self.get_pan_use_separate_lr_colors(),
            'enable_fader_gradient': self.get_fader_gradient_enabled(),
            'master_strip_background_color': self.get_master_strip_background_color()
        }
        self._undo_available = True
        self.undoAvailableChanged.emit(True)
    
    def can_undo(self) -> bool:
        """Check if undo is available"""
        return self._undo_available and self._previous_colors is not None
    
    def undo_last_change(self):
        """Undo the last color change"""
        if not self.can_undo():
            return False
            
        # Temporarily disable undo tracking while restoring
        previous_colors = self._previous_colors
        self._undo_available = False
        self.undoAvailableChanged.emit(False)
        
        # Restore all previous colors without creating new undo states
        self._restore_colors_without_undo(previous_colors)
        
        # Clear undo state
        self._previous_colors = None
        return True
    
    def _restore_colors_without_undo(self, colors: dict):
        """Restore colors without triggering undo system"""
        # Set colors directly in settings without going through the normal setters
        # This prevents creating a new undo state while restoring
        for key, value in colors.items():
            self.settings.set(key, value)
        
        self.settings.save_settings()
        
        # Emit all the change signals to update the UI
        self.backgroundColorChanged.emit(colors['background_color'])
        self.faderZeroDbColorChanged.emit(colors['fader_zero_db_color'])
        self.panCenterColorChanged.emit(colors['pan_center_color'])
        self.panOffCenterColorChanged.emit(colors['pan_off_center_color'])
        self.panLeftColorChanged.emit(colors['pan_left_color'])
        self.panRightColorChanged.emit(colors['pan_right_color'])
        self.panUseSeparateLRColorsChanged.emit(colors['pan_use_separate_lr_colors'])
        self.faderGradientEnabledChanged.emit(colors['enable_fader_gradient'])
        self.masterStripBackgroundColorChanged.emit(colors['master_strip_background_color'])
        self.colorsChanged.emit()

    def get_all_colors(self) -> dict:
        """Get all current color settings as a dictionary"""
        return {
            'background_color': self.get_background_color(),
            'fader_zero_db_color': self.get_fader_zero_db_color(),
            'pan_center_color': self.get_pan_center_color(),
            'enable_fader_gradient': self.get_fader_gradient_enabled()
        }
    
    def get_all_colors(self) -> dict:
        """Get all current color settings as a dictionary"""
        return {
            'background_color': self.get_background_color(),
            'fader_zero_db_color': self.get_fader_zero_db_color(),
            'pan_center_color': self.get_pan_center_color(),
            'pan_off_center_color': self.get_pan_off_center_color(),
            'pan_left_color': self.get_pan_left_color(),
            'pan_right_color': self.get_pan_right_color(),
            'enable_fader_gradient': self.get_fader_gradient_enabled(),
            'pan_use_separate_lr_colors': self.get_pan_use_separate_lr_colors()
        }

    def get_default_colors(self) -> dict:
        """Get default color settings"""
        return self._default_colors.copy()

class ColorPickerWidget(QtWidgets.QWidget):
    """Reusable color picker widget with preview and reset functionality"""
    
    colorChanged = QtCore.pyqtSignal(str)  # Emitted when user selects a new color
    
    def __init__(self, label_text: str, color_manager: ColorThemeManager, 
                 color_key: str, parent=None):
        super().__init__(parent)
        self.color_manager = color_manager
        self.color_key = color_key  # e.g. 'background_color', 'fader_zero_db_color'
        self.current_color = self._get_current_color()
        
        self._setup_ui(label_text)
        self._connect_signals()
    
    def _setup_ui(self, label_text: str):
        """Setup the widget UI"""
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # Label
        self.label = QtWidgets.QLabel(label_text)
        self.label.setMinimumWidth(120)
        layout.addWidget(self.label)
        
        # Color preview button
        self.color_button = QtWidgets.QPushButton()
        self.color_button.setMinimumSize(40, 25)
        self.color_button.setMaximumSize(60, 25)
        self.color_button.clicked.connect(self._open_color_picker)
        self.color_button.setToolTip("Click to choose color")
        self._update_color_button()
        layout.addWidget(self.color_button)
        
        # Color value label
        self.color_label = QtWidgets.QLabel(self.current_color)
        self.color_label.setMinimumWidth(70)
        self.color_label.setStyleSheet("font-family: monospace; color: #666;")
        layout.addWidget(self.color_label)
        
        # Reset button
        self.reset_button = QtWidgets.QPushButton("Reset")
        self.reset_button.setMaximumWidth(60)
        self.reset_button.clicked.connect(self._reset_to_default)
        self.reset_button.setToolTip("Reset to default color")
        layout.addWidget(self.reset_button)
        
        layout.addStretch()  # Push everything to the left
    
    def _connect_signals(self):
        """Connect to color manager signals to update when colors change externally"""
        if self.color_key == 'background_color':
            self.color_manager.backgroundColorChanged.connect(self._on_external_color_change)
        elif self.color_key == 'fader_zero_db_color':
            self.color_manager.faderZeroDbColorChanged.connect(self._on_external_color_change)
        elif self.color_key == 'pan_center_color':
            self.color_manager.panCenterColorChanged.connect(self._on_external_color_change)
        elif self.color_key == 'pan_off_center_color':
            self.color_manager.panOffCenterColorChanged.connect(self._on_external_color_change)
        elif self.color_key == 'pan_left_color':
            self.color_manager.panLeftColorChanged.connect(self._on_external_color_change)
        elif self.color_key == 'pan_right_color':
            self.color_manager.panRightColorChanged.connect(self._on_external_color_change)
        elif self.color_key == 'master_strip_background_color':
            self.color_manager.masterStripBackgroundColorChanged.connect(self._on_external_color_change)

    def _get_current_color(self) -> str:
        """Get current color from color manager"""
        if self.color_key == 'background_color':
            return self.color_manager.get_background_color()
        elif self.color_key == 'fader_zero_db_color':
            return self.color_manager.get_fader_zero_db_color()
        elif self.color_key == 'pan_center_color':
            return self.color_manager.get_pan_center_color()
        elif self.color_key == 'master_strip_background_color':
            return self.color_manager.get_master_strip_background_color()
        else:
            return '#000000'  # Fallback

    def _get_current_color(self) -> str:
        """Get current color from color manager"""
        if self.color_key == 'background_color':
            return self.color_manager.get_background_color()
        elif self.color_key == 'fader_zero_db_color':
            return self.color_manager.get_fader_zero_db_color()
        elif self.color_key == 'pan_center_color':
            return self.color_manager.get_pan_center_color()
        elif self.color_key == 'pan_off_center_color':
            return self.color_manager.get_pan_off_center_color()
        elif self.color_key == 'pan_left_color':
            return self.color_manager.get_pan_left_color()
        elif self.color_key == 'pan_right_color':
            return self.color_manager.get_pan_right_color()
        elif self.color_key == 'master_strip_background_color':
            return self.color_manager.get_master_strip_background_color()
        else:
            return '#000000'  # Fallback

    def _get_default_color(self) -> str:
        """Get default color for this key"""
        defaults = self.color_manager.get_default_colors()
        return defaults.get(self.color_key, '#000000')
    
    def _update_color_button(self):
        """Update the color preview button appearance"""
        # Create a style with the current color as background
        self.color_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.current_color};
                border: 2px solid #666;
                border-radius: 3px;
            }}
            QPushButton:hover {{
                border: 2px solid #999;
            }}
            QPushButton:pressed {{
                border: 2px solid #333;
            }}
        """)
    
    def _open_color_picker(self):
        """Open the native color picker dialog"""
        # Convert current color to QColor
        initial_color = QtGui.QColor(self.current_color)
        
        # Open color dialog
        color = QtWidgets.QColorDialog.getColor(
            initial_color, 
            self, 
            f"Choose {self.label.text()}",
            QtWidgets.QColorDialog.ShowAlphaChannel
        )
        
        # If user selected a color (didn't cancel)
        if color.isValid():
            new_color = color.name()  # Returns #RRGGBB format
            self._set_color(new_color)
    
    def _set_color(self, color: str):
        """Set the color (internal method)"""
        if color != self.current_color:
            self.current_color = color
            self._update_color_button()
            self.color_label.setText(color)
            
            # Update the color manager
            if self.color_key == 'background_color':
                self.color_manager.set_background_color(color)
            elif self.color_key == 'fader_zero_db_color':
                self.color_manager.set_fader_zero_db_color(color)
            elif self.color_key == 'pan_center_color':
                self.color_manager.set_pan_center_color(color)
            elif self.color_key == 'pan_off_center_color':
                self.color_manager.set_pan_off_center_color(color)
            elif self.color_key == 'pan_left_color':
                self.color_manager.set_pan_left_color(color)
            elif self.color_key == 'pan_right_color':
                self.color_manager.set_pan_right_color(color)
            elif self.color_key == 'master_strip_background_color':
                self.color_manager.set_master_strip_background_color(color)
            
            # Emit our signal
            self.colorChanged.emit(color)
    
    def _reset_to_default(self):
        """Reset color to default value"""
        default_color = self._get_default_color()
        self._set_color(default_color)
    
    def _on_external_color_change(self, new_color: str):
        """Handle color changes from external sources (like test menu)"""
        if new_color != self.current_color:
            self.current_color = new_color
            self._update_color_button()
            self.color_label.setText(new_color)
            # Don't emit colorChanged signal here to avoid loops
    
    def set_color(self, color: str):
        """Public method to set color externally"""
        self._set_color(color)
    
    def get_color(self) -> str:
        """Get current color"""
        return self.current_color

class SettingsDialog(QtWidgets.QDialog):
    """Main settings dialog with tabbed interface"""
    
    def __init__(self, color_manager: ColorThemeManager, parent=None):
        super().__init__(parent)
        self.color_manager = color_manager
        self.setWindowTitle("DM4800 Settings")
        self.setModal(True)
        self.resize(500, 400)
        
        self._setup_ui()
        
        # Add keyboard shortcut for undo
        #self.undo_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence.Undo, self)  # Ctrl+Z
        #self.undo_shortcut.activated.connect(self._undo_last_change)
        #self.undo_shortcut.setEnabled(self.color_manager.can_undo())
        #self.color_manager.undoAvailableChanged.connect(self.undo_shortcut.setEnabled)
    
    def _setup_ui(self):
        """Setup the dialog UI with tabs"""
        layout = QtWidgets.QVBoxLayout(self)
        
        # Create tab widget
        self.tab_widget = QtWidgets.QTabWidget()
        layout.addWidget(self.tab_widget)
        
        # Colors tab
        self._create_colors_tab()
        
        # General tab (for future settings)
        self._create_general_tab()
        
        # Button box
        button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Apply
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        button_box.button(QtWidgets.QDialogButtonBox.Apply).clicked.connect(self._apply_settings)
        layout.addWidget(button_box)
    
    def _create_colors_tab(self):
        """Create the colors configuration tab"""
        colors_widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(colors_widget)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)
        
        # Title
        title = QtWidgets.QLabel("Color Customization")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #333;")
        layout.addWidget(title)
        
        # Background section
        bg_group = QtWidgets.QGroupBox("Background")
        bg_layout = QtWidgets.QVBoxLayout(bg_group)
        
        self.bg_picker = ColorPickerWidget("Application Background:", self.color_manager, 'background_color')
        bg_layout.addWidget(self.bg_picker)
        layout.addWidget(bg_group)
        # Master strip background section
        master_strip_group = QtWidgets.QGroupBox("Main Fader Strip")
        master_strip_layout = QtWidgets.QVBoxLayout(master_strip_group)

        self.master_strip_bg_picker = ColorPickerWidget("Main Fader Strip Background:", self.color_manager, 'master_strip_background_color')
        master_strip_layout.addWidget(self.master_strip_bg_picker)
        layout.addWidget(master_strip_group)
        
        # Fader section
        fader_group = QtWidgets.QGroupBox("Faders")
        fader_layout = QtWidgets.QVBoxLayout(fader_group)
        
        self.fader_picker = ColorPickerWidget("Fader Color at 0dB:", self.color_manager, 'fader_zero_db_color')
        fader_layout.addWidget(self.fader_picker)
        
        # Gradient toggle
        gradient_layout = QtWidgets.QHBoxLayout()
        gradient_label = QtWidgets.QLabel("Enable gradient colors above/below 0dB:")
        gradient_label.setMinimumWidth(250)
        
        self.gradient_checkbox = QtWidgets.QCheckBox()
        self.gradient_checkbox.setChecked(self.color_manager.get_fader_gradient_enabled())
        self.gradient_checkbox.toggled.connect(self.color_manager.set_fader_gradient_enabled)
        
        gradient_layout.addWidget(gradient_label)
        gradient_layout.addWidget(self.gradient_checkbox)
        gradient_layout.addStretch()
        
        fader_layout.addLayout(gradient_layout)
        layout.addWidget(fader_group)
        
        # Pan section
        pan_group = QtWidgets.QGroupBox("Pan/Balance Controls")
        pan_layout = QtWidgets.QVBoxLayout(pan_group)
        
        # Pan center color (always visible)
        self.pan_center_picker = ColorPickerWidget("Pan/Bal Color at Center:", self.color_manager, 'pan_center_color')
        pan_layout.addWidget(self.pan_center_picker)
        
        # Separator
        separator_line = QtWidgets.QFrame()
        separator_line.setFrameShape(QtWidgets.QFrame.HLine)
        separator_line.setFrameStyle(QtWidgets.QFrame.Sunken)
        separator_line.setStyleSheet("color: #ccc;")
        pan_layout.addWidget(separator_line)
        
        # Toggle for separate L/R colors
        toggle_layout = QtWidgets.QHBoxLayout()
        toggle_label = QtWidgets.QLabel("Use different colors for Left and Right:")
        toggle_label.setMinimumWidth(250)
        
        self.pan_separate_lr_checkbox = QtWidgets.QCheckBox()
        self.pan_separate_lr_checkbox.setChecked(self.color_manager.get_pan_use_separate_lr_colors())
        self.pan_separate_lr_checkbox.toggled.connect(self.color_manager.set_pan_use_separate_lr_colors)
        self.pan_separate_lr_checkbox.toggled.connect(self._on_pan_toggle_changed)
        
        toggle_layout.addWidget(toggle_label)
        toggle_layout.addWidget(self.pan_separate_lr_checkbox)
        toggle_layout.addStretch()
        
        pan_layout.addLayout(toggle_layout)
        
        # Single off-center color (visible when toggle is off)
        self.pan_single_picker = ColorPickerWidget("Pan/Bal Color when not at Center:", self.color_manager, 'pan_off_center_color')
        pan_layout.addWidget(self.pan_single_picker)
        
        # Left and Right colors (visible when toggle is on)
        self.pan_left_picker = ColorPickerWidget("Pan/Bal Color when turned Left:", self.color_manager, 'pan_left_color')
        pan_layout.addWidget(self.pan_left_picker)
        
        self.pan_right_picker = ColorPickerWidget("Pan/Bal Color when turned Right:", self.color_manager, 'pan_right_color')
        pan_layout.addWidget(self.pan_right_picker)
        
        # Set initial visibility based on toggle state
        self._on_pan_toggle_changed(self.pan_separate_lr_checkbox.isChecked())
        
        layout.addWidget(pan_group)
        
        # Reset and Undo section
        reset_group = QtWidgets.QGroupBox("Reset & Undo")
        reset_layout = QtWidgets.QVBoxLayout(reset_group)
        
        # Undo button
        self.undo_button = QtWidgets.QPushButton("Undo Last Color Change")
        self.undo_button.clicked.connect(self._undo_last_change)
        self.undo_button.setStyleSheet("QPushButton { padding: 8px; background-color: #e8f4f8; }")
        self.undo_button.setEnabled(self.color_manager.can_undo())
        reset_layout.addWidget(self.undo_button)
        
        # Separator between undo and reset
        separator_line = QtWidgets.QFrame()
        separator_line.setFrameShape(QtWidgets.QFrame.HLine)
        separator_line.setFrameStyle(QtWidgets.QFrame.Sunken)
        separator_line.setStyleSheet("color: #ccc;")
        reset_layout.addWidget(separator_line)
        
        # Reset all button
        reset_button = QtWidgets.QPushButton("Reset All Colors to Defaults")
        reset_button.clicked.connect(self._reset_all_colors)
        reset_button.setStyleSheet("QPushButton { padding: 8px; background-color: #f8e8e8; }")
        reset_layout.addWidget(reset_button)
        
        layout.addWidget(reset_group)

        layout.addStretch() 
        
        # Connect signals to keep checkboxes in sync
        self.color_manager.faderGradientEnabledChanged.connect(self.gradient_checkbox.setChecked)
        self.color_manager.panUseSeparateLRColorsChanged.connect(self.pan_separate_lr_checkbox.setChecked)
        self.color_manager.panUseSeparateLRColorsChanged.connect(self._on_pan_toggle_changed)
        self.color_manager.undoAvailableChanged.connect(self.undo_button.setEnabled)
        
        self.tab_widget.addTab(colors_widget, "Colors")
    
    def _create_general_tab(self):
        """Create general settings tab (placeholder for future settings)"""
        general_widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(general_widget)
        layout.setContentsMargins(15, 15, 15, 15)
        
        title = QtWidgets.QLabel("General Settings")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #333;")
        layout.addWidget(title)
        
        # Placeholder for future settings
        placeholder = QtWidgets.QLabel("Additional settings will be added here in future versions.")
        placeholder.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(placeholder)
        
        layout.addStretch()
        
        self.tab_widget.addTab(general_widget, "General")
    
    def _reset_all_colors(self):
        """Reset all colors to defaults with confirmation"""
        reply = QtWidgets.QMessageBox.question(
            self, 
            "Reset Colors", 
            "Are you sure you want to reset all colors to their default values?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No
        )
        
        if reply == QtWidgets.QMessageBox.Yes:
            self.color_manager.reset_all_colors()
    
    def _on_pan_toggle_changed(self, checked: bool):
        """Handle changes to the pan L/R colors toggle"""
        # Show/hide appropriate color pickers based on toggle state
        self.pan_single_picker.setVisible(not checked)
        self.pan_left_picker.setVisible(checked)
        self.pan_right_picker.setVisible(checked)
        
        # Refresh all pan colors to apply the new setting immediately
        if hasattr(self.parent(), '_refresh_all_pan_colors'):
            self.parent()._refresh_all_pan_colors()
    
    def _undo_last_change(self):
        """Undo the last color change with confirmation"""
        reply = QtWidgets.QMessageBox.question(
            self, 
            "Undo Color Change", 
            "Are you sure you want to undo the last color change?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.Yes  # Default to Yes for undo
        )
        
        if reply == QtWidgets.QMessageBox.Yes:
            success = self.color_manager.undo_last_change()
            if success:
                QtWidgets.QMessageBox.information(
                    self, 
                    "Undo Complete", 
                    "The last color change has been undone."
                )
            else:
                QtWidgets.QMessageBox.warning(
                    self, 
                    "Undo Failed", 
                    "No changes available to undo."
                )
    
    def _apply_settings(self):
        """Apply settings without closing dialog"""
        # Settings are applied immediately via signals, so just show confirmation
        QtWidgets.QMessageBox.information(self, "Settings Applied", "Color settings have been applied.")
    
    def show_colors_tab(self):
        """Show dialog with colors tab selected"""
        self.tab_widget.setCurrentIndex(0)  # Colors tab is index 0
        self.show()
        self.raise_()
        self.activateWindow()

class ContextColorMenu(QtWidgets.QMenu):
    """Context-aware color menu that appears on right-click"""
    
    def __init__(self, context_type: str, color_manager: ColorThemeManager, parent=None):
        super().__init__(parent)
        self.context_type = context_type  # "fader", "pan", or "background"
        self.color_manager = color_manager
        
        self._setup_menu()
    
    def _setup_menu(self):
        """Setup menu based on context type"""
        if self.context_type == "fader":
            self.setTitle("Fader Colors")
            self._add_fader_actions()
        elif self.context_type == "pan":
            self.setTitle("Pan/Balance Colors")
            self._add_pan_actions()
        elif self.context_type == "background":
            self.setTitle("Background Color")
            self._add_background_actions()
        
        self.addSeparator()
        
        # Always add "Open Settings..." option
        settings_action = QtWidgets.QAction("Open Settings...", self)
        settings_action.triggered.connect(self._open_settings)
        self.addAction(settings_action)
    
    def _add_fader_actions(self):
        """Add fader-specific color actions"""
        # Quick color change for 0dB
        current_color = self.color_manager.get_fader_zero_db_color()
        
        color_action = QtWidgets.QAction(f"Change 0dB Color ({current_color})", self)
        color_action.triggered.connect(self._change_fader_color)
        self.addAction(color_action)
        
        # Gradient toggle
        gradient_enabled = self.color_manager.get_fader_gradient_enabled()
        gradient_text = "Disable Gradient Colors" if gradient_enabled else "Enable Gradient Colors"
        
        gradient_action = QtWidgets.QAction(gradient_text, self)
        gradient_action.triggered.connect(lambda: self.color_manager.set_fader_gradient_enabled(not gradient_enabled))
        self.addAction(gradient_action)
        
        self.addSeparator()
        
        # Quick color presets
        presets_menu = self.addMenu("Quick Colors")
        
        colors = [
            ("Green (Default)", "#00ff00"),
            ("Blue", "#0088ff"),
            ("Orange", "#ff6600"),
            ("Red", "#ff0000"),
            ("Purple", "#8800ff"),
            ("Yellow", "#ffdd00")
        ]
        
        for name, color in colors:
            action = QtWidgets.QAction(name, self)
            action.triggered.connect(lambda checked, c=color: self.color_manager.set_fader_zero_db_color(c))
            presets_menu.addAction(action)
    
    def _add_pan_actions(self):
        """Add pan-specific color actions"""
        current_color = self.color_manager.get_pan_center_color()
        
        color_action = QtWidgets.QAction(f"Change Center Color ({current_color})", self)
        color_action.triggered.connect(self._change_pan_color)
        self.addAction(color_action)
        
        self.addSeparator()
        
        # Quick color presets
        presets_menu = self.addMenu("Quick Colors")
        
        colors = [
            ("Green (Default)", "#00ff00"),
            ("Blue", "#0088ff"),
            ("Orange", "#ff6600"),
            ("Red", "#ff0000"),
            ("Purple", "#8800ff"),
            ("Cyan", "#00ffff")
        ]
        
        for name, color in colors:
            action = QtWidgets.QAction(name, self)
            action.triggered.connect(lambda checked, c=color: self.color_manager.set_pan_center_color(c))
            presets_menu.addAction(action)
    
    def _add_background_actions(self):
        """Add background-specific color actions"""
        current_color = self.color_manager.get_background_color()
        
        color_action = QtWidgets.QAction(f"Change Background ({current_color})", self)
        color_action.triggered.connect(self._change_background_color)
        self.addAction(color_action)
        
        self.addSeparator()
        
        # Theme presets
        themes_menu = self.addMenu("Quick Themes")
        
        themes = [
            ("Default Gray", "#737374"),
            ("Dark Mode", "#2a2a3e"),
            ("Light Gray", "#e0e0e0"),
            ("Blue Gray", "#4a5568"),
            ("Dark Green", "#2d3748"),
            ("Warm Gray", "#8d7053")
        ]
        
        for name, color in themes:
            action = QtWidgets.QAction(name, self)
            action.triggered.connect(lambda checked, c=color: self.color_manager.set_background_color(c))
            themes_menu.addAction(action)
    
    def _change_fader_color(self):
        """Open color picker for fader color"""
        current_color = QtGui.QColor(self.color_manager.get_fader_zero_db_color())
        color = QtWidgets.QColorDialog.getColor(current_color, self.parent(), "Choose Fader 0dB Color")
        
        if color.isValid():
            self.color_manager.set_fader_zero_db_color(color.name())
    
    def _change_pan_color(self):
        """Open color picker for pan color"""
        current_color = QtGui.QColor(self.color_manager.get_pan_center_color())
        color = QtWidgets.QColorDialog.getColor(current_color, self.parent(), "Choose Pan Center Color")
        
        if color.isValid():
            self.color_manager.set_pan_center_color(color.name())
    
    def _change_background_color(self):
        """Open color picker for background color"""
        current_color = QtGui.QColor(self.color_manager.get_background_color())
        color = QtWidgets.QColorDialog.getColor(current_color, self.parent(), "Choose Background Color")
        
        if color.isValid():
            self.color_manager.set_background_color(color.name())
    
    def _open_settings(self):
        """Open the main settings dialog"""
        # Find the main window and open settings
        main_window = self.parent()
        while main_window and not isinstance(main_window, MixerWindow):
            main_window = main_window.parent()
        
        if main_window and hasattr(main_window, '_show_settings_dialog'):
            main_window._show_settings_dialog()

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
        self.text_area.setFont(QtGui.QFont("Courier New", 9))
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
            max_cc = MASTER_UI_MAX_CC if self.is_master else 127
            t = cc / max_cc
            return int(rect.top() + (1.0 - t) * rect.height())

        painter.setPen(QtGui.QPen(QtCore.Qt.black, 1))
        font = painter.font()
        font.setPointSizeF(max(7.0, BASE_SMALL_FONT * self.font_scale))
        painter.setFont(font)

        scale_points = MASTER_DB_POINTS if self.is_master else CHANNEL_DB_POINTS

        for cc, label in scale_points:
            y = y_for_cc(cc)
    
            # No offset needed - all strips should have consistent scale positioning
            offset = 0
    
            y_adjusted = y + offset
            painter.drawLine(rect.right()-6, y_adjusted, rect.right()-2, y_adjusted)
            text_rect = QtCore.QRect(rect.left(), y_adjusted-8, rect.width()-8, 16)
    
            y_adjusted = y + offset
            painter.drawLine(rect.right()-6, y_adjusted, rect.right()-2, y_adjusted)
            text_rect = QtCore.QRect(rect.left(), y_adjusted-8, rect.width()-8, 16)
    
            # Remove "dB" and other text, show just the numbers
            if "+" in label:
                short_label = label.replace(" dB", "").replace("+", "+")  # Keep the + sign
            elif "-∞" in label or "-âˆž" in label:
                short_label = "-∞"
            else:
                short_label = label.replace(" dB", "")  # Remove " dB" from labels
    
            painter.drawText(text_rect, QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter, short_label)
            
class CustomFaderSlider(QtWidgets.QSlider):
    doubleClicked = QtCore.pyqtSignal()

    def __init__(self, is_master=False, is_stereo=False, parent=None):
        super().__init__(QtCore.Qt.Vertical, parent)
        if is_master:
            self.setRange(0, MASTER_UI_MAX_CC)  # Extended range for master fader cap visibility
        else:
            self.setRange(0, 127)
        self.setInvertedAppearance(False)
        self.setTickPosition(QtWidgets.QSlider.NoTicks)
        self.is_master = is_master
        self.is_stereo = is_stereo
        self.current_value = MASTER_ZERO_DB_CC if is_master else FADER_ZERO_DB_CC
        self.setStyleSheet(self._get_fader_style())

    def mousePressEvent(self, event):
        """Override to constrain master fader movement"""
        if self.is_master and event.button() == QtCore.Qt.LeftButton:
            # Calculate the position where 0dB (127) would be
            opt = QtWidgets.QStyleOptionSlider()
            self.initStyleOption(opt)
            groove = self.style().subControlRect(QtWidgets.QStyle.CC_Slider, opt, QtWidgets.QStyle.SC_SliderGroove, self)
        
            # Calculate the Y position for CC 127 (0dB)
            zero_db_y = groove.bottom() - ((127.0 / MASTER_UI_MAX_CC) * groove.height())
        
            # If click is above 0dB line, don't process the click
            if event.pos().y() < zero_db_y:
                return  # Block the click
    
        # For right-click color menu and normal operation
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        """Override to constrain master fader dragging"""
        if self.is_master:
            # Let the parent handle the move first
            super().mouseMoveEvent(event)
            # Then constrain the value if it exceeds 127
            if self.value() > MASTER_ZERO_DB_CC:
                self.setValue(MASTER_ZERO_DB_CC)
        else:
            super().mouseMoveEvent(event)

    def _interpolate_color(self, start_rgb: tuple, end_rgb: tuple, factor: float) -> str:
        factor = max(0.0, min(1.0, factor))
        r = int(start_rgb[0] + (end_rgb[0] - start_rgb[0]) * factor)
        g = int(start_rgb[1] + (end_rgb[1] - start_rgb[1]) * factor)
        b = int(start_rgb[2] + (end_rgb[2] - start_rgb[2]) * factor)
        return f"rgb({r}, {g}, {b})"

    def _get_fader_color(self, value: int) -> str:
        zero_db_value = MASTER_ZERO_DB_CC if self.is_master else FADER_ZERO_DB_CC
        max_value = MASTER_UI_MAX_CC if self.is_master else 127
    
        # Get color manager from the main window
        main_window = None
        widget = self.parent()
        while widget and not hasattr(widget, 'color_manager'):
            widget = widget.parent()
    
        if widget and hasattr(widget, 'color_manager'):
            zero_db_color = widget.color_manager.get_fader_zero_db_color()
            gradient_enabled = widget.color_manager.get_fader_gradient_enabled()
        else:
            zero_db_color = "#00ff00"  # Fallback
            gradient_enabled = True
    
        # Parse the zero_db_color
        if zero_db_color.startswith('#'):
            hex_color = zero_db_color[1:]
            if len(hex_color) == 6:
                zero_db_rgb = (int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
            else:
                zero_db_rgb = (0, 255, 0)  # Fallback green
        else:
            zero_db_rgb = (0, 255, 0)  # Fallback green
    
        if not gradient_enabled:
            return f"rgb({zero_db_rgb[0]}, {zero_db_rgb[1]}, {zero_db_rgb[2]})"
    
        # Calculate color based on position relative to 0dB
        if value < zero_db_value:
            # Below 0dB: interpolate from dark to zero_db_color
            if zero_db_value > 0:
                factor = value / zero_db_value
            else:
                factor = 0
            dark_rgb = (int(zero_db_rgb[0] * 0.3), int(zero_db_rgb[1] * 0.3), int(zero_db_rgb[2] * 0.3))
            return self._interpolate_color(dark_rgb, zero_db_rgb, factor)
        else:
            # At or above 0dB: interpolate from zero_db_color to red
            if max_value > zero_db_value:
                factor = (value - zero_db_value) / (max_value - zero_db_value)
            else:
                factor = 0
            red_rgb = (255, 0, 0)
            return self._interpolate_color(zero_db_rgb, red_rgb, factor)

    def refresh_colors(self):
        """Refresh fader colors from color manager"""
        self.update_color(self.current_value)

    def _get_fader_style(self):
        color = self._get_fader_color(self.current_value)
    
        # Use different margin for stereo faders
        # Use different margin based on context:
        # -10px for wide stereo strips (when checkbox is checked)
        # -5px for regular strips and individual stereo strips (when checkbox is unchecked)
        margin_top = "-15px"  # Default for all strips
    
        return f"""
            QSlider:groove:vertical {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #333333, stop:0.5 #555555, stop:1 #333333);
                width: 8px; border: 1px solid #222222; border-radius: 2px;
            }}
            QSlider:handle:vertical {{
                background: {color}; border: 1px solid #444444; height: 40px; width: 24px; margin: {margin_top} -8px; border-radius: 3px;
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

    def mousePressEvent(self, event):
        """Handle mouse press events including right-click"""
        if event.button() == QtCore.Qt.RightButton:
            # Find the color manager
            color_manager = None
            widget = self.parent()
            while widget and not hasattr(widget, 'color_manager'):
                widget = widget.parent()
            
            if color_manager := (widget.color_manager if widget and hasattr(widget, 'color_manager') else None):
                menu = ContextColorMenu("fader", color_manager, self)
                menu.exec_(event.globalPos())
                return
        
        # Call parent for normal handling (left-click, etc.)
        super().mousePressEvent(event)

class FaderWithScale(QtWidgets.QWidget):
    valueChanged = QtCore.pyqtSignal(int)

    def __init__(self, initial=FADER_ZERO_DB_CC, h_scale=1.0, v_scale=1.0, is_master=False, is_stereo=False, parent=None):
        super().__init__(parent)
        self.h_scale_factor = h_scale
        self.v_scale_factor = v_scale
        self.is_master = is_master
        self.is_stereo = is_stereo
        self.slider = CustomFaderSlider(is_master=is_master, is_stereo=is_stereo)
        self.slider.setValue(initial)

        self.scale = ScaleWidget(is_master=is_master)
        # Make scale even narrower 
        self.scale.setMinimumWidth(32)  # Keep scale readable

        layout = QtWidgets.QHBoxLayout(self)
        # Use consistent margins for all fader types
        if is_master:
            layout.setContentsMargins(0,8,0,2)  # Master needs different left margin
        else:
            layout.setContentsMargins(1,8,0,2)  # Consistent 8px top margin for all non-master faders
        layout.setSpacing(0)
        
        layout.addWidget(self.scale)
        layout.addWidget(self.slider)
        
        #layout = QtWidgets.QHBoxLayout(self)
        ## Use even tighter spacing
        #layout.setContentsMargins(1,2,0,2)  # Reduced left margin to 0
        #layout.setSpacing(0)  # Negative spacing to overlap slightly
        #layout.addWidget(self.scale)
        #layout.addWidget(self.slider)

        self.slider.valueChanged.connect(self._on_value_changed)
        self.slider.doubleClicked.connect(self.reset_to_zero_db)

        self.apply_scale(self.h_scale_factor, self.v_scale_factor)
        self.slider.update_color(initial)

    def _on_value_changed(self, value):
        self.slider.update_color(value)
        # For master fader, constrain MIDI output to 0dB max (127) even if UI goes higher
        if self.is_master:
            midi_value = min(value, MASTER_ZERO_DB_CC)  # Cap at 127 (0dB) for MIDI
        else:
            midi_value = value
        self.valueChanged.emit(midi_value)

    def reset_to_zero_db(self):
        zero_db_value = MASTER_ZERO_DB_CC if self.is_master else FADER_ZERO_DB_CC
        self.slider.setValue(zero_db_value)

    def setValue(self, v: int):
        self.slider.blockSignals(True)
        if self.is_master:
            # Allow master fader UI to go up to MASTER_UI_MAX_CC for cap visibility
            self.slider.setValue(int(max(0, min(MASTER_UI_MAX_CC, v))))
        else:
            self.slider.setValue(int(max(0, min(127, v))))
        self.slider.blockSignals(False)
        self.slider.update_color(v)

    def value(self) -> int:
        return self.slider.value()

    def apply_scale(self, h_scale: float, v_scale: float):
        self.h_scale_factor = h_scale
        self.v_scale_factor = v_scale
        fader_h = int(BASE_FADER_HEIGHT * v_scale)

        # Add consistent extra height for all fader types to prevent clipping
        if self.is_master:
            fader_h += 15  # Master fader needs extra height
        else:
            fader_h += 12  # All other faders get consistent extra height
    
        self.slider.setFixedHeight(fader_h)
    
        # Only make fader wider for stereo when in wide display mode
        if self.is_stereo:
            fader_w = int(32 * BASE_STEREO_FADER_WIDTH_MULTIPLIER * h_scale)  # Reduced from 40 to 32
            self.slider.setMinimumWidth(fader_w)
        
        self.scale.set_scale(min(h_scale, v_scale))

class PanDial(QtWidgets.QDial):
    doubleClicked = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(self._get_pan_style())

    def _get_pan_style(self):
        """Default pan style - this will be overridden by update_color method"""
        return """
            QDial {
                background-color: #444444;
                border: 2px solid #333333;
                border-radius: 27px;
            }
        """
            
    def update_color(self, value: int):
        # Get color manager from the main window
        color_manager = None
        widget = self.parent()
        while widget and not hasattr(widget, 'color_manager'):
            widget = widget.parent()
        if widget and hasattr(widget, 'color_manager'):
            color_manager = widget.color_manager
        
        if value == PAN_CENTER_CC:
            # Use color manager's pan center color with matching border
            if color_manager:
                center_color = color_manager.get_pan_center_color()
            else:
                center_color = "#00ff00"  # Fallback
                
            # Create a slightly darker border color
            border_color = self._get_darker_color(center_color)
                
            self.setStyleSheet(f"""
                QDial {{
                    background-color: {center_color};
                    border: 2px solid {border_color};
                    border-radius: 27px;
                }}
            """)
        else:
            # Knob is not at center - determine which color to use
            if color_manager:
                use_separate_colors = color_manager.get_pan_use_separate_lr_colors()
                
                if use_separate_colors:
                    # Use separate colors for left and right
                    if value < PAN_CENTER_CC:
                        # Turned left
                        knob_color = color_manager.get_pan_left_color()
                    else:
                        # Turned right
                        knob_color = color_manager.get_pan_right_color()
                else:
                    # Use single off-center color
                    knob_color = color_manager.get_pan_off_center_color()
            else:
                # Fallback to default gray
                knob_color = "#444444"
            
            # Create a slightly darker border color
            border_color = self._get_darker_color(knob_color)
            
            self.setStyleSheet(f"""
                QDial {{
                    background-color: {knob_color};
                    border: 2px solid {border_color};
                    border-radius: 27px;
                }}
            """)

    def _get_darker_color(self, hex_color: str) -> str:
        """Get a darker version of the given hex color for borders"""
        if not hex_color.startswith('#'):
            return "#333333"
        
        try:
            # Remove the # and convert to RGB
            hex_color = hex_color[1:]
            if len(hex_color) == 3:
                hex_color = ''.join([c*2 for c in hex_color])
            
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            
            # Make it darker by reducing each component by 30%
            r = max(0, int(r * 0.7))
            g = max(0, int(g * 0.7))
            b = max(0, int(b * 0.7))
            
            return f"#{r:02x}{g:02x}{b:02x}"
        except (ValueError, IndexError):
            return "#333333"  # Fallback to dark gray

    def refresh_colors(self):
        """Refresh pan colors from color manager"""
        self.update_color(self.value())

    def mouseDoubleClickEvent(self, e: QtGui.QMouseEvent):
        if e.button() == QtCore.Qt.LeftButton:
            self.doubleClicked.emit()
        super().mouseDoubleClickEvent(e)

    def mousePressEvent(self, event):
        """Handle mouse press events including right-click"""
        if event.button() == QtCore.Qt.RightButton:
            # Find the color manager
            color_manager = None
            widget = self.parent()
            while widget and not hasattr(widget, 'color_manager'):
                widget = widget.parent()
            
            if color_manager := (widget.color_manager if widget and hasattr(widget, 'color_manager') else None):
                menu = ContextColorMenu("pan", color_manager, self)
                menu.exec_(event.globalPos())
                return
        
        # Call parent for normal handling
        super().mousePressEvent(event)

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
            self.setStyleSheet("QPushButton { background-color: rgb(249,7,7); color: white; font-weight: bold; padding-left: 2px; padding-right: 2px; }")
        else:
            self.setStyleSheet("QPushButton { background-color: #464646; color: white; font-weight: bold; padding-left: 2px; padding-right: 2px; }")
        # mute button on is: self.setStyleSheet("QPushButton { background-color: rgb(249,7,7) [red]; color: white; font-weight: bold; padding-left: 2px; padding-right: 2px; }")
        # mute button off is: self.setStyleSheet("QPushButton { background-color: #464646 [grey]; color: white; font-weight: bold; padding-left: 2px; padding-right: 2px; }")
        #

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
            if scribble_key and "Channel" in scribble_key:
                channel_num = int(scribble_key.split()[-1])
                is_left = self.mono_stereo_manager.is_stereo_left(scribble_key) if self.mono_stereo_manager else (channel_num % 2 == 1)
                display_title = f"Ch {channel_num} {'L' if is_left else 'R'}"
            elif scribble_key and "Bus" in scribble_key:
                bus_num = int(scribble_key.split()[-1])
                is_left = self.mono_stereo_manager.is_stereo_left(scribble_key) if self.mono_stereo_manager else (bus_num % 2 == 1)
                display_title = f"Bus {bus_num} {'L' if is_left else 'R'}"
            elif scribble_key and "Aux" in scribble_key:
                aux_num = int(scribble_key.split()[-1])
                is_left = self.mono_stereo_manager.is_stereo_left(scribble_key) if self.mono_stereo_manager else (aux_num % 2 == 1)
                display_title = f"Aux {aux_num} {'L' if is_left else 'R'}"
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
        self.v.setContentsMargins(2,2,2,2)
        self.v.setSpacing(2)
        self.setMinimumWidth(int(BASE_STRIP_WIDTH * self.h_scale_factor))

        # Title with stereo indicator
        title_layout = QtWidgets.QHBoxLayout()
        self.title_label = QtWidgets.QLabel(display_title)
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.title_label.setStyleSheet(f"font-weight: bold; color: {title_color};")
        title_layout.addWidget(self.title_label)

        self.v.addLayout(title_layout)

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
            
            # Add pan/balance label based on stereo status
            pan_label_text = "Bal" if is_stereo_pair else "Pan"
            self.pan_label = QtWidgets.QLabel(pan_label_text)
            self.pan_label.setAlignment(QtCore.Qt.AlignCenter)
            self.pan_label.setStyleSheet("font-size: 14px; color: rgb(60,60,60); font-weight: bold;")
            #print(f"Created pan_label with text: {pan_label_text}, is_stereo_pair: {is_stereo_pair}")  # DEBUG
            self.v.addWidget(self.pan_label)
            
    
            self.pan.update_color(PAN_CENTER_CC)
        else:
            self.pan = None
            self.pan_value = None
            self.pan_label = None

        if has_mute:
            self.mute = MuteButton()
            # Center the mute button in the strip
            self.v.addWidget(self.mute, 0, QtCore.Qt.AlignCenter)
        else:
            self.mute = None

        initial_val = MASTER_ZERO_DB_CC if is_master else FADER_ZERO_DB_CC
        self.v.addSpacing(5)  # Add gap before fader
        # Always pass True for is_stereo when this strip is part of a stereo pair
        self.fader = FaderWithScale(initial=initial_val, h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=is_master, is_stereo=is_stereo_pair)
        
        if is_master:
            # Add stretch above the fader to push it to bottom
            self.v.addStretch(1)
            self.v.addWidget(self.fader, 0)
        else:
            self.v.addWidget(self.fader, 1, QtCore.Qt.AlignRight)

        if has_scribble and not is_master:
            self.scribble = ScribbleStripTextEdit()
            self.scribble.textCommitted.connect(self._on_scribble_changed)
            self.v.addWidget(self.scribble)
        else:
            self.scribble = None

           
        self.apply_scale(self.h_scale_factor, self.v_scale_factor)

    #def set_stereo_partner(self, partner_strip): DUPLICATE OF ROW 949
    #    self.stereo_partner_strip = partner_strip

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
            if v == PAN_CENTER_CC:
                self.pan_value.setText("C")
            elif v < PAN_CENTER_CC:
                self.pan_value.setText(f"L{PAN_CENTER_CC - v}")
            else:
                self.pan_value.setText(f"R{v - PAN_CENTER_CC}")

        if self.pan:
            self.pan.update_color(v)

    def _pan_reset(self):
        if self.pan:
            self.pan.setValue(PAN_CENTER_CC)

    def apply_scale(self, h_scale: float, v_scale: float):
        self.h_scale_factor = h_scale
        self.v_scale_factor = v_scale

        # Use appropriate width and font based on strip type
        if self.is_master:
            fixed_width = 60
            title_f = 14  # Increased from 12 for "Main" label
        else:
            fixed_width = 75
            title_f = 10  # Increased from 8 for channel/bus/aux titles

        self.setMinimumWidth(fixed_width)
        self.setMaximumWidth(fixed_width)

        small_f = 9  # Increased from 7 for pan value and labels

        current_color = "rgb(20,20,225)" if self.is_stereo_pair else "black"
        self.title_label.setStyleSheet(f"font-size: {title_f}px; font-weight: bold; color: {current_color};")

        if self.pan_value:
            self.pan_value.setStyleSheet(f"font-size: {small_f}px;")

        if hasattr(self, 'pan_label') and self.pan_label:
            self.pan_label.setStyleSheet(f"font-size: {small_f}px; color: rgb(60,60,60); font-weight: bold;")
            
        if self.mute:
            # Make mute button twice the height and 50% wider than height
            mute_height = int(BASE_BUTTON_H * v_scale * 2.2)  # Double the height
            mute_width = int(mute_height * 1.5)  # 50% wider than height
            self.mute.setFixedSize(mute_width, mute_height)

        if self.pan:
            size = 55  # Fixed smaller pan dial size
            self.pan.setFixedSize(size, size)
            
        if self.scribble:
            scribble_h = int(max(40, 60 * v_scale))
            self.scribble.setMaximumHeight(scribble_h)

        self.fader.apply_scale(h_scale, v_scale)
        
    def set_stereo_partner(self, partner_strip):
        """Set the stereo partner strip for synchronized control"""
        self.stereo_partner_strip = partner_strip

    def _sync_fader_to_partner(self, value: int):
        """Update partner's fader without triggering signals"""
        if self.stereo_partner_strip and hasattr(self.stereo_partner_strip, 'fader'):
            self.stereo_partner_strip.fader.blockSignals(True)
            self.stereo_partner_strip.fader.setValue(value)
            self.stereo_partner_strip.fader.blockSignals(False)

    def _sync_pan_to_partner(self, value: int):
        """Update partner's pan without triggering signals"""
        if self.stereo_partner_strip and hasattr(self.stereo_partner_strip, 'pan') and self.stereo_partner_strip.pan:
            self.stereo_partner_strip.pan.blockSignals(True)
            self.stereo_partner_strip.pan.setValue(value)
            self.stereo_partner_strip.pan.blockSignals(False)
            self.stereo_partner_strip._on_pan_changed(value)

    def _sync_mute_to_partner(self, is_muted: bool):
        """Update partner's mute without triggering signals"""
        if self.stereo_partner_strip and hasattr(self.stereo_partner_strip, 'mute') and self.stereo_partner_strip.mute:
            self.stereo_partner_strip.mute.blockSignals(True)
            self.stereo_partner_strip.mute.setChecked(is_muted)
            self.stereo_partner_strip.mute.blockSignals(False)
            self.stereo_partner_strip.mute.update_style()
    
    def _sync_fader_to_partner(self, value: int):
        """Update partner's fader without triggering signals"""
        if self.stereo_partner_strip and hasattr(self.stereo_partner_strip, 'fader'):
            self.stereo_partner_strip.fader.blockSignals(True)
            self.stereo_partner_strip.fader.setValue(value)
            self.stereo_partner_strip.fader.blockSignals(False)

    def _sync_pan_to_partner(self, value: int):
        """Update partner's pan without triggering signals"""
        if self.stereo_partner_strip and hasattr(self.stereo_partner_strip, 'pan') and self.stereo_partner_strip.pan:
            self.stereo_partner_strip.pan.blockSignals(True)
            self.stereo_partner_strip.pan.setValue(value)
            self.stereo_partner_strip.pan.blockSignals(False)
            self.stereo_partner_strip._on_pan_changed(value)

    def _sync_mute_to_partner(self, is_muted: bool):
        """Update partner's mute without triggering signals"""
        if self.stereo_partner_strip and hasattr(self.stereo_partner_strip, 'mute') and self.stereo_partner_strip.mute:
            self.stereo_partner_strip.mute.blockSignals(True)
            self.stereo_partner_strip.mute.setChecked(is_muted)
            self.stereo_partner_strip.mute.blockSignals(False)
            self.stereo_partner_strip.mute.update_style()
    
    def get_fader_value(self) -> int:
        """Get the current fader value"""
        return self.fader.value()

    def set_fader_value(self, val: int):
        """Set the fader value"""
        self.fader.setValue(val)

    def get_pan_value(self) -> int:
        """Get the current pan value"""
        return self.pan.value() if self.pan else PAN_CENTER_CC

    def set_pan_value(self, val: int):
        """Set the pan value"""
        if self.pan:
            self.pan.setValue(val)

    def get_mute_value(self) -> bool:
        """Get the current mute state"""
        return self.mute.isChecked() if self.mute else False

    def set_mute_value(self, on: bool):
        """Set the mute state"""
        if self.mute:
            self.mute.setChecked(on)
    
    
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
        self.v.setContentsMargins(0,0,0,0)
        self.v.setSpacing(0)
        
        # Double width for wide stereo strips (2x regular width)
        fixed_width = 150  # 2x the regular strip width
        self.setMinimumWidth(fixed_width)
        self.setMaximumWidth(fixed_width)
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

        # Use existing components from ChannelStrip
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
            
            # Always "Bal" for wide stereo strips
            self.pan_label = QtWidgets.QLabel("Bal")
            self.pan_label.setAlignment(QtCore.Qt.AlignCenter)
            self.pan_label.setStyleSheet("font-size: 14px; color: rgb(20,20,225); font-weight: bold;")
            self.v.addWidget(self.pan_label)
            
            self.pan.update_color(PAN_CENTER_CC)
        else:
            self.pan = None
            self.pan_value = None
            self.pan_label = None

        if has_mute:
            self.mute = MuteButton()
            # Center the mute button in the strip
            self.v.addWidget(self.mute, 0, QtCore.Qt.AlignCenter)
        else:
            self.mute = None

        # Use your existing fader - we'll make it wider in apply_scale  
        self.v.addSpacing(5)  # Add gap before fader to match ChannelStrip
        self.fader = FaderWithScale(initial=FADER_ZERO_DB_CC, h_scale=self.h_scale_factor, 
                                  v_scale=self.v_scale_factor, is_master=False, is_stereo=True)
        self.v.addWidget(self.fader, 1, QtCore.Qt.AlignRight)  # Also align right like ChannelStrip


        if has_scribble:
                    self.scribble = ScribbleStripTextEdit()
                    self.scribble.setPlainText(DEFAULT_SCRIBBLE_TEXT)
                    self.scribble.textCommitted.connect(self._on_scribble_changed)
                    self.v.addWidget(self.scribble)
        else:
            self.scribble = None

        # Apply custom styling to the scale widget in wide stereo strips only
        if hasattr(self.fader, 'scale'):
            # Override the paintEvent method for this specific scale widget
            original_paint_event = self.fader.scale.paintEvent

            def custom_paint_event(event):
                # Call original paintEvent but with modified painter
                painter = QtGui.QPainter(self.fader.scale)
                rect = self.fader.scale.rect().adjusted(6, 6, -6, -6)
                painter.setRenderHint(QtGui.QPainter.Antialiasing, True)

                painter.setPen(QtGui.QPen(QtCore.Qt.gray, 1))
                painter.drawLine(rect.right()-2, rect.top(), rect.right()-2, rect.bottom())

                def y_for_cc(cc: int) -> int:
                    t = cc / 127.0
                    return int(rect.top() + (1.0 - t) * rect.height())

                painter.setPen(QtGui.QPen(QtCore.Qt.black, 1))
                font = painter.font()
                font.setPointSizeF(max(7.0, BASE_SMALL_FONT * self.fader.scale.font_scale))
                painter.setFont(font)

                scale_points = MASTER_DB_POINTS if self.fader.scale.is_master else CHANNEL_DB_POINTS

                for cc, label in scale_points:
                    y = y_for_cc(cc) + 3  # Add 3px downward offset for wide stereo strip alignment
                    painter.drawLine(rect.right()-6, y, rect.right()-2, y)
                    text_rect = QtCore.QRect(rect.left(), y-8, rect.width()-8, 16)
        
                    # Remove "dB" text like we did before
                    if "+" in label:
                        short_label = label.replace(" dB", "").replace("+", "+")
                    elif "-∞" in label or "-âˆž" in label:
                        short_label = "-∞"
                    else:
                        short_label = label.replace(" dB", "")
        
                    painter.drawText(text_rect, QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter, short_label)

            # Replace the paint event for this specific scale widget
            self.fader.scale.paintEvent = custom_paint_event

        self.apply_scale(h_scale, v_scale)
    
    # Include all the same methods as ChannelStrip
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
        text = self.scribble.toPlainText() if self.scribble else ""
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
    
        # Set fixed width for wide stereo strips
        fixed_width = 150  # 2x the regular strip width to accommodate wider layout
        self.setMinimumWidth(fixed_width)
        self.setMaximumWidth(fixed_width)
    
        # Use fixed font sizes like regular ChannelStrip (not scaled)
        title_f = 10  # Increased from 8 for wide stereo title font
        small_f = 9   # Increased from 7 for pan value and labels

        self.title_label.setStyleSheet(f"font-size: {title_f}px; font-weight: bold; color: rgb(20,20,225);")

        if self.pan_value:
            self.pan_value.setStyleSheet(f"font-size: {small_f}px;")

        if hasattr(self, 'pan_label') and self.pan_label:
            self.pan_label.setStyleSheet(f"font-size: {small_f}px; color: rgb(20,20,225); font-weight: bold;")

        if self.mute:
            # Make mute button twice the height and 50% wider than height
            mute_height = int(BASE_BUTTON_H * v_scale * 2.2)  # Double the height
            mute_width = int(mute_height * 1.5)  # 50% wider than height
            self.mute.setFixedSize(mute_width, mute_height)

        if self.pan:
            # Use fixed pan dial size like regular ChannelStrip (not scaled)
            size = 55  # Fixed smaller pan dial size  
            self.pan.setFixedSize(size, size)

        if self.scribble:
            scribble_h = int(max(40, 60 * v_scale))
            self.scribble.setMaximumHeight(scribble_h)

        self.fader.apply_scale(h_scale, v_scale)
        
class MixerWindow(QtWidgets.QMainWindow):
    def __init__(self, csv_path: str):
        super().__init__()
        self.setWindowTitle("DM4800 Mixer Controller")
        self.resize(1700, 1000)

        self.settings = SettingsManager()
        self.color_manager = ColorThemeManager(self.settings, self)
        self._connect_color_manager_signals()
        
        # Apply initial styling using color manager
        self._apply_background_colors()

        self.mappings = load_midi_mappings(csv_path)
        self.scribble_manager = ScribbleStripManager()
        self.mono_stereo_manager = MonoStereoManager()
        self.midi = MidiManager()
        self.midi.midi_message_received.connect(self.on_midi_cc)
        
        self.midi_monitor = MidiMonitorWindow()
        self.midi.midi_message_received.connect(lambda ch, cc, val: self.midi_monitor.add_message("IN ", ch, cc, val))
        self.midi.midi_message_sent.connect(lambda ch, cc, val: self.midi_monitor.add_message("OUT", ch, cc, val))

        self.h_scale_factor = 1.0  
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
        self.rows_v.setContentsMargins(1,1,1,1)
        self.rows_v.setSpacing(1)
        scroll.setWidget(container)

        self.channel_widgets: Dict[Tuple[str,int], ChannelStrip] = {}
        self.bus_widgets: Dict[int, ChannelStrip] = {}
        self.aux_widgets: Dict[int, ChannelStrip] = {}
        self.master_widget: Optional[ChannelStrip] = None
        self.stereo_pairs: Dict[str, ChannelStrip] = {}

        # State persistence for display mode switching
        self._saved_control_state = {
            'faders': {},  # Key: (section, number) -> value
            'pans': {},    # Key: (section, number) -> value  
            'mutes': {}    # Key: (section, number) -> value
        }
        self._rebuilding_interface = False  # Guard flag to prevent recursive rebuilds

        self._build_rows()
        self._setup_stereo_pairs()
        self._rewire_all_controls()
        self._load_scribble_strips()
        self._install_zoom_shortcuts()

        # Apply colors at startup
        self._apply_background_colors()
        self._refresh_all_fader_colors()
        self._refresh_all_pan_colors()
        self._apply_master_strip_background_color()

        if self.settings.get('remember_ports', True):
            self._auto_open_midi_ports()

    def closeEvent(self, event):
        self.settings.save_settings()
        self.midi.close()
        self.midi_monitor.close()
        super().closeEvent(event)

    def _apply_background_colors(self):
        """Apply background colors from color manager"""
        bg_color = self.color_manager.get_background_color()
        self.setStyleSheet(f"""
            QMainWindow {{ background-color: {bg_color}; color: black; }}
            QWidget {{ background-color: {bg_color}; color: black; }}
            QScrollArea {{ background-color: {bg_color}; color: black; }}
            QLabel {{ color: black; }}
            QPushButton {{ color: black; }}
            QComboBox {{ color: black; }}
        """)
    
    def _connect_color_manager_signals(self):
        """Connect color manager signals to update UI"""
        self.color_manager.backgroundColorChanged.connect(self._on_background_color_changed)
        self.color_manager.faderZeroDbColorChanged.connect(self._on_fader_color_changed)
        self.color_manager.panCenterColorChanged.connect(self._on_pan_color_changed)
        self.color_manager.panOffCenterColorChanged.connect(self._on_pan_color_changed)
        self.color_manager.panLeftColorChanged.connect(self._on_pan_color_changed)
        self.color_manager.panRightColorChanged.connect(self._on_pan_color_changed)
        self.color_manager.panUseSeparateLRColorsChanged.connect(self._on_pan_color_changed)
        self.color_manager.masterStripBackgroundColorChanged.connect(self._on_master_strip_color_changed)
    
    def _on_background_color_changed(self, new_color: str):
        """Handle background color change"""
        self._apply_background_colors()
    
    def _on_fader_color_changed(self):
        """Handle fader color changes"""
        self._refresh_all_fader_colors()
    
    def _on_pan_color_changed(self):
        """Handle pan color changes"""
        self._refresh_all_pan_colors()
    
    def _on_master_strip_color_changed(self):
        """Handle master strip color changes"""
        self._apply_master_strip_background_color()

    def _on_fader_color_changed(self):
        """Handle fader color changes"""
        self._refresh_all_fader_colors()
    
    def _on_pan_color_changed(self):
        """Handle pan color changes"""
        self._refresh_all_pan_colors()

    def _refresh_all_fader_colors(self):
        """Refresh colors on all faders"""
        # Refresh channel strip faders
        for strip in self.channel_widgets.values():
            if hasattr(strip, 'fader') and hasattr(strip.fader, 'slider'):
                strip.fader.slider.refresh_colors()
        
        # Refresh bus strip faders
        for strip in self.bus_widgets.values():
            if hasattr(strip, 'fader') and hasattr(strip.fader, 'slider'):
                strip.fader.slider.refresh_colors()
        
        # Refresh aux strip faders
        for strip in self.aux_widgets.values():
            if hasattr(strip, 'fader') and hasattr(strip.fader, 'slider'):
                strip.fader.slider.refresh_colors()
        
        # Refresh master fader
        if self.master_widget and hasattr(self.master_widget, 'fader') and hasattr(self.master_widget.fader, 'slider'):
            self.master_widget.fader.slider.refresh_colors()
    
    def _refresh_all_pan_colors(self):
        """Refresh pan colors for all strips"""
        # Refresh channel strips
        for strip in self.channel_widgets.values():
            if hasattr(strip, 'pan') and strip.pan:
                strip.pan.refresh_colors()
        
        # Refresh bus strips
        for strip in self.bus_widgets.values():
            if hasattr(strip, 'pan') and strip.pan:
                strip.pan.refresh_colors()
        
        # Refresh aux strips
        for strip in self.aux_widgets.values():
            if hasattr(strip, 'pan') and strip.pan:
                strip.pan.refresh_colors()
        
        # Refresh wide stereo strips
        for strip in self.stereo_pairs.values():
            if hasattr(strip, 'left_pan') and strip.left_pan:
                strip.left_pan.refresh_colors()
            if hasattr(strip, 'right_pan') and strip.right_pan:
                strip.right_pan.refresh_colors()

    def _apply_master_strip_background_color(self):
        """Apply master strip background color"""
        if self.master_widget:
            master_bg_color = self.color_manager.get_master_strip_background_color()
            self.master_widget.setStyleSheet(f"ChannelStrip {{ background-color: {master_bg_color}; }}")

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
        
        # Settings menu
        settings_menu = menubar.addMenu("&Settings")

        # Stereo display toggle
        self.stereo_display_action = QtWidgets.QAction("Display Stereo Pairs as Single Channel Strips", self)
        self.stereo_display_action.setCheckable(True)
        self.stereo_display_action.setChecked(self.settings.get('stereo_display_mode') == 'wide_fader')
        self.stereo_display_action.triggered.connect(self._toggle_stereo_display_mode)
        settings_menu.addAction(self.stereo_display_action)
        
        # Add separator and color settings
        settings_menu.addSeparator()
        
        self.color_settings_action = QtWidgets.QAction("Colors...", self)
        self.color_settings_action.triggered.connect(self._show_settings_dialog)
        settings_menu.addAction(self.color_settings_action)
        
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

    def _toggle_remember_ports(self, enabled: bool):
        self.settings.set('remember_ports', enabled)
        self.settings.save_settings()

    def _refresh_midi_ports(self):
        # Close existing connections
        self.midi.close()
        # The port lists will be refreshed when the user next opens the selection dialogs
        QtWidgets.QMessageBox.information(self, "MIDI Ports Refreshed", "MIDI port lists have been refreshed. Please reselect your ports if needed.")


    def _show_midi_monitor(self):
        self.midi_monitor.show()
        self.midi_monitor.raise_()
        self.midi_monitor.activateWindow()

    def _show_about(self):
        about_text = """Version: v022a
        Author: GrooveMonkey
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
    
    def _capture_current_state(self):
        """Capture current fader, pan, and mute values before rebuilding interface"""
        # Clear previous state
        self._saved_control_state = {
            'faders': {},
            'pans': {},
            'mutes': {}
        }
        
        # Capture channel states
        for (section, number), strip in self.channel_widgets.items():
            if section == "channel":
                # Store fader value
                self._saved_control_state['faders'][('channel', number)] = strip.get_fader_value()
                
                # Store pan value if strip has pan control
                if hasattr(strip, 'get_pan_value') and strip.get_pan_value is not None:
                    self._saved_control_state['pans'][('channel', number)] = strip.get_pan_value()
                
                # Store mute value if strip has mute control
                if hasattr(strip, 'get_mute_value'):
                    self._saved_control_state['mutes'][('channel', number)] = strip.get_mute_value()
        
        # Capture bus states
        for number, strip in self.bus_widgets.items():
            # Store fader value
            self._saved_control_state['faders'][('bus', number)] = strip.get_fader_value()
            
            # Store pan value if strip has pan control
            if hasattr(strip, 'get_pan_value') and strip.get_pan_value is not None:
                self._saved_control_state['pans'][('bus', number)] = strip.get_pan_value()
            
            # Store mute value if strip has mute control  
            if hasattr(strip, 'get_mute_value'):
                self._saved_control_state['mutes'][('bus', number)] = strip.get_mute_value()
        
        # Capture aux states
        for number, strip in self.aux_widgets.items():
            # Store fader value
            self._saved_control_state['faders'][('aux', number)] = strip.get_fader_value()
            
            # Store pan value if strip has pan control
            if hasattr(strip, 'get_pan_value') and strip.get_pan_value is not None:
                self._saved_control_state['pans'][('aux', number)] = strip.get_pan_value()
            
            # Store mute value if strip has mute control
            if hasattr(strip, 'get_mute_value'):
                self._saved_control_state['mutes'][('aux', number)] = strip.get_mute_value()
        
        # Capture master state
        if self.master_widget:
            self._saved_control_state['faders'][('master', 1)] = self.master_widget.get_fader_value()
            
            # Master typically doesn't have pan or mute, but check just in case
            if hasattr(self.master_widget, 'get_pan_value') and self.master_widget.get_pan_value is not None:
                self._saved_control_state['pans'][('master', 1)] = self.master_widget.get_pan_value()
            
            if hasattr(self.master_widget, 'get_mute_value'):
                self._saved_control_state['mutes'][('master', 1)] = self.master_widget.get_mute_value()
        
        #print(f"Captured state: {len(self._saved_control_state['faders'])} faders, {len(self._saved_control_state['pans'])} pans, {len(self._saved_control_state['mutes'])} mutes")

    def _restore_saved_state(self):
        """Restore previously captured fader, pan, and mute values after rebuilding interface"""
        if not self._saved_control_state:
            return
        
        restored_count = {'faders': 0, 'pans': 0, 'mutes': 0}
        
        # Restore channel states
        for (section, number), strip in self.channel_widgets.items():
            if section == "channel":
                # Restore fader value - set directly on slider widget
                fader_key = ('channel', number)
                if fader_key in self._saved_control_state['faders']:
                    if hasattr(strip, 'fader') and strip.fader:
                        value = self._saved_control_state['faders'][fader_key]
                        strip.fader.setValue(value)  # Direct widget call
                        restored_count['faders'] += 1
                
                # Restore pan value - set directly on pan widget
                pan_key = ('channel', number)
                if pan_key in self._saved_control_state['pans']:
                    if hasattr(strip, 'pan') and strip.pan:
                        value = self._saved_control_state['pans'][pan_key]
                        strip.pan.setValue(value)  # Direct widget call
                        # Update the display text manually
                        if hasattr(strip, 'pan_value') and strip.pan_value:
                            if value == 64:  # PAN_CENTER_CC
                                strip.pan_value.setText("C")
                            elif value < 64:
                                strip.pan_value.setText(f"L{64 - value}")
                            else:
                                strip.pan_value.setText(f"R{value - 64}")
                        # Update pan color manually
                        if hasattr(strip.pan, 'update_color'):
                            strip.pan.update_color(value)
                        restored_count['pans'] += 1
                
                # Restore mute value - set directly on mute widget
                mute_key = ('channel', number)
                if mute_key in self._saved_control_state['mutes']:
                    if hasattr(strip, 'mute') and strip.mute:
                        value = self._saved_control_state['mutes'][mute_key]
                        strip.mute.setChecked(value)  # Direct widget call
                        # Update mute button style manually
                        if hasattr(strip.mute, 'update_style'):
                            strip.mute.update_style()
                        restored_count['mutes'] += 1
        
        # Restore bus states
        for number, strip in self.bus_widgets.items():
            # Restore fader value
            fader_key = ('bus', number)
            if fader_key in self._saved_control_state['faders']:
                if hasattr(strip, 'fader') and strip.fader:
                    value = self._saved_control_state['faders'][fader_key]
                    strip.fader.setValue(value)
                    restored_count['faders'] += 1
            
            # Restore pan value
            pan_key = ('bus', number)
            if pan_key in self._saved_control_state['pans']:
                if hasattr(strip, 'pan') and strip.pan:
                    value = self._saved_control_state['pans'][pan_key]
                    strip.pan.setValue(value)
                    # Update display and color
                    if hasattr(strip, 'pan_value') and strip.pan_value:
                        if value == 64:
                            strip.pan_value.setText("C")
                        elif value < 64:
                            strip.pan_value.setText(f"L{64 - value}")
                        else:
                            strip.pan_value.setText(f"R{value - 64}")
                    if hasattr(strip.pan, 'update_color'):
                        strip.pan.update_color(value)
                    restored_count['pans'] += 1
            
            # Restore mute value
            mute_key = ('bus', number)
            if mute_key in self._saved_control_state['mutes']:
                if hasattr(strip, 'mute') and strip.mute:
                    value = self._saved_control_state['mutes'][mute_key]
                    strip.mute.setChecked(value)
                    if hasattr(strip.mute, 'update_style'):
                        strip.mute.update_style()
                    restored_count['mutes'] += 1
        
        # Restore aux states
        for number, strip in self.aux_widgets.items():
            # Restore fader value
            fader_key = ('aux', number)
            if fader_key in self._saved_control_state['faders']:
                if hasattr(strip, 'fader') and strip.fader:
                    value = self._saved_control_state['faders'][fader_key]
                    strip.fader.setValue(value)
                    restored_count['faders'] += 1
            
            # Restore pan value
            pan_key = ('aux', number)
            if pan_key in self._saved_control_state['pans']:
                if hasattr(strip, 'pan') and strip.pan:
                    value = self._saved_control_state['pans'][pan_key]
                    strip.pan.setValue(value)
                    # Update display and color
                    if hasattr(strip, 'pan_value') and strip.pan_value:
                        if value == 64:
                            strip.pan_value.setText("C")
                        elif value < 64:
                            strip.pan_value.setText(f"L{64 - value}")
                        else:
                            strip.pan_value.setText(f"R{value - 64}")
                    if hasattr(strip.pan, 'update_color'):
                        strip.pan.update_color(value)
                    restored_count['pans'] += 1
            
            # Restore mute value
            mute_key = ('aux', number)
            if mute_key in self._saved_control_state['mutes']:
                if hasattr(strip, 'mute') and strip.mute:
                    value = self._saved_control_state['mutes'][mute_key]
                    strip.mute.setChecked(value)
                    if hasattr(strip.mute, 'update_style'):
                        strip.mute.update_style()
                    restored_count['mutes'] += 1
        
        # Restore master state
        if self.master_widget:
            fader_key = ('master', 1)
            if fader_key in self._saved_control_state['faders']:
                if hasattr(self.master_widget, 'fader') and self.master_widget.fader:
                    value = self._saved_control_state['faders'][fader_key]
                    self.master_widget.fader.setValue(value)
                    restored_count['faders'] += 1
            
            # Master usually doesn't have pan or mute, but handle if it does
            pan_key = ('master', 1)
            if pan_key in self._saved_control_state['pans']:
                if hasattr(self.master_widget, 'pan') and self.master_widget.pan:
                    value = self._saved_control_state['pans'][pan_key]
                    self.master_widget.pan.setValue(value)
                    restored_count['pans'] += 1
            
            mute_key = ('master', 1)
            if mute_key in self._saved_control_state['mutes']:
                if hasattr(self.master_widget, 'mute') and self.master_widget.mute:
                    value = self._saved_control_state['mutes'][mute_key]
                    self.master_widget.mute.setChecked(value)
                    if hasattr(self.master_widget.mute, 'update_style'):
                        self.master_widget.mute.update_style()
                    restored_count['mutes'] += 1
        
        print(f"Restored state: {restored_count['faders']} faders, {restored_count['pans']} pans, {restored_count['mutes']} mutes")

    def _on_stereo_checkbox_changed(self, checked: bool):
        """Handle stereo display checkbox change"""
        new_mode = 'wide_fader' if checked else 'linked_pair'
        self.settings.set('stereo_display_mode', new_mode)
        self.settings.save_settings()
    
        # Update menu action to match
        self.stereo_display_action.blockSignals(True)
        self.stereo_display_action.setChecked(checked)
        self.stereo_display_action.blockSignals(False)
    
        # Rebuild interface with new mode
        self._rebuild_interface()
    
    def _rebuild_interface(self):
        """Rebuild the entire interface when stereo display mode changes"""
        # Prevent recursive rebuilds
        if self._rebuilding_interface:
            #print("DEBUG: Preventing recursive rebuild")
            return
            
        #print("DEBUG: Starting rebuild")
        self._rebuilding_interface = True
        
        try:
            # STEP 1: Capture current state before destroying widgets
            self._capture_current_state()
            
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
            self._setup_stereo_pairs()
            self._rewire_all_controls()
            self._load_scribble_strips()
            
            # CRITICAL: Reapply colors after rebuilding interface
            self._apply_background_colors()
            self._refresh_all_fader_colors()
            self._refresh_all_pan_colors()
            self._apply_master_strip_background_color()
            
            # STEP 2: Restore saved state after rebuilding widgets
            self._restore_saved_state()
            
        finally:
            # Always reset the rebuild flag, even if an error occurred
            self._rebuilding_interface = False
            #print("DEBUG: Rebuild complete")
    
    def _build_zoom_row(self):
        zoom_h = QtWidgets.QHBoxLayout()

        zoom_h.addWidget(QtWidgets.QLabel("V-Zoom"))
        self.v_zoom_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.v_zoom_slider.setRange(50, 125)
        self.v_zoom_slider.setValue(int(self.v_scale_factor * 100))
        self.v_zoom_slider.valueChanged.connect(self._on_v_zoom_changed)
        self.v_zoom_value_lbl = QtWidgets.QLabel(f"{int(self.v_scale_factor * 100)}%")
        zoom_h.addWidget(self.v_zoom_slider, 1)
        zoom_h.addWidget(self.v_zoom_value_lbl)

        # Add stereo display checkbox
        zoom_h.addSpacing(20)
        self.stereo_checkbox = QtWidgets.QCheckBox("Display Stereo Pairs as Single Channel Strips")
        self.stereo_checkbox.setChecked(self.settings.get('stereo_display_mode') == 'wide_fader')
        self.stereo_checkbox.toggled.connect(self._on_stereo_checkbox_changed)
        zoom_h.addWidget(self.stereo_checkbox)

        self.main_v.addLayout(zoom_h)

    def _install_zoom_shortcuts(self):
        zoom_in = QtWidgets.QShortcut(QtGui.QKeySequence.ZoomIn, self)
        zoom_out = QtWidgets.QShortcut(QtGui.QKeySequence.ZoomOut, self)
        zoom_reset = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+0"), self)
        zoom_reset_mac = QtWidgets.QShortcut(QtGui.QKeySequence("Meta+0"), self)

        zoom_in.activated.connect(lambda: self.v_zoom_slider.setValue(min(125, self.v_zoom_slider.value()+10)))
        zoom_out.activated.connect(lambda: self.v_zoom_slider.setValue(max(50, self.v_zoom_slider.value()-10)))
        zoom_reset.activated.connect(lambda: self.v_zoom_slider.setValue(100))
        zoom_reset_mac.activated.connect(lambda: self.v_zoom_slider.setValue(100))

    def _build_rows(self):
        use_wide_stereo = self.settings.get('stereo_display_mode') == 'wide_fader'
        
        def add_channel_row(start, end, include_pan=True, include_master=False):
            row = QtWidgets.QWidget()
            h = QtWidgets.QHBoxLayout(row)
            h.setContentsMargins(0,0,0,0)
            h.setSpacing(1)
            h.setAlignment(QtCore.Qt.AlignLeft)
    
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
                            h.addWidget(strip)
                            self.channel_widgets[("channel", ch)] = strip
                            self._wire_strip_controls("channel", ch, strip)
                            strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                        else:
                            # Create left channel strip
                            left_strip = ChannelStrip(f"Ch {ch}", scribble_key=scribble_key, has_pan=include_pan, 
                                                   has_mute=True, has_scribble=True,
                                                   h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, 
                                                   is_master=False, mono_stereo_manager=self.mono_stereo_manager,
                                                   is_stereo_pair=True, stereo_partner_num=partner_num)
                            h.addWidget(left_strip)
                            self.channel_widgets[("channel", ch)] = left_strip
                            self._wire_strip_controls("channel", ch, left_strip)
                            left_strip.scribbleTextChanged.connect(self._on_scribble_text_changed)

                            # Create right channel strip
                            right_strip = ChannelStrip(f"Ch {partner_num}", scribble_key=partner_key, has_pan=include_pan, 
                                                    has_mute=True, has_scribble=True,
                                                    h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, 
                                                    is_master=False, mono_stereo_manager=self.mono_stereo_manager,
                                                    is_stereo_pair=True, stereo_partner_num=ch)
                            h.addWidget(right_strip)
                            self.channel_widgets[("channel", partner_num)] = right_strip
                            self._wire_strip_controls("channel", partner_num, right_strip)
                            right_strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                        
                        strips_added += 2
                        ch += 2
                    else:
                        ch += 1
                else:
                    strip = ChannelStrip(f"Ch {ch}", scribble_key=scribble_key, has_pan=include_pan, 
                                       has_mute=True, has_scribble=True,
                                       h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, 
                                       is_master=False, mono_stereo_manager=self.mono_stereo_manager)
                    h.addWidget(strip)
                    self.channel_widgets[("channel", ch)] = strip
                    self._wire_strip_controls("channel", ch, strip)
                    strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
                    strips_added += 1
                    ch += 1
            
            if include_master:
                h.addSpacing(1)
                ms = ChannelStrip("Main", has_pan=False, has_mute=False, has_scribble=False,
                                h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=True)
                h.addWidget(ms)
                self.master_widget = ms
                self._wire_strip_controls("master", 1, ms)
            
            self.rows_v.addWidget(row)
            h.addStretch()  # Push all strips to the left
            return ch

        next_ch = add_channel_row(1, 64, include_pan=True, include_master=True)
        
        if next_ch <= 64:
            next_ch = add_channel_row(next_ch, 64, include_pan=True, include_master=False)
        
        if next_ch <= 64:
            add_channel_row(next_ch, 64, include_pan=True, include_master=False)



        bus_row = QtWidgets.QWidget()
        hbus = QtWidgets.QHBoxLayout(bus_row)
        hbus.setContentsMargins(0,0,0,0)
        hbus.setSpacing(1)
        hbus.setAlignment(QtCore.Qt.AlignLeft)
        
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
                        hbus.addWidget(strip)
                        self.bus_widgets[b] = strip
                        self._wire_strip_controls("bus", b, strip)
                        strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
            
                        strips_added += 2
                        b += 2
                    else:
                        # Create both left and right bus strips for linked pair mode
                        left_strip = ChannelStrip(f"Bus {b}", scribble_key=scribble_key, has_pan=False, has_mute=True, has_scribble=True,
                                               h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=False, 
                                               mono_stereo_manager=self.mono_stereo_manager, is_stereo_pair=True, stereo_partner_num=partner_num)
                        hbus.addWidget(left_strip)
                        self.bus_widgets[b] = left_strip
                        self._wire_strip_controls("bus", b, left_strip)
                        left_strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
            
                        right_strip = ChannelStrip(f"Bus {partner_num}", scribble_key=partner_key, has_pan=False, has_mute=True, has_scribble=True,
                                                 h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=False,
                                                 mono_stereo_manager=self.mono_stereo_manager, is_stereo_pair=True, stereo_partner_num=b)
                        hbus.addWidget(right_strip)
                        self.bus_widgets[partner_num] = right_strip
                        self._wire_strip_controls("bus", partner_num, right_strip)
                        right_strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
            
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
        
        hbus.addStretch()  # Push all bus strips to the left
        self.rows_v.addWidget(bus_row)

        aux_row = QtWidgets.QWidget()
        haux = QtWidgets.QHBoxLayout(aux_row)
        haux.setContentsMargins(0,0,0,0)
        haux.setSpacing(1)
        haux.setAlignment(QtCore.Qt.AlignLeft)
       
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
                       haux.addWidget(strip)
                       self.aux_widgets[a] = strip
                       self._wire_strip_controls("aux", a, strip)
                       strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
           
                       strips_added += 2
                       a += 2
                   else:
                       # Create both left and right aux strips for linked pair mode
                       left_strip = ChannelStrip(f"Aux {a}", scribble_key=scribble_key, has_pan=False, has_mute=True, has_scribble=True,
                                               h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=False, 
                                               mono_stereo_manager=self.mono_stereo_manager, is_stereo_pair=True, stereo_partner_num=partner_num)
                       haux.addWidget(left_strip)
                       self.aux_widgets[a] = left_strip
                       self._wire_strip_controls("aux", a, left_strip)
                       left_strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
           
                       right_strip = ChannelStrip(f"Aux {partner_num}", scribble_key=partner_key, has_pan=False, has_mute=True, has_scribble=True,
                                                h_scale=self.h_scale_factor, v_scale=self.v_scale_factor, is_master=False,
                                                mono_stereo_manager=self.mono_stereo_manager, is_stereo_pair=True, stereo_partner_num=a)
                       haux.addWidget(right_strip)
                       self.aux_widgets[partner_num] = right_strip
                       self._wire_strip_controls("aux", partner_num, right_strip)
                       right_strip.scribbleTextChanged.connect(self._on_scribble_text_changed)
           
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
        
        haux.addStretch()  # Push all aux strips to the left
        self.rows_v.addWidget(aux_row)

    def _wire_strip_controls(self, section: str, number: int, strip: ChannelStrip):
        def map_for(t: str) -> Optional[MidiMapping]:
            return self.mappings.get((section, number, t))

        key = f"{section.title()} {number}"
        is_stereo = hasattr(strip, 'is_stereo_pair') and strip.is_stereo_pair
        has_partner = hasattr(strip, 'stereo_partner_strip') and strip.stereo_partner_strip
        mode_check = self.settings.get('stereo_display_mode', 'linked_pair') == 'linked_pair'
        is_wide_stereo = isinstance(strip, WideStereoStrip)
        
        # Disconnect any existing connections to avoid double-wiring
        try:
            strip.fader.valueChanged.disconnect()
        except:
            pass
        if strip.mute:
            try:
                strip.mute.toggledCC.disconnect()
                strip.mute.toggled.disconnect()
            except:
                pass
        if strip.pan:
            try:
                strip.pan.valueChanged.disconnect()
                strip.pan.doubleClicked.disconnect()
            except:
                pass
                
        # Handle WideStereoStrip differently
        if isinstance(strip, WideStereoStrip):
            # Wire MIDI for wide stereo strips using left channel mappings
            left_fader_mapping = self.mappings.get((section, strip.left_num, 'fader'))
            if left_fader_mapping:
                strip.fader.valueChanged.connect(lambda v, m=left_fader_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))
        
            if strip.mute:
                # Reconnect essential mute functions
                strip.mute.toggled.connect(strip.mute._on_toggled)
            
                left_mute_mapping = self.mappings.get((section, strip.left_num, 'mute'))
                if left_mute_mapping:
                    strip.mute.toggledCC.connect(lambda v, m=left_mute_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))
        
            if strip.pan:
                # Reconnect essential pan functions
                strip.pan.doubleClicked.connect(strip._pan_reset)
                strip.pan.valueChanged.connect(strip._on_pan_changed)
            
                left_pan_mapping = self.mappings.get((section, strip.left_num, 'pan'))
                if left_pan_mapping:
                    strip.pan.valueChanged.connect(lambda v, m=left_pan_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))
            return


        # Handle MIDI output for stereo pairs vs regular strips
        if (hasattr(strip, 'is_stereo_pair') and strip.is_stereo_pair and hasattr(strip, 'stereo_partner_strip') and strip.stereo_partner_strip and
            self.settings.get('stereo_display_mode', 'linked_pair') == 'linked_pair' and
            not isinstance(strip, WideStereoStrip)):
        
            # For stereo pairs, both channels send the LEFT channel's MIDI commands
            left_key = key if self.mono_stereo_manager.is_stereo_left(key) else self.mono_stereo_manager.get_stereo_partner(key)
            if left_key:
                left_section, left_number = parse_section_and_number(left_key)
            
                # Use left channel's MIDI mappings for both channels
                left_fader_mapping = self.mappings.get((left_section, left_number, 'fader'))
                if left_fader_mapping:
                    strip.fader.valueChanged.connect(lambda v, m=left_fader_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))
                    # Also sync to partner
                    strip.fader.valueChanged.connect(lambda v: strip.sync_to_partner("fader", v))

                if strip.mute:
                    strip.mute.toggled.connect(strip.mute._on_toggled)
                    left_mute_mapping = self.mappings.get((left_section, left_number, 'mute'))
                    if left_mute_mapping:
                        strip.mute.toggledCC.connect(lambda v, m=left_mute_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))
                        strip.mute.toggledCC.connect(lambda v: strip.sync_to_partner("mute", v))

                if strip.pan:
                    strip.pan.doubleClicked.connect(strip._pan_reset)
                    strip.pan.valueChanged.connect(strip._on_pan_changed)
                    left_pan_mapping = self.mappings.get((left_section, left_number, 'pan'))
                    if left_pan_mapping:
                        strip.pan.valueChanged.connect(lambda v, m=left_pan_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))
                        strip.pan.valueChanged.connect(lambda v: strip.sync_to_partner("pan", v))
        else:
            # Regular strip wiring
            fader_mapping = map_for('fader')
            if fader_mapping:
                strip.fader.valueChanged.connect(lambda v, m=fader_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))
            
            if strip.mute:
                strip.mute.toggled.connect(strip.mute._on_toggled)
                mute_mapping = map_for('mute')
                if mute_mapping:
                    strip.mute.toggledCC.connect(lambda v, m=mute_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))
                    
            if strip.pan:
                strip.pan.doubleClicked.connect(strip._pan_reset)
                strip.pan.valueChanged.connect(strip._on_pan_changed)
                pan_mapping = map_for('pan')
                if pan_mapping:
                    strip.pan.valueChanged.connect(lambda v, m=pan_mapping: self.midi.send_cc(m.midi_channel, m.cc_number, v))

    def _setup_stereo_pairs(self):
        """Setup stereo pair relationships after all strips are created"""
        if self.settings.get('stereo_display_mode') != 'linked_pair':
            return
            
        # Link channel pairs
        for (section, number), strip in self.channel_widgets.items():
            if section == "channel":
                key = f"Channel {number}"
                if (hasattr(strip, 'is_stereo_pair') and strip.is_stereo_pair and 
                    hasattr(strip, 'stereo_partner_num') and strip.stereo_partner_num):
                    partner_strip = self.channel_widgets.get(("channel", strip.stereo_partner_num))
                    if partner_strip:
                        strip.set_stereo_partner(partner_strip)
                        
        # Link bus pairs
        for number, strip in self.bus_widgets.items():
            key = f"Bus {number}"
            if (hasattr(strip, 'is_stereo_pair') and strip.is_stereo_pair and 
                hasattr(strip, 'stereo_partner_num') and strip.stereo_partner_num):
                partner_strip = self.bus_widgets.get(strip.stereo_partner_num)
                if partner_strip:
                    strip.set_stereo_partner(partner_strip)
                    
        # Link aux pairs  
        for number, strip in self.aux_widgets.items():
            key = f"Aux {number}"
            if (hasattr(strip, 'is_stereo_pair') and strip.is_stereo_pair and 
                hasattr(strip, 'stereo_partner_num') and strip.stereo_partner_num):
                partner_strip = self.aux_widgets.get(strip.stereo_partner_num)
                if partner_strip:
                    strip.set_stereo_partner(partner_strip)

    def _rewire_all_controls(self):
        for key, strip in self.channel_widgets.items():
            if isinstance(key, tuple):
                section, number = key
                self._wire_strip_controls(section, number, strip)
            
        for number, strip in self.bus_widgets.items():
            self._wire_strip_controls("bus", number, strip)
            
        for number, strip in self.aux_widgets.items():
            self._wire_strip_controls("aux", number, strip)
            
        if self.master_widget:
            self._wire_strip_controls("master", 1, self.master_widget)

    def _load_scribble_strips(self):
        for key, strip in self.channel_widgets.items():
            if isinstance(key, tuple):
                section, number = key
                scribble_key = f"Channel {number}"
            else:
                scribble_key = f"Channel {key}"
            text = self.scribble_manager.get_scribble_text(scribble_key)
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

    def _on_v_zoom_changed(self, val: int):
        self.v_scale_factor = val / 100.0
        self.v_zoom_value_lbl.setText(f"{val}%")
        self.settings.set('vertical_zoom', val)
        self._apply_all_scales()

    def _apply_all_scales(self):
        for key, strip in self.channel_widgets.items():
            strip.apply_scale(1.0, self.v_scale_factor)
        for strip in self.bus_widgets.values():
            strip.apply_scale(1.0, self.v_scale_factor)
        if self.master_widget:
            self.master_widget.apply_scale(1.0, self.v_scale_factor)
        for strip in self.aux_widgets.values():
            strip.apply_scale(1.0, self.v_scale_factor)
        # Reapply master strip background color after scaling
        self._apply_master_strip_background_color()

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
            strip.set_fader_value(value)
        elif typ == "mute":
            strip.set_mute_value(value >= 64)
        elif typ == "pan":
            strip.set_pan_value(value)


    def _show_settings_dialog(self):
        """Show the main settings dialog"""
        if not hasattr(self, '_settings_dialog'):
            self._settings_dialog = SettingsDialog(self.color_manager, self)
        
        self._settings_dialog.show_colors_tab()


    def mousePressEvent(self, event):
        """Handle right-click on background areas"""
        if event.button() == QtCore.Qt.RightButton:
            menu = ContextColorMenu("background", self.color_manager, self)
            menu.exec_(event.globalPos())
            return
        
        super().mousePressEvent(event)

# Main entry point
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    try:
        window = MixerWindow(CSV_DEFAULT)
        window.show()
        sys.exit(app.exec_())
    except FileNotFoundError as e:
        QtWidgets.QMessageBox.critical(None, "File Not Found", str(e))
        sys.exit(1)
    except Exception as e:
        QtWidgets.QMessageBox.critical(None, "Error", f"An error occurred: {str(e)}")
        sys.exit(1)

