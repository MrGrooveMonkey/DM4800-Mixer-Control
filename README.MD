#DM4800-Mixer-Control
- Python-based control interface for the Tascam DM-4800 digital mixer

##Overview
- DM4800-Mixer-Control is a Python application that enables external control of the Tascam DM-4800 mixer’s faders, pan pots, and mute switches.
- Ir runs on MacOS and Windows. Might run on Unix but haven't tested it there.
- It provides a GUI for hands-on control via mouse/keyboard (All OS's) or touchscreen (Windows only)

##✨ Features
- GUI control via PyQt5 interface.
- Mixer fader automation — adjust channel levels programmatically or interactively.
- Pan control — modify stereo positioning remotely.
- Mute toggling — programmatically mute or unmute mixer channels.
- Stereo pair display improvements — this branch enhances the way paired channels are shown.
- MIDI communication — leverages mido with python-rtmidi as the backend.
- Editable CSVs allow custom mapping and naming for flexible integration.

#🚀 Requirements
Install Python 3.x and the following dependencies:
```pip install PyQt5 mido python-rtmidi ```

#📝 Installation & Setup
## Clone the repository
git clone -b MrGrooveMonkey-StereoPair-Display-improvements https://github.com/MrGrooveMonkey/DM4800-Mixer-Control.git
cd DM4800-Mixer-Control

## (Optional) Create virtual environment
python3 -m venv venv
source venv/bin/activate   # On Windows: venv\Scripts\activate

## Install dependencies
```pip install PyQt5 mido python-rtmidi ``` 

# Usage
### Connect the Tascam DM-4800 to your computer via MIDI (USB or DIN MIDI ports).
- Run the main script:
- python DM4800v021a.py (as of 1 Sep 2025)
- Select the MIDI In and MIDI Out ports to use from the MIDI menu in the app
### Use the GUI to:
- Adjust faders
- Pan channels
- Toggle mute
- View stereo pairs

# Module Listing
- DM4800v021a.py — Main application (latest version).
- DM4800_CC_Values.csv - CC values file that must be in the same folder where the app is run from. 

# Notes & Further Development
- The application relies on the DM-4800’s MIDI implementation for all remote control functions. It uses MIDI CC Values to control the DM4800 and to respond to changes made to Faders, Pan/Balance, and Mute on the DM4800.
- The DM4800_CC_Values.csv contains all of the MIDI CC mapping for the app to communicate with the DM4800. These should line up with your mixer settings. See MIDI CONFIG NOTES section below for more details.

# MIDI Config Notes
- Compare the values in the DM4800_CC_Values.csv file to your mixer settings located at ALT-->MIDI-->Ctrl Chg MIX tab
- Be sure to check against all the targets in the POD 4 scroll section

#Controls / GUI Layout
## The GUI is organized into a mixer window containing multiple channel strips:
| Element                 | Function                                       |
| ------------------      | ---------------------------------------        |
| **MixerWindow**         | Scrollable view and control of channel strips. |
| **Fader Slider**        | Adjusts channel level (dB scale shown).        |
| **Mono/Stereo label**   | Clearly shows channel format.                  |
| **Pan/Balance Dial**    | Sets stereo position, center in green.         |
| **Mute Button**         | Toggles mute on/off.                           |
| **Scribble Strip**      | Editable channel label text.                   |
| **Stereo Strip**        | Wide strip for stereo-paired channels.         |
| **MIDI Monitor**        | Displays live MIDI messages.                   |
| **MIDI Port Selection** |  Choose MIDI in/out ports                      |
| **Settings Menu**       |  Choose colors for various UI elements.        |

# Configuration Files                                                 
| ```DM4800_CC_Values.csv ```|  MIDI CC mapping values.                   |
| ```scribblestrip.csv    ```|  custom channel names for scribble strips. |
| ```DM4800MonoStereo.csv ```|  mono/stereo configuration data.           |
| ```dm4800_settings.json ```|  saved settings (GUI state, ports, etc.).  |
| ```DM4800_MIDI_Log.txt  ```|  log of recent MIDI messages.              |
