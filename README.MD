# DM4800-Mixer-Control
- Python-based control interface for the Tascam DM-4800 digital mixer

## Overview
- DM4800-Mixer-Control is a Python application that enables external control of the Tascam DM-4800 mixer’s faders, pan pots, and mute switches.
- Ir runs on MacOS and Windows. Should run on Unix but haven't tested it there.
- It provides a GUI for hands-on control via mouse/keyboard (All OS's) or touchscreen (Windows only)

## ✨ Features
- Fader control — adjust channel levels.
- Pan control — modify stereo positioning.
- Mute control —  mute or unmute mixer channels.
- Stereo pair display — enhances the way paired channels are shown.
- MIDI communication — uses MIDI CC commands to control the DM4800 and to receive updates to Faders, Pan/Balance knobs, and Mutes from the DM4800.
- Editable CSVs - allow custom mapping and naming for stereo channel pairs and scribble strips.
- Color customization - change the colors of various UI elements such as Fader caps, Pan/Balance knobs, and application background.

## 📷 Screenshots
![Screenshot 1](https://github.com/MrGrooveMonkey/DM4800-Mixer-Control/blob/820b3f0feadbc2fae3c5d5a1070d270938bce256/4800mixer1.jpeg)
![Screenshot 2](https://github.com/MrGrooveMonkey/DM4800-Mixer-Control/blob/fd778ac7b03ea1913dfbe2f545f97e8c0f6f6d2b/4800mixer2.jpeg)
![Screenshot 3](https://github.com/MrGrooveMonkey/DM4800-Mixer-Control/blob/820b3f0feadbc2fae3c5d5a1070d270938bce256/4800mixer3.jpeg)

# 🚀 Requirements
Install Python 3.x and the following dependencies:
```pip install PyQt5 mido python-rtmidi ```

# 📝 Installation & Setup
## Clone the repository
git clone -b MrGrooveMonkey-StereoPair-Display-improvements https://github.com/MrGrooveMonkey/DM4800-Mixer-Control.git
cd DM4800-Mixer-Control

## (Optional) Create virtual environment
python3 -m venv venv
source venv/bin/activate   # On Windows: venv\Scripts\activate

## Install dependencies
```pip install PyQt5 mido python-rtmidi ``` 

# Usage
## **NOTE:** 
- There is no way that I have found to query the DM4800 for the status of stereo pairings, current values of the faders, pan/bal, or mutes. Therefore we need to update this info manually.
### Update the DM4800MonoStereo.csv file to match your DM4800 mono/stereo pairings. 
- The format should be evident when you look at the file so just stick to the format and the app will recognize it. 
- If you make changes to stereo relationships on the DM4800 be sure to update this file and restart the app.
### Update the scribblestrip.csv file if you want to preload the scribble strips under the channels, buses, and auxes.
- This is an optional feature I added to the app to give you scribble strips. 
- You can edit these in the app and they will be saved to the file automatically.
### Enusre the MIDI CONFIG matches your DM4800 
- See the Notes sections below for more details.
### Connect the Tascam DM-4800 to your computer via MIDI (USB or DIN MIDI ports).
- Run the main script:
-- python DM4800v0.9.0.py (as of 2 Sep 2025). 
- Accept the license agreement (MIT License which is 2 short paragraphs)
- Select the MIDI In and MIDI Out ports to use from the MIDI menu in the app.
- You should now be able to control the DM4800 and have the app update itself any time you adjust the Faders, Pan/Balance knobs, and Mutes on the DM4800.
### Use the GUI to:
- Adjust faders
- Pan mono channels
- Balance stereo channels
- Toggle mute
- View stereo pairs
- Customize colors

# Module Listing
- DM4800v0.9.0.py — Main application (latest version).
- DM4800_CC_Values.csv - CC values file that must be in the same folder where the app is run from. 

# Notes & Further Development
- The application relies on the DM4800’s MIDI implementation for all remote control functions. It uses MIDI CC Values to control the DM4800 and to respond to changes made to Faders, Pan/Balance, and Mute on the DM4800.
- The DM4800_CC_Values.csv contains all of the MIDI CC mapping for the app to communicate with the DM4800. These should line up with your mixer settings. See MIDI CONFIG NOTES section below for more details.

# MIDI Config Notes
- Compare the values in the DM4800_CC_Values.csv file to your mixer settings located at ALT-->MIDI-->Ctrl Chg MIX tab
- Be sure to check against all the targets in the POD 4 scroll section

# Controls / GUI Layout
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
| ```DM4800_CC_Values.csv```|  MIDI CC mapping values.                   |
| ```scribblestrip.csv```   |  custom channel names for scribble strips. |
| ```DM4800MonoStereo.csv```|  mono/stereo configuration data.           |
| ```dm4800_settings.json```|  saved settings (GUI state, ports, etc.).  |
| ```DM4800_MIDI_Log.txt ```|  log of recent MIDI messages.              |
| ```License.txt```         |  License info file.                        |
| ```DM4800_license.txt```  |  License acceptance record.                |
