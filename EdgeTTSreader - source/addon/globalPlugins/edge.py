import globalPluginHandler
from scriptHandler import script
import ui
import logging
import os
import asyncio
import threading
import json
import wx
import io
import tempfile
from collections import deque
from . import edge_tts
import api
from NVDAObjects import treeInterceptorHandler, textInfos
import gui
from tones import beep  # Import beep function from tones module

# Add addonHandler import for NVDA-specific requirements
import addonHandler

# Change imports according to the provided instructions
import os
import sys
import time  # Add time module for sleep function

dirAddon = os.path.dirname(__file__)
sys.path.append(dirAddon)
if sys.version.startswith("3.11"):
    sys.path.append(os.path.join(dirAddon, "_311"))
    import ctypes
    ctypes.__path__.append(os.path.join(dirAddon, "_311", "ctypes"))
else:
    sys.path.append(os.path.join(dirAddon, "_37"))
    import ctypes
    ctypes.__path__.append(os.path.join(dirAddon, "_37", "ctypes"))

os.environ['PYTHON_VLC_MODULE_PATH'] = os.path.abspath(os.path.dirname(__file__))
os.environ['PYTHON_VLC_LIB_PATH'] = os.path.abspath(os.path.join(os.path.dirname(__file__), "libvlc.dll"))
curDir = os.getcwd()
os.chdir(dirAddon)
from . import vlc
os.chdir(curDir)
del sys.path[-2:]

# Get the directory of the addon code
addon_dir = os.path.dirname(__file__)

# Construct the absolute file paths for voicelist.json, user_settings.json, and multilingual.json
voicelist_file_path = os.path.join(addon_dir, 'voicelist.json')
user_settings_file_path = os.path.join(addon_dir, 'user_settings.json')
multilingual_file_path = os.path.join(addon_dir, 'multilingual.json')

# Load options from JSON file
try:
    with open(voicelist_file_path) as f:
        voicelist_data = json.load(f)
except FileNotFoundError:
    logging.error("voicelist.json file not found. Make sure the file exists in the addon directory.")
    voicelist_data = []

# Load user settings from JSON file
try:
    with open(user_settings_file_path) as f:
        user_settings_data = json.load(f)
except FileNotFoundError:
    logging.error("user_settings.json file not found. Make sure the file exists in the addon directory.")
    user_settings_data = {"VoiceName": "", "Rate": "+100%", "IncludeExperimental": False}

# Load multilingual voices from JSON file
try:
    with open(multilingual_file_path) as f:
        multilingual_data = json.load(f)
except FileNotFoundError:
    logging.error("multilingual.json file not found. Make sure the file exists in the addon directory.")
    multilingual_data = []

# Default selected option index
selected_option_index = 0

class OptionsPanel(gui.settingsDialogs.SettingsPanel):
    title = _("EdgeTTS voice settings")

    def makeSettings(self, settingsSizer):
        sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

        # Add a label for voice selection
        voice_label = wx.StaticText(self, label=_("Neural voice"))
        sHelper.addItem(voice_label)

        # Extract option labels from JSON data
        self.option_labels = [option['VoiceName'] for option in voicelist_data]
        self.optionChoice = sHelper.addItem(
            wx.Choice(self, choices=self.option_labels)
        )

        # Add a checkbox to enable/disable experimental voices
        self.includeExperimentalCheckbox = sHelper.addItem(
            wx.CheckBox(self, label=_("Include multilingual Neural voices (experimental)"))
        )
        # Bind an event handler to the checkbox
        self.includeExperimentalCheckbox.Bind(wx.EVT_CHECKBOX, self.onIncludeExperimentalChanged)

        # Set initial state of experimental voices checkbox
        if "IncludeExperimental" in user_settings_data:
            self.includeExperimentalCheckbox.SetValue(user_settings_data["IncludeExperimental"])
        else:
            self.includeExperimentalCheckbox.SetValue(False)

        # Include experimental voices if checkbox is checked
        if self.includeExperimentalCheckbox.GetValue():
            experimental_voices = [voice["VoiceName"] for voice in multilingual_data]
            # Append experimental voices to the option labels
            self.option_labels.extend(experimental_voices)
            self.optionChoice.Set(self.option_labels)

        # Set initial value of selected voice
        selected_voice = user_settings_data.get("VoiceName")
        if selected_voice:
            try:
                selected_index = self.option_labels.index(selected_voice)
                self.optionChoice.SetSelection(selected_index)
            except ValueError:
                logging.warning("Selected voice not found in options.")

        # Add a label for speech rate selection
        rate_label = wx.StaticText(self, label=_("Speech rate"))
        sHelper.addItem(rate_label)

        # Add rate selection control with custom display values
        rate_choices = ["-20%", "-10%", "+0%", "+10%", "+20%", "+30%", "+40%", "+50%", "+60%", "+70%", "+80%", "+90%", "+100%", "+110%", "+120%", "+130%", "+140%", "+150%", "+160%", "+170%", "+180%", "+190%", "+200%"]
        self.rateChoice = sHelper.addItem(
            wx.Choice(self, choices=rate_choices)
        )
        # Set initial value of selected rate
        selected_rate = user_settings_data.get("Rate", "+100%")
        rate_index = rate_choices.index(selected_rate)
        self.rateChoice.SetSelection(rate_index)

    def onIncludeExperimentalChanged(self, event):
        user_settings_data["IncludeExperimental"] = event.IsChecked()

        if event.IsChecked():
            # Include experimental voices
            experimental_voices = [voice["VoiceName"] for voice in multilingual_data]
            # Append experimental voices to the option labels
            self.option_labels.extend(experimental_voices)
        else:
            # Exclude experimental voices
            experimental_voices = [voice["VoiceName"] for voice in multilingual_data]
            # Remove experimental voices from the option labels
            self.option_labels = [option for option in self.option_labels if option not in experimental_voices]
            # Set default voice value to 0
            user_settings_data["VoiceName"] = self.option_labels[0] if self.option_labels else ""

        # Update the choices in the choice control
        self.optionChoice.Set(self.option_labels)

        # Restore the previously selected voice
        selected_voice = user_settings_data.get("VoiceName")
        if selected_voice:
            try:
                selected_index = self.option_labels.index(selected_voice)
                self.optionChoice.SetSelection(selected_index)
            except ValueError:
                logging.warning("Selected voice not found in options.")

    def onSave(self):
        rate_choices = ["-20%", "-10%", "+0%", "+10%", "+20%", "+30%", "+40%", "+50%", "+60%", "+70%", "+80%", "+90%", "+100%", "+110%", "+120%", "+130%", "+140%", "+150%", "+160%", "+170%", "+180%", "+190%", "+200%"]
        rate_index = self.rateChoice.GetSelection()
        user_settings_data["Rate"] = rate_choices[rate_index]

        selected_option_index = self.optionChoice.GetSelection()
        if 0 <= selected_option_index < len(self.option_labels):
            user_settings_data["VoiceName"] = self.option_labels[selected_option_index]
        else:
            logging.error("Selected option index out of range.")
            ui.message("Error: Selected option not available.")

        # Save user settings to JSON file
        with open(user_settings_file_path, 'w') as f:
            json.dump(user_settings_data, f)

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

    def __init__(self):
        super(GlobalPlugin, self).__init__()
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(OptionsPanel)
        self.temp_mp3_queue = deque()  # Queue to manage temporary MP3 files
        self.mp3_buffer = io.BytesIO()  # Initialize MP3 buffer
        self.player_lock = threading.Lock()
        self.player = None
        self.stop_event = threading.Event()
        self.playback_in_progress = False  # Flag to indicate if playback is in progress

        # Initialize Python VLC player
        self.tts_player = TTSPlayer()

    async def stream_audio(self, text, voice):
        rate = user_settings_data.get("Rate", "+100%")  # Get the rate setting with default value
        text_to_speak = text
        communicate = edge_tts.Communicate(text_to_speak, voice, rate=rate)

        # Reset and truncate the main buffer to prepare for streaming
        self.mp3_buffer.seek(0)
        self.mp3_buffer.truncate(0)

        async for chunk in communicate.stream():
            if chunk["type"] == "audio":
                self.mp3_buffer.write(chunk["data"])  # Write MP3 data directly to the main buffer

        # Wait for the completion of writing to the buffer
        await asyncio.sleep(0.1)

    async def process_text(self, text):
        ui.message("Processing text...")
        selected_option = user_settings_data.get("VoiceName")
        voice = selected_option
        self.mp3_buffer.seek(0)  # Reset the in-memory buffer position

        # Start continuous beeping during text processing
        beep_thread = threading.Thread(target=self.continuous_beep, args=())
        beep_thread.start()

        await self.stream_audio(text, voice)  # Start streaming audio
        await self.prepare_next_mp3()  # Prepare the next MP3 file for playback
        await self.play_mp3()  # Call play_mp3 asynchronously

        # Stop continuous beeping after text processing is finished
        self.stop_event.set()
        beep_thread.join()
        self.stop_event.clear()

    def continuous_beep(self):
        # Beep at intervals until text processing is complete
        while not self.stop_event.is_set():
            beep(100,25)
            time.sleep(1)  # Adjust the interval to 1 second


    async def prepare_next_mp3(self):
        temp_file = tempfile.NamedTemporaryFile(suffix=".mp3", delete=False)
        temp_file_path = temp_file.name
        temp_file.write(self.mp3_buffer.getvalue())
        temp_file.close()
        self.temp_mp3_queue.append(temp_file_path)

    async def play_mp3(self):
        with self.player_lock:
            if self.playback_in_progress:
                return  # If playback is already in progress, skip

            self.playback_in_progress = True

            if self.player is None:
                self.player = self.tts_player  # Use Python VLC player

            try:
                # Check if the queue is empty
                if not self.temp_mp3_queue:
                    logging.warning("No MP3 file available for playback.")
                    return

                # Clear stop event flag
                self.stop_event.clear()

                # Start playback
                await self.player.play(self.temp_mp3_queue.popleft())

            except Exception as e:
                logging.error(f"Error during playback: {e}")

            finally:
                # Reset playback flag after playback is finished or on error
                self.playback_in_progress = False

    def cleanup_on_exit(self):
        """
        Cleanup method to be called when the addon or NVDA exits.
        """
        with self.player_lock:
            try:
                # Stop playback
                if self.player is not None:
                    self.player.stop()
                    self.player = None  # Reset player instance

                # Cleanup temporary MP3 files
                for temp_file_path in self.temp_mp3_queue:
                    if os.path.exists(temp_file_path):
                        os.unlink(temp_file_path)

            except Exception as e:
                logging.error(f"Error during cleanup: {e}")

        logging.info("Resource cleanup completed.")

    def terminate(self):
        """
        Method called when NVDA is exiting.
        """
        logging.info("Addon is exiting. Cleaning up resources.")
        self.cleanup_on_exit()
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(OptionsPanel)

    def onShutdown(self):
        """
        Method called when NVDA is shutting down.
        """
        logging.info("NVDA is shutting down. Cleaning up resources.")
        self.cleanup_on_exit()
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(OptionsPanel)

    def __del__(self):
        """
        Destructor method.
        """
        logging.info("Destructor called. Cleaning up resources.")
        self.cleanup_on_exit()

    @script(gesture="kb:NVDA+Shift+E")
    def script_readSelectedTextWithAzureVoice(self, gesture):
        obj = api.getFocusObject()
        treeInterceptor = obj.treeInterceptor

        if isinstance(treeInterceptor, treeInterceptorHandler.DocumentTreeInterceptor) and not treeInterceptor.passThrough:
            obj = treeInterceptor

        try:
            info = obj.makeTextInfo(textInfos.POSITION_SELECTION)
        except (RuntimeError, NotImplementedError):
            info = None

        if not info or info.isCollapsed:
            ui.message("No text selected.")
        else:
            selected_text = info.text
            asyncio.run(self.process_text(selected_text))  # Asynchronously process text and play audio

    @script(gesture="kb:NVDA+Shift+X")
    def script_stop_audio(self, gesture):
        with self.player_lock:
            if self.player is not None:
                self.stop_event.set()  # Signal playback stop
                self.player.stop()
                self.player = None  # Reset player instance

    @script(gesture="kb:NVDA+Shift+P")
    def script_toggle_audio(self, gesture):
        """
        Toggle between pausing and resuming the audio playback.
        """
        with self.player_lock:
            if self.player is not None:
                if self.player.get_state() == vlc.State.Playing:
                    self.player.pause()
                    ui.message("Pause")
                elif self.player.get_state() == vlc.State.Paused:
                    self.player.resume()

# Define TTSPlayer class for Python VLC integration
class TTSPlayer:
    def __init__(self):
        self.instance = vlc.Instance()
        self.player = self.instance.media_player_new()
        self.state = vlc.State.Stopped

    async def play(self, file_path):
        media = self.instance.media_new(file_path)
        self.player.set_media(media)
        self.player.play()
        while True:
            await asyncio.sleep(0.1)
            state = self.player.get_state()
            if state == vlc.State.Playing:
                self.state = state
                break

    def pause(self):
        if self.state == vlc.State.Playing:
            self.player.pause()
            self.state = vlc.State.Paused

    def resume(self):
        if self.state == vlc.State.Paused:
            self.player.play()
            self.state = vlc.State.Playing

    def stop(self):
        self.player.stop()
        self.state = vlc.State.Stopped

    def get_state(self):
        return self.state
