# speech_utils.py

import speech
import speechViewer
import os
import globalVars
from collections import deque, namedtuple
import logging
from eventHandler import FocusLossCancellableSpeechCommand

log = logging.getLogger("nvda.simpleCopy.speech")

# Snapshot of one intercepted speak() call.
# text is extracted immediately inside _my_speak, the same moment record.txt
# extracts it, so later reads never depend on the original sequence object
# still being intact (NVDA may reuse/clear that list once speak() returns).
SpeechHistoryEntry = namedtuple("SpeechHistoryEntry", ["sequence", "text"])


class SpeechNavigator:
	def __init__(self, history_deque):
		self.history = history_deque
		self.current_index = -1

	def reset(self):
		self.current_index = -1

	def move_backward(self):
		target_index = self.current_index + 1
		if target_index < len(self.history):
			self.current_index = target_index
			return self.history[self.current_index]
		return None

	def move_forward(self):
		target_index = self.current_index - 1
		if target_index >= -1:
			self.current_index = target_index
			if self.current_index == -1:
				return None
			return self.history[self.current_index]
		return None

	def has_previous(self):
		return (self.current_index + 1) < len(self.history)

	def has_next(self):
		return self.current_index > -1


class SpeechHistoryHandler:
	def __init__(self, maxlen=500, callback=None):
		self.history = deque(maxlen=maxlen)
		self.callback = callback
		self._orig_speak = None
		self._patched = False
		self._setup_storage()
		self.navigator = SpeechNavigator(self.history)
		self.patch_speech()

	def _setup_storage(self):
		config_path = globalVars.appArgs.configPath
		self.history_dir = os.path.join(config_path, "ChaiChaimee", "simpleCopy")
		self.history_file = os.path.join(self.history_dir, "speech_log.txt")
		if not os.path.exists(self.history_dir):
			os.makedirs(self.history_dir)

		try:
			with open(self.history_file, "w", encoding="utf-8") as f:
				f.write("")
		except Exception as e:
			log.error(f"Failed to initialize log file: {e}")

	def patch_speech(self):
		try:
			if hasattr(speech, 'speech') and hasattr(speech.speech, 'speak'):
				self._orig_speak = speech.speech.speak
				speech.speech.speak = self._my_speak
				self._patched = True
			elif hasattr(speech, 'speak'):
				self._orig_speak = speech.speak
				speech.speak = self._my_speak
				self._patched = True
			log.info("Speech interception active")
		except Exception as e:
			log.error(f"Failed to patch speech: {e}")

	def restore_patch(self):
		if self._patched:
			if hasattr(speech, 'speech'):
				speech.speech.speak = self._orig_speak
			else:
				speech.speak = self._orig_speak
			self._patched = False
			log.info("Speech interception restored")

	def _get_sequence_text(self, seq):
		if not seq:
			return ""

		if isinstance(seq, str):
			return seq

		if isinstance(seq, (list, tuple)):
			valid_commands = [
				command for command in seq
				if not isinstance(command, FocusLossCancellableSpeechCommand)
			]
			return speechViewer.SPEECH_ITEM_SEPARATOR.join(
				[item for item in valid_commands if isinstance(item, str)]
			)

		return ""

	def _my_speak(self, sequence, *args, **kwargs):
		if self._orig_speak:
			self._orig_speak(sequence, *args, **kwargs)

		try:
			# Extract the full text right now, while sequence is still the
			# exact list NVDA just built for this utterance. Once speak()
			# returns, NVDA is free to reuse/clear/mutate that list, so any
			# extraction done later (e.g. when the user presses F9) can
			# silently come back short. Freezing the string here - the same
			# moment record.txt does it - is what guarantees the full line.
			text_version = self._get_sequence_text(sequence)
			if not text_version:
				return

			clean_text = text_version.strip()
			if not clean_text:
				return

			self.history.appendleft(SpeechHistoryEntry(sequence=sequence, text=clean_text))

			if self.callback:
				callback_text = clean_text
				if "\n" not in callback_text:
					callback_text += "\n"
				self.callback(callback_text)
		except Exception as e:
			log.error(f"MySpeak processing error: {e}")

	def get_latest_sequence(self):
		return self.history[0].sequence if self.history else None

	def get_latest_text(self):
		# Return the text frozen at capture time - never re-derive it from
		# the stored sequence, since that sequence may no longer hold the
		# same content it did when it was spoken.
		return self.history[0].text if self.history else ""

	def open_history_file(self):
		lines = []
		for entry in self.history:
			if entry.text:
				lines.append(entry.text)

		seen = set()
		unique_lines = []
		for line in lines:
			if line and line not in seen:
				seen.add(line)
				unique_lines.append(line)

		try:
			with open(self.history_file, "w", encoding="utf-8") as f:
				f.write("\n".join(unique_lines))
			os.startfile(self.history_file)
		except Exception as e:
			log.error(f"File access error: {e}")

	def get_previous_sequence(self):
		if not self.history or not self.navigator.has_previous():
			return None
		entry = self.navigator.move_backward()
		return entry.sequence if entry else None

	def get_next_sequence(self):
		if not self.history or not self.navigator.has_next():
			return None
		entry = self.navigator.move_forward()
		return entry.sequence if entry else None

	def reset_navigation(self):
		self.navigator.reset()
