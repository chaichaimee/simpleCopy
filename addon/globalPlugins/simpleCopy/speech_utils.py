# speech_utils.py

import speech
import speechViewer
import os
import globalVars
from collections import deque
import logging
from eventHandler import FocusLossCancellableSpeechCommand

log = logging.getLogger("nvda.simpleCopy.speech")


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

	def _clean_sequence(self, seq):
		"""
		Remove FocusLossCancellableSpeechCommand and keep only strings.
		Returns a new list that can be stored and spoken later.
		"""
		if isinstance(seq, (list, tuple)):
			return [item for item in seq if not isinstance(item, FocusLossCancellableSpeechCommand)]
		return seq

	def _sequence_to_text(self, seq):
		"""Convert a sequence (list/tuple of strings) to a single string for clipboard/file."""
		if not seq:
			return ""
		if isinstance(seq, str):
			return seq
		if isinstance(seq, (list, tuple)):
			strings = [item for item in seq if isinstance(item, str)]
			return speechViewer.SPEECH_ITEM_SEPARATOR.join(strings)
		return ""

	def _my_speak(self, sequence, *args, **kwargs):
		if self._orig_speak:
			self._orig_speak(sequence, *args, **kwargs)

		cleaned_seq = self._clean_sequence(sequence)
		if not cleaned_seq:
			return

		self.history.appendleft(cleaned_seq)
		if self.callback:
			text_version = self._sequence_to_text(cleaned_seq)
			if text_version:
				self.callback(text_version)

	def get_latest_sequence(self):
		"""Return the raw sequence of the most recent speech."""
		return self.history[0] if self.history else None

	def get_latest_text(self):
		"""Return the string version of the most recent speech (for clipboard)."""
		seq = self.get_latest_sequence()
		if seq is None:
			return ""
		return self._sequence_to_text(seq)

	def open_history_file(self):
		lines = []
		for seq in reversed(self.history):
			text_line = self._sequence_to_text(seq)
			if text_line.strip():
				lines.append(text_line)

		# Remove duplicates while preserving order
		seen = set()
		unique_lines = []
		for line in lines:
			clean = line.strip()
			if clean and clean not in seen:
				seen.add(clean)
				unique_lines.append(line)

		try:
			with open(self.history_file, "w", encoding="utf-8") as f:
				f.write("\n".join(unique_lines))
			os.startfile(self.history_file)
		except Exception as e:
			log.error(f"File access error: {e}")

	def get_previous_sequence(self):
		"""Return the raw sequence of the previous speech item."""
		if not self.history:
			return None
		if not self.navigator.has_previous():
			return None
		return self.navigator.move_backward()

	def get_next_sequence(self):
		"""Return the raw sequence of the next speech item."""
		if not self.history:
			return None
		if not self.navigator.has_next():
			return None
		return self.navigator.move_forward()

	def reset_navigation(self):
		self.navigator.reset()