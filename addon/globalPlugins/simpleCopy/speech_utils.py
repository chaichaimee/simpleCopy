# speech_utils.py

import speech
import speechViewer
from speech.extensions import filter_speechSequence
import os
import globalVars
from collections import deque, namedtuple
import logging
from eventHandler import FocusLossCancellableSpeechCommand

log = logging.getLogger("nvda.simpleCopy.speech")

SpeechHistoryEntry = namedtuple("SpeechHistoryEntry", ["sequence", "text"])

_unrecordedSpeechDepth = 0


def speakUnrecorded(sequence, *args, **kwargs):
	global _unrecordedSpeechDepth
	_unrecordedSpeechDepth += 1
	try:
		speech.speak(sequence, *args, **kwargs)
	finally:
		_unrecordedSpeechDepth -= 1


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
		self._filterRegistered = False
		self._setup_storage()
		self.navigator = SpeechNavigator(self.history)
		self.register_filter()

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

	def register_filter(self):
		if self._filterRegistered:
			return
		filter_speechSequence.register(self._recordSequence)
		self._filterRegistered = True
		log.info("Speech interception active")

	def unregister_filter(self):
		if not self._filterRegistered:
			return
		filter_speechSequence.unregister(self._recordSequence)
		self._filterRegistered = False
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

	def _get_replayable_sequence(self, seq):
		if not isinstance(seq, (list, tuple)):
			return seq
		return [
			command for command in seq
			if not isinstance(command, FocusLossCancellableSpeechCommand)
		]

	def _is_state_echo_duplicate(self, previous_text, current_text):
		if not previous_text or not current_text:
			return False
		previous_without_checkbox_role = previous_text.replace(" check box", " ").strip()
		previous_tokens = set(previous_without_checkbox_role.lower().split())
		current_tokens = set(current_text.lower().split())
		return bool(previous_tokens) and previous_tokens.issubset(current_tokens)

	def _recordSequence(self, sequence):
		if _unrecordedSpeechDepth > 0:
			return sequence

		try:
			text_version = self._get_sequence_text(sequence)
			if not text_version:
				return sequence

			clean_text = text_version.strip()
			if not clean_text:
				return sequence

			replaySequence = self._get_replayable_sequence(sequence)
			if self.history and self._is_state_echo_duplicate(self.history[0].text, clean_text):
				self.history[0] = SpeechHistoryEntry(sequence=replaySequence, text=clean_text)
			else:
				self.history.appendleft(SpeechHistoryEntry(sequence=replaySequence, text=clean_text))

			if self.callback:
				callback_text = clean_text
				if "\n" not in callback_text:
					callback_text += "\n"
				self.callback(callback_text)
		except Exception as e:
			log.error(f"Speech record error: {e}")
		return sequence

	def get_latest_sequence(self):
		return self.history[0].sequence if self.history else None

	def get_latest_text(self):
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

