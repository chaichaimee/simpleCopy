# speech_utils.py

import speech
import speechViewer
import os
import globalVars
from collections import deque
import logging

log = logging.getLogger("nvda.simpleCopy.speech")

class SpeechHistoryHandler:
	def __init__(self, maxlen=500, callback=None):
		self.history = deque(maxlen=maxlen)
		self.callback = callback
		self._orig_speak = None
		self._patched = False
		self._setup_storage()
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

	def _my_speak(self, sequence, *args, **kwargs):
		if self._orig_speak:
			self._orig_speak(sequence, *args, **kwargs)
		
		# Extract string items and join with standard separator
		text_parts = [str(item) for item in sequence if isinstance(item, str)]
		text = speechViewer.SPEECH_ITEM_SEPARATOR.join(text_parts)
		
		if text.strip():
			self.history.appendleft(text)
			if self.callback:
				self.callback(text)

	def get_latest(self):
		return self.history[0] if self.history else ""

	def open_history_file(self):
		lines = list(reversed(self.history))
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