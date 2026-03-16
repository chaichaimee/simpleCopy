# init.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import globalPluginHandler
import api
import speech
import keyboardHandler
import scriptHandler
import addonHandler
import controlTypes
import textInfos
import time
import logging
import gui
import wx
import ctypes
from ui import message as ui_message
import tones
from . import url_utils
from . import clipboard_utils
from . import speech_utils

log = logging.getLogger("nvda.simpleCopy")

addonHandler.initTranslation()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("Simple Copy")
	
	isTextCopied = False
	_double_tap_threshold = 0.5
	
	_ctrl_shift_a_tap_count = 0
	_ctrl_shift_a_last_tap_time = 0
	_ctrl_shift_a_timer = None
	
	_ctrl_shift_c_tap_count = 0
	_ctrl_shift_c_last_tap_time = 0
	_ctrl_shift_c_timer = None
	
	_f9_tap_count = 0
	_f9_last_tap_time = 0
	_f9_timer = None
	
	_captured_speech_buffer = []
	_is_recording_active = False
	
	def __init__(self):
		super().__init__()
		self.clipboard_handler = clipboard_utils.ClipboardHandler()
		self.url_handler = url_utils.URLHandler()
		self.speech_history = speech_utils.SpeechHistoryHandler(callback=self._on_speech_received)
		log.info("SimpleCopy: Module initialized")

	def _on_speech_received(self, text):
		if self._is_recording_active and text.strip():
			self._captured_speech_buffer.append(text)

	def _performAppendAction(self, obj):
		try:
			text_to_append = self.clipboard_handler.get_selected_text(obj)
			if not text_to_append:
				speech.speak([_("No text selected to append")])
				return
				
			result = self.clipboard_handler.append_to_clipboard(text_to_append)
			if result["success"]:
				self.isTextCopied = True
				speech.speak([_(result["message"])])
			else:
				speech.speak([_(result["message"])])
		except Exception as e:
			log.error(f"Append action failed: {e}")
			speech.speak([_("Error during append operation")])

	@scriptHandler.script(
		description=_("copy URL (single) copy hyper link (double)"),
		gesture="kb:control+shift+a",
		category=scriptCategory
	)
	def script_copyUrlOrHyperlink(self, gesture):
		current_time = time.time()
		if current_time - self._ctrl_shift_a_last_tap_time > self._double_tap_threshold:
			self._ctrl_shift_a_tap_count = 0
		
		self._ctrl_shift_a_tap_count += 1
		self._ctrl_shift_a_last_tap_time = current_time
		
		if self._ctrl_shift_a_timer and self._ctrl_shift_a_timer.IsRunning():
			self._ctrl_shift_a_timer.Stop()
		
		self._ctrl_shift_a_timer = wx.CallLater(int(self._double_tap_threshold * 1000), self._execute_a_action)
	
	def _execute_a_action(self):
		if self._ctrl_shift_a_tap_count == 1:
			self._copyBrowserUrl()
		elif self._ctrl_shift_a_tap_count >= 2:
			self._copyHyperlinkUrl()
		self._ctrl_shift_a_tap_count = 0

	def _copyBrowserUrl(self):
		obj = api.getFocusObject()
		if (obj.role in (controlTypes.Role.EDITABLETEXT, controlTypes.Role.TEXTFRAME) or controlTypes.State.EDITABLE in obj.states):
			keyboardHandler.KeyboardInputGesture.fromName("control+shift+a").send()
			return
		
		if self.url_handler.is_browser_app(obj):
			url = self.url_handler.get_current_url()
			if url and api.copyToClip(url):
				speech.speak([_("Copy"), url])
			else:
				ui_message(_("No URL"))
		else:
			keyboardHandler.KeyboardInputGesture.fromName("control+shift+a").send()

	def _copyHyperlinkUrl(self):
		obj = api.getNavigatorObject()
		if self.url_handler.is_browser_app(obj):
			try:
				url = self.url_handler.get_link_url(obj)
				if url and api.copyToClip(url):
					speech.speak([_("Copy"), url])
				else:
					ui_message(_("No link found"))
			except Exception as e:
				log.error(f"Hyperlink copy error: {e}")
		else:
			keyboardHandler.KeyboardInputGesture.fromName("control+shift+a").send()

	@scriptHandler.script(
		description=_("copy and append (single) clear (double)"),
		gesture="kb:control+shift+c",
		category=scriptCategory
	)
	def script_appendOrClear(self, gesture):
		current_time = time.time()
		if current_time - self._ctrl_shift_c_last_tap_time > self._double_tap_threshold:
			self._ctrl_shift_c_tap_count = 0
		
		self._ctrl_shift_c_tap_count += 1
		self._ctrl_shift_c_last_tap_time = current_time
		
		if self._ctrl_shift_c_timer and self._ctrl_shift_c_timer.IsRunning():
			self._ctrl_shift_c_timer.Stop()
		
		self._ctrl_shift_c_timer = wx.CallLater(int(self._double_tap_threshold * 1000), self._execute_c_action)

	def _execute_c_action(self):
		if self._ctrl_shift_c_tap_count == 1:
			obj = api.getFocusObject()
			if not self.clipboard_handler.get_selected_text(obj):
				keyboardHandler.KeyboardInputGesture.fromName("control+shift+c").send()
			else:
				self._performAppendAction(obj)
		elif self._ctrl_shift_c_tap_count >= 2:
			self._clearClipboard()
		self._ctrl_shift_c_tap_count = 0

	def _clearClipboard(self):
		try:
			user32 = ctypes.windll.user32
			if user32.OpenClipboard(None):
				user32.EmptyClipboard()
				user32.CloseClipboard()
				self.isTextCopied = False
				self._captured_speech_buffer.clear()
				self._is_recording_active = False
				speech.speak([_("Clean")])
			else:
				tones.beep(200, 100)
		except Exception as e:
			log.error(f"Clipboard clear failed: {e}")
			tones.beep(200, 100)

	@scriptHandler.script(
		description=_("copy last speech (single) append last speech (double) copy until last speech (triple) open record speech (quadruple)"),
		gesture="kb:f9",
		category=scriptCategory
	)
	def script_copySpeech(self, gesture):
		current_time = time.time()
		if current_time - self._f9_last_tap_time > self._double_tap_threshold:
			self._f9_tap_count = 0
		self._f9_tap_count += 1
		self._f9_last_tap_time = current_time
		
		if self._f9_timer and self._f9_timer.IsRunning():
			self._f9_timer.Stop()
		self._f9_timer = wx.CallLater(int(self._double_tap_threshold * 1000), self._execute_f9_action)

	def _execute_f9_action(self):
		if self._f9_tap_count == 1:
			self._handle_f9_single()
		elif self._f9_tap_count == 2:
			self._handle_f9_double()
		elif self._f9_tap_count == 3:
			self._handle_f9_triple()
		elif self._f9_tap_count >= 4:
			self.speech_history.open_history_file()
		self._f9_tap_count = 0

	def _handle_f9_single(self):
		text = self.speech_history.get_latest()
		if not text:
			tones.beep(200, 100)
			return
		if api.copyToClip(text):
			tones.beep(1500, 100)
			# Start accumulation for triple-tap
			self._captured_speech_buffer.clear()
			self._captured_speech_buffer.append(text)
			self._is_recording_active = True

	def _handle_f9_double(self):
		text = self.speech_history.get_latest()
		if not text:
			tones.beep(200, 100)
			return
		if self.clipboard_handler.append_text_silent(text):
			speech.speak([_("Append")])

	def _handle_f9_triple(self):
		if not self._is_recording_active or not self._captured_speech_buffer:
			tones.beep(200, 100)
			return
		
		combined_text = "\n".join(self._captured_speech_buffer)
		if api.copyToClip(combined_text):
			speech.speak([_("Copy Until Last")])
			self._captured_speech_buffer.clear()
			self._is_recording_active = False

	def terminate(self):
		self.speech_history.restore_patch()
		super().terminate()