# __init__.py
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
import core
from ui import message as ui_message
import tones
from . import url_utils
from . import clipboard_utils
from . import speech_utils
from . import reviewCursor
from . import url_history

log = logging.getLogger("nvda.simpleCopy")

addonHandler.initTranslation()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("Simple Copy")

	isTextCopied = False
	_double_tap_threshold = 0.5

	_ctrl_shift_a_tap_count = 0
	_ctrl_shift_a_last_tap_time = 0
	_ctrl_shift_a_timer = None

	_ctrl_shift_v_tap_count = 0
	_ctrl_shift_v_last_tap_time = 0
	_ctrl_shift_v_timer = None

	_f9_tap_count = 0
	_f9_last_tap_time = 0
	_f9_timer = None

	_shift_f9_tap_count = 0
	_shift_f9_last_tap_time = 0
	_shift_f9_timer = None

	_captured_speech_buffer = []
	_is_recording_active = False

	def __init__(self):
		super().__init__()
		self.clipboard_handler = clipboard_utils.ClipboardHandler()
		self.url_handler = url_utils.URLHandler()
		self.speech_history = speech_utils.SpeechHistoryHandler(callback=self._on_speech_received)
		self.review_cursor_handler = reviewCursor.ReviewCursorHandler()
		self.url_history_handler = url_history.URLHistoryManager()
		self.url_history_dialog = None
		log.info("SimpleCopy: Module initialized")

	def _is_state_echo_duplicate(self, previous_text, current_text):
		if not previous_text or not current_text:
			return False
		previous_without_checkbox_role = previous_text.replace(" check box", " ").strip()
		previous_tokens = set(previous_without_checkbox_role.lower().split())
		current_tokens = set(current_text.lower().split())
		return bool(previous_tokens) and previous_tokens.issubset(current_tokens)

	def _on_speech_received(self, text):
		clean_text = text.strip()
		if not self._is_recording_active or not clean_text:
			return

		if self._captured_speech_buffer and self._is_state_echo_duplicate(
			self._captured_speech_buffer[-1], clean_text
		):
			self._captured_speech_buffer[-1] = clean_text
		else:
			self._captured_speech_buffer.append(clean_text)

	def _performAppendAction(self, obj):
		try:
			selected_text = self.clipboard_handler.get_selected_text(obj)
			if not selected_text:
				speech.speak([_("No text selected to append")])
				return

			result = self.clipboard_handler.append_to_clipboard(selected_text)
			if result["success"]:
				self.isTextCopied = True
				appended_char_count = len(selected_text)
				total_char_count = len(result.get("fullText", selected_text))
				char_label = _("character") if appended_char_count == 1 else _("characters")
				total_label = _("character") if total_char_count == 1 else _("characters")

				if result.get("appended"):
					speech.speak([
						result["message"],
						f"{appended_char_count} {char_label}, ",
						_("total"),
						f"{total_char_count} {total_label}"
					])
				else:
					speech.speak([
						result["message"],
						f"{appended_char_count} {char_label}"
					])
			else:
				speech.speak([result["message"]])
		except Exception as e:
			log.error(f"Append action failed: {e}")
			speech.speak([_("Error during append operation")])

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
		description=_("Copy/Append selected text (single) Copy from review cursor (double) Clear clipboard (triple)"),
		gesture="kb:control+shift+v",
		category=scriptCategory
	)
	def script_handleTextCopy(self, gesture):
		current_time = time.time()
		if current_time - self._ctrl_shift_v_last_tap_time > self._double_tap_threshold:
			self._ctrl_shift_v_tap_count = 0

		self._ctrl_shift_v_tap_count += 1
		self._ctrl_shift_v_last_tap_time = current_time

		if self._ctrl_shift_v_timer and self._ctrl_shift_v_timer.IsRunning():
			self._ctrl_shift_v_timer.Stop()

		self._ctrl_shift_v_timer = core.callLater(int(self._double_tap_threshold * 1000), self._execute_v_action)

	def _execute_v_action(self):
		if self._ctrl_shift_v_tap_count == 1:
			self._handle_single_tap()
		elif self._ctrl_shift_v_tap_count == 2:
			self._handle_double_tap()
		elif self._ctrl_shift_v_tap_count >= 3:
			self._clearClipboard()
		self._ctrl_shift_v_tap_count = 0

	def _handle_single_tap(self):
		obj = api.getFocusObject()
		selected_text = self.clipboard_handler.get_selected_text(obj)
		if not selected_text:
			keyboardHandler.KeyboardInputGesture.fromName("control+shift+c").send()
		else:
			self._performAppendAction(obj)

	def _handle_double_tap(self):
		text = self.review_cursor_handler.copy_from_review_cursor()
		if text:
			if api.copyToClip(text):
				tones.beep(1500, 100)
				char_count = len(text)
				char_label = _("character") if char_count == 1 else _("characters")
				speech.speak([_("Copy from review"), f"{char_count} {char_label}"])
				log.info(f"Review cursor copied: {char_count} characters")
			else:
				tones.beep(200, 100)
				speech.speak([_("Copy failed")])
		else:
			tones.beep(200, 100)
			speech.speak([_("No text at review cursor")])
			log.warning("No text retrieved from review cursor")

	@scriptHandler.script(
		description=_("copy URL (single) copy hyper link (double) open URL history (triple)"),
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

		self._ctrl_shift_a_timer = core.callLater(int(self._double_tap_threshold * 1000), self._execute_a_action)

	def _execute_a_action(self):
		if self._ctrl_shift_a_tap_count == 1:
			self._copyBrowserUrl()
		elif self._ctrl_shift_a_tap_count == 2:
			self._copyHyperlinkUrl()
		elif self._ctrl_shift_a_tap_count >= 3:
			self.show_url_history()
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
				self.url_history_handler.add_item(url)
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
					self.url_history_handler.add_item(url)
				else:
					ui_message(_("No link found"))
			except Exception as e:
				log.error(f"Hyperlink copy error: {e}")
		else:
			keyboardHandler.KeyboardInputGesture.fromName("control+shift+a").send()

	@scriptHandler.script(
		description=_("copy last speech (single) append last speech (double) copy until last speech (triple)"),
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

		self._f9_timer = core.callLater(int(self._double_tap_threshold * 1000), self._execute_f9_action)

	def _execute_f9_action(self):
		if self._f9_tap_count == 1:
			self._handle_f9_single()
		elif self._f9_tap_count == 2:
			self._handle_f9_double()
		elif self._f9_tap_count >= 3:
			self._handle_f9_triple()
		self._f9_tap_count = 0

	def _handle_f9_single(self):
		seq = self.speech_history.get_latest_sequence()
		if seq is None:
			tones.beep(200, 100)
			return

		text_version = self.speech_history.get_latest_text()
		if api.copyToClip(text_version):
			tones.beep(1500, 100)
			self._captured_speech_buffer.clear()
			self._captured_speech_buffer.append(text_version)
			self._is_recording_active = True

	def _handle_f9_double(self):
		text = self.speech_history.get_latest_text()
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

	@scriptHandler.script(
		description=_("Navigate speech history: previous (single) next (double) open log file (triple)"),
		gesture="kb:shift+f9",
		category=scriptCategory
	)
	def script_navigateSpeechHistory(self, gesture):
		current_time = time.time()
		if current_time - self._shift_f9_last_tap_time > self._double_tap_threshold:
			self._shift_f9_tap_count = 0

		self._shift_f9_tap_count += 1
		self._shift_f9_last_tap_time = current_time

		if self._shift_f9_timer and self._shift_f9_timer.IsRunning():
			self._shift_f9_timer.Stop()

		self._shift_f9_timer = core.callLater(int(self._double_tap_threshold * 1000), self._execute_shift_f9_action)

	def _execute_shift_f9_action(self):
		if self._shift_f9_tap_count == 1:
			self._navigate_history_backward()
		elif self._shift_f9_tap_count == 2:
			self._navigate_history_forward()
		elif self._shift_f9_tap_count >= 3:
			self.speech_history.open_history_file()
		self._shift_f9_tap_count = 0

	def _navigate_history_backward(self):
		seq = self.speech_history.get_previous_sequence()
		if seq is not None:
			speech.speak(seq)
		else:
			tones.beep(200, 100)

	def _navigate_history_forward(self):
		seq = self.speech_history.get_next_sequence()
		if seq is not None:
			speech.speak(seq)
		else:
			tones.beep(200, 100)

	def show_url_history(self):
		if self.url_history_dialog and self.url_history_dialog.IsShown():
			self.url_history_dialog.Raise()
			return
		self.url_history_dialog = url_history.URLHistoryDialog(gui.mainFrame, self.url_history_handler, self)
		gui.mainFrame.prePopup()
		self.url_history_dialog.Show()
		self.url_history_dialog.CentreOnScreen()
		self.url_history_dialog.Raise()
		gui.mainFrame.postPopup()
		log.info("URL History dialog shown")

	def terminate(self):
		self.speech_history.restore_patch()
		if self.url_history_dialog:
			self.url_history_dialog.Destroy()
		if hasattr(self, "url_history_handler"):
			self.url_history_handler.save(immediate=True)
		super().terminate()