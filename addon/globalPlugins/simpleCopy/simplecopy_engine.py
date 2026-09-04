# simplecopy_engine.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import globalPluginHandler
import api
import speech
import keyboardHandler
import scriptHandler
import controlTypes
import time
import logging
import gui
import wx
import ctypes
import core
import synthDriverHandler
from ui import message as ui_message
import tones

from . import url_utils
from . import clipboard_utils
from . import speech_utils
from . import reviewCursor
from . import url_history

log = logging.getLogger("nvda.simpleCopy")

class ReviewTextDialog(wx.Dialog):
	def __init__(self, parent, text):
		super().__init__(
			parent,
			title=_("Copied Text"),
			style=wx.CAPTION | wx.CLOSE_BOX | wx.RESIZE_BORDER
		)
		self.SetSize(500, 400)
		sizer = wx.BoxSizer(wx.VERTICAL)
		self.textControl = wx.TextCtrl(
			self,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
			value=text
		)
		sizer.Add(self.textControl, proportion=1, flag=wx.EXPAND | wx.ALL, border=10)
		closeButton = wx.Button(self, label=_("Close"))
		closeButton.Bind(wx.EVT_BUTTON, lambda evt: self.Close())
		sizer.Add(closeButton, flag=wx.ALIGN_CENTER | wx.ALL, border=10)
		self.SetSizer(sizer)
		self.CentreOnScreen()
		self.Bind(wx.EVT_CHAR_HOOK, self._onCharHook)

	def _onCharHook(self, evt):
		if evt.GetKeyCode() == wx.WXK_ESCAPE:
			self.Close()
		else:
			evt.Skip()


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

	_pending_review_text = None
	_dialog_fallback_timer = None

	def __init__(self):
		super().__init__()
		self.clipboard_handler = clipboard_utils.ClipboardHandler()
		self.url_handler = url_utils.URLHandler()
		self.speech_history = speech_utils.SpeechHistoryHandler(callback=self._on_speech_received)
		self.review_cursor_handler = reviewCursor.ReviewCursorHandler()
		self.url_history_handler = url_history.URLHistoryManager()
		self.url_history_dialog = None
		self.review_text_dialog = None
		log.info("SimpleCopy: Engine initialized")

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

	def _performAppendAction(self, selected_text):
		try:
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
		description=_("Copy/Append selected text (single) Copy from review cursor and show text (double) Clear clipboard (triple)"),
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
		self.clipboard_handler.get_selected_text(obj, self._onSingleTapTextResolved)

	def _onSingleTapTextResolved(self, selected_text):
		if not selected_text:
			keyboardHandler.KeyboardInputGesture.fromName("control+shift+c").send()
			return
		self._performAppendAction(selected_text)

	def _handle_double_tap(self):
		try:
			text = self.review_cursor_handler.copy_full_window_text()
			if not text:
				text = self.review_cursor_handler.copy_from_review_cursor()

			if not text:
				tones.beep(200, 100)
				return

			if not api.copyToClip(text):
				tones.beep(200, 100)
				return

			char_count = len(text)
			line_count = text.count('\n') + 1 if text else 0

			char_label = _("character") if char_count == 1 else _("characters")
			line_label = _("line") if line_count == 1 else _("lines")

			tones.beep(1500, 100)
			speech.speak([
				_("Copied from review cursor"),
				f"{char_count} {char_label}, ",
				f"{line_count} {line_label}"
			])

			self._scheduleCopiedTextDialog(text)

		except Exception as e:
			log.error(f"Copy from review cursor failed: {e}")
			tones.beep(200, 100)

	def _scheduleCopiedTextDialog(self, text):
		"""
		Defers opening the review dialog until the synthesizer has actually
		finished speaking the character/line count announcement. A fixed
		delay is unreliable because speech duration varies with rate and
		message length; the dialog stealing focus mid-utterance flushes the
		queue via CancellableSpeech and cuts the announcement short.
		"""
		self._cancelPendingDialog()
		self._pending_review_text = text
		synthDriverHandler.synthDoneSpeaking.register(self._onCountSpeechFinished)
		self._dialog_fallback_timer = core.callLater(6000, self._onDialogFallbackTimeout)

	def _onCountSpeechFinished(self):
		try:
			synthDriverHandler.synthDoneSpeaking.unregister(self._onCountSpeechFinished)
		except Exception:
			pass
		if self._dialog_fallback_timer and self._dialog_fallback_timer.IsRunning():
			self._dialog_fallback_timer.Stop()
		self._dialog_fallback_timer = None

		text = self._pending_review_text
		self._pending_review_text = None
		if text:
			core.callLater(150, self._showCopiedTextDialog, text)

	def _onDialogFallbackTimeout(self):
		self._dialog_fallback_timer = None
		try:
			synthDriverHandler.synthDoneSpeaking.unregister(self._onCountSpeechFinished)
		except Exception:
			pass

		text = self._pending_review_text
		self._pending_review_text = None
		if text:
			self._showCopiedTextDialog(text)

	def _cancelPendingDialog(self):
		if self._dialog_fallback_timer and self._dialog_fallback_timer.IsRunning():
			self._dialog_fallback_timer.Stop()
		self._dialog_fallback_timer = None
		try:
			synthDriverHandler.synthDoneSpeaking.unregister(self._onCountSpeechFinished)
		except Exception:
			pass
		self._pending_review_text = None

	def _showCopiedTextDialog(self, text):
		try:
			if self.review_text_dialog and self.review_text_dialog.IsShown():
				self.review_text_dialog.Close()
			self.review_text_dialog = ReviewTextDialog(gui.mainFrame, text)
			self.review_text_dialog.Show()
			self.review_text_dialog.Raise()
		except Exception as e:
			log.error(f"Failed to show copied text dialog: {e}")

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
		self._cancelPendingDialog()
		self.speech_history.restore_patch()
		if self.url_history_dialog:
			self.url_history_dialog.Destroy()
		if self.review_text_dialog and self.review_text_dialog.IsShown():
			self.review_text_dialog.Close()
		if hasattr(self, "url_history_handler"):
			self.url_history_handler.save(immediate=True)
		super().terminate()
