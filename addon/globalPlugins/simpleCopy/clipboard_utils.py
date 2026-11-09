# clipboard_utils.py

import addonHandler
addonHandler.initTranslation()

import api
import winUser
import gui
import core
import keyboardHandler
import textInfos
import browseMode
import logging
import hashlib
import sys


class ClipboardHandler:

	def __init__(self):
		self.logger = logging.getLogger(__name__)

	def normalize_text(self, text):
		if not text:
			return ""
		text = "".join(char for char in text if char.isprintable() or char in {"\r", "\n", " "})
		text = text.replace("\r\n", "\n").replace("\r", "\n")
		return text

	def calculate_sha256(self, text):
		normalized_text = self.normalize_text(text)
		return hashlib.sha256(normalized_text.encode('utf-8')).hexdigest()

	def get_selected_text(self, obj_param, callback):
		if sys.version_info >= (3, 13):
			self._get_selected_text_2026(obj_param, callback)
		else:
			self._get_selected_text_2025(obj_param, callback)

	def _try_make_text_info(self, obj_param):
		try:
			target_obj_for_text = None
			if hasattr(obj_param, 'treeInterceptor') and isinstance(obj_param.treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
				target_obj_for_text = obj_param.treeInterceptor
				self.logger.info("Using treeInterceptor for text info.")
			elif hasattr(obj_param, 'makeTextInfo'):
				target_obj_for_text = obj_param

			if not target_obj_for_text:
				self.logger.info("No makeTextInfo available.")
				return None

			info = target_obj_for_text.makeTextInfo(textInfos.POSITION_SELECTION)
			if info and not info.isCollapsed:
				selected_text = info.clipboardText if hasattr(info, 'clipboardText') else info.text
				if selected_text:
					self.logger.info(f"makeTextInfo retrieved: {repr(selected_text[:50])}...")
					return selected_text.replace('\r\n', '\n').replace('\r', '\n').strip()
		except (RuntimeError, NotImplementedError) as e:
			self.logger.warning(f"makeTextInfo selection failed: {str(e)}")
		except Exception as e_info:
			self.logger.error(f"Error with makeTextInfo attempt: {str(e_info)}")
		return None

	def _try_tree_interceptor_selection(self, obj_param):
		try:
			if hasattr(obj_param, 'treeInterceptor') and obj_param.treeInterceptor:
				treeInterceptor = obj_param.treeInterceptor
				if hasattr(treeInterceptor, 'selection'):
					selection = treeInterceptor.selection
					if selection and hasattr(selection, 'isCollapsed') and not selection.isCollapsed:
						selected_text = None
						if hasattr(selection, 'clipboardText'):
							selected_text = selection.clipboardText
						elif hasattr(selection, 'text'):
							selected_text = selection.text
						if selected_text:
							self.logger.info(f"treeInterceptor.selection.clipboardText/text success: {repr(selected_text[:50])}")
							return selected_text.replace('\r\n', '\n').replace('\r', '\n').strip()
		except Exception as e:
			self.logger.warning(f"treeInterceptor selection failed: {e}")
		return None

	def _get_selected_text_2025(self, obj_param, callback):
		self.logger.info("SimpleCopy: get_selected_text started (2025).")
		directText = self._try_make_text_info(obj_param)
		if directText:
			callback(directText)
			return

		self.logger.info("makeTextInfo failed, attempting Ctrl+C fallback.")
		self._startCtrlCFallback(callback, maxAttempts=1, retryDelayMs=50)

	def _get_selected_text_2026(self, obj_param, callback):
		self.logger.info("SimpleCopy: get_selected_text started (2026).")
		selectionText = self._try_tree_interceptor_selection(obj_param)
		if selectionText:
			callback(selectionText)
			return

		directText = self._try_make_text_info(obj_param)
		if directText:
			callback(directText)
			return

		self.logger.info("Falling back to Ctrl+C method (2026).")
		self._startCtrlCFallback(callback, maxAttempts=3, retryDelayMs=100)

	def _startCtrlCFallback(self, callback, maxAttempts, retryDelayMs):
		originalClipboardData = ""
		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				originalClipboardData = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
				winUser.emptyClipboard()
		except Exception as e_snapshot:
			self.logger.error(f"Error snapshotting clipboard: {str(e_snapshot)}")

		keyboardHandler.KeyboardInputGesture.fromName("control+c").send()
		core.callLater(
			retryDelayMs, self._checkCtrlCResult,
			originalClipboardData, callback, 1, maxAttempts, retryDelayMs
		)

	def _checkCtrlCResult(self, originalClipboardData, callback, attempt, maxAttempts, retryDelayMs):
		clipboardText = ""
		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				clipboardText = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
		except Exception as e_fallback:
			self.logger.warning(f"Ctrl+C fallback attempt {attempt} failed: {str(e_fallback)}")

		if clipboardText:
			self.logger.info(f"Ctrl+C fallback attempt {attempt} retrieved: {repr(clipboardText[:50])}...")
			self._restoreClipboard(originalClipboardData)
			callback(clipboardText.replace('\r\n', '\n').replace('\r', '\n').strip())
			return

		if attempt < maxAttempts:
			core.callLater(
				retryDelayMs, self._checkCtrlCResult,
				originalClipboardData, callback, attempt + 1, maxAttempts, retryDelayMs
			)
			return

		self.logger.info("No selected text found after Ctrl+C fallback.")
		self._restoreClipboard(originalClipboardData)
		callback(None)

	def _restoreClipboard(self, originalClipboardData):
		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				winUser.emptyClipboard()
				if originalClipboardData:
					winUser.setClipboardData(winUser.CF_UNICODETEXT, originalClipboardData)
		except Exception as e_restore:
			self.logger.warning(f"Failed to restore clipboard: {str(e_restore)}")

	def append_to_clipboard(self, text_to_append):
		clipData = ""
		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				clipData = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
				if clipData and not isinstance(clipData, str):
					return {
						"success": False,
						"appended": False,
						"message": _("Cannot append to non-text clipboard content"),
						"fullText": ""
					}
		except Exception as e:
			self.logger.error(f"Error reading clipboard: {str(e)}")
			clipData = ""

		processed_text_to_append = text_to_append

		if clipData:
			clipData_normalized = clipData.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n')
			processed_text_to_append_normalized = processed_text_to_append.replace('\r\n', '\n').replace('\r', '\n').lstrip('\n')
			newText = clipData_normalized + "\n" + processed_text_to_append_normalized
			appended = True
		else:
			newText = processed_text_to_append.replace('\r\n', '\n').replace('\r', '\n')
			appended = False

		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				winUser.emptyClipboard()
				winUser.setClipboardData(winUser.CF_UNICODETEXT, newText.replace('\n', '\r\n'))
			return {
				"success": True,
				"appended": appended,
				"message": _("Appended") if appended else _("Copied"),
				"fullText": newText
			}
		except Exception as e:
			self.logger.error(f"Error writing to clipboard: {str(e)}")
			return {
				"success": False,
				"appended": False,
				"message": _("Error writing to clipboard"),
				"fullText": ""
			}

	def read_text(self):
		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				clipData = winUser.getClipboardData(winUser.CF_UNICODETEXT)
		except OSError as e:
			self.logger.warning(f"read_text: clipboard unavailable: {e}")
			return None
		return clipData if isinstance(clipData, str) else ""

	def write_text(self, text):
		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				winUser.emptyClipboard()
				winUser.setClipboardData(winUser.CF_UNICODETEXT, text)
			return True
		except OSError as e:
			self.logger.error(f"write_text: failed: {e}")
			return False

	def join_appended_text(self, existingText, newText):
		newNormalized = newText.replace("\r\n", "\n").replace("\r", "\n")
		if not existingText:
			return newNormalized.replace("\n", "\r\n")
		existingNormalized = existingText.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
		return (existingNormalized + "\n" + newNormalized.lstrip("\n")).replace("\n", "\r\n")

	def split_paste_lines(self, text):
		if not text:
			return []
		normalized = text.replace("\r\n", "\n").replace("\r", "\n")
		return [line for line in normalized.split("\n") if line.strip()]

	def append_text_silent(self, text_to_append):
		clipData = ""
		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				clipData = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
				if clipData and not isinstance(clipData, str):
					self.logger.warning("Clipboard contains non-text data, cannot append")
					return False
		except Exception as e:
			self.logger.error(f"append_text_silent: Error reading clipboard: {e}")
			clipData = ""

		processed_text_to_append = text_to_append

		if clipData:
			clipData_normalized = clipData.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n')
			processed_text_to_append_normalized = processed_text_to_append.replace('\r\n', '\n').replace('\r', '\n').lstrip('\n')
			newText = clipData_normalized + "\n" + processed_text_to_append_normalized
		else:
			newText = processed_text_to_append.replace('\r\n', '\n').replace('\r', '\n')

		try:
			with winUser.openClipboard(gui.mainFrame.Handle):
				winUser.emptyClipboard()
				winUser.setClipboardData(winUser.CF_UNICODETEXT, newText.replace('\n', '\r\n'))
			return True
		except Exception as e:
			self.logger.error(f"append_text_silent: Error writing to clipboard: {e}")
			return False


