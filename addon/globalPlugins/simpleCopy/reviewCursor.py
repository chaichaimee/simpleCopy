# reviewCursor.py

import api
import textInfos
import logging

logger = logging.getLogger("nvda.simpleCopy.reviewCursor")

class ReviewCursorHandler:
	def __init__(self):
		self.logger = logging.getLogger("nvda.simpleCopy.reviewCursor")

	def copy_from_review_cursor(self):
		"""
		Enhanced Review Cursor Copy logic.
		Prioritizes copying only the selected text (POSITION_SELECTION).
		If no selection exists, falls back to copying all text (POSITION_ALL).
		"""
		try:
			review_pos = api.getReviewPosition()
			if not review_pos:
				self.logger.debug("No review position available")
				return None

			focus_obj = api.getFocusObject()
			selected_text = self._get_selected_text(focus_obj)
			if selected_text:
				self.logger.info(f"Copied selected text only: {len(selected_text)} chars")
				return selected_text

			self.logger.debug("No selection found, falling back to copying all text.")
			all_text = self._get_all_text(review_pos)
			if all_text:
				self.logger.info(f"Copied all text from review context: {len(all_text)} chars")
				return all_text

			self.logger.warning("No text found at review cursor")
			return None

		except Exception as e:
			self.logger.error(f"Error copying from review cursor: {e}", exc_info=True)
			return None

	def _get_selected_text(self, focus_obj):
		"""
		Attempts to retrieve only the currently selected text chunk.
		Uses the textInfos.POSITION_SELECTION constant which respects NVDA's virtual buffer.
		"""
		if not focus_obj:
			return None
		try:
			tree_interceptor = getattr(focus_obj, 'treeInterceptor', None)
			info_source = tree_interceptor if tree_interceptor else focus_obj

			if hasattr(info_source, 'makeTextInfo'):
				info = info_source.makeTextInfo(textInfos.POSITION_SELECTION)
				if info and not info.isCollapsed:
					return self._extract_text(info, "POSITION_SELECTION")
		except (RuntimeError, NotImplementedError) as e:
			self.logger.debug(f"POSITION_SELECTION not supported or failed: {e}")
		return None

	def _get_all_text(self, review_pos):
		"""
		Fallback method to copy all visible text in the review context.
		"""
		try:
			info = review_pos.copy()
			info.expand(textInfos.UNIT_STORY)
			return self._extract_text(info, "UNIT_STORY")
		except Exception as e:
			self.logger.debug(f"Could not expand to UNIT_STORY: {e}")

		try:
			focus_obj = api.getFocusObject()
			if hasattr(focus_obj, 'makeTextInfo'):
				info = focus_obj.makeTextInfo(textInfos.POSITION_ALL)
				return self._extract_text(info, "POSITION_ALL")
		except Exception as e:
			self.logger.debug(f"POSITION_ALL fallback failed: {e}")

		return None

	def _extract_text(self, text_info, method_name):
		"""
		Safely extracts and cleans text from a TextInfo object.
		Uses clipboardText first to preserve formatting like newlines accurately.
		"""
		if not text_info:
			return None
		try:
			raw_text = text_info.clipboardText if hasattr(text_info, 'clipboardText') else text_info.text
			if raw_text:
				return raw_text.replace('\r\n', '\n').replace('\r', '\n').strip()
		except Exception as e:
			self.logger.error(f"Text extraction failed using {method_name}: {e}")
		return None