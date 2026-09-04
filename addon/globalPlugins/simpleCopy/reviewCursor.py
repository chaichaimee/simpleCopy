# reviewCursor.py

import api
import textInfos
import controlTypes
import logging
import wx

logger = logging.getLogger("nvda.simpleCopy.reviewCursor")

class ReviewCursorHandler:
	def __init__(self):
		self.logger = logging.getLogger("nvda.simpleCopy.reviewCursor")

	def copy_from_review_cursor(self):
		try:
			focus_obj = api.getFocusObject()
			if not focus_obj:
				self.logger.debug("No focus object available")
				return None

			review_pos = self._open_review_on_object(focus_obj)
			if not review_pos:
				self.logger.debug("No review position available")
				return None

			all_text = self._get_all_text(review_pos, focus_obj)
			if all_text:
				self.logger.info(f"Copied all text from review cursor: {len(all_text)} chars")
				return all_text

			self.logger.warning("No text found at review cursor")
			return None

		except Exception as e:
			self.logger.error(f"Error copying from review cursor: {e}", exc_info=True)
			return None

	def copy_full_window_text(self):
		try:
			obj = api.getForegroundObject()
			if not obj:
				obj = api.getFocusObject()
				if obj:
					while obj.parent and obj.parent.role != controlTypes.Role.DESKTOP:
						obj = obj.parent

			if not obj:
				self.logger.debug("No top-level object found")
				return None

			try:
				api.setNavigatorObject(obj)
				info = obj.makeTextInfo(textInfos.POSITION_ALL)
				api.setReviewPosition(info)
			except Exception:
				pass

			text_parts = []
			try:
				info = obj.makeTextInfo(textInfos.POSITION_ALL)
				if info and not info.isCollapsed:
					raw = info.clipboardText if hasattr(info, 'clipboardText') else info.text
					if raw:
						text_parts.append(raw.strip())
			except Exception:
				pass

			if sum(len(p) for p in text_parts) < 100:
				visited_count = [0]
				self._collect_text_bounded(obj, text_parts, visited_count)

			if not text_parts:
				self.logger.debug("No text collected from window")
				return None

			full_text = "\n".join(part.strip() for part in text_parts if part.strip())
			if full_text:
				self.logger.info(f"Collected {len(full_text)} chars from window")
				return full_text

			return None

		except Exception as e:
			self.logger.error(f"Error collecting full window text: {e}", exc_info=True)
			return None

	def _collect_text_bounded(self, obj, text_parts, visited_count, max_objects=200):
		if not obj or visited_count[0] >= max_objects:
			return

		visited_count[0] += 1
		
		if visited_count[0] % 20 == 0:
			wx.YieldIfNeeded()

		name = getattr(obj, "name", "")
		value = getattr(obj, "value", "")
		description = getattr(obj, "description", "")

		if name and name not in text_parts:
			text_parts.append(name)
		if value and value != name and value not in text_parts:
			text_parts.append(value)
		if description and description not in (name, value) and description not in text_parts:
			text_parts.append(description)

		try:
			if obj.role in (controlTypes.Role.EDITABLETEXT, controlTypes.Role.TEXTFRAME):
				info = obj.makeTextInfo(textInfos.POSITION_ALL)
				if info and not info.isCollapsed:
					text = info.clipboardText if hasattr(info, 'clipboardText') else info.text
					if text and text.strip() not in (name, value, description) and text.strip() not in text_parts:
						text_parts.append(text.strip())
		except Exception:
			pass

		try:
			children = obj.children
		except Exception:
			children = []

		for child in children:
			if visited_count[0] >= max_objects:
				break
			self._collect_text_bounded(child, text_parts, visited_count, max_objects)

	def _open_review_on_object(self, focus_obj):
		try:
			info = focus_obj.makeTextInfo(textInfos.POSITION_ALL)
			api.setReviewPosition(info)
			return info
		except (RuntimeError, NotImplementedError) as e:
			self.logger.debug(f"Could not open review on focus object: {e}")
		except Exception as e:
			self.logger.warning(f"Unexpected error opening review on focus object: {e}")

		try:
			return api.getReviewPosition()
		except Exception as e:
			self.logger.debug(f"getReviewPosition fallback failed: {e}")
			return None

	def _get_all_text(self, review_pos, focus_obj):
		if review_pos:
			text = self._extract_text(review_pos, "POSITION_ALL")
			if text:
				return text

		try:
			info = review_pos.copy()
			info.expand(textInfos.UNIT_STORY)
			text = self._extract_text(info, "UNIT_STORY")
			if text:
				return text
		except Exception as e:
			self.logger.debug(f"Could not expand to UNIT_STORY: {e}")

		try:
			if focus_obj and hasattr(focus_obj, 'makeTextInfo'):
				info = focus_obj.makeTextInfo(textInfos.POSITION_ALL)
				return self._extract_text(info, "POSITION_ALL")
		except Exception as e:
			self.logger.debug(f"POSITION_ALL fallback failed: {e}")

		return None

	def _extract_text(self, text_info, method_name):
		if not text_info:
			return None
		try:
			raw_text = text_info.clipboardText if hasattr(text_info, 'clipboardText') else text_info.text
			if raw_text:
				return raw_text.replace('\r\n', '\n').replace('\r', '\n').strip()
		except Exception as e:
			self.logger.error(f"Text extraction failed using {method_name}: {e}")
		return None
