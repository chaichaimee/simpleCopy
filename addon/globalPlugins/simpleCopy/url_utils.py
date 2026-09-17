# url_utils.py

import addonHandler
addonHandler.initTranslation()

import api
import browseMode
import controlTypes
import NVDAObjects
import UIAHandler
import logHandler
import logging
import re

class URLHandler:

	# Matches any scheme://... URI, not just the previously hardcoded
	# http/https/file/ftp whitelist, so a valid URL in an unexpected
	# scheme (e.g. an internal browser scheme) is still recognized.
	_urlSchemePattern = re.compile(r'^[a-zA-Z][a-zA-Z0-9+\-.]*://\S')

	def __init__(self):
		self.logger = logging.getLogger(__name__)
		# Keep this module's logger in step with NVDA's own configured level.
		# A plain logging.getLogger() otherwise defaults to WARNING regardless
		# of NVDA's own logging-level setting.
		self.logger.setLevel(logHandler.log.getEffectiveLevel())
		self.browser_apps = ["chrome", "firefox", "edge", "msedge", "opera", "safari", "brave"]

	def is_browser_app(self, obj):
		if not obj or not obj.appModule:
			return False
		return obj.appModule.appName.lower() in self.browser_apps

	def _is_valid_url(self, url):
		if not url or not isinstance(url, str):
			return False
		return bool(self._urlSchemePattern.match(url.strip()))

	def get_current_url(self):
		# Diagnostic history from debug/info-level logging has proven unreliable
		# to surface in log.txt across several sessions, regardless of the
		# per-logger level fix above. Rather than continue depending on that,
		# every attempt below now returns a (url, failureReason) pair, and a
		# single aggregated summary is logged at WARNING level - a level
		# already confirmed to reach log.txt reliably in every prior session -
		# so the exact failure point is visible on the very next capture.
		focus = api.getFocusObject()
		if not focus:
			self.logger.warning("Could not retrieve current URL: no focus object")
			return None

		attempts = []

		url, reason = self._get_url_from_iaccessible_document(focus)
		if url:
			return url
		attempts.append(f"iaccessibleDocument[{reason}]")

		url, reason = self._get_url_from_tree_interceptor(focus)
		if url:
			return url
		attempts.append(f"treeInterceptor[{reason}]")

		url, reason = self._get_url_from_uia(focus)
		if url:
			return url
		attempts.append(f"uia[{reason}]")

		url, reason = self._get_url_from_appmodule(focus)
		if url:
			return url
		attempts.append(f"appModule[{reason}]")

		focusRole = getattr(focus, "role", None)
		appName = getattr(focus.appModule, "appName", None) if focus.appModule else None
		self.logger.warning(
			f"Could not retrieve current URL (focus role={focusRole!r}, appModule={appName!r}) :: "
			+ " | ".join(attempts)
		)
		return None

	def _get_url_from_tree_interceptor(self, focus):
		try:
			if not (hasattr(focus, "treeInterceptor") and focus.treeInterceptor):
				return None, "no active treeInterceptor"
			if not isinstance(focus.treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
				return None, f"treeInterceptor is {type(focus.treeInterceptor).__name__}, not browse-mode"
			# BrowseModeDocumentTreeInterceptor never had a `.URL` attribute; the
			# live log confirmed hasattr(..., "URL") is False on this NVDA version.
			# The real property, per NVDA's own browseMode.py, is
			# documentConstantIdentifier - explicitly documented there as tracking
			# the current URL even as it changes on single-page apps.
			if not hasattr(focus.treeInterceptor, "documentConstantIdentifier"):
				return None, "treeInterceptor has no documentConstantIdentifier attribute"
			url = focus.treeInterceptor.documentConstantIdentifier
			if self._is_valid_url(url):
				return url, None
			return None, f"treeInterceptor.documentConstantIdentifier={url!r}"
		except Exception as e:
			return None, f"exception: {e}"

	def _get_url_from_uia(self, focus):
		# Walks the ancestor chain the same way _get_url_from_iaccessible_document
		# does. UIA_UrlPropertyId is typically exposed on the document-level UIA
		# element, not on whatever descendant control happens to have focus
		# (e.g. a toolbar button deep inside a page's UI), so checking only the
		# focused object itself misses it whenever focus sits on a nested control.
		try:
			current = focus
			levelsWalked = 0
			uiaElementsSeen = 0
			lastRawUrl = None
			for _ in range(20):
				levelsWalked += 1
				if isinstance(current, NVDAObjects.UIA.UIA):
					uiaElementsSeen += 1
					try:
						lastRawUrl = current.UIAElement.getCurrentPropertyValue(UIAHandler.UIA_UrlPropertyId)
					except Exception as e_prop:
						lastRawUrl = f"<property query failed: {e_prop}>"
					if self._is_valid_url(lastRawUrl):
						return lastRawUrl, None
				if hasattr(current, "parent") and current.parent:
					current = current.parent
				else:
					break
			return None, f"no UIA Url property within {levelsWalked} level(s), {uiaElementsSeen} UIA element(s) seen, last value={lastRawUrl!r}"
		except Exception as e:
			return None, f"exception: {e}"

	def _get_url_from_iaccessible_document(self, focus):
		try:
			current = focus
			levelsWalked = 0
			# Widened from 15 to 20 ancestor levels: deeply nested single-page
			# app layouts (toolbars/menus wrapped in several ARIA group/region
			# containers) can place the focused control further from the
			# document root than a typical static page.
			for _ in range(20):
				levelsWalked += 1
				if current.role == controlTypes.Role.DOCUMENT:
					if hasattr(current, "IAccessibleObject"):
						rawUrl = current.IAccessibleObject.accValue(0)
						if self._is_valid_url(rawUrl):
							return rawUrl, None
						return None, f"document role at level {levelsWalked}, accValue={rawUrl!r}"
					return None, f"document role at level {levelsWalked} has no IAccessibleObject"
				if hasattr(current, "parent") and current.parent:
					current = current.parent
				else:
					return None, f"no document role within {levelsWalked} level(s), last role={current.role!r}, reached top of chain"
			return None, f"no document role within {levelsWalked} level(s), last role={current.role!r}"
		except Exception as e:
			return None, f"exception: {e}"

	def _get_url_from_appmodule(self, focus):
		try:
			if not (focus.appModule and hasattr(focus.appModule, "getCurrentURL")):
				return None, "appModule has no getCurrentURL method"
			url = focus.appModule.getCurrentURL()
			if self._is_valid_url(url):
				return url, None
			return None, f"appModule.getCurrentURL()={url!r}"
		except Exception as e:
			return None, f"exception: {e}"

	def get_link_url(self, obj):
		if not obj:
			return None

		if obj.role == controlTypes.Role.LINK:
			url = self._extract_link_url(obj)
			if url:
				return url

		current = obj
		for _ in range(5):
			if current.role == controlTypes.Role.LINK:
				url = self._extract_link_url(current)
				if url:
					return url
			if hasattr(current, "parent") and current.parent:
				current = current.parent
			else:
				break

		return None

	def _extract_link_url(self, link_obj):
		try:
			if hasattr(link_obj, "value") and link_obj.value:
				url = link_obj.value
				if self._is_valid_url(url):
					return url

			if hasattr(link_obj, "UIAElement"):
				url = link_obj.UIAElement.currentValue
				if self._is_valid_url(url):
					return url

			if hasattr(link_obj, "IAccessibleObject"):
				url = link_obj.IAccessibleObject.accValue(0)
				if self._is_valid_url(url):
					return url
		except Exception as e:
			self.logger.warning(f"Extract link URL failed: {e}")
		return None
