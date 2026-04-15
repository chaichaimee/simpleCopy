# url_utils.py

import addonHandler
addonHandler.initTranslation()

import api
import browseMode
import controlTypes
import NVDAObjects
import UIAHandler
import logging
import sys

class URLHandler:
	
	def __init__(self):
		self.logger = logging.getLogger(__name__)
		self.browser_apps = ["chrome", "firefox", "edge", "msedge", "opera", "safari", "brave"]
	
	def is_browser_app(self, obj):
		return (
			obj and obj.appModule and 
			obj.appModule.appName.lower() in self.browser_apps
		)
	
	def get_current_url(self):
		if sys.version_info >= (3, 13):
			return self._get_current_url_2026()
		else:
			return self._get_current_url_2025()
	
	def _get_current_url_2025(self):
		focus = api.getFocusObject()
		url_to_copy = None
		
		url_to_copy = api.getCurrentURL()
		
		if not url_to_copy and hasattr(focus, 'treeInterceptor') and isinstance(focus.treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
			try:
				url_to_copy = focus.treeInterceptor.URL
			except Exception:
				pass
		
		if not url_to_copy and isinstance(focus, NVDAObjects.UIA.UIA):
			try:
				url_to_copy = focus.UIAElement.cachedAutomationID
				if url_to_copy and not url_to_copy.startswith('http'):
					url_to_copy = None
			except Exception:
				pass
		
		if not url_to_copy and hasattr(focus, 'IAccessibleObject'):
			try:
				url_to_copy = focus.IAccessibleObject.accValue(0)
			except Exception:
				pass
		
		return url_to_copy
	
	def _get_current_url_2026(self):
		focus = api.getFocusObject()
		
		# Traverse up to find the document object
		current = focus
		for _ in range(15):
			if current.role == controlTypes.Role.DOCUMENT:
				if hasattr(current, 'IAccessibleObject'):
					try:
						url = current.IAccessibleObject.accValue(0)
						if url and isinstance(url, str) and url.startswith(('http://', 'https://')):
							self.logger.info(f"2026 URL from document IAccessible: {url}")
							return url
					except Exception as e:
						self.logger.warning(f"Document IAccessible accValue failed: {e}")
				break
			
			if hasattr(current, 'parent') and current.parent:
				current = current.parent
			else:
				break
		
		# Fallback: try treeInterceptor
		if hasattr(focus, 'treeInterceptor') and focus.treeInterceptor:
			try:
				if hasattr(focus.treeInterceptor, 'URL'):
					url = focus.treeInterceptor.URL
					if url:
						self.logger.info(f"2026 URL from treeInterceptor: {url}")
						return url
			except Exception as e:
				self.logger.warning(f"treeInterceptor URL failed: {e}")
		
		# Fallback: try UIA Url property
		if isinstance(focus, NVDAObjects.UIA.UIA):
			try:
				url = focus.UIAElement.getCurrentPropertyValue(UIAHandler.UIA_UrlPropertyId)
				if url and isinstance(url, str) and url.startswith(('http://', 'https://')):
					self.logger.info(f"2026 URL from UIA Url: {url}")
					return url
			except Exception as e:
				self.logger.warning(f"UIA Url failed: {e}")
		
		self.logger.warning("Could not retrieve current URL in 2026")
		return None
	
	def get_link_url(self, obj):
		url = None
		
		if obj.role == controlTypes.Role.LINK:
			url = obj.value
			
			if not url and hasattr(obj, 'UIAElement'):
				try:
					url = obj.UIAElement.currentValue
				except Exception:
					pass
			
			if not url and hasattr(obj, 'IAccessibleObject'):
				try:
					url = obj.IAccessibleObject.accValue(0)
				except Exception:
					pass
		else:
			current = obj
			max_iterations = 5
			iterations = 0
			while current and current != api.getDesktopObject() and iterations < max_iterations:
				if current.role == controlTypes.Role.LINK:
					url = current.value
					
					if not url and hasattr(current, 'UIAElement'):
						try:
							url = current.UIAElement.currentValue
						except Exception:
							pass
					
					if not url and hasattr(current, 'IAccessibleObject'):
						try:
							url = current.IAccessibleObject.accValue(0)
						except Exception:
							pass
					
					if url:
						break
				current = current.parent
				iterations += 1
		
		return url