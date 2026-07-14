# url_utils.py

import addonHandler
addonHandler.initTranslation()

import api
import browseMode
import controlTypes
import NVDAObjects
import UIAHandler
import logging

class URLHandler:
	
	def __init__(self):
		self.logger = logging.getLogger(__name__)
		self.browser_apps = ["chrome", "firefox", "edge", "msedge", "opera", "safari", "brave"]
	
	def is_browser_app(self, obj):
		if not obj or not obj.appModule:
			return False
		return obj.appModule.appName.lower() in self.browser_apps
	
	def _is_valid_url(self, url):
		if not url or not isinstance(url, str):
			return False
		url_lower = url.lower()
		return (url_lower.startswith('http://') or 
		        url_lower.startswith('https://') or 
		        url_lower.startswith('file://') or 
		        url_lower.startswith('ftp://'))
	
	def get_current_url(self):
		focus = api.getFocusObject()
		if not focus:
			return None
		
		# Try document IAccessible first (covers file:// and other protocols)
		url = self._get_url_from_iaccessible_document(focus)
		if url:
			return url
		
		# Try treeInterceptor
		url = self._get_url_from_tree_interceptor(focus)
		if url:
			return url
		
		# Try UIA
		url = self._get_url_from_uia(focus)
		if url:
			return url
		
		# Try appModule
		url = self._get_url_from_appmodule(focus)
		if url:
			return url
		
		self.logger.warning("Could not retrieve current URL")
		return None
	
	def _get_url_from_tree_interceptor(self, focus):
		try:
			if hasattr(focus, 'treeInterceptor') and focus.treeInterceptor:
				if isinstance(focus.treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
					if hasattr(focus.treeInterceptor, 'URL'):
						url = focus.treeInterceptor.URL
						if self._is_valid_url(url):
							self.logger.info(f"URL from treeInterceptor: {url}")
							return url
		except Exception as e:
			self.logger.warning(f"treeInterceptor URL failed: {e}")
		return None
	
	def _get_url_from_uia(self, focus):
		try:
			if isinstance(focus, NVDAObjects.UIA.UIA):
				url = focus.UIAElement.getCurrentPropertyValue(UIAHandler.UIA_UrlPropertyId)
				if self._is_valid_url(url):
					self.logger.info(f"URL from UIA Url property: {url}")
					return url
		except Exception as e:
			self.logger.warning(f"UIA Url property failed: {e}")
		return None
	
	def _get_url_from_iaccessible_document(self, focus):
		try:
			current = focus
			for _ in range(15):
				if current.role == controlTypes.Role.DOCUMENT:
					if hasattr(current, 'IAccessibleObject'):
						url = current.IAccessibleObject.accValue(0)
						if self._is_valid_url(url):
							self.logger.info(f"URL from document IAccessible: {url}")
							return url
					break
				if hasattr(current, 'parent') and current.parent:
					current = current.parent
				else:
					break
		except Exception as e:
			self.logger.warning(f"IAccessible document URL failed: {e}")
		return None
	
	def _get_url_from_appmodule(self, focus):
		try:
			if focus.appModule and hasattr(focus.appModule, 'getCurrentURL'):
				url = focus.appModule.getCurrentURL()
				if self._is_valid_url(url):
					return url
		except Exception as e:
			self.logger.warning(f"appModule URL failed: {e}")
		return None
	
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
			if hasattr(current, 'parent') and current.parent:
				current = current.parent
			else:
				break
		
		return None
	
	def _extract_link_url(self, link_obj):
		try:
			if hasattr(link_obj, 'value') and link_obj.value:
				url = link_obj.value
				if self._is_valid_url(url):
					return url
			
			if hasattr(link_obj, 'UIAElement'):
				url = link_obj.UIAElement.currentValue
				if self._is_valid_url(url):
					return url
			
			if hasattr(link_obj, 'IAccessibleObject'):
				url = link_obj.IAccessibleObject.accValue(0)
				if self._is_valid_url(url):
					return url
		except Exception as e:
			self.logger.warning(f"Extract link URL failed: {e}")
		return None