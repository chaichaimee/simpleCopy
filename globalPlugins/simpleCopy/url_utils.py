# url_utils.py

import addonHandler
addonHandler.initTranslation()

import api
import browseMode
import controlTypes
import NVDAObjects
import UIAHandler
import IAccessibleHandler
import logging
from comtypes import COMError

class URLHandler:
    """Handler for URL-related operations"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.browser_apps = ["chrome", "firefox", "edge", "msedge", "opera", "safari", "brave"]
    
    def is_browser_app(self, obj):
        """Check if the current app is a browser"""
        return (
            obj and obj.appModule and 
            obj.appModule.appName.lower() in self.browser_apps
        )
    
    def get_current_url(self):
        """Get current URL from browser using multiple methods"""
        focus = api.getFocusObject()
        url_to_copy = None
        
        # Method 1: Standard API
        url_to_copy = api.getCurrentURL()
        
        # Method 2: Tree interceptor fallback
        if not url_to_copy and hasattr(focus, 'treeInterceptor') and isinstance(focus.treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
            try:
                url_to_copy = focus.treeInterceptor.URL
            except Exception:
                pass
        
        # Method 3: UIA for modern browsers
        if not url_to_copy and isinstance(focus, NVDAObjects.UIA.UIA):
            try:
                url_to_copy = focus.UIAElement.cachedAutomationID
                if url_to_copy and not url_to_copy.startswith('http'):
                    url_to_copy = None
            except Exception:
                pass
        
        # Method 4: IAccessible fallback
        if not url_to_copy and hasattr(focus, 'IAccessibleObject'):
            try:
                url_to_copy = focus.IAccessibleObject.accValue(0)
            except Exception:
                pass
        
        return url_to_copy
    
    def get_link_url(self, obj):
        """Get URL from a link object using multiple methods"""
        url = None
        
        if obj.role == controlTypes.Role.LINK:
            # Standard method
            url = obj.value
            
            # UIA fallback
            if not url and hasattr(obj, 'UIAElement'):
                try:
                    url = obj.UIAElement.currentValue
                except Exception:
                    pass
            
            # IAccessible fallback
            if not url and hasattr(obj, 'IAccessibleObject'):
                try:
                    url = obj.IAccessibleObject.accValue(0)
                except Exception:
                    pass
        else:
            # Search parents for link
            current = obj
            max_iterations = 5
            iterations = 0
            while current and current != api.getDesktopObject() and iterations < max_iterations:
                if current.role == controlTypes.Role.LINK:
                    # Standard
                    url = current.value
                    
                    # UIA fallback
                    if not url and hasattr(current, 'UIAElement'):
                        try:
                            url = current.UIAElement.currentValue
                        except Exception:
                            pass
                    
                    # IAccessible fallback
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

