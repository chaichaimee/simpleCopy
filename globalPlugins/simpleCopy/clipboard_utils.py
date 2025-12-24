# clipboard_utils.py

import addonHandler
addonHandler.initTranslation()

import api
import winUser
import gui
import time
import keyboardHandler
import textInfos
import browseMode
import logging
import hashlib

class ClipboardHandler:
    """Handler for clipboard-related operations"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def normalize_text(self, text):
        """Normalize text by removing non-printable characters and standardizing line endings"""
        if not text:
            return ""
        text = "".join(char for char in text if char.isprintable() or char in {"\r", "\n", " "})
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        return text
    
    def calculate_sha256(self, text):
        """Calculate SHA256 hash of normalized text"""
        normalized_text = self.normalize_text(text)
        return hashlib.sha256(normalized_text.encode('utf-8')).hexdigest()
    
    def get_selected_text(self, obj_param):
        """Get selected text using multiple methods"""
        self.logger.info("SimpleCopy: get_selected_text started.")
        current_obj = obj_param
        selected_text = None
        
        try:
            target_obj_for_text = None
            if hasattr(current_obj, 'treeInterceptor') and isinstance(current_obj.treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
                target_obj_for_text = current_obj.treeInterceptor
                self.logger.info("SimpleCopy: (get_selected_text) Using treeInterceptor for text info.")
            elif hasattr(current_obj, 'makeTextInfo'):
                target_obj_for_text = current_obj
            
            if target_obj_for_text:
                try:
                    info = target_obj_for_text.makeTextInfo(textInfos.POSITION_SELECTION)
                    if info and not info.isCollapsed:
                        selected_text = info.clipboardText
                        if selected_text:
                            self.logger.info(f"SimpleCopy: (get_selected_text) makeTextInfo (Position Selection) retrieved: {repr(selected_text[:50])}...")
                            return selected_text.replace('\r\n', '\n').replace('\r', '\n').strip()
                except (RuntimeError, NotImplementedError) as e:
                    self.logger.warning(f"SimpleCopy: (get_selected_text) makeTextInfo selection failed on {target_obj_for_text.name if target_obj_for_text else 'unknown'}: {str(e)}")
            else:
                self.logger.info("SimpleCopy: (get_selected_text) No makeTextInfo available on focus/treeInterceptor object.")
        
        except Exception as e_info:
            self.logger.error(f"SimpleCopy: (get_selected_text) Error with makeTextInfo attempt: {str(e_info)}")
        
        self.logger.info("SimpleCopy: (get_selected_text) makeTextInfo failed or no text, attempting Ctrl+C fallback.")
        original_clipboard_data = ""
        try:
            with winUser.openClipboard(gui.mainFrame.Handle):
                original_clipboard_data = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
                winUser.emptyClipboard()
            
            keyboardHandler.injectKey("control+c")
            time.sleep(0.05)
            
            with winUser.openClipboard(gui.mainFrame.Handle):
                clipboard_text = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
            
            if clipboard_text:
                selected_text = clipboard_text
                self.logger.info(f"SimpleCopy: (get_selected_text) Fallback Ctrl+C retrieved: {repr(selected_text[:50])}...")
                return selected_text.replace('\r\n', '\n').replace('\r', '\n').strip()
        
        except Exception as e_fallback:
            self.logger.error(f"SimpleCopy: (get_selected_text) Fallback Ctrl+C failed: {str(e_fallback)}")
        finally:
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    winUser.emptyClipboard()
                    if original_clipboard_data:
                        winUser.setClipboardData(winUser.CF_UNICODETEXT, original_clipboard_data)
            except Exception as e_restore:
                self.logger.error(f"SimpleCopy: (get_selected_text) Failed to restore clipboard: {str(e_restore)}")
        
        self.logger.info("SimpleCopy: (get_selected_text) No selected text found after all attempts. Returning None.")
        return None
    
    def append_to_clipboard(self, text_to_append):
        """Append text to current clipboard content"""
        clipData = ""
        try:
            with winUser.openClipboard(gui.mainFrame.Handle):
                clipData = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
                if clipData and not isinstance(clipData, str):
                    return {
                        "success": False,
                        "appended": False,
                        "message": _("Cannot append to non-text clipboard content")
                    }
        except Exception as e:
            self.logger.error(f"SimpleCopy: (append_to_clipboard) Error reading clipboard: {str(e)}")
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
                "message": _("Appended") if appended else _("Copied")
            }
        except Exception as e:
            self.logger.error(f"SimpleCopy: (append_to_clipboard) Error writing to clipboard: {str(e)}")
            return {
                "success": False,
                "appended": False,
                "message": _("Error writing to clipboard")
            }


