# Copyright (C) 2025 chai chaimee
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
import re
import logging
import UIAHandler
import winUser
import gui
from ui import message as ui_message
from datetime import datetime
import comtypes.client
import IAccessibleHandler
import hashlib

# Configure logging for debugging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

addonHandler.initTranslation()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("Simple Copy")
    isEnabled = True
    isCopyAppendEnabled = True
    isTextCopied = False
    lastKeyPress = {"key": None, "time": 0}
    pendingCKeyPress = None
    pendingShiftCKeyPress = None

    def __init__(self):
        super().__init__()

    def normalize_text(self, text):
        if not text:
            return ""
        # Remove non-printable characters except common whitespace and newlines
        text = "".join(char for char in text if char.isprintable() or char in {"\r", "\n", " "})
        text = text.replace("\r\n", "\n").replace("\r", "\n") # Convert all to consistent \n
        return text

    def calculate_sha256(self, text):
        normalized_text = self.normalize_text(text)
        return hashlib.sha256(normalized_text.encode('utf-8')).hexdigest()

    def getSelectedText(self, obj_param):
        """
        Gets selected text using info.clipboardText, which is designed for clipboard compatibility
        and is expected to handle multi-line preservation better.
        """
        logging.info("SimpleCopy: getSelectedText started (using info.clipboardText).")
        current_obj = obj_param
        selected_text = None

        is_browser_app = (
            current_obj.appModule and
            current_obj.appModule.appName in ("chrome", "firefox", "edge", "msedge", "opera", "safari")
        )

        try:
            # Prioritize treeInterceptor for browser apps if available and capable
            # This logic is consistent with clipContentsDesigner's approach 
            if is_browser_app and hasattr(current_obj, 'treeInterceptor') and current_obj.treeInterceptor and not current_obj.treeInterceptor.passThrough:
                # In clipContentsDesigner, it specifically checks for BrowseModeDocumentTreeInterceptor 
                # We'll adapt by just checking for makeTextInfo on treeInterceptor
                if hasattr(current_obj.treeInterceptor, 'makeTextInfo'):
                    current_obj = current_obj.treeInterceptor
                    logging.info(f"SimpleCopy: (getSelectedText) Switched to treeInterceptor for app: {obj_param.appModule.appName}")
                else:
                    logging.info(f"SimpleCopy: (getSelectedText) TreeInterceptor found but no makeTextInfo, sticking with original obj.")
            else:
                logging.info(f"SimpleCopy: (getSelectedText) Not a browser app or no valid treeInterceptor. Using original obj.")

            # Attempt to get text from selection using makeTextInfo
            if hasattr(current_obj, 'makeTextInfo'):
                try:
                    info = current_obj.makeTextInfo(textInfos.POSITION_SELECTION)
                    if info and not info.isCollapsed:
                        # CRITICAL CHANGE: Use info.clipboardText as suggested by clipContentsDesigner 
                        selected_text = info.clipboardText
                        logging.info(f"SimpleCopy: (getSelectedText) makeTextInfo (Position Selection) using clipboardText: {repr(selected_text)}")
                        if selected_text:
                            # Apply final normalization: convert line endings and strip overall
                            return selected_text.replace('\r\n', '\n').replace('\r', '\n').strip()
                        
                except (RuntimeError, NotImplementedError) as e:
                    logging.warning(f"SimpleCopy: (getSelectedText) makeTextInfo selection failed on {current_obj.name}: {str(e)}")
            else:
                logging.info(f"SimpleCopy: (getSelectedText) Object does not have makeTextInfo method.")
            
            # Fallback to simulating Ctrl+C if direct makeTextInfo failed or yielded no useful text
            logging.info("SimpleCopy: (getSelectedText) No text from makeTextInfo, attempting Ctrl+C fallback.")
            original_clipboard_data = ""
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    original_clipboard_data = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
                    winUser.emptyClipboard()
                
                keyboardHandler.feedKeyPress("control+c")
                time.sleep(0.05) # Small delay for clipboard to update

                with winUser.openClipboard(gui.mainFrame.Handle):
                    clipboard_text = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
                
                if clipboard_text:
                    selected_text = clipboard_text
                    logging.info(f"SimpleCopy: (getSelectedText) Fallback Ctrl+C retrieved: {repr(selected_text)}")
                    
                    # Apply basic normalization for Ctrl+C text
                    return selected_text.replace('\r\n', '\n').replace('\r', '\n').strip()

            except Exception as e_fallback:
                logging.error(f"SimpleCopy: (getSelectedText) Fallback Ctrl+C failed: {str(e_fallback)}")
            finally:
                # Restore original clipboard content
                try:
                    with winUser.openClipboard(gui.mainFrame.Handle):
                        winUser.emptyClipboard()
                        if original_clipboard_data:
                            winUser.setClipboardData(winUser.CF_UNICODETEXT, original_clipboard_data)
                except Exception as e_restore:
                    logging.error(f"SimpleCopy: (getSelectedText) Failed to restore clipboard: {str(e_restore)}")

        except Exception as e:
            logging.error(f"SimpleCopy: (getSelectedText) Unexpected general error: {str(e)}")
            
        logging.info("SimpleCopy: (getSelectedText) No selected text found. Returning None.")
        return None

    @scriptHandler.script(
        description=_("Append text"),
        gesture="kb(desktop):shift+c",
        category=scriptCategory
    )
    def script_appendToClipboard(self, gesture):
        if not self.isEnabled or not self.isCopyAppendEnabled:
            gesture.send()
            return
            
        currentTime = time.time()
        key = "shift+c"
            
        if self.lastKeyPress["key"] == key and (currentTime - self.lastKeyPress["time"]) < 0.5:
            if self.pendingShiftCKeyPress:
                scriptHandler.cancelScript(self.pendingShiftCKeyPress)
                self.pendingShiftCKeyPress = None

            obj = api.getFocusObject()
            text_to_append = None
                
            try:
                text_to_append = self.getSelectedText(obj)
                logging.info(f"SimpleCopy: (script_appendToClipboard) Text obtained for append: {repr(text_to_append)}")
            except Exception as e:
                logging.error(f"SimpleCopy: (script_appendToClipboard) Error retrieving text: {str(e)}")

            if not text_to_append:
                self.lastKeyPress = {"key": None, "time": 0}
                speech.speak([_("No text selected to append")])
                return
                
            clipData = ""
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    clipData = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
                    if clipData and not isinstance(clipData, str):
                        speech.speak([_("Cannot append to non-text clipboard content")])
                        self.lastKeyPress = {"key": None, "time": 0}
                        return
            except Exception as e:
                logging.error(f"SimpleCopy: (script_appendToClipboard) Error reading clipboard: {str(e)}")
                clipData = "" # Reset if error

            processed_text_to_append = text_to_append
            
            if clipData:
                # Ensure a clear separation with double newline
                # Normalize line endings from both clipData and processed_text_to_append
                clipData_normalized = clipData.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n')
                processed_text_to_append_normalized = processed_text_to_append.replace('\r\n', '\n').replace('\r', '\n').lstrip('\n')
                
                newText = clipData_normalized + "\r\n\r\n" + processed_text_to_append_normalized
            else:
                newText = processed_text_to_append.replace('\r\n', '\n').replace('\r', '\n') # Just normalize if first text
            
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    winUser.emptyClipboard()
                    # Always set clipboard data with \r\n for universal compatibility on Windows
                    winUser.setClipboardData(winUser.CF_UNICODETEXT, newText.replace('\n', '\r\n'))
                self.isTextCopied = True
                speech.speak([_("Appended") if clipData else _("Copied")])
            except Exception as e:
                logging.error(f"SimpleCopy: (script_appendToClipboard) Error writing to clipboard: {str(e)}")
                speech.speak([_("Error writing to clipboard")])
            finally:
                self.lastKeyPress = {"key": None, "time": 0}
        else:
            self.lastKeyPress = {"key": key, "time": currentTime}
            self.pendingShiftCKeyPress = scriptHandler.runScript(self._resetKeyPressAndSendShiftC, gesture, 0.5)

    def _resetKeyPressAndSendC(self, original_gesture):
        if self.lastKeyPress["key"] == "c" and (time.time() - self.lastKeyPress["time"]) >= 0.5:
            self.lastKeyPress = {"key": None, "time": 0}
            original_gesture.send()
        self.pendingCKeyPress = None

    def _resetKeyPressAndSendShiftC(self, original_gesture):
        if self.lastKeyPress["key"] == "shift+c" and (time.time() - self.lastKeyPress["time"]) >= 0.5:
            self.lastKeyPress = {"key": None, "time": 0}
            original_gesture.send()
        self.pendingShiftCKeyPress = None

    @scriptHandler.script(
        description=_("Toggle append functionality"),
        gesture="kb:windows+c",
        category=scriptCategory
    )
    def script_toggleCopyAppend(self, gesture):
        self.isCopyAppendEnabled = not self.isCopyAppendEnabled
        state = _("on") if self.isCopyAppendEnabled else _("off")
        speech.speak([_("Append {}").format(state)])

    @scriptHandler.script(
        description=_("Clear clipboard"),
        gesture="kb:windows+z",
        category=scriptCategory
    )
    def script_clearClipboard(self, gesture):
        if not self.isEnabled:
            gesture.send()
            return
        try:
            with winUser.openClipboard(gui.mainFrame.Handle):
                winUser.emptyClipboard()
            self.isTextCopied = False
            speech.speak([_("Clean")])
        except Exception as e:
            logging.error(f"SimpleCopy: Error clearing clipboard: {str(e)}")
            speech.speak([_("Cannot clean")])

    @scriptHandler.script(
        description=_("Copy URL web browser"),
        gesture="kb:shift+a",
        category=scriptCategory
    )
    def script_copyBrowserAddress(self, gesture):
        if not self.isEnabled:
            gesture.send()
            return
        currentTime = time.time()
        if self.lastKeyPress["key"] == "shift+a" and (currentTime - self.lastKeyPress["time"]) < 0.5:
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    winUser.emptyClipboard()

                time.sleep(0.1)

                obj = api.getFocusObject()
                if not obj or not obj.appModule or obj.appModule.appName not in ("chrome", "firefox", "edge", "msedge", "opera", "safari"):
                    ui_message(_("Not in a supported browser"))
                    return

                url = api.getCurrentURL()

                if url is None or not re.match(r'^https?://', url, re.IGNORECASE):
                    ui_message(_("No valid URL available"))
                    return

                if api.copyToClip(url):
                    ui_message(_("Copy"))
                else:
                    ui_message(_("Failed to copy"))
            except Exception as e:
                logging.error(f"SimpleCopy: Error in copyBrowserAddress: {str(e)}")
                ui_message(_("Error copy"))
        self.lastKeyPress["key"] = "shift+a"
        self.lastKeyPress["time"] = currentTime

    @scriptHandler.script(
        description=_("Copy file & folder name"),
        gesture="kb:shift+f",
        category=scriptCategory
    )
    def script_copyExplorerFileName(self, gesture):
        if not self.isEnabled:
            gesture.send()
            return
        currentTime = time.time()
        if self.lastKeyPress["key"] == "shift+f" and (currentTime - self.lastKeyPress["time"]) < 0.5:
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    winUser.emptyClipboard()

                obj = api.getFocusObject()
                if not obj or obj.appModule.appName != "explorer":
                    ui_message(_("Not in File Explorer"))
                    return
                name = obj.name
                if name and name.strip():
                    if api.copyToClip(name):
                        ui_message(_("Copy"))
                    else:
                        ui_message(_("Failed copy"))
                else:
                    ui_message(_("No valid name available"))
            except Exception as e:
                logging.error(f"SimpleCopy: Error in copyExplorerFileName: {str(e)}")
                ui_message(_("Error copy"))
        self.lastKeyPress["key"] = "shift+f"
        self.lastKeyPress["time"] = currentTime

    @scriptHandler.script(
        description=_("Copy the hyperlink URL"),
        gesture="kb:shift+l",
        category=scriptCategory
    )
    def script_copyLink(self, gesture):
        if not self.isEnabled:
            gesture.send()
            return
        currentTime = time.time()
        if self.lastKeyPress["key"] == "shift+l" and (currentTime - self.lastKeyPress["time"]) < 0.5:
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    winUser.emptyClipboard()

                obj = api.getNavigatorObject()
                if not obj or not obj.appModule or obj.appModule.appName not in ("chrome", "firefox", "edge", "msedge", "opera", "safari"):
                    ui_message(_("Not in a web browser"))
                    return
                url = None
                if obj.role == controlTypes.Role.LINK:
                    url = obj.value or (obj.IAccessibleObject.accValue(0) if hasattr(obj, 'IAccessibleObject') and obj.IAccessibleObject else None)
                else:
                    current = obj
                    max_iterations = 5
                    iterations = 0
                    while current and current != api.getDesktopObject() and iterations < max_iterations:
                        if current.role == controlTypes.Role.LINK:
                            url = current.value or (current.IAccessibleObject.accValue(0) if hasattr(current, 'IAccessibleObject') and current.IAccessibleObject else None)
                            break
                        current = current.parent
                        iterations += 1
                if url and url.strip():
                    if api.copyToClip(url):
                        ui_message(_("Copy"))
                    else:
                        ui_message(_("Failed copy"))
                else:
                    ui_message(_("No link found"))
            except Exception as e:
                logging.error(f"SimpleCopy: Error in copyLink: {str(e)}")
                ui_message(_("Cannot copy link"))
        self.lastKeyPress["key"] = "shift+l"
        self.lastKeyPress["time"] = currentTime

    @scriptHandler.script(
        description=_("Copy current date and time"),
        gesture="kb:shift+d",
        category=scriptCategory
    )
    def script_copyDateTime(self, gesture):
        if not self.isEnabled:
            gesture.send()
            return
        currentTime = time.time()
        if self.lastKeyPress["key"] == "shift+d" and (currentTime - self.lastKeyPress["time"]) < 0.5:
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    winUser.emptyClipboard()

                date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if api.copyToClip(date_time):
                    ui_message(_("Date and time copy"))
                else:
                    ui_message(_("Failed copy"))
            except Exception as e:
                logging.error(f"SimpleCopy: Error in copyDateTime: {str(e)}")
                ui_message(_("Error copy"))
        self.lastKeyPress["key"] = "shift+d"
        self.lastKeyPress["time"] = currentTime
