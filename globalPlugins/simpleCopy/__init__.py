# __init__.py
# Copyright (C) 2025 ['CHAI CHAIMEE']
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
import wx
import browseMode
import NVDAObjects
from . import url_utils
from . import clipboard_utils

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

addonHandler.initTranslation()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("Simple Copy")
    isTextCopied = False
    
    # Tap detection variables for CTRL+Shift+A
    _ctrl_shift_a_tap_count = 0
    _ctrl_shift_a_last_tap_time = 0
    _ctrl_shift_a_timer = None
    
    # Tap detection variables for CTRL+Shift+C
    _ctrl_shift_c_tap_count = 0
    _ctrl_shift_c_last_tap_time = 0
    _ctrl_shift_c_timer = None
    
    _double_tap_threshold = 0.5  # seconds
    
    def __init__(self):
        super().__init__()
        self.clipboard_handler = clipboard_utils.ClipboardHandler()
        self.url_handler = url_utils.URLHandler()

    def normalize_text(self, text):
        return self.clipboard_handler.normalize_text(text)

    def calculate_sha256(self, text):
        normalized_text = self.normalize_text(text)
        return hashlib.sha256(normalized_text.encode('utf-8')).hexdigest()

    def getSelectedText(self, obj_param):
        logging.info("SimpleCopy: getSelectedText started.")
        return self.clipboard_handler.get_selected_text(obj_param)

    def _isTextSelected(self, obj_param):
        logging.info("SimpleCopy: _isTextSelected called.")
        selected_text = self.getSelectedText(obj_param)
        return bool(selected_text)
    
    def _performAppendAction(self, obj):
        text_to_append = None
        try:
            text_to_append = self.getSelectedText(obj) 
            logging.info(f"SimpleCopy: (_performAppendAction) Text obtained for append: {repr(text_to_append[:50]) if text_to_append else 'None'}...")
        except Exception as e:
            logging.error(f"SimpleCopy: (_performAppendAction) Error retrieving text: {str(e)}")

        if not text_to_append:
            speech.speak([_("No text selected to append")])
            return
            
        try:
            result = self.clipboard_handler.append_to_clipboard(text_to_append)
            if result["success"]:
                self.isTextCopied = True
                speech.speak([_("Appended") if result["appended"] else _("Copied")])
            else:
                speech.speak([_(result["message"])])
        except Exception as e:
            logging.error(f"SimpleCopy: (_performAppendAction) Error in clipboard operation: {str(e)}")
            speech.speak([_("Error during append operation")])

    @scriptHandler.script(
        description=_("Copy URL (single tap), Copy hyperlink (double tap)"),
        gesture="kb:control+shift+a",
        category=scriptCategory
    )
    def script_copyUrlOrHyperlink(self, gesture):
        """Handle CTRL+Shift+A: single tap for URL, double tap for hyperlink"""
        current_time = time.time()
        
        # Reset tap count if too much time has passed
        if current_time - self._ctrl_shift_a_last_tap_time > self._double_tap_threshold:
            self._ctrl_shift_a_tap_count = 0
        
        self._ctrl_shift_a_tap_count += 1
        self._ctrl_shift_a_last_tap_time = current_time
        
        # Cancel any existing timer
        if self._ctrl_shift_a_timer and self._ctrl_shift_a_timer.IsRunning():
            self._ctrl_shift_a_timer.Stop()
        
        # Schedule action after threshold
        self._ctrl_shift_a_timer = wx.CallLater(
            int(self._double_tap_threshold * 1000),
            self._executeCtrlShiftAAction
        )
    
    def _executeCtrlShiftAAction(self):
        """Execute the appropriate action for CTRL+Shift+A based on tap count"""
        if self._ctrl_shift_a_tap_count == 1:
            # Single tap: Copy URL web browser
            self._copyBrowserUrl()
        elif self._ctrl_shift_a_tap_count >= 2:
            # Double tap: Copy hyperlink URL
            self._copyHyperlinkUrl()
        
        # Reset tap count
        self._ctrl_shift_a_tap_count = 0
    
    def _copyBrowserUrl(self):
        """Copy current browser URL"""
        obj = api.getFocusObject()
        is_editable = (obj.role in (controlTypes.Role.EDITABLETEXT, controlTypes.Role.TEXTFRAME) or 
                       controlTypes.State.EDITABLE in obj.states)
        is_browser_app = self.url_handler.is_browser_app(obj)

        # If editable field, send original gesture
        if is_editable:
            keyboardHandler.KeyboardInputGesture.fromName("control+shift+a").send()
            return
        
        if is_browser_app:
            try:
                url_to_copy = self.url_handler.get_current_url()
                if url_to_copy is None:
                    ui_message(_("No URL"))
                    return
                if api.copyToClip(url_to_copy):
                    ui_message(_("Copy"))
                else:
                    ui_message(_("Failed to copy"))
            except Exception as e:
                logging.error(f"SimpleCopy: Error in copyBrowserAddress: {str(e)}")
                ui_message(_("Error copy"))
        else:
            # Not in browser, send original gesture
            keyboardHandler.KeyboardInputGesture.fromName("control+shift+a").send()
    
    def _copyHyperlinkUrl(self):
        """Copy hyperlink URL"""
        obj = api.getFocusObject()
        is_browser_app = self.url_handler.is_browser_app(obj)
        is_editable = (obj.role in (controlTypes.Role.EDITABLETEXT, controlTypes.Role.TEXTFRAME) or 
                       controlTypes.State.EDITABLE in obj.states)

        if is_browser_app and is_editable:
            keyboardHandler.KeyboardInputGesture.fromName("control+shift+a").send()
            return

        if is_browser_app:
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    winUser.emptyClipboard()

                obj = api.getNavigatorObject()
                if not obj or not obj.appModule or not self.url_handler.is_browser_app(obj):
                    ui_message(_("Not in a web browser"))
                    return
                
                url = self.url_handler.get_link_url(obj)
                
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
        else:
            # Not in browser, send original gesture
            keyboardHandler.KeyboardInputGesture.fromName("control+shift+a").send()

    @scriptHandler.script(
        description=_("Append text (single tap), Clear clipboard (double tap)"),
        gesture="kb:control+shift+c",
        category=scriptCategory
    )
    def script_appendOrClear(self, gesture):
        """Handle CTRL+Shift+C: single tap for append, double tap for clear"""
        current_time = time.time()
        
        # Reset tap count if too much time has passed
        if current_time - self._ctrl_shift_c_last_tap_time > self._double_tap_threshold:
            self._ctrl_shift_c_tap_count = 0
        
        self._ctrl_shift_c_tap_count += 1
        self._ctrl_shift_c_last_tap_time = current_time
        
        # Cancel any existing timer
        if self._ctrl_shift_c_timer and self._ctrl_shift_c_timer.IsRunning():
            self._ctrl_shift_c_timer.Stop()
        
        # Schedule action after threshold
        self._ctrl_shift_c_timer = wx.CallLater(
            int(self._double_tap_threshold * 1000),
            self._executeCtrlShiftCAction
        )
    
    def _executeCtrlShiftCAction(self):
        """Execute the appropriate action for CTRL+Shift+C based on tap count"""
        if self._ctrl_shift_c_tap_count == 1:
            # Single tap: Append text
            self._appendToClipboard()
        elif self._ctrl_shift_c_tap_count >= 2:
            # Double tap: Clear clipboard
            self._clearClipboard()
        
        # Reset tap count
        self._ctrl_shift_c_tap_count = 0
    
    def _appendToClipboard(self):
        """Append selected text to clipboard"""
        obj = api.getFocusObject()
        has_selection = self._isTextSelected(obj)
        
        if not has_selection:
            # No selection, send original gesture
            keyboardHandler.KeyboardInputGesture.fromName("control+shift+c").send()
            return
        
        self._performAppendAction(obj)
    
    def _clearClipboard(self):
        """Clear clipboard content"""
        try:
            with winUser.openClipboard(gui.mainFrame.Handle):
                winUser.emptyClipboard()
            self.isTextCopied = False
            speech.speak([_("Clean")])
        except Exception as e:
            logging.error(f"SimpleCopy: Error clearing clipboard: {str(e)}")
            speech.speak([_("Cannot clean")])

