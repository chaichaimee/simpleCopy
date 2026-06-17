# simpleCopy

**Copy, Append, and Manage Text Efficiently with NVDA**

**author:** chai chaimee  
**url:** https://github.com/chaichaimee/simpleCopy

---

## Description

**simpleCopy** is a lightweight NVDA add-on that simplifies text copying, URL extraction, and speech history management.

This tool helps you capture and organize information quickly without interrupting your workflow. Whether you are copying text, grabbing web links, or saving spoken content, simpleCopy provides intuitive keyboard shortcuts that work seamlessly with NVDA.

---

## Hot Keys

All commands use a multi-tap system. Press the key combination once, twice, or three times in quick succession to perform different actions.

### CTRL+Shift+A — URL and Link Capture

- **Single Tap:** Copies the current webpage URL.
- **Double Tap:** Copies the destination URL of the focused hyperlink.

### CTRL+Shift+V — Text Copy, Append, and Clipboard Management

- **Single Tap:** Copies the selected text. If text already exists in the clipboard, the new selection is appended to it.
- **Double Tap:** Copies text from the current review cursor position. This works with any selection made using NVDA's review cursor, including multi-line selections and full document selection.
- **Triple Tap:** Clears all content from the clipboard.

### F9 — Speech Capture and Management

- **Single Tap:** Copies the most recent speech output from NVDA.
- **Double Tap:** Appends the most recent speech output to existing clipboard content.
- **Triple Tap:** Copies all speech output accumulated since the first F9 press.

### Shift+F9 — Speech History Navigation

- **Single Tap:** Navigates to the previous speech history item.
- **Double Tap:** Navigates to the next speech history item.
- **Triple Tap:** Opens the complete speech history log file.

---

## Features

Here is how each feature works in practice:

### 1. Webpage URL Copy

Press **CTRL+Shift+A once** while browsing any website. The current page URL is copied to your clipboard. NVDA confirms by speaking the copied URL.

### 2. Hyperlink URL Extraction

Focus on any link and press **CTRL+Shift+A twice**. The destination URL is extracted and copied without opening the link.

### 3. Text Copy and Append

Select text and press **CTRL+Shift+V once**. If the clipboard is empty, the text is copied. If the clipboard already contains text, the new selection is appended with a line break.

### 4. Review Cursor Copy

Use NVDA's review cursor to select text (using NVDA+Shift+Down or NVDA+Ctrl+Shift+Down to select multiple lines), then press **CTRL+Shift+V twice**. All selected text from the review cursor position is copied to the clipboard. This works with any selection size, from a single word to an entire document.

### 5. Clipboard Clear

Press **CTRL+Shift+V three times** to instantly clear all clipboard content. NVDA confirms with the message "Clean".

### 6. Last Speech Copy

When NVDA speaks something you want to save, press **F9 once**. The last spoken phrase is copied to your clipboard.

### 7. Speech Append

Press **F9 twice** to append the last spoken phrase to existing clipboard content.

### 8. Speech History Logging

Press **F9 three times** to copy all speech output accumulated during your current session.

### 9. Speech History Navigation

Use **Shift+F9 once** to move backward through speech history, and **Shift+F9 twice** to move forward. This lets you review past speech output without changing your current focus.

### 10. Speech Log File Access

Press **Shift+F9 three times** to open the complete speech history file in your default text editor for reviewing, searching, or copying.

### 11. Smart Context Awareness

When you are typing in editable fields, simpleCopy does not interfere. Commands only activate when they are useful, preserving your normal workflow.

---

## Support Me

If this add-on helps you work more efficiently, consider supporting future development with a small donation.

[![Donate](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

Your support helps keep this project alive and improving.

---

© 2026 Chai Chaimee NVDA Add-on Released under GNU General Public License