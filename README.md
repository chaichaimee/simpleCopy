<div align="center">

<img src="https://www.nvaccess.org/files/nvda/documentation/userGuide/images/nvda.ico" alt="NVDA Logo" width="120" style="display: block; margin: 0 auto 20px; height: auto;">

# simpleCopy

<br>

<p style="font-size: 1.2em; color: #2c3e50;"><b>Copy, Append, and Manage Text Efficiently with NVDA</b></p>

</div>

<br>

<div align="center">

**author:** chai chaimee

**url:** https://github.com/chaichaimee/simpleCopy

</div>

<hr>

## Description

<br>

**simpleCopy** is a lightweight NVDA add-on that simplifies text copying, URL extraction, speech history management, and URL history organization.

<br>

This tool helps you capture and organize information quickly without interrupting your workflow. Whether you are copying text, grabbing web links, saving spoken content, or managing frequently used URLs, simpleCopy provides intuitive keyboard shortcuts that work seamlessly with NVDA.

<br>

## Hot Keys

<br>

All commands use a multi-tap system. Press the key combination once, twice, or three times in quick succession to perform different actions.

<br>

> ### CTRL+Shift+A — URL, Link, and URL History
> 
> **Single Tap:** Copies the current webpage URL.
> 
> **Double Tap:** Copies the destination URL of the focused hyperlink.
> 
> **Triple Tap:** Opens the URL History dialog to manage saved URLs and hyperlinks.

<br>

> ### CTRL+Shift+V — Text Copy, Append, and Clipboard Management
> 
> **Single Tap:** Copies the selected text. If text already exists in the clipboard, the new selection is appended to it.
> 
> **Double Tap:** Copies text from the current review cursor position. This works with any selection made using NVDA's review cursor, including multi-line selections and full document selection.
> 
> **Triple Tap:** Clears all content from the clipboard.

<br>

> ### F9 — Speech Capture and Management
> 
> **Single Tap:** Copies the most recent speech output from NVDA.
> 
> **Double Tap:** Appends the most recent speech output to existing clipboard content.
> 
> **Triple Tap:** Copies all speech output accumulated since the first F9 press.

<br>

> ### Shift+F9 — Speech History Navigation
> 
> **Single Tap:** Navigates to the previous speech history item.
> 
> **Double Tap:** Navigates to the next speech history item.
> 
> **Triple Tap:** Opens the complete speech history log file.

<br>

## Features

<br>

Here is how each feature works in practice:

<br>

### 1. Webpage URL Copy

Press **CTRL+Shift+A once** while browsing any website. The current page URL is copied to your clipboard. NVDA confirms by speaking the copied URL.

<br>

### 2. Hyperlink URL Extraction

Focus on any link and press **CTRL+Shift+A twice**. The destination URL is extracted and copied without opening the link.

<br>

### 3. Text Copy and Append

Select text and press **CTRL+Shift+V once**. If the clipboard is empty, the text is copied. If the clipboard already contains text, the new selection is appended with a line break.

<br>

### 4. Review Cursor Copy

Use NVDA's review cursor to select text (using NVDA+Shift+Down or NVDA+Ctrl+Shift+Down to select multiple lines), then press **CTRL+Shift+V twice**. All selected text from the review cursor position is copied to the clipboard. This works with any selection size, from a single word to an entire document.

<br>

### 5. Clipboard Clear

Press **CTRL+Shift+V three times** to instantly clear all clipboard content. NVDA confirms with the message "Clean".

<br>

### 6. Last Speech Copy

When NVDA speaks something you want to save, press **F9 once**. The last spoken phrase is copied to your clipboard.

<br>

### 7. Speech Append

Press **F9 twice** to append the last spoken phrase to existing clipboard content.

<br>

### 8. Speech History Logging

Press **F9 three times** to copy all speech output accumulated during your current session.

<br>

### 9. Speech History Navigation

Use **Shift+F9 once** to move backward through speech history, and **Shift+F9 twice** to move forward. This lets you review past speech output without changing your current focus.

<br>

### 10. Speech Log File Access

Press **Shift+F9 three times** to open the complete speech history file in your default text editor for reviewing, searching, or copying.

<br>

### 11. URL History Management

Press **CTRL+Shift+A three times** to open the URL History dialog. This dialog stores every URL and hyperlink you copy using **CTRL+Shift+A** (single or double tap), so you can access them again later.

<br>

**Using the URL History dialog:**

- Press **Enter** on any item to copy that URL to the clipboard. NVDA will announce the copied URL.
- Press the **Applications key** or right-click to open the context menu.
- From the context menu, you can **Pin**, **Edit**, **Move Up**, **Move Down**, or **Delete** an item.
- Press the **Delete** key to quickly remove the selected item.
- Use **Clear** from the context menu to delete all items that are not pinned.

<br>

**Pinning and display names:**

- Pinned items are protected from automatic cleanup and from the Clear command.
- You can assign a custom display name to any URL by choosing **Edit**. This is especially useful for URLs you use often, making them easier to identify in the list.

<br>

**Storage limit:**

- The URL history is limited to **300 items**. When the limit is reached, the oldest non-pinned items are automatically removed to make room for new ones.

<br>

### 12. Smart Context Awareness

When you are typing in editable fields, simpleCopy does not interfere. Commands only activate when they are useful, preserving your normal workflow.

<br><br>

## Support Me

If this add-on helps you work more efficiently, consider supporting future development with a small donation.

<br>

[![Support me](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

<br>

Your support helps keep this project alive and improving.

<br>

&copy; 2026 Chai Chaimee NVDA Add-on Released under GNU General Public License