<p align="center">
  <img src="https://www.nvaccess.org/files/nvda/documentation/userGuide/images/nvda.ico" alt="NVDA Logo" width="120">
  <br>
  <h1 align="center">Simple Copy NVDA Add-on</h1>
  <p align="center">Enhanced clipboard management and data copying for NVDA</p>
</p>

<p align="center">
  <b>Author:</b> chai chaimee<br>
  <b>URL:</b> https://github.com/chaichaimee/simpleCopy
</p>

---

## ?? Description

**Simple Copy** is an NVDA (NonVisual Desktop Access) add-on designed to significantly improve and extend clipboard management and data copying capabilities for screen reader users.

This add-on provides streamlined workflows for common tasks:
- Appending selected text to existing clipboard content with a single command
- Copying browser URLs and hyperlink addresses with intelligent browser detection
- Utilizing an intuitive **single-tap/double-tap** system for accessing multiple functions from the same keystrokes

The add-on integrates seamlessly with NVDA and supports all major browsers including Chrome, Firefox, Edge, Safari, Opera, and Brave. It's designed specifically for compatibility with the NVDA 2025.x API.

---

## ?? Hot Keys

The add-on uses an efficient **tap-based system** where single and double taps of the same keystroke perform different but related functions.

### `Control+Shift+A` - URL Management

```
Single Tap (press once):    Copy the current browser URL to clipboard
Double Tap (press twice):   Copy the URL of the focused hyperlink
```

> **Note:** When pressed in an editable text field, the original keystroke is passed through to the application.

### `Control+Shift+C` - Clipboard Management

```
Single Tap (press once):    Copy or append selected text to clipboard
Double Tap (press twice):   Clear all clipboard content
```

> **Note:** If no text is selected when using the append function, the original keystroke is passed through.

### ? Tap Timing
For a double tap to be recognized, the second press must occur within **0.5 seconds** of the first press. If more time elapses, NVDA will treat them as separate single taps.

---

## ? Features

### 1. Smart Browser URL Copying
- **Gesture:** `Control+Shift+A` (Single Tap)
- Copies the complete URL of the current webpage to clipboard
- Automatically detects browser environment (Chrome, Firefox, Edge, Opera, Safari, Brave)
- Passes through original keystroke in editable fields

### 2. Hyperlink URL Extraction
- **Gesture:** `Control+Shift+A` (Double Tap)
- Extracts URL from hyperlinks using multiple methods:
  - Standard `obj.value` property access
  - UIA (UI Automation) fallback for modern browsers
  - IAccessible fallback for legacy support
  - Parent element searching (up to 5 levels)
- Works only in browser contexts

### 3. Text Append with Intelligent Formatting
- **Gesture:** `Control+Shift+C` (Single Tap)
- **Appends to existing content:** Selected text is appended with proper line separation
- **Creates new content:** If clipboard is empty, text is copied as new content
- **Format normalization:** Automatically normalizes line endings (CRLF, CR, LF)
- **Non-text protection:** Detects and prevents appending to non-text clipboard content

### 4. Clipboard Clearing
- **Gesture:** `Control+Shift+C` (Double Tap)
- Clears all content from Windows clipboard
- Provides audio confirmation ("Clean")
- Useful for security or preparing clipboard for new content

### 5. Advanced Text Selection Detection
- Uses `makeTextInfo(textInfos.POSITION_SELECTION)` from treeInterceptor or focus objects
- Fallback method simulating Ctrl+C when primary method fails
- Preserves original clipboard content during extraction
- Normalizes text by removing non-printable characters

### 6. Browser Detection & Context Awareness
- Automatically identifies when in a web browser
- Adjusts URL-related functions based on current context
- Passes through original keystrokes in non-browser contexts
- Uses multiple URL retrieval methods (standard API, treeInterceptor, UIA, IAccessible)

### 7. Robust Link URL Resolution
- Checks if focused object has `Role.LINK`
- Searches parent hierarchy for link elements if not direct link
- Attempts multiple property access methods for maximum compatibility
- Provides clear audio feedback for all operations

---

## ?? Compatibility

- **NVDA Version:** 2025.x and compatible versions
- **Browsers Supported:**
  - Google Chrome
  - Mozilla Firefox
  - Microsoft Edge / MSEdge
  - Opera
  - Safari
  - Brave
- **Operating Systems:** Microsoft Windows
- **License:** GNU General Public License

---

## ?? Technical Notes

This add-on is built specifically for the NVDA 2025.x API. It uses proper Python tools and linters to ensure code quality and follows NVDA add-on development best practices.

The source code is structured into modular components:
- `__init__.py` - Main plugin initialization and script handlers
- `clipboard_utils.py` - Clipboard-related operations
- `url_utils.py` - URL handling and browser detection

---

## ?? License

Copyright (C) 2025 CHAI CHAIMEE

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

NVDA (NonVisual Desktop Access) is a free and open-source screen reader for Microsoft Windows.
