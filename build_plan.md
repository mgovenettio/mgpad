# MGPad Build Plan (v0.1–v0.3)

## Goal

A minimalistic personal text editor for Windows:
- Slightly more powerful than Notepad
- Only the features I actually use
- Optimized for English/Japanese writing
- Supports simple formatting, markdown preview, PDF export, and timestamps

---

## Tech Stack

- Language: C#
- Framework: .NET 8
- UI: WPF desktop app
- Main control: RichTextBox (for rich text) + TextBox/FlowDocument for modes as needed

---

## Phase 0: Project skeleton

1. Create a new WPF .NET 8 app called `MGPad`.
2. MainWindow layout:
   - Menu bar at top (File, Edit, View, Format, Tools, Help).
   - Toolbar below the menu.
   - Central editor region.
   - Status bar at the bottom.
3. Central editor in v0.1:
   - Single RichTextBox docked in the main area.

---

## Phase 1: Core editor + file I/O

4. Implement File menu:
   - New: clears editor and resets state.
   - Open: supports `.txt`, `.rtf`, `.md` (for now treat `.md` as plain text).
   - Save: saves to current filename.
   - Save As: lets user pick `.txt`, `.rtf`, or `.md`.
5. Implement basic document handling:
   - `.txt` → plain text to/from RichTextBox.
   - `.rtf` → use RichTextBox built-in RTF load/save.
   - `.md` → plain text, no formatting at this stage.
6. Track "dirty" state:
   - Mark document as modified on changes.
   - Prompt to save before closing or opening a new file.

7. Edit menu:
   - Undo, Redo
   - Cut, Copy, Paste, Select All
   - Keyboard shortcuts wired up.

---

## Phase 2: Formatting (bold/underline only)

8. Add toolbar buttons:
   - Bold (B)
   - Underline (U)
9. Add keyboard shortcuts:
   - Ctrl+B for bold toggle
   - Ctrl+U for underline toggle
10. Implement formatting behavior:
   - If there is a selection, toggle bold/underline on the selected text.
   - If no selection, toggle the current typing style.
11. For `.txt` and `.md` files:
   - Decide behavior:
     - Either disable formatting buttons, **or**
     - Allow formatting visually but show a warning that `.txt` / `.md` won’t store formatting.
   - For now, simplest: disable formatting buttons for `.txt` / `.md` and only enable for `.rtf`.

---

## Phase 3: JP/EN input toggle (Mode A)

12. Status bar:
   - Right side: display current input language code (`EN` or `JP`).
13. Implement language detection:
   - Query current input language from Windows.
   - Map to a friendly code ("EN", "JP").
14. Implement toggle action:
   - Toolbar button "JP/EN" that flips between:
     - Preferred English input language
     - Preferred Japanese IME language
   - Only these two languages are considered.
15. Keyboard shortcut:
   - Assign a shortcut like Ctrl+J to trigger the same toggle.
16. After toggling:
   - Update the status bar indicator.

---

## Phase 4: Markdown mode with split view

17. Update layout to support two modes:
   - Normal mode: single editor (RichTextBox).
   - Markdown mode: split view (left editor, right preview).
18. View menu:
   - Checkbox item: "Markdown Mode".
19. In Markdown mode:
   - Left pane: plain text editor (can reuse RichTextBox or TextBox).
   - Right pane: Markdown preview.
20. Markdown preview behavior:
   - On text change (with a short debounce), re-render Markdown and show it.
   - Start with headings, lists, bold/italic, code blocks.
21. File integration:
   - If opening a `.md` file, automatically enable Markdown mode.
   - Saving a `.md` file always saves the raw text from the left pane.

---

## Phase 5: PDF export (simple)

22. Add File → "Export as PDF…" menu item.
23. Implement export pipeline:
   - For `.rtf` or rich documents:
     - Convert document content into a simple sequence of paragraphs with formatting.
   - Generate a PDF using a .NET PDF library.
24. Layout:
   - Single column of text.
   - Standard margins.
   - Preserve bold/underline.
25. Show confirmation or error messages for export.

---

## Phase 6: Timestamp insertion

26. Tools menu:
   - Add "Insert Timestamp" item.
27. Toolbar:
   - Add "TS" button for inserting timestamp.
28. Keyboard shortcut:
   - Assign Ctrl+T (or other free combo).
29. Behavior:
   - Insert timestamp at caret using default format: `YYYY-MM-DD HH:MM`.
30. Optional later: make timestamp format configurable in a simple Settings dialog.

---

## Phase 7: Small quality-of-life polish

31. Status bar:
   - Left: current file name (or "Untitled").
   - Center: word count (optional).
   - Right: input language (EN/JP).
32. Remember:
   - Last window size and position.
   - Last used open/save folder.
33. Optional:
   - Recent files list under File menu.
   - Simple dark mode toggle.
   - Font and font size preferences.
