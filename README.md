# MGPad Run Guide

MGPad is a WPF text editor targeting Windows that supports multiple document types (plain text, rich text, and Markdown) along with standard editing commands and status indicators.

## Project layout

| Path | Purpose |
| --- | --- |
| `MGPad/MGPad.csproj` | .NET 8 Windows desktop project that enables WPF (`UseWPF`) and Windows-only targeting. |
| `MGPad/MainWindow.xaml` | Defines the window layout: menu bar with File/Edit entries, toolbar buttons, the main `RichTextBox`, and the status bar labels. |
| `MGPad/MainWindow.xaml.cs` | Implements file management, document state tracking, save/load logic, and command handlers for the editor.

## Prerequisites

1. **Windows 10/11** – The project targets `net8.0-windows`, requires `UseWPF`, and enables Windows-specific APIs.
2. **.NET 8 SDK** – Provides the `dotnet` CLI for restore/build/run operations.
3. **Visual Studio 2022 (optional)** – For a full IDE experience, open `MGPad.sln`.

## First-time setup

1. Clone the repository:
   ```bash
   git clone <repo-url>
   cd mgpad
   ```
2. Restore dependencies and build:
   ```bash
   dotnet restore
   dotnet build MGPad/MGPad.csproj
   ```

## Running the app

### Option 1 – `dotnet run`

```bash
dotnet run --project MGPad/MGPad.csproj
```

The command launches the MGPad window. Because the project produces a WinExe, the app will open directly on Windows.

### Option 2 – Visual Studio

1. Double-click `MGPad.sln`.
2. Set `MGPad` as the startup project if necessary.
3. Press <kbd>F5</kbd> (Debug) or <kbd>Ctrl</kbd>+<kbd>F5</kbd> (Run without debugging).

## Using MGPad

## Feature overview

* **Multi-format editing** – Create, open, and save plain text, rich text, Markdown, and OpenDocument Text (`.odt`) files with automatic extension selection in the dialogs. Unsaved changes are tracked and prompt before being discarded.
* **Recent documents** – The File menu keeps a list of the last ten opened files for quick access.
* **Exports** – Export the current document to PDF or OpenDocument Text from the File menu while preserving rich text styling.
* **Markdown mode** – Toggle a side-by-side Markdown preview (View → Markdown Mode) that updates after you stop typing for a moment.
* **Rich formatting tools** – Bold, italic, underline, strikethrough, monospaced toggles, and font family/size pickers apply to the current selection. Buttons exist for bulleted and numbered lists plus left/center/right alignment controls.
* **Language and theme toggles** – A JP/EN switch (or <kbd>Ctrl</kbd>+<kbd>J</kbd>) flips between available English and Japanese input languages; a Night/Day toggle adjusts the editor, preview, and chrome colors.
* **Timestamp insertion** – Insert a formatted timestamp via the toolbar button or <kbd>Alt</kbd>+<kbd>Shift</kbd>+<kbd>T</kbd>.
* **Selection math helper** – The editor context menu can sum all numbers in the current selection and append the result.
* **Zoom controls** – Status bar buttons zoom the editor in or out in 10% steps and show the current percentage.
* **Status bar indicators** – The footer shows the active file name (with an asterisk when unsaved) and the current language indicator.

## Using MGPad

* **Menus & shortcuts** – The File menu exposes New/Open/Save/Save As and Exit entries with the expected Ctrl-based shortcuts. The Edit menu wires Undo/Redo, Cut/Copy/Paste, and Select All commands directly to the `RichTextBox` control, and `Ctrl+Shift+S` is bound globally to Save As.
* **Toolbar & editor** – A toolbar hosts the formatting toggles and utilities described above, and the centered `RichTextBox` offers auto vertical scrolling.
* **Status bar** – The bottom bar shows the current file name (with an asterisk when unsaved changes exist) and a language indicator defaulting to "EN".
* **File formats** – When opening or saving, the dialogs provide filters for `.txt`, `.rtf`, `.md`, and `.odt`. Rich text documents load/save using RTF streams, while plain text and Markdown are handled via string reads/writes. The code automatically selects the proper extension and filter index based on the active document type and prompts before discarding unsaved changes.

When you close the window, MGPad confirms whether to save any dirty document before exiting.
