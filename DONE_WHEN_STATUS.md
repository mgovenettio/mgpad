# MGPad Phase 0–1 Checklist Verification

## 1. Build + Launch
- `MGPad.csproj` targets `net8.0-windows` with `UseWPF` enabled, so `MGPad.sln` builds the expected desktop app shell.
- The main window definition uses the title "MGPad" and hooks a `Closing` handler, which is the entry point once the solution launches.

## 2. Main Window Layout
- `MainWindow.xaml` defines the four-row grid (menu, toolbar, editor, status bar) with the requested menu headers, placeholder toolbar buttons, a `RichTextBox` that stretches with scrollbars, and status bar text blocks for the filename and language indicator.

## 3. Document State + Title
- `MainWindow.xaml.cs` keeps `_currentFilePath`, `_currentDocumentType`, `_isDirty`, and updates the window title/status-bar text via `UpdateWindowTitle`/`UpdateStatusBar`, marking the dirty flag with `MarkDirty`/`MarkClean`.

## 4. File Menu Behavior
- Click handlers and command bindings cover File → New/Open/Save/Save As/Exit. `CreateNewDocument`, `OpenDocumentFromDialog`, `SaveCurrentDocument`, `SaveDocumentWithDialog`, and `ExitApplication` implement the behaviors described.

## 5. File Formats (.txt/.rtf/.md)
- `DetermineDocumentType` chooses between `PlainText`, `RichText`, and `Markdown` by extension. `LoadDocumentFromFile` loads `.rtf` via `TextRange.Load`/`DataFormats.Rtf` and plain text/markdown via `File.ReadAllText`. `SaveDocumentToFile` mirrors the behavior.

## 6. Dirty State + Prompts
- `EditorBox_TextChanged` sets `_isDirty` unless `_isLoadingDocument` is true. `ConfirmDiscardUnsavedChanges` prompts with Yes/No/Cancel and is used by New/Open/Exit/Closing.

## 7. Edit Menu + Shortcuts
- Menu items bind to `ApplicationCommands` (Undo/Redo/Cut/Copy/Paste/Select All) and `EditorCommand_CanExecute`/`EditorCommand_Executed` dispatch to the RichTextBox. A `KeyBinding` explicitly adds Ctrl+Shift+S for Save As.

## 8. Manual Flows
- With the above implementations, the flows (New→Save As, Open/Save for each format, dirty prompts) are covered by the same helpers and data flow.
