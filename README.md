# Easier Paper Work

A Python desktop tool for comparing two files paragraph by paragraph and helping clean up mismatched revisions.

在作者们和审稿人之间辗转，在response letter和正文间辗转，我已厌倦，，，逐步更新，如果codex不继续砍额度的话

## Ver2 features

- Select two files from the GUI instead of pasting raw text.
- Support `docx` and `txt` input, with `docx` as the main workflow.
- Match highly similar paragraphs by natural paragraph units.
- Preview the standard paragraph and target paragraph side by side in the UI.
- Count pending differences in the right sidebar.
- Correct one item at a time or apply all corrections in one click.
- Choose whether file A or file B is treated as the standard.
- Export a corrected copy instead of overwriting the original file.

## Files

- `text_compare_tool.py`: main Tkinter GUI application.
- `launch_text_compare_tool.bat`: launcher for running the Python script directly.
- `TextCompareTool.spec`: PyInstaller spec file used for packaging.
- `build_exe.bat`: build script for generating the executable.

## Run locally

```bat
python text_compare_tool.py
```

Or double-click `launch_text_compare_tool.bat`.

## Build the exe

```bat
build_exe.bat
```

The packaged executable is written to the parent directory.
