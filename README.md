# PPTX Editor

Professional PowerPoint-style Presentation Editor built with C++ and GDI+.

## Features
- Ribbon UI (Home / Insert / Design tabs)
- Multiple slide management with thumbnail sidebar
- Shapes: Rectangle, Rounded Rect, Ellipse, Triangle, Diamond, Star
- Text formatting: Bold, Italic, Underline, Font size, Alignment
- Drag & drop elements
- Undo / Redo (Ctrl+Z / Ctrl+Y)
- Color themes

## Build Locally (MinGW)

```bash
g++ -std=c++17 -O2 -o pptx_editor.exe main.cpp ^
    -lgdiplus -lole32 -loleaut32 -luuid -lshlwapi -lcomctl32 ^
    -mwindows -static-libgcc -static-libstdc++
```

## Build via GitHub Actions

Push to `main` or `master` branch → Actions tab → Download EXE from Artifacts.

To create a GitHub Release with the EXE attached, push a tag:
```bash
git tag v1.0.0
git push origin v1.0.0
```

## Controls
| Action | How |
|--------|-----|
| Select element | Click on it |
| Move element | Click + drag |
| Type text | Select element → type |
| Delete char | Backspace |
| Undo | Ctrl+Z |
| Redo | Ctrl+Y |
| Zoom | Ctrl + Scroll |
| Scroll slides | Scroll in sidebar |
