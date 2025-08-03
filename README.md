# Ivy.Tools.CaptureWindow

A command-line tool for capturing Windows application screenshots with interactive window selection.

When background option is set to `transparent` (default), the resulting image will have transparent shadows due to dual capture technique.

## Installation

Install as a global .NET tool:

```bash
dotnet tool install -g Ivy.Tools.CaptureWindow
```

## Usage

### Interactive Mode (Default)

Simply run the tool to get an interactive window selector:

```bash
capture-window
```

### Specific Window by Title

```bash
capture-window capture --title "Notepad" --output "notepad.png"
```

### Specific Window by Class

```bash
capture-window capture --class "Notepad" --decorations
```

### List All Windows

```bash
capture-window list
```

## Command Options

### `capture` (default command)

- `-c, --class <CLASS>`: Window class identifier
- `-t, --title <TITLE>`: Window title
- `-o, --output <FILE>`: Output filename (defaults to desktop with window title + .png)
- `-i, --interactive`: Interactive window selection (default: true)
- `-r, --resize <SIZE>`: Resize window before capture (format: widthxheight, e.g., 1920x1080)
- `-m, --margins <MARGINS>`: Capture margins in pixels for transparent shadows (format: left,top,right,bottom or single value for all sides)
- `-b, --background <COLOR>`: Background color (transparent, black, white, or hex color like #FF0000). Default: transparent

### `list`

Lists all visible windows with their class names and titles.

## Examples

```bash
# Interactive selection (default)
capture-window

# Capture specific window with custom output
capture-window capture --title "Visual Studio Code" --output "vscode.png"

# Capture with window resize
capture-window capture --title "Notepad" --resize "1920x1080"

# Capture with custom margins for shadows
capture-window capture --title "Calculator" --margins "30,30,30,30"

# Capture with colored background instead of transparent
capture-window capture --title "Paint" --background "white"

# List all windows
capture-window list

# Capture by class name
capture-window capture --class "Chrome_WidgetWin_1"
```

## Supported Image Formats

- PNG (default)
- JPEG/JPG
- BMP
- GIF
- TIFF

The format is automatically determined by the file extension.

## Requirements

- Windows OS
- .NET 8.0 or later