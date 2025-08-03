using Spectre.Console;
using Spectre.Console.Cli;
using System.ComponentModel;
using System.Runtime.Versioning;

namespace Ivy.Tools.CaptureWindow.Commands;

[SupportedOSPlatform("windows")]
public class CaptureCommand : Command<CaptureCommand.Settings>
{
    public class Settings : CommandSettings
    {
        [CommandOption("-c|--class")]
        [Description("Window class identifier")]
        public string? ClassId { get; set; }

        [CommandOption("-t|--title")]
        [Description("Window title")]
        public string? Title { get; set; }

        [CommandOption("-o|--output")]
        [Description("Output filename")]
        public string? Output { get; set; }

        [CommandOption("-i|--interactive")]
        [Description("Interactive window selection")]
        public bool Interactive { get; set; } = true;

        [CommandOption("-r|--resize")]
        [Description("Resize window before capture (format: widthxheight, e.g., 1920x1080)")]
        public string? Resize { get; set; }

        [CommandOption("-m|--margins")]
        [Description("Capture margins in pixels for transparent shadows (format: left,top,right,bottom or single value for all sides)")]
        public string? Margins { get; set; }

        [CommandOption("-b|--background")]
        [Description("Background color (transparent, black, white, or hex color like #FF0000). Default: transparent")]
        public string? Background { get; set; } = "transparent";


    }

    public override int Execute(CommandContext context, Settings settings)
    {
        try
        {
            WindowInfo? selectedWindow = null;

            if (!string.IsNullOrEmpty(settings.ClassId) || !string.IsNullOrEmpty(settings.Title))
            {
                settings.Interactive = false;
            }

            if (settings.Interactive)
            {
                selectedWindow = SelectWindowInteractively();
                if (selectedWindow == null)
                {
                    AnsiConsole.MarkupLine("[red]No window selected.[/]");
                    return 1;
                }
            }

            string filename = DetermineFilename(settings, selectedWindow);
            var resizeParams = ParseResizeParameter(settings.Resize);
            var margins = ParseMargins(settings.Margins);
            var backgroundColor = ParseBackground(settings.Background);

            bool success;
            if (selectedWindow != null)
            {
                success = WindowCapture.CaptureWindow(selectedWindow, filename, resizeParams, margins, backgroundColor);
                AnsiConsole.MarkupLine($"[green]Captured window: {selectedWindow.Title}[/]");
            }
            else
            {
                success = WindowCapture.CaptureWindow(settings.ClassId!, settings.Title!, filename, resizeParams, margins, backgroundColor);
            }

            if (!success)
            {
                AnsiConsole.MarkupLine("[red]Window not found or capture failed.[/]");
                return 2;
            }

            AnsiConsole.MarkupLine($"[green]Screenshot saved to: {filename}[/]");
            return 0;
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteException(ex);
            return 3;
        }
    }

    private static WindowInfo? SelectWindowInteractively()
    {
        var windows = WindowEnumerator.GetVisibleWindows();
        
        if (windows.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]No visible windows found.[/]");
            return null;
        }

        var selection = AnsiConsole.Prompt(
            new SelectionPrompt<WindowInfo>()
                .Title("Select a window to capture:")
                .PageSize(10)
                .MoreChoicesText("[grey](Move up and down to reveal more windows)[/]")
                .AddChoices(windows)
                .UseConverter(window => $"{window.ClassName.EscapeMarkup()} - {window.Title.EscapeMarkup()}"));

        return selection;
    }

    private static string DetermineFilename(Settings settings, WindowInfo? window)
    {
        string filename;
        
        if (!string.IsNullOrEmpty(settings.Output))
        {
            filename = settings.Output;
        }
        else
        {
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            if (window != null)
            {
                var sanitized = string.Join("_", window.Title.Split(Path.GetInvalidFileNameChars()));
                filename = Path.Combine(desktopPath, $"{sanitized}.png");
            }
            else if (!string.IsNullOrEmpty(settings.Title))
            {
                var sanitized = string.Join("_", settings.Title.Split(Path.GetInvalidFileNameChars()));
                filename = Path.Combine(desktopPath, $"{sanitized}.png");
            }
            else if (!string.IsNullOrEmpty(settings.ClassId))
            {
                filename = Path.Combine(desktopPath, $"{settings.ClassId}.png");
            }
            else
            {
                filename = Path.Combine(desktopPath, "screenshot.png");
            }
        }

        return GetAvailableFilename(filename);
    }

    private static string GetAvailableFilename(string originalPath)
    {
        if (!File.Exists(originalPath))
        {
            return originalPath;
        }

        var directory = Path.GetDirectoryName(originalPath) ?? "";
        var nameWithoutExtension = Path.GetFileNameWithoutExtension(originalPath);
        var extension = Path.GetExtension(originalPath);

        int counter = 1;
        string newPath;
        
        do
        {
            var newName = $"{nameWithoutExtension}_{counter}{extension}";
            newPath = Path.Combine(directory, newName);
            counter++;
        }
        while (File.Exists(newPath));

        return newPath;
    }

    private static (int width, int height)? ParseResizeParameter(string? resize)
    {
        if (string.IsNullOrEmpty(resize))
            return null;

        var parts = resize.ToLowerInvariant().Split('x');
        if (parts.Length != 2)
        {
            AnsiConsole.MarkupLine("[red]Invalid resize format. Use: widthxheight (e.g., 1920x1080)[/]");
            return null;
        }

        if (int.TryParse(parts[0], out int width) && int.TryParse(parts[1], out int height))
        {
            if (width > 0 && height > 0)
            {
                return (width, height);
            }
        }

        AnsiConsole.MarkupLine("[red]Invalid resize dimensions. Width and height must be positive numbers.[/]");
        return null;
    }

    private static (int left, int top, int right, int bottom) ParseMargins(string? margins)
    {
        // Default margins for Windows 11 drop shadows
        var defaultMargins = (left: 50, top: 50, right: 50, bottom: 50);
        
        if (string.IsNullOrEmpty(margins))
            return defaultMargins;

        var parts = margins.Split(',');
        
        if (parts.Length == 1)
        {
            // Single value for all sides
            if (int.TryParse(parts[0].Trim(), out int value) && value >= 0)
            {
                return (value, value, value, value);
            }
        }
        else if (parts.Length == 4)
        {
            // left,top,right,bottom format
            if (int.TryParse(parts[0].Trim(), out int left) && left >= 0 &&
                int.TryParse(parts[1].Trim(), out int top) && top >= 0 &&
                int.TryParse(parts[2].Trim(), out int right) && right >= 0 &&
                int.TryParse(parts[3].Trim(), out int bottom) && bottom >= 0)
            {
                return (left, top, right, bottom);
            }
        }

        AnsiConsole.MarkupLine("[yellow]Invalid shadow margins format. Using defaults. Use: single value or left,top,right,bottom[/]");
        return defaultMargins;
    }

    private static System.Drawing.Color? ParseBackground(string? background)
    {
        if (string.IsNullOrEmpty(background) || background.Equals("transparent", StringComparison.OrdinalIgnoreCase))
            return null; // Transparent (use dual capture)

        if (background.Equals("black", StringComparison.OrdinalIgnoreCase))
            return System.Drawing.Color.Black;

        if (background.Equals("white", StringComparison.OrdinalIgnoreCase))
            return System.Drawing.Color.White;

        // Try to parse hex color
        if (background.StartsWith("#") && background.Length == 7)
        {
            if (uint.TryParse(background.Substring(1), System.Globalization.NumberStyles.HexNumber, null, out uint color))
            {
                return System.Drawing.Color.FromArgb((int)(0xFF000000 | color));
            }
        }

        AnsiConsole.MarkupLine("[yellow]Invalid background color format. Using transparent. Use: transparent, black, white, or #RRGGBB[/]");
        return null; // Default to transparent
    }

}