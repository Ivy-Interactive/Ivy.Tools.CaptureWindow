using Ivy.Tools.CaptureWindow.Commands;
using Spectre.Console.Cli;
using System.Runtime.Versioning;

[SupportedOSPlatform("windows")]
static CommandApp CreateApp()
{
    var app = new CommandApp();

    app.Configure(config =>
    {
        config.SetApplicationName("capture-window");
        config.SetApplicationVersion("1.0.0");
        
        config.AddCommand<CaptureCommand>("capture")
            .WithDescription("Capture a window screenshot")
            .WithExample(new[] { "capture", "--interactive" })
            .WithExample(new[] { "capture", "--title", "Notepad", "--output", "notepad.png" })
            .WithExample(new[] { "capture", "--class", "Notepad" })
            .WithExample(new[] { "capture", "--resize", "1920x1080" })
            .WithExample(new[] { "capture", "--shadows" });
            
        config.AddCommand<ListCommand>("list")
            .WithDescription("List all visible windows")
            .WithExample(new[] { "list" });

        config.AddCommand<CaptureCommand>("")
            .WithDescription("Capture a window screenshot (default)");
    });

    return app;
}

if (!OperatingSystem.IsWindows())
{
    Console.WriteLine("This tool is only supported on Windows.");
    return 1;
}

var app = CreateApp();
return app.Run(args);
