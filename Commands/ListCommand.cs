using Spectre.Console;
using Spectre.Console.Cli;
using System.Runtime.Versioning;

namespace Ivy.Tools.CaptureWindow.Commands;

[SupportedOSPlatform("windows")]
public class ListCommand : Command
{
    public override int Execute(CommandContext context)
    {
        var windows = WindowEnumerator.GetVisibleWindows();
        
        if (windows.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]No visible windows found.[/]");
            return 0;
        }

        var table = new Table();
        table.AddColumn("Class Name");
        table.AddColumn("Window Title");

        foreach (var window in windows)
        {
            table.AddRow(
                window.ClassName.EscapeMarkup(),
                window.Title.EscapeMarkup()
            );
        }

        AnsiConsole.Write(table);
        return 0;
    }
}