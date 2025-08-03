using System.Text;
using System.Runtime.Versioning;

namespace Ivy.Tools.CaptureWindow;

public record WindowInfo(IntPtr Handle, string ClassName, string Title)
{
    public override string ToString() => $"{ClassName} - {Title}";
    
    public bool IsForeground => Win32Api.GetForegroundWindow() == Handle;
    public bool IsMinimized => Win32Api.IsIconic(Handle);
    
    public int ZOrder { get; init; }
}

[SupportedOSPlatform("windows")]
public static class WindowEnumerator
{
    public static List<WindowInfo> GetVisibleWindows()
    {
        var windows = new List<WindowInfo>();
        var zOrder = 0;
        
        Win32Api.EnumWindows((hWnd, lParam) =>
        {
            if (Win32Api.IsWindowVisible(hWnd))
            {
                var className = new StringBuilder(80);
                var title = new StringBuilder(80);
                
                Win32Api.GetClassName(hWnd, className, className.Capacity);
                Win32Api.GetWindowText(hWnd, title, title.Capacity);
                
                var titleStr = title.ToString();
                var classNameStr = className.ToString();
                
                if (!string.IsNullOrWhiteSpace(titleStr) && !classNameStr.Contains("Grammarly.Desktop.exe"))
                {
                    windows.Add(new WindowInfo(hWnd, classNameStr, titleStr) { ZOrder = zOrder++ });
                }
            }
            
            return true;
        }, IntPtr.Zero);
        
        return SortWindowsByImportance(windows);
    }

    private static List<WindowInfo> SortWindowsByImportance(List<WindowInfo> windows)
    {
        return windows
            .OrderBy(w => w.IsForeground ? 0 : 1)           // Foreground first
            .ThenBy(w => w.IsMinimized ? 1 : 0)             // Non-minimized before minimized
            .ThenBy(w => w.ZOrder)                          // Then by Z-order (top to bottom)
            .ThenBy(w => w.Title.ToLowerInvariant())        // Finally alphabetical
            .ToList();
    }
}