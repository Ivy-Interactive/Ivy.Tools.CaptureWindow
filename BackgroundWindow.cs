using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Ivy.Tools.CaptureWindow;

[SupportedOSPlatform("windows")]
public class BackgroundWindow : IDisposable
{
    private IntPtr _hwnd = IntPtr.Zero;
    private readonly string _className = "CaptureWindowBackground_" + Guid.NewGuid().ToString("N");
    private Win32Api.WndProc? _wndProc;
    private bool _disposed = false;

    public IntPtr Handle => _hwnd;

    public bool Create(Win32Api.RECT bounds, uint backgroundColor)
    {
        try
        {
            // Register window class
            _wndProc = WndProc;
            var wndClassEx = new Win32Api.WNDCLASSEX
            {
                cbSize = (uint)Marshal.SizeOf<Win32Api.WNDCLASSEX>(),
                style = 0,
                lpfnWndProc = _wndProc,
                cbClsExtra = 0,
                cbWndExtra = 0,
                hInstance = Win32Api.GetModuleHandle(null),
                hIcon = IntPtr.Zero,
                hCursor = IntPtr.Zero,
                hbrBackground = Win32Api.CreateSolidBrush(backgroundColor & 0x00FFFFFF), // Remove alpha channel for GDI
                lpszMenuName = null,
                lpszClassName = _className,
                hIconSm = IntPtr.Zero
            };

            if (!Win32Api.RegisterClassEx(ref wndClassEx))
            {
                return false;
            }

            int width = bounds.Right - bounds.Left;
            int height = bounds.Bottom - bounds.Top;
            
            // Create window (without TOPMOST so it stays behind)
            _hwnd = Win32Api.CreateWindowEx(
                Win32Api.WS_EX_TOOLWINDOW | Win32Api.WS_EX_TOPMOST,
                _className,
                "Background",
                Win32Api.WS_POPUP | Win32Api.WS_VISIBLE,
                bounds.Left,
                bounds.Top,
                width,
                height,
                IntPtr.Zero,
                IntPtr.Zero,
                Win32Api.GetModuleHandle(null),
                IntPtr.Zero
            );

            if (_hwnd == IntPtr.Zero)
            {
                return false;
            }

            // Show the window first
            Win32Api.ShowWindow(_hwnd, Win32Api.SW_SHOW);
            
            // Position it as topmost initially so it's visible
            Win32Api.SetWindowPos(_hwnd, new IntPtr(-1), // HWND_TOPMOST
                0, 0, 0, 0, 
                Win32Api.SWP_NOMOVE | Win32Api.SWP_NOSIZE | Win32Api.SWP_NOACTIVATE);
            
            // Force window to paint
            Win32Api.UpdateWindow(_hwnd);
            
            return true;
        }
        catch
        {
            Dispose();
            return false;
        }
    }

    private IntPtr WndProc(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        const uint WM_PAINT = 0x000F;
        const uint WM_ERASEBKGND = 0x0014;
        
        if (msg == WM_ERASEBKGND)
        {
            // Return non-zero to indicate we erased the background
            return new IntPtr(1);
        }
        
        return Win32Api.DefWindowProc(hwnd, msg, wParam, lParam);
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            if (_hwnd != IntPtr.Zero)
            {
                Win32Api.DestroyWindow(_hwnd);
                _hwnd = IntPtr.Zero;
            }

            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }

    ~BackgroundWindow()
    {
        Dispose();
    }
}