using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Spectre.Console;

namespace Ivy.Tools.CaptureWindow;

[SupportedOSPlatform("windows6.1")]
public static class WindowCapture
{
    public static bool CaptureWindow(WindowInfo window, string filename, (int width, int height)? resize = null, (int left, int top, int right, int bottom) margins = default(ValueTuple<int, int, int, int>), System.Drawing.Color? backgroundColor = null)
    {
        // Validate window exists and is visible
        if (!Win32Api.IsWindow(window.Handle))
        {
            return false;
        }

        // Always try to bring window to foreground for consistent capture
        if (!Win32Api.IsWindowVisible(window.Handle))
        {
            // Try to restore and bring to foreground
            Win32Api.ShowWindow(window.Handle, Win32Api.SW_RESTORE);
            Win32Api.SetForegroundWindow(window.Handle);
            Thread.Sleep(200); // Give time for window to become visible
            
            // Check if it's visible now
            if (!Win32Api.IsWindowVisible(window.Handle))
            {
                return false;
            }
        }
        else
        {
            // Window is visible, bring to front for clean capture
            Win32Api.SetForegroundWindow(window.Handle);
            Thread.Sleep(100);
        }

        Win32Api.RECT? originalRect = null;
        
        if (resize.HasValue)
        {
            Win32Api.GetWindowRect(window.Handle, out var currentRect);
            originalRect = currentRect;
            
            var (width, height) = resize.Value;
            
            // If using margins, adjust window size so final image matches requested dimensions
            if (margins.left > 0 || margins.top > 0 || margins.right > 0 || margins.bottom > 0)
            {
                width -= (margins.left + margins.right);
                height -= (margins.top + margins.bottom);
                
                // Ensure window dimensions are still positive
                if (width <= 0 || height <= 0)
                {
                    return false;
                }
            }
            
            Win32Api.SetWindowPos(window.Handle, IntPtr.Zero, 
                currentRect.Left, currentRect.Top, width, height, 
                Win32Api.SWP_NOZORDER | Win32Api.SWP_NOACTIVATE);
                
            Thread.Sleep(100);
        }

        try
        {
            Bitmap? bitmap;
            if (backgroundColor.HasValue)
            {
                // Single capture with specified background color
                bitmap = CaptureWithSingleBackground(window.Handle, margins, backgroundColor.Value);
            }
            else
            {
                // Dual capture technique for transparent shadows
                bitmap = CaptureWindowWithTransparency(window.Handle, margins);
            }

            if (bitmap != null)
            {
                using (bitmap)
                {
                    SaveBitmap(bitmap, filename);
                    return true;
                }
            }
        }
        finally
        {
            if (originalRect.HasValue)
            {
                var rect = originalRect.Value;
                Win32Api.SetWindowPos(window.Handle, IntPtr.Zero,
                    rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top,
                    Win32Api.SWP_NOZORDER | Win32Api.SWP_NOACTIVATE);
            }
        }

        return false;
    }

    public static bool CaptureWindow(string classId, string title, string filename, (int width, int height)? resize = null, (int left, int top, int right, int bottom) margins = default, System.Drawing.Color? backgroundColor = null)
    {
        IntPtr hWnd = Win32Api.FindWindow(
            string.IsNullOrEmpty(classId) ? null : classId,
            string.IsNullOrEmpty(title) ? null : title
        );

        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        var window = new WindowInfo(hWnd, classId ?? "", title ?? "");
        return CaptureWindow(window, filename, resize, margins, backgroundColor);
    }


    private static void SaveBitmap(Bitmap bitmap, string filename)
    {
        var format = GetImageFormat(filename);
        
        // For PNG format, preserve transparency
        if (format == ImageFormat.Png && bitmap.PixelFormat == System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        {
            bitmap.Save(filename, ImageFormat.Png);
        }
        else
        {
            bitmap.Save(filename, format);
        }
    }

    private static ImageFormat GetImageFormat(string filename)
    {
        string extension = Path.GetExtension(filename).ToLowerInvariant();
        return extension switch
        {
            ".png" => ImageFormat.Png,
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            ".bmp" => ImageFormat.Bmp,
            ".gif" => ImageFormat.Gif,
            ".tiff" => ImageFormat.Tiff,
            _ => ImageFormat.Png
        };
    }






    private static Bitmap? CaptureWithSingleBackground(IntPtr hWnd, (int left, int top, int right, int bottom) margins, System.Drawing.Color backgroundColor)
    {
        try
        {
            // Get DWM extended bounds for accurate shadow dimensions
            if (Win32Api.DwmGetWindowAttribute(hWnd, Win32Api.DWMWA_EXTENDED_FRAME_BOUNDS, 
                out Win32Api.RECT bounds, Marshal.SizeOf<Win32Api.RECT>()) != 0)
            {
                // Fallback to regular window rect
                Win32Api.GetWindowRect(hWnd, out bounds);
            }

            // Apply margins to expand capture area
            bounds.Left -= margins.left;
            bounds.Top -= margins.top;
            bounds.Right += margins.right;
            bounds.Bottom += margins.bottom;

            int width = bounds.Right - bounds.Left;
            int height = bounds.Bottom - bounds.Top;

            if (width <= 0 || height <= 0)
                return null;

            // Single capture with specified background color
            return CaptureWithSolidBackground(hWnd, bounds, backgroundColor);
        }
        catch
        {
            return null;
        }
    }

    private static Bitmap? CaptureWindowWithTransparency(IntPtr hWnd, (int left, int top, int right, int bottom) margins)
    {
        try
        {
            // Get DWM extended bounds for accurate shadow dimensions
            if (Win32Api.DwmGetWindowAttribute(hWnd, Win32Api.DWMWA_EXTENDED_FRAME_BOUNDS, 
                out Win32Api.RECT bounds, Marshal.SizeOf<Win32Api.RECT>()) != 0)
            {
                // Fallback to regular window rect
                Win32Api.GetWindowRect(hWnd, out bounds);
            }

            // Apply margins to expand capture area
            bounds.Left -= margins.left;
            bounds.Top -= margins.top;
            bounds.Right += margins.right;
            bounds.Bottom += margins.bottom;

            int width = bounds.Right - bounds.Left;
            int height = bounds.Bottom - bounds.Top;

            if (width <= 0 || height <= 0)
                return null;

            // Always use dual capture technique for true transparency
            return CaptureWithDualBackgrounds(hWnd, bounds);
        }
        catch
        {
            return null;
        }
    }

    private static Bitmap? CaptureWithDualBackgrounds(IntPtr hWnd, Win32Api.RECT bounds)
    {
        try
        {
            int width = bounds.Right - bounds.Left;
            int height = bounds.Bottom - bounds.Top;
            
            // First capture with white background
            var whiteCapture = CaptureWithSolidBackground(hWnd, bounds, System.Drawing.Color.White);
            
            if (whiteCapture == null)
                return null;
            
            // Second capture with black background  
            var blackCapture = CaptureWithSolidBackground(hWnd, bounds, System.Drawing.Color.Black);
            
            if (blackCapture == null)
            {
                whiteCapture.Dispose();
                return null;
            }
            
            // Calculate transparency using dual capture algorithm
            var result = CalculateTransparency(blackCapture, whiteCapture, width, height);
            
            whiteCapture.Dispose();
            blackCapture.Dispose();
            
            return result;
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteException(ex);
            return null;
        }
    }
    
    private static Bitmap? CaptureWithSolidBackground(IntPtr hWnd, Win32Api.RECT bounds, System.Drawing.Color backgroundColor)
    {
        try
        {
            // Store original mouse position
            Win32Api.GetCursorPos(out Win32Api.POINT originalPos);
            
            // Move mouse out of capture area (to top-left corner)
            Win32Api.SetCursorPos(0, 0);
            
            // Create a simple background form with solid color
            using (var backgroundForm = new SimpleBackgroundForm())
            {
                if (!backgroundForm.Create(bounds, backgroundColor))
                {
                    // Restore mouse position
                    Win32Api.SetCursorPos(originalPos.X, originalPos.Y);
                    return null;
                }
                
                // Give background form time to render
                Thread.Sleep(500);
                
                // Make sure target window is above the background form
                Win32Api.SetWindowPos(hWnd, new IntPtr(-1), 0, 0, 0, 0, 
                    Win32Api.SWP_NOMOVE | Win32Api.SWP_NOSIZE | Win32Api.SWP_NOACTIVATE);
                
                // DON'T focus the window to avoid blinking cursors
                // Just ensure it's visible without focusing
                Win32Api.ShowWindow(hWnd, Win32Api.SW_SHOWNOACTIVATE);
                
                // Force window to repaint
                Win32Api.UpdateWindow(hWnd);
                
                // Longer delay for everything to settle and repaint
                Thread.Sleep(2000);
                
                // Capture the area
                var result = CaptureFromDesktop(bounds);
                
                // Restore original mouse position
                Win32Api.SetCursorPos(originalPos.X, originalPos.Y);
                
                return result;
            }
        }
        catch
        {
            return null;
        }
    }
    

    
    private static Bitmap? CaptureFromDesktop(Win32Api.RECT bounds)
    {
        try
        {
            int width = bounds.Right - bounds.Left;
            int height = bounds.Bottom - bounds.Top;
            
            var bitmap = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, new System.Drawing.Size(width, height), CopyPixelOperation.SourceCopy);
            }
            return bitmap;
        }
        catch
        {
            return null;
        }
    }

    private static Bitmap? CaptureWindowWithAlpha(IntPtr hWnd, Win32Api.RECT bounds, (int left, int top, int right, int bottom) shadowMargins)
    {
        try
        {
            int width = bounds.Right - bounds.Left;
            int height = bounds.Bottom - bounds.Top;
            
            // Create bitmap with alpha channel
            var bitmap = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            
            // Make the entire bitmap transparent first
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(System.Drawing.Color.Transparent);
            }
            
            // Get window bounds without margins
            Win32Api.GetWindowRect(hWnd, out Win32Api.RECT windowRect);
            
            // Calculate window position within our capture area
            int offsetX = windowRect.Left - bounds.Left;
            int offsetY = windowRect.Top - bounds.Top;
            int windowWidth = windowRect.Right - windowRect.Left;
            int windowHeight = windowRect.Bottom - windowRect.Top;
            
            // Create a DC for the window
            IntPtr hdcScreen = Win32Api.GetDC(IntPtr.Zero);
            IntPtr hdcMem = Win32Api.CreateCompatibleDC(hdcScreen);
            IntPtr hBitmap = Win32Api.CreateCompatibleBitmap(hdcScreen, windowWidth, windowHeight);
            Win32Api.SelectObject(hdcMem, hBitmap);
            
            // Use PrintWindow to capture just the window content
            bool success = Win32Api.PrintWindow(hWnd, hdcMem, Win32Api.PW_RENDERFULLCONTENT);
            
            if (success)
            {
                // Create bitmap from the captured content
                var windowBitmap = Image.FromHbitmap(hBitmap);
                
                // Draw the window onto our transparent bitmap at the correct offset
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.DrawImage(windowBitmap, offsetX, offsetY);
                }
                
                windowBitmap.Dispose();
            }
            
            Win32Api.DeleteDC(hdcMem);
            Win32Api.ReleaseDC(IntPtr.Zero, hdcScreen);
            
            // Note: This captures the window content but not the shadows
            // Windows shadows are rendered by DWM and not part of the window itself
            AnsiConsole.MarkupLine("[yellow]Note: True transparent shadows require desktop composition. Using window content only.[/]");
            
            return bitmap;
        }
        catch
        {
            return null;
        }
    }

    private static Bitmap? ExtractWindowFromBackground(Bitmap withWindow, Bitmap backgroundOnly)
    {
        try
        {
            int width = withWindow.Width;
            int height = withWindow.Height;
            
            var result = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            // Lock bitmap data for direct pixel access
            var withWindowData = withWindow.LockBits(new Rectangle(0, 0, width, height), 
                ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var backgroundData = backgroundOnly.LockBits(new Rectangle(0, 0, width, height), 
                ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var resultData = result.LockBits(new Rectangle(0, 0, width, height), 
                ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            unsafe
            {
                byte* windowPtr = (byte*)withWindowData.Scan0;
                byte* bgPtr = (byte*)backgroundData.Scan0;
                byte* resultPtr = (byte*)resultData.Scan0;

                for (int i = 0; i < width * height; i++)
                {
                    int idx = i * 4;
                    
                    // Get pixel values (BGRA format)
                    float bw = windowPtr[idx + 0] / 255.0f;
                    float gw = windowPtr[idx + 1] / 255.0f;
                    float rw = windowPtr[idx + 2] / 255.0f;
                    
                    float bb = bgPtr[idx + 0] / 255.0f;
                    float gb = bgPtr[idx + 1] / 255.0f;
                    float rb = bgPtr[idx + 2] / 255.0f;
                    
                    // Calculate luminance-weighted difference for better shadow detection
                    float diffB = Math.Abs(bw - bb);
                    float diffG = Math.Abs(gw - gb);
                    float diffR = Math.Abs(rw - rb);
                    
                    // Use perceptual weighting (similar to human vision)
                    float weightedDiff = 0.299f * diffR + 0.587f * diffG + 0.114f * diffB;
                    
                    // Higher threshold to avoid capturing noise as transparency
                    if (weightedDiff > 0.02f) // Less sensitive to avoid window content transparency
                    {
                        float alpha;
                        
                        // Strong differences = solid window content (fully opaque)
                        if (weightedDiff > 0.15f)
                        {
                            alpha = 1.0f; // Fully opaque for window content
                        }
                        // Medium differences = likely shadows or soft edges
                        else if (weightedDiff > 0.08f)
                        {
                            alpha = 0.6f + (weightedDiff - 0.08f) * 5.0f; // Gradual opacity 60%-100%
                        }
                        // Subtle differences = soft shadows
                        else
                        {
                            alpha = Math.Min(0.5f, weightedDiff * 12.0f); // Light shadows 0%-50%
                        }
                        
                        // Use original window colors
                        resultPtr[idx + 0] = windowPtr[idx + 0]; // B
                        resultPtr[idx + 1] = windowPtr[idx + 1]; // G  
                        resultPtr[idx + 2] = windowPtr[idx + 2]; // R
                        resultPtr[idx + 3] = (byte)(alpha * 255); // A
                    }
                    else
                    {
                        // Fully transparent background (no significant difference)
                        resultPtr[idx + 0] = 0;
                        resultPtr[idx + 1] = 0;
                        resultPtr[idx + 2] = 0;
                        resultPtr[idx + 3] = 0;
                    }
                }
            }

            withWindow.UnlockBits(withWindowData);
            backgroundOnly.UnlockBits(backgroundData);
            result.UnlockBits(resultData);

            return result;
        }
        catch
        {
            return null;
        }
    }

    private static Bitmap? CalculateTransparency(Bitmap blackBg, Bitmap whiteBg, int width, int height)
    {
        try
        {
            var result = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            // Lock bitmap data for direct pixel access
            var blackData = blackBg.LockBits(new Rectangle(0, 0, width, height), 
                ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var whiteData = whiteBg.LockBits(new Rectangle(0, 0, width, height), 
                ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var resultData = result.LockBits(new Rectangle(0, 0, width, height), 
                ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            unsafe
            {
                byte* blackPtr = (byte*)blackData.Scan0;
                byte* whitePtr = (byte*)whiteData.Scan0;
                byte* resultPtr = (byte*)resultData.Scan0;

                for (int i = 0; i < width * height; i++)
                {
                    int idx = i * 4;
                    
                    // Get RGB values from both captures (BGRA format)
                    float b1 = blackPtr[idx + 0] / 255.0f;
                    float g1 = blackPtr[idx + 1] / 255.0f;
                    float r1 = blackPtr[idx + 2] / 255.0f;
                    
                    float b2 = whitePtr[idx + 0] / 255.0f;
                    float g2 = whitePtr[idx + 1] / 255.0f;
                    float r2 = whitePtr[idx + 2] / 255.0f;
                    
                    // Calculate alpha: alpha = 1 - (white - black)
                    // Since backgrounds are pure black (0) and white (1), 
                    // the difference tells us how much the original color contributed
                    float alphaB = 1.0f - (b2 - b1);
                    float alphaG = 1.0f - (g2 - g1);
                    float alphaR = 1.0f - (r2 - r1);
                    
                    // Take the average alpha (they should be similar)
                    float alpha = (alphaB + alphaG + alphaR) / 3.0f;
                    alpha = Math.Max(0.0f, Math.Min(1.0f, alpha));
                    
                    if (alpha > 0.001f)
                    {
                        // Recover original color using: original = (black - (1-alpha)*0) / alpha
                        // Since black background = 0, this simplifies to: original = black / alpha
                        float b = b1 / alpha;
                        float g = g1 / alpha;
                        float r = r1 / alpha;
                        
                        // Clamp values
                        b = Math.Max(0.0f, Math.Min(1.0f, b));
                        g = Math.Max(0.0f, Math.Min(1.0f, g));
                        r = Math.Max(0.0f, Math.Min(1.0f, r));
                        
                        // Write to result (NOT pre-multiplied for PNG)
                        resultPtr[idx + 0] = (byte)(b * 255); // B
                        resultPtr[idx + 1] = (byte)(g * 255); // G  
                        resultPtr[idx + 2] = (byte)(r * 255); // R
                        resultPtr[idx + 3] = (byte)(alpha * 255); // A
                    }
                    else
                    {
                        // Fully transparent pixel
                        resultPtr[idx + 0] = 0;
                        resultPtr[idx + 1] = 0;
                        resultPtr[idx + 2] = 0;
                        resultPtr[idx + 3] = 0;
                    }
                }
            }

            blackBg.UnlockBits(blackData);
            whiteBg.UnlockBits(whiteData);
            result.UnlockBits(resultData);

            return result;
        }
        catch
        {
            return null;
        }
    }
}