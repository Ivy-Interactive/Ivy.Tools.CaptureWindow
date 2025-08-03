using System.Drawing;
using System.Runtime.Versioning;
using System.Windows.Forms;

namespace Ivy.Tools.CaptureWindow;

[SupportedOSPlatform("windows")]
public class SimpleBackgroundForm : IDisposable
{
    private Form? _form;
    private Thread? _formThread;
    private bool _disposed = false;
    private readonly ManualResetEvent _formCreated = new ManualResetEvent(false);

    public bool Create(Win32Api.RECT bounds, Color backgroundColor)
    {
        try
        {
            _formThread = new Thread(() =>
            {
                _form = new Form
                {
                    StartPosition = FormStartPosition.Manual,
                    Text = string.Empty,
                    BackColor = backgroundColor,
                    Left = bounds.Left,
                    Top = bounds.Top,
                    Width = bounds.Right - bounds.Left,
                    Height = bounds.Bottom - bounds.Top,
                    ControlBox = false,
                    FormBorderStyle = FormBorderStyle.None,
                    TopMost = false,
                    ShowInTaskbar = false
                };

                _form.Show();
                _formCreated.Set();
                
                Application.EnableVisualStyles();
                Application.Run();
            });
            _formThread.SetApartmentState(ApartmentState.STA);
            _formThread.Start();

            // Wait for form to be created
            return _formCreated.WaitOne(2000);
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _form?.Invoke(new Action(() =>
            {
                _form?.Close();
                Application.Exit();
            }));
            
            _formThread?.Join(1000);
            _formCreated.Dispose();
            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }

    ~SimpleBackgroundForm()
    {
        Dispose();
    }
}