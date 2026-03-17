using System.Collections.Generic;
using System.Resources;
using System.Windows.Forms;
using VirtualDesktopNumber.Helpers;
using WindowsDesktop;

namespace VirtualDesktopNumber.UI;

public class TrayContext : ApplicationContext
{
    private readonly NotifyIcon trayIcon;
    private readonly ContextMenuStrip contextMenu;
    private readonly ResourceManager resourceManager = Properties.Resources.ResourceManager;

    public TrayContext()
    {
        contextMenu = GetContextMenuStrip();

        trayIcon = new NotifyIcon()
        {
            Text = "Virtual Desktop Number",
            Visible = true
        };
        trayIcon.MouseUp += TrayIcon_MouseUp;

        // VirtualDesktop APIs can throw in some environments (e.g. missing COM mappings).
        // Catch and fall back to a safe default to avoid crashing the app.
        TryUpdateDesktopNumber();
        Application.Idle += ApplicationIdle;
    }

    private void TryUpdateDesktopNumber()
    {
        try
        {
            UpdateDesktopNumber(VirtualDesktop.Current);
        }
        catch (KeyNotFoundException)
        {
            // If we can't resolve the current desktop number, fall back to the first desktop.
            SetIcon(1);
        }
    }

    private void ApplicationIdle(object? sender, EventArgs e)
    {
        Application.Idle -= ApplicationIdle;
        VirtualDesktop.CurrentChanged += OnCurrentDesktopChanged;
    }

    private void OnCurrentDesktopChanged(object? sender, VirtualDesktopChangedEventArgs e)
    {
        try
        {
            UpdateDesktopNumber(e.NewDesktop);
        }
        catch (KeyNotFoundException)
        {
            // Ignore failures while the desktop system is in an inconsistent state.
            SetIcon(1);
        }
    }

    private void UpdateDesktopNumber(VirtualDesktop desktop)
    {
        SetIcon(desktop.GetNumber());
    }

    private void SetIcon(int desktopNumber)
    {
        string windowsTheme = ThemingHelpers.IsLightTheme() ? "Light" : "Dark";
        trayIcon.Icon = resourceManager.GetObject($"TrayIcon{windowsTheme}{desktopNumber}") as Icon;
    }

    private void TrayIcon_MouseUp(object? sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Right)
            return;

        // Show the context menu at the mouse cursor position so it appears where the user clicked.
        contextMenu.Show(Cursor.Position);
    }

    private ContextMenuStrip GetContextMenuStrip()
    {
        ContextMenuStrip context = new();
        ToolStripMenuItem itemExit = new("Exit", null, (obj, args) => Exit());
        context.Items.Add(itemExit);

        return context;
    }

    private void Exit()
    {
        trayIcon.MouseUp -= TrayIcon_MouseUp;
        contextMenu.Dispose();
        trayIcon.Visible = false;
        trayIcon.Dispose();
        Application.Exit();
    }
}