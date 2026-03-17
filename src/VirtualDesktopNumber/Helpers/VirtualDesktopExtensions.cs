using System.Collections.Generic;
using WindowsDesktop;

namespace VirtualDesktopNumber.Helpers;

public static class VirtualDesktopExtensions
{
    public static int GetNumber(this VirtualDesktop desktop)
    {
        try
        {
            return Array.FindIndex(VirtualDesktop.GetDesktops(), d => d == desktop) + 1;
        }
        catch (KeyNotFoundException)
        {
            // Some Windows builds / configurations can throw when accessing the desktop COM mappings.
            // Fall back to a safe default instead of crashing the app.
            return 1;
        }
    }
}
