Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace WinApi
{
    public class Wallpaper
    {
	    public static void Main()
		{
			var wallpaper = new Wallpaper();
			wallpaper.SetSlideShow(@"c:\b", TimeSpan.FromMinutes(10), DesktopWallpaperPosition.Stretch);
		}
	
        public void SetSlideShow(string folderPath, TimeSpan intervall, DesktopWallpaperPosition position)
        {
            var fullPath = Path.GetFullPath(folderPath);

            var pidlList = NativeMethods.ILCreateFromPathW(fullPath);

            IntPtr pitemArr;
            NativeMethods.SHCreateShellItemArrayFromIDLists(1, out pidlList, out pitemArr);

            IDesktopWallpaper wallpaper = GetWallpaper();

            wallpaper.SetSlideshow(pitemArr);
            wallpaper.SetSlideshowOptions(DesktopSlideshowShuffle.Enabled, (uint) intervall.TotalMilliseconds);
            wallpaper.SetPosition(position);

            Marshal.ReleaseComObject(wallpaper);

            NativeMethods.ILFree(pidlList);
        }

        private static readonly Guid CLSID_DesktopWallpaper = new Guid("{C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD}");

        public static IDesktopWallpaper GetWallpaper()
        {
            Type typeDesktopWallpaper = Type.GetTypeFromCLSID(CLSID_DesktopWallpaper);
            return (IDesktopWallpaper) Activator.CreateInstance(typeDesktopWallpaper);
        }

        static class NativeMethods
        {
            [DllImport("shell32.dll", ExactSpelling = true)]
            public static extern void ILFree(IntPtr pidlList);

            [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
            public static extern IntPtr ILCreateFromPathW(string pszPath);

            [DllImport("shell32.dll")]
            public static extern int SHCreateShellItemArrayFromIDLists(uint cidl, out IntPtr rgpidl,
                out IntPtr ppsiItemArray);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public enum DesktopSlideshowOptions
        {
            ShuffleImages = 0x01
        }

        public enum DesktopSlideshowState
        {
            Enabled = 0x01,
            Slideshow = 0x02,
            DisabledByRemoteSession = 0x04
        }

        public enum DesktopSlideshowDirection
        {
            Forward = 0,
            Backward = 1
        }

        public enum DesktopSlideshowShuffle
        {
            Disabled = 0,
            Enabled = 1
        }

        public enum DesktopWallpaperPosition
        {
            Center = 0,
            Tile = 1,
            Stretch = 2,
            Fit = 3,
            Fill = 4,
            Span = 5
        }

        [ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDesktopWallpaper
        {
            void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID,
                [MarshalAs(UnmanagedType.LPWStr)] string wallpaper);

            [return: MarshalAs(UnmanagedType.LPWStr)]
            string GetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID);

            [return: MarshalAs(UnmanagedType.LPWStr)]
            string GetMonitorDevicePathAt(uint monitorIndex);

            [return: MarshalAs(UnmanagedType.U4)]
            uint GetMonitorDevicePathCount();

            [return: MarshalAs(UnmanagedType.Struct)]
            Rect GetMonitorRECT([MarshalAs(UnmanagedType.LPWStr)] string monitorID);

            void SetBackgroundColor([MarshalAs(UnmanagedType.U4)] uint color);

            [return: MarshalAs(UnmanagedType.U4)]
            uint GetBackgroundColor();

            void SetPosition([MarshalAs(UnmanagedType.I4)] DesktopWallpaperPosition position);

            [return: MarshalAs(UnmanagedType.I4)]
            DesktopWallpaperPosition GetPosition();

            void SetSlideshow(IntPtr items);

            IntPtr GetSlideshow();

            void SetSlideshowOptions(DesktopSlideshowShuffle options, uint slideshowTick);

            [PreserveSig]

            uint GetSlideshowOptions(out DesktopSlideshowShuffle options, out uint slideshowTick);

            void AdvanceSlideshow([MarshalAs(UnmanagedType.LPWStr)] string monitorID,
                [MarshalAs(UnmanagedType.I4)] DesktopSlideshowDirection direction);

            DesktopSlideshowDirection GetStatus();

            bool Enable();
        }
    }
}
"@ -ReferencedAssemblies 'System.Drawing.dll', System.Windows.Forms

[WinApi.Wallpaper]::Main()
