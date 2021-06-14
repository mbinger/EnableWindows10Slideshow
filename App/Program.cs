using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Slideshow
{
    class Program
    {
        static void Main(string[] args)
        {
            var wp = new Wallpaper();
            wp.SetSlideShow(@"C:\Wallpapers\Wallpapers", TimeSpan.FromMinutes(10), Wallpaper.DesktopWallpaperPosition.Stretch);
        }
    }
}
