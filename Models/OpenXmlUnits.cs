using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TocBuilder_dotnet_framework.Models
{
    public static class OpenXmlUnits
    {
        public const long EmuPerPoint = 12700;
        public const long EmuPerPixelAt96Dpi = 9525;
        public const long EmuPerInch = 914400;

        public static long PixelsToEmu(double pixels) => (long)(pixels * EmuPerPixelAt96Dpi);
        public static long PointsToEmu(double points) => (long)(points * EmuPerPoint);
    }
}
