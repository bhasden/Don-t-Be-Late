using System.Drawing;
using System.Windows.Forms;

namespace RiotousLabs.DontBeLate
{
    // Derived from http://msdn.microsoft.com/en-us/library/ms268747(VS.80).aspx
    sealed public class ConvertImage : AxHost
    {
        private ConvertImage()
            : base(null)
        {
        }

        public static stdole.IPictureDisp Convert(Image Image)
        {
            return (stdole.IPictureDisp)AxHost.GetIPictureDispFromPicture(Image);
        }
    }
}
