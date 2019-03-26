using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSTO {
    internal class PictureConverter : AxHost {
        protected PictureConverter() : base(string.Empty) { }

        static public stdole.IPictureDisp ImageToPictureDisp(Image image) {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }

        static public stdole.IPictureDisp IconToPictureDisp(Icon icon) {
            return ImageToPictureDisp(icon.ToBitmap());
        }

        static public Image PictureDispToImage(stdole.IPictureDisp picture) {
            return GetPictureFromIPicture(picture);
        }
    }
}
