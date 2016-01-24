using System;
using System.Windows.Media.Imaging;
using System.Windows.Controls;
using System.Drawing;
using System.Windows;


namespace KickOff
{
    public static class ico2bmap
    {
        public static IconBitMap GetBitmapFromFileIcon(string file)
        {
            System.Drawing.Icon ico = System.Drawing.Icon.ExtractAssociatedIcon(file);

            Bitmap bm = ico.ToBitmap();

            BitmapSource bmsource =
              System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
              bm.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty,
              BitmapSizeOptions.FromEmptyOptions());

            WriteableBitmap wbitmap =
                new WriteableBitmap(bmsource);

            IconBitMap ibm = new IconBitMap()
            {
                bitmap = bm,
                bitmapsource = bmsource,
                writeablebitmap = wbitmap
            }; //, iconImage = new System.Windows.Controls.Image() };

            //ibm.iconImage.Width = bm.Width;
            //ibm.iconImage.Source = ibm.bitmapsource;
            //ibm.iconImage.Visibility = Visibility.Visible;

            return ibm;
        }

    
    }

    public class IconBitMap
    {
        public IconBitMap()
        {

        }

        public Bitmap bitmap { get; set; }

        public BitmapSource bitmapsource { get; set; }

        public WriteableBitmap writeablebitmap { get; set; }

        //public System.Windows.Controls.Image iconImage{ get; set; }
    }
}
