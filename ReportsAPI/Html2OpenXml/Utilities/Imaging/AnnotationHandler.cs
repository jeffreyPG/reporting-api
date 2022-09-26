using ImageMagick;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HtmlToOpenXml.Utilities.Imaging
{
    public class AnnotationHandler
    {

        public byte[] AnnotateImage(byte[] imageBytes, string timestampString)
        {
            //Assuming annotation isn't required
            if (string.IsNullOrEmpty(timestampString)) return imageBytes;

            return AnnotateByPlugin(imageBytes, timestampString);
        }
        private byte[] AnnotateByPlugin(byte[] imageBytes, string timestampString)
        {
            var tempimage = new MagickImage(imageBytes);

            int textWidth = tempimage.Width - 10;
            int widthOfImage = GetFontSizeBasedOnImageWidth(tempimage.Width);
            MagickReadSettings settings = new MagickReadSettings()
            {
                FillColor = MagickColor.FromRgb(228, 52, 12),
                FontFamily = "Calibri",
                FontPointsize = widthOfImage - 10
            };


            using (var image = new MagickImage(imageBytes, settings))
            {
                image.Annotate(timestampString, Gravity.Southeast);
                return image.ToByteArray();
            }
        }
        private int GetFontSizeBasedOnImageWidth(int width)
        {


            if (width > 480 && width <= 680)
            {
                return 40;
                // return 20;
            }

            if (width > 680 && width <= 800)
            {
                return 44;
                // return 24;
            }

            if (width > 800 && width <= 1024)
            {
                return 52;
                // return 32;
            }

            if (width > 1024 && width <= 1600)
            {
                return 64;
                // return 44;
            }

            if (width > 1600 && width <= 2048)
            {
                return 70;
                // return 50;
            }

            if (width > 2048 && width <= 2560)
            {
                return 86;
                // return 66;
            }

            if (width > 2560 && width <= 6000)
            {
                return 100;
                // return 80;
            }

            return 26;
            // return 16;
        }
        private  DateTime UnixTimeStampToDateTime(double unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dateTime = dateTime.AddSeconds(unixTimeStamp);
            return dateTime;
        }
    }
}
