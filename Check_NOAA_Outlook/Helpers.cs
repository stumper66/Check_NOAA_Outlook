using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ImageMagick;
using System.IO;

namespace Check_NOAA_Outlook
{
    public static class Helpers
    {
        public static Exception GetSingleExceptionFromAggregateException(Exception ex)
        {
            if (ex == null) return null;
            if (ex is not AggregateException) return ex;

            AggregateException ex2 = (AggregateException)ex;
            if (ex2.InnerExceptions != null && ex2.InnerExceptions.Count > 0)
                return ex2.InnerExceptions[0];
            else
                return ex2;
        }
        public static DateTime? CheckIfDateTimeN(object Input)
        {
            if (Input == null || Input == DBNull.Value)
                return null;
            else
                return Convert.ToDateTime(Input);
        }

        public static int CheckIfInt(object Input, int DefaultValue)
        {
            if (Input == null || Input == DBNull.Value)
                return DefaultValue;
            else
                return (int)Input;
        }

        public static int? CheckIfIntN(object Input)
        {
            if (Input == null || Input == DBNull.Value)
                return null;
            else
                return (int)Input;
        }

        public static string CheckIfStr(object Input)
        {
            if (Input == null || Input == DBNull.Value)
                return null;
            else
                return Input.ToString();
        }

        public static byte[] ConvertGIFToPNG(byte[] SourceImage)
        {
            MagickImage image = new(SourceImage);
            byte[] Results;

            using MemoryStream MS = new();

            image.Write(MS, MagickFormat.Png);
            Results = new byte[MS.Length];
            MS.Write(Results, 0, (int)MS.Length);

            //using (MemoryStream MS = new())
            //{
            //    image.Write(MS, MagickFormat.Png);
            //    Results = new byte[MS.Length];
            //    MS.Write(Results, 0, (int)MS.Length);
            //}

            return Results;
        }

        //public static byte[] ConvertImageToPng(byte[] SourceImage)
        //{
        //    //create a new byte array
        //    byte[] bin = new byte[0];

        //    //check if there is data
        //    if (SourceImage == null || SourceImage.Length == 0)
        //    {
        //        return bin;
        //    }

        //    //convert the byte array to a bitmap
        //    MagickImage NewImage;
        //    using (MemoryStream ms = new MemoryStream(SourceImage))
        //    {
        //        NewImage = new MagickImage(ms);
        //    }

        //    //set some properties
        //    MagickImage TempImage = new MagickImage(NewImage.Width, NewImage.Height);
        //    using (Graphics g = Graphics.FromImage(TempImage))
        //    {
        //        g.CompositingMode = CompositingMode.SourceCopy;
        //        g.CompositingQuality = CompositingQuality.HighQuality;
        //        g.SmoothingMode = SmoothingMode.HighQuality;
        //        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        //        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        //        g.DrawImage(NewImage, 0, 0, NewImage.Width, NewImage.Height);
        //    }
        //    NewImage = TempImage;

        //    //save the image to a stream
        //    using (MemoryStream ms = new MemoryStream())
        //    {

        //        EncoderParameters encoderParameters = new EncoderParameters(1);
        //        encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 80L);

        //        NewImage.Save(ms, GetEncoderInfo("image/png"), encoderParameters);
        //        bin = ms.ToArray();
        //    }

        //    //cleanup
        //    NewImage.Dispose();
        //    TempImage.Dispose();

        //    //return data
        //    return bin;
        //}


        //get the correct encoder info
        //private static ImageCodecInfo GetEncoderInfo(string MimeType)
        //{
        //    ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
        //    for (int j = 0; j < encoders.Length; ++j)
        //    {
        //        if (encoders[j].MimeType.ToLower() == MimeType.ToLower())
        //            return encoders[j];
        //    }
        //    return null;
        //}
    }
}
