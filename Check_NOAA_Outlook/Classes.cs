using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace Check_NOAA_Outlook
{
    public class LocationsInfo
    {
        public Point Location, MesoMapLocation;
        public string Label, EmailRecipients, EmailRecipients_BCC;
        public int MinSeverity, LocationId, MinSevForPicture;
        public bool CheckExtended;
    }

    public enum MesoParseResult
    {
        Found_Mesos,
        No_Meso,
        Had_Error
    }

    public class PreviousResultsInfo
    {
        public PreviousResultsInfo() { }

        public PreviousResultsInfo(int LocationId)
        {
            this.LocationId = LocationId;
        }

        public int LocationId;
        public DateTime? LastCheckTime;
        public Dictionary<int, SeveritiesEnum> LastSeverities;
        public List<string> Mesos;
    }

    public class SevereCategoryValues
    {
        public SevereCategoryValues(Bitmap bm, Bitmap cities)
        {
            this.bm = bm;
            this.cities = cities;

            InitializeStuff();
        }

        public SevereCategoryValues(Bitmap bm, Bitmap cities, Bitmap day48_legend)
        {
            this.bm = bm;
            this.cities = cities;
            this.day48_legend = day48_legend;
            this.IsDay48 = day48_legend != null;

            InitializeStuff();
        }

        public void InitializeStuff()
        {
            Point[] Pts = new Point[] {
                    new Point(590, 530), // TSTM
                    new Point(670, 530), // MRGL
                    new Point(750, 530), // SLGT
                    new Point(590, 547), // ENH
                    new Point(670, 547), // MDT
                    new Point(750, 547)  // HIGH
                };

            TSTM = bm.GetPixel(Pts[0].X, Pts[0].Y);
            MRGL = bm.GetPixel(Pts[1].X, Pts[1].Y);
            SLGT = bm.GetPixel(Pts[2].X, Pts[2].Y);
            ENH = bm.GetPixel(Pts[3].X, Pts[3].Y);
            MDT = bm.GetPixel(Pts[4].X, Pts[4].Y);
            HIGH = bm.GetPixel(Pts[5].X, Pts[5].Y);
            Nothing = Color.FromArgb(255, 255, 255, 255);

            ColorToSeverity = new Dictionary<Color, SeveritiesEnum>();
            ColorToSeverity.Add(TSTM, SeveritiesEnum.TSTM);
            ColorToSeverity.Add(MRGL, SeveritiesEnum.MGRL);
            ColorToSeverity.Add(SLGT, SeveritiesEnum.SLGT);
            ColorToSeverity.Add(ENH, SeveritiesEnum.ENH);
            ColorToSeverity.Add(MDT, SeveritiesEnum.MDT);
            ColorToSeverity.Add(HIGH, SeveritiesEnum.HIGH);
            ColorToSeverity.Add(Nothing, SeveritiesEnum.Nothing);
            ColorToSeverity.Add(bm.GetPixel(20, 20), SeveritiesEnum.Water);
            ColorToSeverity.Add(bm.GetPixel(308, 36), SeveritiesEnum.Outside_US);

            if (this.IsDay48)
            {

            }
        }

        public Dictionary<Color, SeveritiesEnum> ColorToSeverity, m_ColorToSeverity2;
        public Color TSTM, MRGL, SLGT;
        public Color ENH, MDT, HIGH, Nothing;
        public Bitmap bm, cities, day48_legend;
        public bool IsDay48;
        private const double ColorTolerance = 5.0D;

        public SeveritiesEnum WhatSeverityIsThis(Point p, bool IsDay48)
        {
            if (IsDay48)
            {

            }

            SeveritiesEnum HighestSev = SeveritiesEnum.Unknown;

            int UseX = p.X;
            int UseY = p.Y;

            // compare against the passed point, but also compare against 4 other points
            // by moving 3 pixels in each direction to form a box
            // this way if the severity line ends just a few pixels away, we count it

            //       X       X
            //
            //  
            //           X
            //
            //
            //       X       X

            for (int i = 0; i < 5; i++)
            {
                switch (i)
                {
                    case 1:
                        UseX -= 3; UseY -= 3; break;
                    case 2:
                        UseX += 3; UseY -= 3; break;
                    case 3:
                        UseX -= 3; UseY += 3; break;
                    case 4:
                        UseX += 3; UseY += 3; break;
                }

                Color c = bm.GetPixel(UseX, UseY);

                foreach (Color co in ColorToSeverity.Keys)
                {
                    double ColorDiff = ColourDistance(c, co);
                    if (ColorDiff <= ColorTolerance && ColorToSeverity[co] > HighestSev)
                        HighestSev = ColorToSeverity[co];
                }
            }

            return HighestSev;
        }

        public Bitmap GetCropOfMyArea(Point p)
        {
            const int XOffset = 150;
            const int YOffset = 150;

            Bitmap MyClone = CombineCities(this.bm);
            //Bitmap MyClone = new Bitmap(this.bm);
            Rectangle rect = new Rectangle(p.X - XOffset, p.Y - YOffset, XOffset * 2, YOffset * 2);

            if (rect.X < 0) rect.X -= rect.X;
            if (rect.Y < 0) rect.Y -= rect.Y;
            if (rect.Width + rect.X > MyClone.Width)
            {
                int Offset = rect.Width + rect.X - MyClone.Width;
                rect.X -= Offset;
            }
            if (rect.Height + rect.Y > MyClone.Height)
            {
                int Offset = rect.Height + rect.Y - MyClone.Height;
                rect.Y -= Offset;
            }

            Bitmap Crop = MyClone.Clone(rect,
                System.Drawing.Imaging.PixelFormat.DontCare);

            return Crop;
        }

        public Bitmap CombineCities(Bitmap OutlookBitmap)
        {

            using (Image i = new Bitmap(OutlookBitmap))
            {
                float maxWidth = i.Width;
                float maxHeight = i.Height;

                float imageWidth = i.PhysicalDimension.Width;
                float imageHeight = i.PhysicalDimension.Height;
                float percentage = maxWidth / imageWidth;
                float newWidth = imageWidth * percentage;
                float newHeight = imageHeight * percentage;

                if (newHeight > maxHeight)
                {
                    percentage = maxHeight / newHeight;

                    newWidth = newWidth * percentage;
                    newHeight = newHeight * percentage;
                }

                using (Bitmap b = new Bitmap((int)newWidth, (int)newHeight))
                {
                    using (Graphics g = Graphics.FromImage(b))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                        g.DrawImage(i, new Rectangle(0, 0, b.Width, b.Height));

                        using (Image j = new Bitmap(this.cities))
                        {
                            g.DrawImage(j, new Rectangle(0, 0, j.Width, j.Height));

                        }

                        Bitmap newImage = Image.FromHbitmap(b.GetHbitmap());
                        return newImage;
                        //string file3 = @"C:\Users\stephen\test\test.gif";
                        //newImage.Save(file3, System.Drawing.Imaging.ImageFormat.Gif);
                    }
                }

            }
        }

        public static double ColourDistance(Color e1, Color e2)
        {
            long rmean = ((long)e1.R + (long)e2.R) / 2;
            long r = (long)e1.R - (long)e2.R;
            long g = (long)e1.G - (long)e2.G;
            long b = (long)e1.B - (long)e2.B;
            return Math.Sqrt((((512 + rmean) * r * r) >> 8) + 4 * g * g + (((767 - rmean) * b * b) >> 8));
        }
    }

    public class OutlookResults
    {
        public OutlookResults() { }

        public SeverityChangesEnum SeverityKind = SeverityChangesEnum.NA;
        public bool WasGood, WasSevere;
        public string EmailMessage;
        public SeveritiesEnum HowSevere;
        public Dictionary<Guid, Bitmap> AttachedImages;
        public Dictionary<int, SeveritiesEnum> Severities;
    }

    public class MesoDiscussionDetails
    {
        public string Filename;
        public string PolyCoords;
        public string Title;
        public string GifName;
        public string MesoText;
        public Bitmap MesoPicture;
        public List<Point> Poly;
        public bool IsDownloaded;
        public Guid Id;

        public MesoDiscussionDetails()
        {
            this.Id = Guid.NewGuid();
        }

        public bool IsPointInPolygon(Point point)
        {
            if (Poly == null) throw new NullReferenceException("Poly must have a value before calling IsPointInPolygon");
            if (Poly.Count == 0) throw new InvalidOperationException("Poly count must be greater than 0");

            var intersects = new List<int>();
            var a = Poly.Last();
            foreach (var b in Poly)
            {
                if (b.X == point.X && b.Y == point.Y)
                {
                    return true;
                }

                if (b.X == a.X && point.X == a.X && point.X >= Math.Min(a.Y, b.Y) && point.Y <= Math.Max(a.Y, b.Y))
                {
                    return true;
                }

                if (b.Y == a.Y && point.Y == a.Y && point.X >= Math.Min(a.X, b.X) && point.X <= Math.Max(a.X, b.X))
                {
                    return true;
                }

                if ((b.Y < point.Y && a.Y >= point.Y) || (a.Y < point.Y && b.Y >= point.Y))
                {
                    var px = (int)(b.X + 1.0 * (point.Y - b.Y) / (a.Y - b.Y) * (a.X - b.X));
                    intersects.Add(px);
                }

                a = b;
            }

            intersects.Sort();
            return intersects.IndexOf(point.X) % 2 == 0 || intersects.Count(x => x < point.X) % 2 == 1;
        }

    }

    public enum SeverityChangesEnum
    {
        New_Severity, Severity_Increased, Both, NA
    }

    public enum SeveritiesEnum
    {
        Outside_US = -3,
        Water = -2,
        Unknown = -1,
        Nothing = 0,
        TSTM = 1,
        MGRL = 2,
        SLGT = 3,
        ENH = 4,
        MDT = 5,
        HIGH = 6
    }

    public class DownloadedImageClass
    {
        public DownloadedImageClass()
        {
            this.Reset = new System.Threading.ManualResetEvent(false);
        }

        public Bitmap TheGif;
        public System.Threading.ManualResetEvent Reset;
        public bool Completed, IsThisForHTML, IsMesoDiscussion;
        public string HTML_Source;
        public byte[] Results;
    }
}
