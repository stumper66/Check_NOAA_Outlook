
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Drawing;
using ImageMagick;

namespace Check_NOAA_Outlook
{
    public class ComparisonAgainPrevResults
    {
        public bool SeverityChangedFromPrev;
        public int HighestSev;
        public bool DayChanged;

    }
    public class LocationsInfo
    {
        // serialized fields:
        public string Coordinates { get; set; }
        public string Label { get; set; }
        public string Email_Recipients { get; set; }
        public string Email_Recipients_BCC { get; set; }
        public string MesoMap_Coordinates { get; set; }
        public int Min_Severity { get; set; }
        public int? Min_Severity_For_Picture { get; set; }

        // fields not serialized:
        public System.Drawing.Point Location, MesoMapLocation;
        public string LocationId;
        public bool CheckExtended;
    }

    public class TrackingInfo
    {
        public DateTime Last_Check { get; set; }
        public string Last_Severities { get; set; }
        public string Last_Mesos { get; set; }
        public string Last_Severities2 { get; set; }
        public DateTime? Last_Severities2_Time { get; set; }
    }
    public enum MesoParseResult
    {
        Found_Mesos,
        No_Meso,
        Had_Error
    }

    public class MyBooleanConverter : System.Text.Json.Serialization.JsonConverter<bool>
    {
        public override bool Read(ref System.Text.Json.Utf8JsonReader reader, Type typeToConvert, System.Text.Json.JsonSerializerOptions options)
        {
            if (reader.TokenType == System.Text.Json.JsonTokenType.String)
            {
                string test = reader.GetString();
                if (string.IsNullOrEmpty(test)) return false;
                test = test.ToLower();

                if (test == "0" || test == "false")
                    return false;

                if (test == "1" || test == "true")
                    return true;
            }
            else if (reader.TokenType == System.Text.Json.JsonTokenType.True) return true;

            return false;
        }

        public override void Write(System.Text.Json.Utf8JsonWriter writer, bool value, System.Text.Json.JsonSerializerOptions options)
        {
            writer.WriteBooleanValue(value);
        }
    }

    public class PreviousResultsInfo
    {
        public PreviousResultsInfo() { }

        public PreviousResultsInfo(string LocationId)
        {
            this.LocationId = LocationId;
        }

        public string LocationId;
        public DateTime? LastCheckTime;
        public Dictionary<int, SeveritiesEnum> LastSeverities;
        public Dictionary<int, SeveritiesEnum> Last2Severities;
        public DateTime Last2SevsTime;
        public List<string> Mesos;
    }
    public class SevereCategoryValues
    {
        public SevereCategoryValues(MagickImage bm, MagickImage cities)
        {
            this.bm = bm;
            this.cities = cities;

            InitializeStuff();
        }

        public SevereCategoryValues(MagickImage bm, MagickImage cities, MagickImage day48_legend)
        {
            this.bm = bm;
            this.cities = cities;
            this.day48_legend = day48_legend;
            this.IsDay48 = day48_legend != null;

            InitializeStuff();
        }

        public void InitializeStuff()
        {
            System.Drawing.Point[] Pts = new System.Drawing.Point[] {
                    new System.Drawing.Point(590, 530), // TSTM
                    new System.Drawing.Point(670, 530), // MRGL
                    new System.Drawing.Point(750, 530), // SLGT
                    new System.Drawing.Point(590, 547), // ENH
                    new System.Drawing.Point(670, 547), // MDT
                    new System.Drawing.Point(750, 547)  // HIGH
                };


            IPixelCollection<byte> Pixels = bm.GetPixels();
            TSTM = new MagickColor(Pixels[Pts[0].X, Pts[0].Y].ToColor());
            MRGL = new MagickColor(Pixels[Pts[1].X, Pts[1].Y].ToColor());
            SLGT = new MagickColor(Pixels[Pts[2].X, Pts[2].Y].ToColor());
            ENH = new MagickColor(Pixels[Pts[3].X, Pts[3].Y].ToColor());
            MDT = new MagickColor(Pixels[Pts[4].X, Pts[4].Y].ToColor());
            HIGH = new MagickColor(Pixels[Pts[5].X, Pts[5].Y].ToColor());
            Nothing = new MagickColor(255, 255, 255, 255);

            ColorToSeverity = new()
            {
                { TSTM, SeveritiesEnum.TSTM },
                { MRGL, SeveritiesEnum.MGRL },
                { SLGT, SeveritiesEnum.SLGT },
                { ENH, SeveritiesEnum.ENH },
                { MDT, SeveritiesEnum.MDT },
                { HIGH, SeveritiesEnum.HIGH },
                { Nothing, SeveritiesEnum.Nothing },
                { new MagickColor(Pixels[20, 20].ToColor()), SeveritiesEnum.Water },
                { new MagickColor(Pixels[308, 36].ToColor()), SeveritiesEnum.Outside_US }
            };

            if (this.IsDay48)
            {

            }
        }

        public Dictionary<MagickColor, SeveritiesEnum> ColorToSeverity, m_ColorToSeverity2;
        public MagickColor TSTM, MRGL, SLGT;
        public MagickColor ENH, MDT, HIGH, Nothing;
        public MagickImage bm, cities, day48_legend;
        public bool IsDay48;
        private const double ColorTolerance = 5.0D;

        public SeveritiesEnum WhatSeverityIsThis(System.Drawing.Point p, bool IsDay48)
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

            IPixelCollection<byte> Pixels = bm.GetPixels();
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

                // TSTM = new MagickColor(Pixels[Pts[0].X, Pts[0].Y].ToColor());
                MagickColor c = new(Pixels[UseX, UseY].ToColor());
                //Color c = bm.GetPixel(UseX, UseY);

                foreach (MagickColor co in ColorToSeverity.Keys)
                {
                    double ColorDiff = ColourDistance(c, co);
                    if (ColorDiff <= ColorTolerance && ColorToSeverity[co] > HighestSev)
                        HighestSev = ColorToSeverity[co];
                }
            }

            return HighestSev;
        }

        public MagickImage GetCropOfMyArea(System.Drawing.Point p)
        {
            const int XOffset = 150;
            const int YOffset = 150;

            MagickImage MyClone = CombineCities(this.bm);
            //Bitmap MyClone = new Bitmap(this.bm);
            System.Drawing.Rectangle rect = new(p.X - XOffset, p.Y - YOffset, XOffset * 2, YOffset * 2);

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

            MagickGeometry g = new(rect.X, rect.Y, rect.Width, rect.Height);
            MagickImage Crop = new(MyClone.Clone(g));
            //MagickImage Crop = MyClone.Clone(rect,
            //    System.Drawing.Imaging.PixelFormat.DontCare);

            return Crop;
        }

        public MagickImage CombineCities(MagickImage OutlookBitmap)
        {
            MagickImage i = new(OutlookBitmap.Clone());

            MagickImageCollection col = new()
            {
                i,
                cities
            };
            return new MagickImage(col.Mosaic());
        }

        public static double ColourDistance(MagickColor e1, MagickColor e2)
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
        public Dictionary<Guid, MagickImage> AttachedImages;
        public Dictionary<int, SeveritiesEnum> Severities;
    }

    public class MesoDiscussionDetails
    {
        public string Filename;
        public string PolyCoords;
        public string Title;
        public string GifName;
        public string MesoText;
        public MagickImage MesoPicture;
        public List<System.Drawing.Point> Poly;
        public bool IsDownloaded;
        public Guid Id;

        public MesoDiscussionDetails()
        {
            this.Id = Guid.NewGuid();
        }

        public bool IsPointInPolygon(System.Drawing.Point point)
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

        public MagickImage TheGif;
        public System.Threading.ManualResetEvent Reset;
        public bool Completed, IsThisForHTML, IsMesoDiscussion;
        public string HTML_Source;
        public byte[] Results;
    }
}
