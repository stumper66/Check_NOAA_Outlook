
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
        public bool severityChangedFromPrev;
        public int highestSev;
        public bool dayChanged;

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
        public System.Drawing.Point location, mesoMapLocation;
        public string locationId;
        public bool checkExtended;
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

        public PreviousResultsInfo(string locationId)
        {
            this.locationId = locationId;
        }

        public string locationId;
        public DateTime? lastCheckTime;
        public Dictionary<int, SeveritiesEnum> lastSeverities;
        public Dictionary<int, SeveritiesEnum> last2Severities;
        public DateTime last2SevsTime;
        public List<string> mesos;
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
            this.isDay48 = day48_legend != null;

            InitializeStuff();
        }

        public void InitializeStuff()
        {
            System.Drawing.Point[] pts = [
                    new System.Drawing.Point(590, 530), // TSTM
                    new System.Drawing.Point(670, 530), // MRGL
                    new System.Drawing.Point(750, 530), // SLGT
                    new System.Drawing.Point(590, 547), // ENH
                    new System.Drawing.Point(670, 547), // MDT
                    new System.Drawing.Point(750, 547)  // HIGH
                ];


            IPixelCollection<byte> pixels = bm.GetPixels();
            tstm = new MagickColor(pixels[pts[0].X, pts[0].Y].ToColor());
            mrgl = new MagickColor(pixels[pts[1].X, pts[1].Y].ToColor());
            slgt = new MagickColor(pixels[pts[2].X, pts[2].Y].ToColor());
            enh = new MagickColor(pixels[pts[3].X, pts[3].Y].ToColor());
            mdt = new MagickColor(pixels[pts[4].X, pts[4].Y].ToColor());
            high = new MagickColor(pixels[pts[5].X, pts[5].Y].ToColor());
            nothing = new MagickColor(255, 255, 255, 255);

            colorToSeverity = new()
            {
                { tstm, SeveritiesEnum.TSTM },
                { mrgl, SeveritiesEnum.MGRL },
                { slgt, SeveritiesEnum.SLGT },
                { enh, SeveritiesEnum.ENH },
                { mdt, SeveritiesEnum.MDT },
                { high, SeveritiesEnum.HIGH },
                { nothing, SeveritiesEnum.Nothing },
                { new MagickColor(pixels[20, 20].ToColor()), SeveritiesEnum.Water },
                { new MagickColor(pixels[308, 36].ToColor()), SeveritiesEnum.Outside_US }
            };

            if (this.isDay48)
            {

            }
        }

        public Dictionary<MagickColor, SeveritiesEnum> colorToSeverity, m_ColorToSeverity2;
        public MagickColor tstm, mrgl, slgt;
        public MagickColor enh, mdt, high, nothing;
        public MagickImage bm, cities, day48_legend;
        public bool isDay48;
        private const double colorTolerance = 5.0D;

        public SeveritiesEnum WhatSeverityIsThis(System.Drawing.Point p, bool isDay48)
        {
            if (isDay48)
            {

            }

            SeveritiesEnum highestSev = SeveritiesEnum.Unknown;

            int useX = p.X;
            int useY = p.Y;

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

            IPixelCollection<byte> pixels = bm.GetPixels();
            for (int i = 0; i < 5; i++)
            {
                switch (i)
                {
                    case 1:
                        useX -= 3; useY -= 3; break;
                    case 2:
                        useX += 3; useY -= 3; break;
                    case 3:
                        useX -= 3; useY += 3; break;
                    case 4:
                        useX += 3; useY += 3; break;
                }

                // TSTM = new MagickColor(Pixels[Pts[0].X, Pts[0].Y].ToColor());
                MagickColor c = new(pixels[useX, useY].ToColor());
                //Color c = bm.GetPixel(UseX, UseY);

                foreach (MagickColor co in colorToSeverity.Keys)
                {
                    double colorDiff = ColourDistance(c, co);
                    if (colorDiff <= colorTolerance && colorToSeverity[co] > highestSev)
                        highestSev = colorToSeverity[co];
                }
            }

            return highestSev;
        }

        public MagickImage GetCropOfMyArea(System.Drawing.Point p)
        {
            const int xOffset = 150;
            const int yOffset = 150;

            MagickImage myClone = CombineCities(this.bm);
            //Bitmap MyClone = new Bitmap(this.bm);
            System.Drawing.Rectangle rect = new(p.X - xOffset, p.Y - yOffset, xOffset * 2, yOffset * 2);

            if (rect.X < 0) rect.X -= rect.X;
            if (rect.Y < 0) rect.Y -= rect.Y;
            if (rect.Width + rect.X > myClone.Width)
            {
                int offset = rect.Width + rect.X - (int)myClone.Width;
                rect.X -= offset;
            }
            if (rect.Height + rect.Y > myClone.Height)
            {
                int offset = rect.Height + rect.Y - (int)myClone.Height;
                rect.Y -= offset;
            }

            MagickGeometry g = new(rect.X, rect.Y, (uint)rect.Width, (uint)rect.Height);
            //MagickImage crop = new(myClone.Clone(g));
            MagickImage crop = new(myClone.CloneArea(g));
            //MagickImage Crop = MyClone.Clone(rect,
            //    System.Drawing.Imaging.PixelFormat.DontCare);

            return crop;
        }

        public MagickImage CombineCities(MagickImage outlookBitmap)
        {
            MagickImage i = new(outlookBitmap.Clone());

            MagickImageCollection col =
            [
                i,
                cities
            ];
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

        public SeverityChangesEnum severityKind = SeverityChangesEnum.NA;
        public bool wasGood, wasSevere;
        public string emailMessage;
        public SeveritiesEnum howSevere;
        public Dictionary<Guid, MagickImage> attachedImages;
        public Dictionary<int, SeveritiesEnum> severities;
    }

    public class MesoDiscussionDetails
    {
        public string filename;
        public string polyCoords;
        public string title;
        public string gifName;
        public string mesoText;
        public MagickImage mesoPicture;
        public List<System.Drawing.Point> poly;
        public bool isDownloaded;
        public Guid id;

        public MesoDiscussionDetails()
        {
            this.id = Guid.NewGuid();
        }

        public bool IsPointInPolygon(System.Drawing.Point point)
        {
            if (poly == null) throw new NullReferenceException("Poly must have a value before calling IsPointInPolygon");
            if (poly.Count == 0) throw new InvalidOperationException("Poly count must be greater than 0");

            var intersects = new List<int>();
            var a = poly.Last();
            foreach (var b in poly)
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
            this.reset = new System.Threading.ManualResetEvent(false);
        }

        public MagickImage theGif;
        public System.Threading.ManualResetEvent reset;
        public bool completed, isThisForHTML, isMesoDiscussion;
        public string htmlSource;
        public byte[] results;
    }
}
