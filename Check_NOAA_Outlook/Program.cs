using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Data;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace Check_NOAA_Outlook
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
            (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += new System.UnhandledExceptionEventHandler(HandleTheUnhandled);

            MainProgram MP = new MainProgram();
            if (MP.LoadSettings())
            {
                int ExitCode = MP.RunMain();
                Environment.ExitCode = ExitCode;
            }
            else
                Environment.ExitCode = 10;

#if DEBUG
            Console.WriteLine("Press any key");
            Console.ReadKey();
#endif
        }

        private static void HandleTheUnhandled(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = (Exception)e.ExceptionObject;
            string StrCD = Environment.CurrentDirectory;
            if (!StrCD.EndsWith("\\")) StrCD += "\\";

            System.Text.StringBuilder sb = new System.Text.StringBuilder("Program: Check_NOAA_Outlook\r\n");
            sb.AppendFormat("Version: {0}\r\n", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version);
            Exception ex2 = null;
            if (ex.InnerException != null) ex2 = (Exception)ex.InnerException;

            sb.AppendFormat("exception message: {0}\r\n" +
                "exception object: {1}\r\n" +
                "exception type: {2}\r\n",
                ex.Message, ex, ex.GetType());

            if (ex2 != null)
            {
                sb.AppendFormat("inner exception message: {0}\r\n" +
                "inner exception object: {1}\r\n" +
                "inner exception type: {2}\r\n",
                ex2.Message, ex2, ex2.GetType());
            }

            string file = StrCD + "debug.txt";

            try
            {
                System.IO.StreamWriter SW = new System.IO.StreamWriter(file);
                SW.Write(sb.ToString());
                SW.Close();
            }
            catch { }

            Console.WriteLine("Unhandled exception.");
            Console.WriteLine(ex.Message);
            log.Fatal("Unhandled exception.", ex);
        }

        private class MainProgram
        {
            public MainProgram()
            {
                m_strCD = Environment.CurrentDirectory;
                if (!m_strCD.EndsWith("\\")) m_strCD += "\\";

                m_Locations = new List<LocationsInfo>();
                m_DS_Tracking = BuildDS2();

                m_SourceCodes = new string[3];
                //m_GetExtendedOutlook = true;
            }

            private List<LocationsInfo> m_Locations;
            private DataSet m_DS_Tracking;
            private Bitmap m_Legend, m_Cities, m_Day48_Lenend;
            private string[] m_SourceCodes;
            private List<string> m_URLs;
            private string m_strCD, m_URL_Extended, m_MesoHtml;
            //private Dictionary<string, MesoDiscussionDetails> m_PolyStrToMeso;
            private Dictionary<Guid, MesoDiscussionDetails> m_AllMesoDetails;
            private string m_EmailHost, m_EmailFrom, m_EmailReplyTo, m_NOAA_URL, m_EmailTestAddress, m_MesoBaseURL;
            private int m_Email_HiPri_MinSev, m_Download_MaxRetries, m_Download_Timeout;
            private bool m_EmailIsEnabled, m_GetExtendedOutlook, m_IncludeMeso;

            private static string ReadAppSetting(string key, string DefaultValue)
            {
                string result = DefaultValue;
                try
                {
                    System.Collections.Specialized.NameValueCollection appSettings =
                        System.Configuration.ConfigurationManager.AppSettings;
                    result = appSettings[key] ?? DefaultValue;
                }
                catch (System.Configuration.ConfigurationErrorsException ex)
                {
                    WriteToConsoleAndLogItWithException(WhichLogType.Error, "Error loading config key: " + key, ex);
                }

                return result;
            }

            public bool LoadSettings()
            {
                m_EmailHost = ReadAppSetting("Email_Host", "localhost");
                m_EmailFrom = ReadAppSetting("Email_From", null);
                m_EmailReplyTo = ReadAppSetting("Email_ReplyTo", null);
                m_Email_HiPri_MinSev = Convert.ToInt32(ReadAppSetting("Email_HighPri_Min_Criteria", "2"));
                m_EmailIsEnabled = Convert.ToBoolean(ReadAppSetting("Email_IsEnabled", "true"));
                m_Download_MaxRetries = Convert.ToInt32(ReadAppSetting("Download_Max_Retries", "3"));
                m_Download_Timeout = Convert.ToInt32(ReadAppSetting("Download_Timeout_Seconds", "10"));
                m_NOAA_URL = ReadAppSetting("NOAA_Main_URL", "http://www.spc.noaa.gov/products/outlook/");
                m_MesoBaseURL = ReadAppSetting("NOAA_Meso_URL", "http://www.spc.noaa.gov/products/outlook/");
                m_EmailTestAddress = ReadAppSetting("Email_Test_Address", null);
                m_IncludeMeso = Convert.ToBoolean(ReadAppSetting("IncludeMesoDiscussions", "false"));

                if (!m_NOAA_URL.EndsWith("/")) m_NOAA_URL += "/";

                m_URLs = new List<string>();
                m_URLs.Add(m_NOAA_URL + "day1otlk.html");
                m_URLs.Add(m_NOAA_URL + "day2otlk.html");
                m_URLs.Add(m_NOAA_URL + "day3otlk.html");
                if (m_IncludeMeso) m_URLs.Add(m_MesoBaseURL);

                m_URL_Extended = "http://www.spc.noaa.gov/products/exper/day4-8/";

                if (System.IO.File.Exists(m_strCD + "legend.jpg"))
                    m_Legend = new Bitmap(m_strCD + "legend.jpg");
                else
                {
                    WriteToConsoleAndLogIt(WhichLogType.Fatal, "required file was not present: legend.jpg");
                    return false;
                }


                if (System.IO.File.Exists(m_strCD + "cities.gif"))
                    m_Cities = new Bitmap(m_strCD + "cities.gif");
                else
                {
                    WriteToConsoleAndLogIt(WhichLogType.Fatal, "required file was not present: cities.jpg");
                    return false;
                }

                if (System.IO.File.Exists(m_strCD + "day48_legend.gif"))
                    m_Day48_Lenend = new Bitmap(m_strCD + "day48_legend.gif");
                else
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "file was not present: day48_legend.jpg");
                

                string LocationsFile = m_strCD + "Locations.xml";
                if (!System.IO.File.Exists(LocationsFile))
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "File not found: " + LocationsFile);
                    return false;
                }

                string TrackingFile = m_strCD + "Tracking.xml";
                if (System.IO.File.Exists(TrackingFile))
                    m_DS_Tracking.ReadXml(TrackingFile);

                DataSet DS = BuildDS();
                DS.ReadXml(LocationsFile);

                DataTable DT = DS.Tables[0];
                if (DT.Rows.Count < 1)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Fatal, "No locations were defined");
                    return false;
                }

                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    DataRow Row = DT.Rows[i];
                    LocationsInfo LI = new LocationsInfo();

                    LI.LocationId = (int)Row["Id"];
                    string Loc = Row["Coordinates"].ToString();
                    string[] Points = Loc.Split(',');
                    if (Points.Length != 2)
                    {
                        WriteToConsoleAndLogIt(WhichLogType.Fatal, "Invalid coordinates: " + Loc);
                        return false;
                    }

                    LI.Location = new Point(
                        Convert.ToInt32(Points[0]),
                        Convert.ToInt32(Points[1])
                        );
                    LI.Label = Row["Label"].ToString();
                    LI.EmailRecipients = Helpers.CheckIfStr(Row["Email_Recipients"]);
                    LI.EmailRecipients_BCC = Helpers.CheckIfStr(Row["Email_Recipients_BCC"]);
                    LI.MinSeverity = Convert.ToInt32(Row["Min_Severity"]);
                    LI.MinSevForPicture = Helpers.CheckIfInt(Row["Min_Severity_For_Picture"], 2);

                    if (Row["MesoMap_Coordinates"] != DBNull.Value && !(string.IsNullOrWhiteSpace((string)Row["MesoMap_Coordinates"])))
                    {
                        Points = ((string)Row["MesoMap_Coordinates"]).Split(',');
                        if (Points.Length == 2)
                        {
                            LI.MesoMapLocation = new Point(
                                Convert.ToInt32(Points[0]),
                                Convert.ToInt32(Points[1])
                                );
                        }
                    }

                    m_Locations.Add(LI);
                }

                return true;
            }

            private static DataSet BuildDS()
            {
                DataSet DS = new DataSet("Settings");
                DataTable DT = new DataTable("Locations");

                DataColumn DC = new DataColumn("Id", typeof(int));
                DC.AllowDBNull = false;
                DC.Unique = true;
                DT.Columns.Add(DC);

                DC = new DataColumn("Coordinates");
                DT.Columns.Add(DC);

                DC = new DataColumn("Label");
                DT.Columns.Add(DC);

                DC = new DataColumn("Email_Recipients");
                DT.Columns.Add(DC);

                DC = new DataColumn("Min_Severity", typeof(int));
                DT.Columns.Add(DC);

                DC = new DataColumn("Min_Severity_For_Picture", typeof(int));
                DT.Columns.Add(DC);

                DC = new DataColumn("Email_Recipients_BCC");
                DT.Columns.Add(DC);

                DC = new DataColumn("MesoMap_Coordinates");
                DT.Columns.Add(DC);

                DS.Tables.Add(DT);
                return DS;
            }

            private static DataSet BuildDS2()
            {
                DataSet DS = new DataSet("Settings");
                DataTable DT = new DataTable("Results");

                DataColumn DC = new DataColumn("Location_Id", typeof(int));
                DC.AllowDBNull = false;
                DC.Unique = true;
                DT.Columns.Add(DC);

                DC = new DataColumn("Last_Check", typeof(DateTime));
                DC.DateTimeMode = DataSetDateTime.Utc;
                DT.Columns.Add(DC);

                DC = new DataColumn("Last_Severities");
                DT.Columns.Add(DC);

                DC = new DataColumn("Last_Mesos");
                DT.Columns.Add(DC);

                DS.Tables.Add(DT);
                return DS;
            }

            // --------------------------------------------- bread and butter here: --------------------------------------------:
            public int RunMain()
            {
                WriteToConsoleAndLogIt(WhichLogType.Info, string.Format("Starting program, locations: {0}", m_Locations.Count));

                // step 1, download the 3 days HTML and pull the image URLs from them
                List<string> ImageURLs = GetAllImageURLs();
                if (ImageURLs == null)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Fatal, "ImageURLs was null");
                    return 3;
                }

                if (m_IncludeMeso)
                {
                    MesoParseResult MesoResult = ParseMainMesoDiscussion();
                    if (MesoResult != MesoParseResult.Found_Mesos)
                    {
                        if (MesoResult == MesoParseResult.Had_Error) WriteToConsoleAndLogIt(WhichLogType.Error, "Had Error parsing main meso discussion");
                        else WriteToConsoleAndLogIt(WhichLogType.Info, "No meso discussions present");
                        m_IncludeMeso = false;
                    }
                    else
                        WriteToConsoleAndLogIt(WhichLogType.Debug, "Mesos found: " + m_AllMesoDetails.Count.ToString());
                }

                // step 2, download the GIFs for each of the outlook days
                List<Bitmap> AllTheGifs = DownloadAllTheGIFs(ImageURLs);
                if (AllTheGifs == null)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Fatal, "AllTheGIFs was null");
                    return 4;
                }

                // step 3, loop thru each defined location and compare the coordinates from each to the 
                // colors on the downloaded GIFs
                for (int i = 0; i < m_Locations.Count; i++)
                {
                    LocationsInfo LI = m_Locations[i];
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Checking " + LI.Label);

                    // ------------------------------ here is our payload logic: ------------------------------
                    OutlookResults Results = ParseTheGIFs(AllTheGifs, LI);
                    bool UseHiPri = false;
                    if ((int)Results.HowSevere >= m_Email_HiPri_MinSev)
                        UseHiPri = true;

                    // pull previous found info from tracking.xml for comparison
                    PreviousResultsInfo PRI = GetPreviousResults(LI);

                    int HighestSevFromPrev = 0;
                    bool SeverityChangedFromPrev = false;
                    bool DayChanged = false;
                    // compare previous results to see if we've already notified for these severities
                    // or if the sevs have changed, email on that
                    bool HasPrevResults = CheckPreviousResults(PRI, Results, out SeverityChangedFromPrev, out HighestSevFromPrev, out DayChanged);

                    bool EmailGotSent = false;
                    if (!log.IsDebugEnabled)
                    {
                        StringBuilder sb = new StringBuilder(LI.Label);
                        sb.AppendFormat(" ({0})", LI.LocationId);

                        foreach (int Day in Results.Severities.Keys)
                        {
                            SeveritiesEnum Sev = Results.Severities[Day];
                            sb.AppendFormat(", d{0}: {1}", Day + 1, Sev);
                        }
                        WriteToConsoleAndLogIt(WhichLogType.Info, sb.ToString());
                    }

                    if (SeverityChangedFromPrev && HighestSevFromPrev < LI.MinSeverity)
                    {
                        WriteToConsoleAndLogIt(WhichLogType.Debug,
                            "Got sev changes but min sev didn't meet requirements of at least " + LI.MinSeverity.ToString());
                    }

#if DEBUG
                    //SeverityChangedFromPrev = true;
#endif

                    List<MesoDiscussionDetails> ApplicableMesos = null;
                    bool MesoChanged = false;
                    if (m_IncludeMeso) ApplicableMesos = GetMesoDiscussionForLocation(LI, PRI, out MesoChanged);
                    bool HasMesos = ApplicableMesos != null && ApplicableMesos.Count > 0;
                    string MesoEmail = null;

                    if (HasMesos)
                    {
                        for (int t = 0; t < ApplicableMesos.Count; t++)
                        {
                            MesoDiscussionDetails Detail = ApplicableMesos[t];
                            Results.AttachedImages.Add(Detail.Id, Detail.MesoPicture);
                            if (t == 0) MesoEmail = "\r\n<p><strong>Mesoscale Discussion Below:</strong></p>";
                            MesoEmail += string.Format("\r\n<p><img src=cid:{0} ></p>\r\n<p>{1}</p>", Detail.Id, Detail.MesoText.Replace(
                                "\r\n\r\n", "</p>\r\n\r\n<p>")); //.Replace("\r\n", "<br>\r\n"));
                        }
                        UseHiPri = true;
                    }

                    WriteToConsoleAndLogIt(WhichLogType.Debug, string.Format(
                        "HasPrev: {0}, HowSevere: {1}, HighestSevFromPrev: {2}, MinSev: {3}, SevChangedFromPrev: {4}, MesoChanged: {5}, DayChanged: {6}",
                         HasPrevResults, (int)Results.HowSevere, HighestSevFromPrev, LI.MinSeverity, SeverityChangedFromPrev, MesoChanged, DayChanged));

                    if (!HasPrevResults && (int)Results.HowSevere >= LI.MinSeverity ||
                        (SeverityChangedFromPrev || DayChanged) && HighestSevFromPrev >= LI.MinSeverity || MesoChanged)
                    {
                        string EmailSubject = "Severe Weather Forecast in " + LI.Label;
                        if (SeverityChangedFromPrev)
                        {
                            WriteToConsoleAndLogIt(WhichLogType.Info, "Sev changed from last run, sending email, l: {0}", LI.LocationId);
                            EmailSubject = string.Format("Severe Weather Forecast in {0} Has Changed!", LI.Label);
                        }

                        if (HasMesos)
                        {
                            EmailSubject += " Mesoscale Discussion Included";
                        }

                        // -------------------------------------- emails get sent here ------------------------------------
                        if (m_EmailIsEnabled)
                            EmailGotSent = SendEmail(EmailSubject, Results.EmailMessage, Results.AttachedImages, UseHiPri, LI, MesoEmail);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.Info, "I wanted to send an email but emails are disabled");
                        //EmailGotSent = true;
                    }
                    else if (HasPrevResults && (int)Results.HowSevere > LI.MinSeverity)
                    {
                        WriteToConsoleAndLogIt(WhichLogType.Debug, "Not sending email due to previous results");
                    }

                    if (PRI == null)
                        PRI = new PreviousResultsInfo(LI.LocationId);
                    
                    PRI.LastSeverities = Results.Severities;
                    if (HasMesos)
                    {
                        PRI.Mesos = new List<string>();
                        for (int t = 0; t < ApplicableMesos.Count; t++)
                            PRI.Mesos.Add(ApplicableMesos[t].Filename);
                    }
                    else
                        PRI.Mesos = null;

                    AddDataForNextTime(PRI);
                } // next location

                string TrackingFile = m_strCD + "Tracking.xml";

                m_DS_Tracking.WriteXml(TrackingFile);

                return 0;
            }

            private MesoParseResult ParseMainMesoDiscussion()
            {
                m_AllMesoDetails = new Dictionary<Guid, MesoDiscussionDetails>();

                if (m_MesoHtml == null)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Meso HTML was null");
                    return MesoParseResult.Had_Error;
                }

                string Html = m_MesoHtml;
                const StringComparison SC = StringComparison.CurrentCultureIgnoreCase;

                int StartingPt = Html.IndexOf("<!-- Contents below-->", SC);
                if (StartingPt < 1)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Could not find starting point (0)");
                    return MesoParseResult.Had_Error;
                }

                int NoMesos = Html.IndexOf("<center>No Mesoscale Discussions are currently in effect.</center>", StartingPt, SC);
                if (NoMesos > 0)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Info, "No Mesoscale Discussions are currently in effect");
                    return MesoParseResult.No_Meso;
                }

                StartingPt = Html.IndexOf("<map name=\"mdimgmap\">", SC);
                if (StartingPt < 1)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Could not find starting point (2)");
                    return MesoParseResult.Had_Error;
                }

                int EndPoint = Html.IndexOf("</map>", StartingPt + 21, SC);
                if (EndPoint < 1)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Could not find end point");
                    return MesoParseResult.Had_Error;
                }

                Html = Html.Substring(StartingPt + 21, EndPoint - StartingPt - 21);
                /*
                 <map name="mdimgmap">
     <area shape="poly" coords="463,122,475,125,494,128,493,138,483,142,462,141,453,130,463,122,463,122" href="https://www.spc.noaa.gov/products/md/md0302.html" title="Mesoscale Discussion # 302">
     <area shape="poly" coords="398,142,406,135,419,132,427,131,439,129,451,123,462,126,454,144,441,149,415,152,388,144,398,142,398,142" href="https://www.spc.noaa.gov/products/md/md0301.html" title="Mesoscale Discussion # 301">
     </map>
                */

                const string RegexStr = @"(\""poly\"")(.*?)(coords=\"")(.*?)(\"")(.*?)(href=\"")(.*?)(\"")";
                RegexOptions Opts = RegexOptions.IgnoreCase | RegexOptions.Multiline;
                MatchCollection Matches = Regex.Matches(Html, RegexStr, Opts);

                if (Matches == null || Matches.Count == 0)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "No regex matches for foly poly coordinates");
                    return MesoParseResult.Had_Error;
                }

                //WriteToConsoleAndLogIt(WhichLogType.Debug, "Match count: " + Matches.Count.ToString());
                for (int i = 0; i < Matches.Count; i++)
                {
                    MesoDiscussionDetails Details = new MesoDiscussionDetails();

                    Match M = Matches[i];

                    if (M.Groups.Count < 9)
                    {
                        WriteToConsoleAndLogIt(WhichLogType.Debug, string.Format("Match {0}, group count was {1}, expected >= 9", i, M.Groups.Count));
                        return MesoParseResult.Had_Error;
                    }

                    Details.PolyCoords = M.Groups[4].Value;
                    Details.Filename = M.Groups[8].Value;

                    if (!string.IsNullOrEmpty(Details.PolyCoords))
                    {
                        string[] Split = Details.PolyCoords.Split(',');
                        List<Point> Points = new List<Point>(Split.Length);
                        for (int t = 0; t < Split.Length; t += 2)
                        {
                            Points.Add(new Point(
                                Convert.ToInt32(Split[t].Trim()),
                                Convert.ToInt32(Split[t + 1].Trim())
                                ));
                        }
                        Details.Poly = Points;
                    }

                    //WriteToConsoleAndLogIt(WhichLogType.Debug, (string.Format("{0}: Poly: {1}, link: {2}", i, Poly, Link));
                    //DownloadMesoDiscussion(BaseURL, Link);

                    if (Details.Poly != null && Details.Poly.Count > 0)
                        m_AllMesoDetails.Add(Details.Id, Details);
                    else
                        WriteToConsoleAndLogIt(WhichLogType.Info, string.Format("Skipping meso {0} as there was no coords", Details.Filename));
                }
                return MesoParseResult.Found_Mesos;
            }

            private bool DownloadMesoDiscussionDetails(Guid WhichOne)
            {
                System.Net.WebClient WC = new System.Net.WebClient();
                WC.DownloadStringCompleted += WC_DownloadStringCompleted;
                MesoDiscussionDetails Details = m_AllMesoDetails[WhichOne];
                string URL = m_MesoBaseURL + Details.Filename;

                WriteToConsoleAndLogIt(WhichLogType.Debug, "Downloading Meso: " + URL);

                int RetryCount = 0;
                Uri TheSite = new Uri(URL);
                string Html = null;

                while (RetryCount <= m_Download_MaxRetries)
                {
                    DownloadedImageClass DIC = new DownloadedImageClass();
                    WC.DownloadStringAsync(TheSite, DIC);

                    DIC.Reset.WaitOne(m_Download_Timeout * 1000);

                    if (DIC.Completed)
                    {
                        Html = DIC.HTML_Source;
                        break;
                    }
                    else
                    {
                        WC.CancelAsync();
                        RetryCount++;
                        if (RetryCount <= m_Download_MaxRetries)
                            WriteToConsoleAndLogIt(WhichLogType.Debug, "Timed out, retry {0}", RetryCount);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.Error, "Timed out, cancelling");
                    }
                } // next retry

                if (RetryCount > m_Download_MaxRetries)
                    return false;

                // we got the HTML, need to parse it for the the text and image.  Will download image
                if (!ParseMesoDetails(Details, Html))
                    return false;

                List<string> DownloadUrl = new List<string>();
                DownloadUrl.Add(m_MesoBaseURL + Details.GifName);

                List<Bitmap> Result = DownloadAllTheGIFs(DownloadUrl, true);
                if (Result == null || Result.Count == 0)
                    return false;

                Details.MesoPicture = Result[0];
                Details.IsDownloaded = true;

                return true;
            }

            private bool ParseMesoDetails(MesoDiscussionDetails Details, string MesoHtml)
            {
                string Html = MesoHtml;
                string RegexStr = @"(<title>)(.*?)(</title>)";
                RegexOptions Opts = RegexOptions.IgnoreCase | RegexOptions.Multiline;
                Match Match = Regex.Match(Html, RegexStr, Opts);

                Details.Title = Details.Filename; // if we can't find the title we'll use the filename
                if (Match.Success && Match.Groups.Count > 2)
                    Details.Title = Match.Groups[2].Value.Trim();

                const StringComparison SC = StringComparison.CurrentCultureIgnoreCase;
                int StartingPt = Html.IndexOf("a name=\"contents\"", SC);

                if (StartingPt < 1)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, Details.Filename + ", could not find starting point (0)");
                    return false;
                }

                // <img src="mcd0330.gif"
                // download gif:
                RegexStr = @"(img src=\"")(.*?)(\"")";

                Html = Html.Substring(StartingPt + 18);

                Match = Regex.Match(Html, RegexStr, Opts);
                if (!Match.Success)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Unable to locate gif in html " + Details.Filename);
                    return false;
                }

                Details.GifName = Match.Groups[2].Value;

                StartingPt = Html.IndexOf("<pre>", SC);

                if (StartingPt < 1)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, Details.Filename + ", could not find starting point (1)");
                    return false;
                }

                int EndPt = Html.IndexOf("</pre>", 0, SC);

                if (EndPt < 1)
                {
                    WriteToConsoleAndLogIt(WhichLogType.Debug, Details.Filename + ", could not find end point");
                    return false;
                }

                int len = Html.Length;
                Html = Html.Substring(StartingPt + 5, EndPt - StartingPt - 5);
                //Html = Html.Substring(0, EndPt);
                StringBuilder sb = new StringBuilder();
                using (System.IO.StringReader SR = new System.IO.StringReader(Html))
                {
                    bool FoundStart = false;
                    while (SR.Peek() >= 0)
                    {
                        string Line = SR.ReadLine();
                        if (Line == null) continue;
                        Line = Line.Trim();

                        if (!FoundStart)
                        {
                            if (Line.Length == 0) continue;
                            FoundStart = true;
                        }

                        sb.AppendLine(Line);
                    }
                }

                Details.MesoText = sb.ToString();

                // remove trailing carriage returns, up to 5
                for (int i = 0; i < 5; i++)
                {
                    if (Details.MesoText.EndsWith("\n"))
                        Details.MesoText = Details.MesoText.Substring(0, Details.MesoText.Length - 1);
                    else
                        break;
                }

                return true;
            }

            private List<MesoDiscussionDetails> GetMesoDiscussionForLocation(LocationsInfo LI, PreviousResultsInfo PRI, out bool MesoChanged)
            {
                MesoChanged = false; // only true if there are mesos present now that wasn't before

                if (m_AllMesoDetails.Count == 0) return null;
                if (LI.MesoMapLocation.IsEmpty) return null;

                List<MesoDiscussionDetails> Mesos = null;

                foreach (Guid Id in m_AllMesoDetails.Keys)
                {
                    MesoDiscussionDetails Detail = m_AllMesoDetails[Id];
                    if (Detail.IsPointInPolygon(LI.MesoMapLocation))
                    {
                        // found a meso that is for this location

                        // first see if we've previously seen this filename
                        if (PRI.Mesos == null || !PRI.Mesos.Contains(Detail.Filename))
                            MesoChanged = true;

                        WriteToConsoleAndLogIt(WhichLogType.Debug, string.Format(
                            "Found applicabled meso for Id: {0}, filename: {1}, IsDownloaded: {2}",
                            LI.LocationId, Detail.Filename, Detail.IsDownloaded));
                        if (!Detail.IsDownloaded)
                        {
                            bool DownloadedIt = DownloadMesoDiscussionDetails(Id);
                            if (!DownloadedIt)
                            {
                                WriteToConsoleAndLogIt(WhichLogType.Info, "Was unable to fully download meso discussion");
                                return null;
                            }
                        }

                        if (Mesos == null) Mesos = new List<MesoDiscussionDetails>();
                        Mesos.Add(Detail);
                    }
                }

                return Mesos;               
            }

            private bool CheckPreviousResults(PreviousResultsInfo PRI, OutlookResults OR, out bool SeverityChangedFromPrev, out int HighestSev, out bool DayChanged)
            {
                DayChanged = false;
                HighestSev = 0;
                SeveritiesEnum HighestSevEnum = SeveritiesEnum.Nothing;
                SeverityChangedFromPrev = false;
                DateTime Now = DateTime.Now;
                DateTime UseDate = new DateTime(Now.Year, Now.Month, Now.Day, 0, 0, 0, DateTimeKind.Local);
                DateTime Yesterday = UseDate.AddDays(-1);
                string[] IOrD = new string[] { "{IncreaseOrDecrease0}", "{IncreaseOrDecrease1}", "{IncreaseOrDecrease2}" };

                if (PRI == null || PRI.LastSeverities == null || !PRI.LastCheckTime.HasValue)
                {
                    if (OR.EmailMessage != null)
                    {
                        foreach (string Text in IOrD)
                            OR.EmailMessage = OR.EmailMessage.Replace(Text, string.Empty);
                    }

                    return false;
                }

                //if (PRI != null && PRI.LastSeverities != null &&
                //    (PRI.LastCheckTime.HasValue && PRI.LastCheckTime.Value.ToLocalTime() > UseDate) ||
                //    PRI != null && !PRI.LastCheckTime.HasValue)

                //bool LastCheckWasPrevDay = false;

                if (PRI.LastCheckTime.HasValue && 
                    PRI.LastCheckTime.Value.ToLocalTime() < UseDate &&
                    PRI.LastCheckTime.Value.ToLocalTime() > Yesterday)
                {
                    DayChanged = true;

                    // we last checked some time yesterday.  Compare the prev day 2 and 3 to see if the severity changed overnight
                    // to do so, we'll shift days prev days 2 and 3 to look like today's 1 and 2
                }

                foreach (int UseDay in PRI.LastSeverities.Keys)
                {
                    int UsePrevDay = UseDay;
                    if (DayChanged)
                    {
                        if (PRI.LastSeverities.ContainsKey(UsePrevDay + 1))
                            UsePrevDay++;
                        else
                        {
                            if (OR.EmailMessage != null)
                                OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay], string.Empty);
                            continue;
                        }
                    }

                    SeveritiesEnum LastSev = PRI.LastSeverities[UsePrevDay];
                    SeveritiesEnum CurrentSev = OR.Severities[UseDay];

                    if (LastSev > HighestSevEnum)
                        HighestSevEnum = LastSev;
                    if (CurrentSev > HighestSevEnum)
                        HighestSevEnum = CurrentSev;

                    if (LastSev > CurrentSev)
                    {
                        // severity decreased
                        SeverityChangedFromPrev = true;
                        WriteToConsoleAndLogIt(WhichLogType.Info, "Day {0} Sev decreased from {1} to {2}", UseDay, LastSev, CurrentSev);
                        OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay],
                            " (<span style=\"background-color: #FFFF00\">decreased</span> from " + LastSev.ToString() + ")");
                    }
                    else if (LastSev < CurrentSev)
                    {
                        // severity increased
                        if (LastSev <= SeveritiesEnum.Nothing) // we don't care if it increased from nothing
                        {
                            WriteToConsoleAndLogIt(WhichLogType.Debug, "Found increase from nothing, so treating it as new");
                            OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay], string.Empty);
                        }
                        else
                        {
                            WriteToConsoleAndLogIt(WhichLogType.Info, "Day {0} Sev increased from {1} to {2}", UseDay, LastSev, CurrentSev);
                            OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay],
                                " <span style=\"background-color: #FFFF00\">(increased from " + LastSev.ToString() + ")</span>");
                        }
                        SeverityChangedFromPrev = true;
                    }
                    else
                    {
                        // no change in sev
                        if (OR.EmailMessage != null)
                            OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay], string.Empty);
                    }
                } // next day

                HighestSev = (int)HighestSevEnum;
                return true;
            }

            private void AddDataForNextTime(PreviousResultsInfo PRI)
            {
                DataView DV = new DataView(m_DS_Tracking.Tables[0]);
                DV.RowFilter = string.Format("Location_Id = {0}", PRI.LocationId);

                if (DV.Count > 0)
                    DV.Delete(0);

                string Mesos = null;
                if (PRI.Mesos != null && PRI.Mesos.Count > 0)
                {
                    for (int i = 0; i < PRI.Mesos.Count; i++)
                    {
                        if (i == 0) Mesos = PRI.Mesos[i];
                        else Mesos += "; " + PRI.Mesos[i];
                    }
                }

                DataRow Row = DV.Table.NewRow();
                Row[0] = PRI.LocationId;
                Row[1] = DateTime.UtcNow;
                Row[3] = Mesos ?? (object)DBNull.Value;

                StringBuilder sb = new StringBuilder();
                foreach (int Day in PRI.LastSeverities.Keys)
                {
                    if (sb.Length > 0) sb.Append(";");
                    sb.Append((int)PRI.LastSeverities[Day]);
                }

                Row[2] = sb.ToString();
                DV.Table.Rows.Add(Row);
            }

            private PreviousResultsInfo GetPreviousResults(LocationsInfo LI)
            {
                if (m_DS_Tracking.Tables[0].Rows.Count < 1) return null;

                DataView DV = new DataView(m_DS_Tracking.Tables[0]);
                DV.RowFilter = string.Format("Location_Id = {0}", LI.LocationId);

                if (DV.Count > 0)
                {
                    PreviousResultsInfo PRI = new PreviousResultsInfo();
                    PRI.LocationId = LI.LocationId;
                    PRI.LastCheckTime = Helpers.CheckIfDateTimeN(DV[0]["Last_Check"]);
                    string Temp = Helpers.CheckIfStr(DV[0]["Last_Severities"]);
                    if (!string.IsNullOrEmpty(Temp)) {
                        string[] Days = Temp.Split(';');
                        PRI.LastSeverities = new Dictionary<int, SeveritiesEnum>();
                        for (int i = 0; i < Days.Length; i++)
                        {
                            string TheDay = Days[i].Trim();
                            SeveritiesEnum Sev = (SeveritiesEnum)Enum.Parse(typeof(SeveritiesEnum), TheDay);
                            PRI.LastSeverities.Add(i, Sev);
                        }
                    }
                    if (DV[0]["Last_Mesos"] != DBNull.Value)
                    {
                        PRI.Mesos = new List<string>();
                        string[] Temp2 = ((string)DV[0]["Last_Mesos"]).Split(';');
                        for (int i = 0; i < Temp2.Length; i++)
                        {
                            PRI.Mesos.Add(Temp2[i].Trim());
                        }
                    }
                    return PRI;
                } // end if had a record

                return null;
            }

            private OutlookResults ParseTheGIFs(List<Bitmap> TheGIFs, LocationsInfo LI)
            {
                OutlookResults Results = new OutlookResults();
                StringBuilder sb = new StringBuilder();
                Results.Severities = new Dictionary<int, SeveritiesEnum>();

                for (int i = 0; i < TheGIFs.Count; i++)
                {
                    Bitmap Gif = TheGIFs[i];

                    bool WasSevere = false;
                    SevereCategoryValues SCV = new SevereCategoryValues(Gif, m_Cities);

                    Point UsePoint = new Point(LI.Location.X, LI.Location.Y);
                    SeveritiesEnum Severity = SCV.WhatSeverityIsThis(UsePoint, false);
                    Results.Severities.Add(i, Severity);

                    if (Severity > Results.HowSevere)
                        Results.HowSevere = Severity;

                    int Day = i + 1;
                    Guid UseGuid = Guid.Empty;

                    if ((int)Severity >= LI.MinSevForPicture) // MRGL and up default
                    {
                        UseGuid = Guid.NewGuid();
                        Bitmap TheCrop = SCV.GetCropOfMyArea(UsePoint);
                        if (Results.AttachedImages == null)
                            Results.AttachedImages = new Dictionary<Guid, Bitmap>();
                        Results.AttachedImages.Add(UseGuid, TheCrop);
                        Results.WasSevere = true;
                        WasSevere = true;
                    }

                    // <img src=""cid:{0}"" />

                    string SummaryText = null;
                    if (WasSevere)
                    {
                        Results.WasSevere = true;
                        SummaryText = ExtractSummaryFromSource(i);
                    }

                    string ImageInfo = WasSevere ?
                        string.Format("<br><b>Summary:</b> {2}<br>\r\n<a href=\"{0}\"><img src=cid:{1} ></a>",
                        m_URLs[i], UseGuid, SummaryText) :
                        string.Empty;

                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Day {0}, Severity: {1}",
                        Day, Severity);
                    string IncreaseOrDecrease = "{IncreaseOrDecrease" + i.ToString() + "}";
                    DateTime UseDate = DateTime.Now.AddDays(i);
                    sb.AppendLine(string.Format("<p><a href=\"{0}\">Day {1}</a> - {2}, Severity: {3}{4}{5}</p>",
                       m_URLs[i] , Day, UseDate.DayOfWeek, Severity, IncreaseOrDecrease ,ImageInfo));
                }

                //if (Results.WasSevere)
                Results.EmailMessage = sb.ToString();

                return Results;
            }

            private List<Bitmap> DownloadAllTheGIFs(List<string> URLs, bool IsMeso = false)
            {
                List<Bitmap> Results = new List<Bitmap>();

                List<string> TheURLs = URLs;
                if (m_GetExtendedOutlook)
                {
                    for (int i = 4; i <= 8; i++)
                    {
                        string Url = m_URL_Extended + string.Format("/day{0}prob.gif", i);

                        TheURLs.Add(m_URL_Extended);
                    }
                }
               
                for (int i = 0; i < TheURLs.Count; i++)
                {
                    string URL = TheURLs[i];

                    System.Net.WebClient WC = new System.Net.WebClient();
                    WC.DownloadDataCompleted += WC_DownloadDataCompleted;

                    Uri TheSite = new Uri(URL);
                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Downloading: " + URL);
                    int RetryCount = 0;
                    while (RetryCount <= m_Download_MaxRetries)
                    {
                        DownloadedImageClass DIC = new DownloadedImageClass();
                        DIC.IsMesoDiscussion = true;
                        WC.DownloadDataAsync(TheSite, DIC);

                        DIC.Reset.WaitOne(m_Download_Timeout * 1000);

                        if (DIC.Completed)
                        {
                            Results.Add(DIC.TheGif);
                            break;
                        }
                        else
                        {
                            WC.CancelAsync();
                            RetryCount++;
                            if (RetryCount <= m_Download_MaxRetries)
                                WriteToConsoleAndLogIt(WhichLogType.Debug, "Timed out, retry {0}", RetryCount);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.Error, "Timed out, cancelling");
                        }
                    } // next retry

                    if (RetryCount > m_Download_MaxRetries)
                        return null;
                } // next URL

                return Results;
            }

            private static List<MailAddress> ParseEmailRecipients(string Input)
            {
                if (string.IsNullOrEmpty(Input)) return null;

                string[] ToPeople = Input.Split(';');
                List<MailAddress> Addresses = new List<MailAddress>();

                for (int i = 0; i < ToPeople.Length; i++)
                {
                    string To = ToPeople[i].Trim();
                    Match M = Regex.Match(To, @"\(.*\)");
                    if (M.Success)
                    {
                        string Displayname = M.Value.Substring(1, M.Value.Length - 2);
                        string Email = To.Substring(0, M.Index).Trim();

                        MailAddress MA = new MailAddress(Email, Displayname);
                        Addresses.Add(MA);
                    }
                    else
                        Addresses.Add(new MailAddress(To));
                }

                if (Addresses.Count > 0) return Addresses;
                else return null;
            }

            private bool SendEmail(string Subject, string Body, Dictionary<Guid, Bitmap> Images, bool UseHighPri, LocationsInfo LI, string Meso)
            {
                SmtpClient Client = new SmtpClient(m_EmailHost);

                MailMessage Msg = new MailMessage();
                Msg.From = new MailAddress(m_EmailFrom);
                if (UseHighPri) Msg.Priority = MailPriority.High;

#if DEBUG
                //if (!string.IsNullOrEmpty(m_EmailTestAddress))
                //    Recips = m_EmailTestAddress;
#endif

                // optionally email address can have a display name.  Syntax:
                // email@address.com (Display Name)
                List<MailAddress> ToPeople = ParseEmailRecipients(LI.EmailRecipients);
                if (ToPeople != null)
                {
                    for (int i = 0; i < ToPeople.Count; i++)
                        Msg.To.Add(ToPeople[i]);
                }

                List<MailAddress> BCCPeople = ParseEmailRecipients(LI.EmailRecipients_BCC);
                if (BCCPeople != null)
                {
                    for (int i = 0; i < BCCPeople.Count; i++)
                        Msg.Bcc.Add(BCCPeople[i]);
                }

                string Extra = string.Empty;
                if (Images != null && Images.Count > 0)
                {
                    Guid g = Guid.NewGuid();

                    Body += string.Format("\r\n<p><img src=cid:{0} ></p>", g);
                    Images.Add(g, m_Legend);
                }

                if (!string.IsNullOrEmpty(m_EmailReplyTo))
                    Msg.ReplyToList.Add(m_EmailReplyTo);


                string Meso2 = Meso ?? string.Empty;
                string UseBody = "<html><body>\r\n" + Body + "\r\n" + Meso2 + "</body></html>";
                //string UseBody = Body;

                Msg.IsBodyHtml = true;
                Msg.Subject = Subject;
                if (Images == null || Images.Count < 1)
                    Msg.Body = UseBody;

                if (Images != null)
                {
                    List<LinkedResource> LRs = new List<LinkedResource>();

                    foreach (Guid g in Images.Keys)
                    {
                        Bitmap bm = Images[g];
                        System.IO.MemoryStream MS = new System.IO.MemoryStream();
                        System.Drawing.Imaging.ImageFormat imgFormat = System.Drawing.Imaging.ImageFormat.Png;
                        string strImageFormat = "image/jpeg";

                        if (m_AllMesoDetails != null && m_AllMesoDetails.ContainsKey(g))
                        {
                            //imgFormat = System.Drawing.Imaging.ImageFormat.Gif;
                            //strImageFormat = "image/gif";
#if DEBUG
                            bm.Save(@"c:\temp\test.gif");
                            bm.Save(@"c:\temp\test2.gif", System.Drawing.Imaging.ImageFormat.Gif);
#endif
                        }

                        bm.Save(MS, imgFormat);
                        MS.Seek(0, System.IO.SeekOrigin.Begin);

                        LinkedResource LR = new LinkedResource(MS, strImageFormat);
                        LR.ContentId = g.ToString();

                        LRs.Add(LR);
                    }

                    if (LRs.Count > 0)
                    {
                        AlternateView AV = AlternateView.CreateAlternateViewFromString(UseBody, null, "text/html");
                        for (int i = 0; i < LRs.Count; i++)
                            AV.LinkedResources.Add(LRs[i]);

                        Msg.AlternateViews.Add(AV);
                    }
                }

#if DEBUG
                System.IO.StreamWriter SW = new System.IO.StreamWriter(@"c:\temp\test.html", false);
                SW.Write(UseBody);
                SW.Close();
#endif

                bool ItGotSent = false;
                try
                {
                    Client.Send(Msg);
                    ItGotSent = true;
                }
                catch (Exception ex)
                {
                    WriteToConsoleAndLogItWithException(WhichLogType.Error, "Error sending email:", ex);
                }

                if (ItGotSent)
                    WriteToConsoleAndLogIt(WhichLogType.Info, string.Format("{0} ({1}), Email got sent", LI.Label, LI.LocationId));

                return ItGotSent;
            }

            private string ExtractSummaryFromSource(int WhichDay)
            {
                string Source = m_SourceCodes[WhichDay];

                const StringComparison SC = StringComparison.CurrentCultureIgnoreCase;
                int Start = Source.IndexOf("...SUMMARY...", SC);
                if (Start < 0) return null;

                StringBuilder sb = new StringBuilder();
                string Source2 = Source.Substring(Start);
                System.IO.StringReader SR = new System.IO.StringReader(Source2);
                SR.ReadLine();
                while (SR.Peek() > 0)
                {
                    string Line = SR.ReadLine();
                    //Line = Line.Substring(27).Trim();
                    // 0        1         2         3
                    // 123456789012345678901234567890
                    // <span id="line478"></span>
                    if (string.IsNullOrWhiteSpace(Line))
                        break;
                    else
                        sb.AppendLine(Line);
                }

                return sb.ToString();
            }

            private List<string> GetAllImageURLs()
            {
                string URL_Gif = m_NOAA_URL + "day{0}{1}.gif";

                List<string> Results = new List<string>();
                System.Net.WebClient WC = new System.Net.WebClient();
                WC.DownloadStringCompleted += WC_DownloadStringCompleted;

                List<string> TheURLs = m_URLs;
                //if (m_GetExtendedOutlook)
                //{
                //    Array.Resize(ref TheURLs, TheURLs.Length + 1);
                //    TheURLs[TheURLs.Length - 1] = m_URL_Extended;
                //}

                for (int i = 0; i < TheURLs.Count; i++)
                {
                    string URL = TheURLs[i];

                    WriteToConsoleAndLogIt(WhichLogType.Debug, "Downloading: " + URL);
                    int RetryCount = 0;
                    while (RetryCount <= m_Download_MaxRetries)
                    {
                        DownloadedImageClass DIC = new DownloadedImageClass();
                        Uri TheSite = new Uri(URL);
                        WC.DownloadStringAsync(TheSite, DIC);

                        DIC.Reset.WaitOne(m_Download_Timeout * 1000);

                        if (DIC.Completed)
                        {
                            string Source = DIC.HTML_Source;
                            if (i < 3)
                            {
                                // day 1-3 outlook
                                m_SourceCodes[i] = Source;

                                // a title="Categorical Outlook"
                                Match m = Regex.Match(Source, ".*title=.*\"Categorical Outlook\"");
                                if (!m.Success) return null;

                                // <td OnClick="show_tab('otlk_0100')" OnMouseOver="show_tab('otlk_0100')"><a title="Categorical Outlook"
                                m = Regex.Match(m.Value, @"\((.*?)\)");
                                if (!m.Success) return null;

                                //('otlk_0100')
                                string Value = m.Value.Substring(2, m.Value.Length - 4);
                                // otlk_0100

                                int UseDay = i + 1;
                                string UseURL = string.Format(URL_Gif, UseDay, Value);
                                Results.Add(UseURL);
                            }
                            else
                            {
                                // meso discussion
                                m_MesoHtml = Source;
                            }

                            break;
                        }
                        else
                        {
                            WC.CancelAsync();
                            RetryCount++;
                            if (RetryCount <= m_Download_MaxRetries)
                                WriteToConsoleAndLogIt(WhichLogType.Warn, "Timed out, retry {0}", RetryCount);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.Fatal, "Timed out, cancelling");
                        }
                    } // next retry

                    if (RetryCount > m_Download_MaxRetries)
                        return null;
                }

                return Results;
            }

            private void WC_DownloadDataCompleted(object sender, System.Net.DownloadDataCompletedEventArgs e)
            {
                if (e.Cancelled) return;

                DownloadedImageClass DIC = (DownloadedImageClass)e.UserState;

                if (e.Error == null)
                {
                    DIC.Results = e.Result;

                    if (DIC.IsMesoDiscussion)
                    {
                        DIC.Results = Helpers.ConvertImageToPng(DIC.Results);
                    }

                    using (System.IO.MemoryStream MS = new System.IO.MemoryStream(DIC.Results))
                        DIC.TheGif = new Bitmap(MS);

                    DIC.Completed = true;
                }
                else
                    WriteToConsoleAndLogItWithException(WhichLogType.Error, "Download exception: ", e.Error);

                DIC.Reset.Set();
            }

            private void WC_DownloadStringCompleted(object sender, System.Net.DownloadStringCompletedEventArgs e)
            {
                if (e.Cancelled) return;

                DownloadedImageClass DIC = (DownloadedImageClass)e.UserState;

                DIC.HTML_Source = e.Result;

                if (e.Error == null)
                    DIC.Completed = true;
                else
                    WriteToConsoleAndLogItWithException(WhichLogType.Error, "Download exception", e.Error);

                DIC.Reset.Set();
            }
        }

#region Logging stuff
        private static void WriteToConsoleAndLogIt(WhichLogType Logtype, string Text)
        {
            switch (Logtype)
            {
                case WhichLogType.Debug:
                    log.Debug(Text);
                    if (log.IsDebugEnabled) Console.WriteLine(Text);
                    break;
                case WhichLogType.Error:
                    log.Error(Text);
                    if (log.IsErrorEnabled) Console.WriteLine(Text);
                    break;
                case WhichLogType.Fatal:
                    log.Fatal(Text);
                    if (log.IsFatalEnabled) Console.WriteLine(Text);
                    break;
                case WhichLogType.Info:
                    log.Info(Text);
                    if (log.IsInfoEnabled) Console.WriteLine(Text);
                    break;
                case WhichLogType.Warn:
                    log.Warn(Text);
                    if (log.IsWarnEnabled) Console.WriteLine(Text);
                    break;
            }
        }

        private static void WriteToConsoleAndLogIt(WhichLogType Logtype, string Text, params object[] args)
        {
            switch (Logtype)
            {
                case WhichLogType.Debug:
                    log.DebugFormat(Text, args);
                    if (log.IsDebugEnabled) Console.WriteLine(Text, args);
                    break;
                case WhichLogType.Error:
                    log.ErrorFormat(Text, args);
                    if (log.IsErrorEnabled) Console.WriteLine(Text, args);
                    break;
                case WhichLogType.Fatal:
                    log.FatalFormat(Text, args);
                    if (log.IsFatalEnabled) Console.WriteLine(Text, args);
                    break;
                case WhichLogType.Info:
                    log.InfoFormat(Text, args);
                    if (log.IsInfoEnabled) Console.WriteLine(Text, args);
                    break;
                case WhichLogType.Warn:
                    log.WarnFormat(Text, args);
                    if (log.IsWarnEnabled) Console.WriteLine(Text, args);
                    break;
            }

        }

        private static void WriteToConsoleAndLogItWithException(WhichLogType Logtype, string Text, Exception ex)
        {
            switch (Logtype)
            {
                case WhichLogType.Debug:
                    log.Debug(Text, ex);
                    if (log.IsDebugEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.Error:
                    log.Error(Text, ex);
                    if (log.IsErrorEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.Fatal:
                    log.Fatal(Text, ex);
                    if (log.IsFatalEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.Info:
                    log.Info(Text, ex);
                    if (log.IsInfoEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.Warn:
                    log.Warn(Text, ex);
                    if (log.IsWarnEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
            }
        }

        private enum WhichLogType
        {
            Info, Debug, Error, Fatal, Warn
        }
#endregion

    }
}
