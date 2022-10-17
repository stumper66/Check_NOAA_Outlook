using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using ImageMagick;
using MailKit.Net.Smtp;
using MimeKit;
using MailKit.Security;

namespace Check_NOAA_Outlook
{
    public class MainProgram
    {
        public MainProgram()
        {
            m_strCD = Environment.CurrentDirectory;
            if (!m_strCD.EndsWith("\\")) m_strCD += "\\";

            m_Locations = new List<LocationsInfo>();

            m_SourceCodes = new string[3];
            //m_GetExtendedOutlook = true;
            m_JsonOpts = new System.Text.Json.JsonSerializerOptions()
            {
                WriteIndented = true,
                ReadCommentHandling = System.Text.Json.JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            };
            m_JsonOpts.Converters.Add(new MyBooleanConverter());
        }

        private Logging log;
        private readonly List<LocationsInfo> m_Locations;
        private readonly string[] m_SourceCodes;
        private readonly System.Text.Json.JsonSerializerOptions m_JsonOpts;
        private readonly string m_strCD;
        private Dictionary<string, TrackingInfo> m_Tracking;
        private Dictionary<string, TrackingInfo> m_Tracking2;
        private MagickImage m_Legend, m_Cities; // m_Day48_Lenend;
        private List<string> m_URLs;
        private string m_MesoHtml; // m_URL_Extended
        private Dictionary<Guid, MesoDiscussionDetails> m_AllMesoDetails;
        private string m_EmailHost, m_EmailFrom, m_EmailReplyTo, m_NOAA_URL, m_MesoBaseURL; // m_EmailTestAddress
        private string m_EmailUser, m_EmailPassword;
        private bool m_UseSSL;
        private int m_EmailPort;
        private int m_Email_HiPri_MinSev, m_Download_MaxRetries, m_Download_Timeout;
        private bool m_EmailIsEnabled, m_IncludeMeso; // m_GetExtendedOutlook
        private System.Net.Http.HttpClient m_MainClient;
        private int m_LocationsInSevere, m_LocationsEmailed, m_MesosIncluded;

        private bool LoadSettings()
        {
            string AppSettingsFile = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

            if (!System.IO.File.Exists(AppSettingsFile))
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "File not found: " + AppSettingsFile);
                return false;
            }

            string AppSettingsText;
            try
            { AppSettingsText = System.IO.File.ReadAllText(AppSettingsFile); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to read " + AppSettingsFile, ex);
                return false;
            }

            Dictionary<string, object> AppSettings;
            try
            {
                AppSettings = new Dictionary<string, object>(
                    System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(
                    AppSettingsText, m_JsonOpts), StringComparer.CurrentCultureIgnoreCase);
            }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to deserialize config file " + AppSettingsFile, ex);
                return false;
            }

            string temp = (string)GetAppSetting("LoggingLevel", "INFO", AppSettings);
            if (string.IsNullOrEmpty(temp)) temp = "INFO";
            log.m_LogLevel = (temp.ToLower()) switch
            {
                "debug" => WhichLogType.DEBUG,
                "error" => WhichLogType.ERROR,
                "fatal" => WhichLogType.FATAL,
                "none" => WhichLogType.NONE,
                "warn" => WhichLogType.WARN,
                _ => WhichLogType.INFO,
            };
            string Temp = (string)GetAppSetting("LogFile", null, AppSettings);
            if (!string.IsNullOrEmpty(Temp)) this.log.LogFile = Environment.ExpandEnvironmentVariables(Temp);

            AppDomain.CurrentDomain.SetData("LogDir", System.IO.Path.GetDirectoryName(this.log.LogFile));

            int MaxLogSizeMB = Convert.ToInt32(GetAppSetting("LogMaxFileSizeMB", 0, AppSettings));
            int MaxLogRotations = Convert.ToInt32(GetAppSetting("LogFilesToKeep", 0, AppSettings));

            if (MaxLogSizeMB > 0)
            {
                bool DidRotate = this.log.CheckIfRotationNeeded(MaxLogRotations, MaxLogSizeMB, out string RotationMsg);
                if (DidRotate)
                    log.Info("Log rotation: " + RotationMsg);
                else if (RotationMsg != null)
                    log.Debug("Log rotation: " + RotationMsg);
            }

            string VersionStr = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            log.InfoFormat("------------------------------ Program starting.  Version {0} -----------------------------------------", VersionStr);

            m_EmailHost = (string)GetAppSetting("Email_Host", "localhost", AppSettings);
            m_EmailUser = (string)GetAppSetting("Email_Username", null, AppSettings);
            m_EmailPassword = (string)GetAppSetting("Email_Password", null, AppSettings);
            m_EmailPort = Convert.ToInt32(GetAppSetting("Email_Port", 587, AppSettings));
            m_UseSSL = Convert.ToBoolean(GetAppSetting("Email_UseSSL", false, AppSettings));
            m_EmailFrom = (string)GetAppSetting("Email_From", null, AppSettings);
            m_EmailReplyTo = (string)GetAppSetting("Email_ReplyTo", null, AppSettings);
            m_Email_HiPri_MinSev = Convert.ToInt32(GetAppSetting("Email_HighPri_Min_Criteria", "2", AppSettings));
            m_EmailIsEnabled = Convert.ToBoolean(GetAppSetting("Email_IsEnabled", "true", AppSettings));
            m_Download_MaxRetries = Convert.ToInt32(GetAppSetting("Download_Max_Retries", "3", AppSettings));
            m_Download_Timeout = Convert.ToInt32(GetAppSetting("Download_Timeout_Seconds", "10", AppSettings));
            m_NOAA_URL = (string)GetAppSetting("NOAA_Main_URL", "http://www.spc.noaa.gov/products/outlook/", AppSettings);
            m_MesoBaseURL = (string)GetAppSetting("NOAA_Meso_URL", "http://www.spc.noaa.gov/products/outlook/", AppSettings);
            //m_EmailTestAddress = (string)GetAppSetting("Email_Test_Address", null, AppSettings);
            m_IncludeMeso = Convert.ToBoolean(GetAppSetting("IncludeMesoDiscussions", "false", AppSettings));

            if (!m_NOAA_URL.EndsWith("/")) m_NOAA_URL += "/";

            m_URLs = new List<string>
            {
                m_NOAA_URL + "day1otlk.html",
                m_NOAA_URL + "day2otlk.html",
                m_NOAA_URL + "day3otlk.html"
            };
            if (m_IncludeMeso) m_URLs.Add(m_MesoBaseURL);

            //m_URL_Extended = "http://www.spc.noaa.gov/products/exper/day4-8/";

            if (System.IO.File.Exists(m_strCD + "legend.jpg"))
                m_Legend = new MagickImage(m_strCD + "legend.jpg");
            else
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "required file was not present: legend.jpg");
                return false;
            }


            if (System.IO.File.Exists(m_strCD + "cities.gif"))
                m_Cities = new MagickImage(m_strCD + "cities.gif");
            else
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "required file was not present: cities.jpg");
                return false;
            }

            //if (System.IO.File.Exists(m_strCD + "day48_legend.gif"))
            //    m_Day48_Lenend = new MagickImage(m_strCD + "day48_legend.gif");
            //else
            //    WriteToConsoleAndLogIt(WhichLogType.DEBUG, "file was not present: day48_legend.jpg");

            if (!LoadJsonFiles()) return false;

            return true;
        }

        private bool LoadJsonFiles()
        {
            string LocationsFile = m_strCD + "Locations.json";
            string TrackingFile = m_strCD + "Tracking.json";

            if (!System.IO.File.Exists(LocationsFile))
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "File not found: " + LocationsFile);
                return false;
            }

            string FileContents;
            try
            { FileContents = System.IO.File.ReadAllText(LocationsFile); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to read " + LocationsFile, ex);
                return false;
            }

            Dictionary<string, LocationsInfo> LocationsDictionary;
            try
            { LocationsDictionary = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, LocationsInfo>>(FileContents, m_JsonOpts); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to deserialize " + LocationsFile, ex);
                return false;
            }

            if (LocationsDictionary == null || LocationsDictionary.Count == 0)
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "No locations were defined");
                return false;
            }

            foreach (string LocationId in LocationsDictionary.Keys)
            {
                LocationsInfo LI = LocationsDictionary[LocationId];
                LI.LocationId = LocationId;

                //LI.LocationId = (int)Row["Id"];
                string Loc = LI.Coordinates;
                string[] Points = Loc.Split(',');
                if (Points.Length != 2)
                {
                    WriteToConsoleAndLogIt(WhichLogType.FATAL, "Invalid coordinates: " + Loc);
                    return false;
                }

                LI.Location = new System.Drawing.Point(
                    Convert.ToInt32(Points[0]),
                    Convert.ToInt32(Points[1])
                    );

                if (!string.IsNullOrWhiteSpace(LI.MesoMap_Coordinates))
                {
                    Points = LI.MesoMap_Coordinates.Split(',');
                    if (Points.Length == 2)
                    {
                        LI.MesoMapLocation = new System.Drawing.Point(
                            Convert.ToInt32(Points[0]),
                            Convert.ToInt32(Points[1])
                            );
                    }
                }

                m_Locations.Add(LI);
            }

            if (!System.IO.File.Exists(TrackingFile))
                return true;

            try
            { FileContents = System.IO.File.ReadAllText(TrackingFile); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to read " + TrackingFile, ex);
                return false;
            }

            try
            { m_Tracking = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, TrackingInfo>>(FileContents, m_JsonOpts); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to deserialize " + TrackingFile, ex);
                return false;
            }

            return true;
        }

        private static object GetAppSetting(string PropertyName, object DefaultValue, Dictionary<string, object> SettingsDictionary)
        {
            if (!SettingsDictionary.ContainsKey(PropertyName))
                return DefaultValue;

            object temp = SettingsDictionary[PropertyName];
            if (temp is not null and System.Text.Json.JsonElement)
            {
                System.Text.Json.JsonElement e = (System.Text.Json.JsonElement)temp;
                return e.ValueKind switch
                {
                    System.Text.Json.JsonValueKind.String => e.ToString(),
                    System.Text.Json.JsonValueKind.Number => e.GetDouble(),
                    System.Text.Json.JsonValueKind.Null => null,
                    System.Text.Json.JsonValueKind.False => false,
                    System.Text.Json.JsonValueKind.True => true,
                    _ => throw new System.IO.InvalidDataException(e.ValueKind.ToString() + " type not supported in appsettings.json"),
                };
            }

            return SettingsDictionary[PropertyName];
        }

        public int RunMain()
        {
            log = new Logging("Check_NOAA_Outlook.log");

            int Result = RunMain2();
            log.Dispose();

            return Result;
        }

        private int RunMain2()
        {
            if (!LoadSettings()) return 10;

            m_MainClient = new System.Net.Http.HttpClient();
            //WriteToConsoleAndLogIt(WhichLogType.INFO, string.Format("Starting program, locations: {0}", m_Locations.Count));

            // step 1, download the 3 days HTML and pull the image URLs from them
            List<string> ImageURLs = GetAllImageURLs();
            if (ImageURLs == null)
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "ImageURLs was null");
                return 3;
            }

            if (m_IncludeMeso)
            {
                MesoParseResult MesoResult = ParseMainMesoDiscussion();
                if (MesoResult != MesoParseResult.Found_Mesos)
                {
                    if (MesoResult == MesoParseResult.Had_Error) WriteToConsoleAndLogIt(WhichLogType.ERROR, "Had Error parsing main meso discussion");
                    else WriteToConsoleAndLogIt(WhichLogType.INFO, "No meso discussions present");
                    m_IncludeMeso = false;
                }
                else
                    WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Mesos found: " + m_AllMesoDetails.Count.ToString());
            }

            // step 2, download the GIFs for each of the outlook days
            List<MagickImage> AllTheGifs = DownloadAllTheGIFs(ImageURLs);
            if (AllTheGifs == null)
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "AllTheGIFs was null");
                return 4;
            }

            m_Tracking2 = new Dictionary<string, TrackingInfo>();
            // step 3, loop thru each defined location and compare the coordinates from each to the 
            // colors on the downloaded GIFs
            for (int i = 0; i < m_Locations.Count; i++)
            {
                LocationsInfo LI = m_Locations[i];
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Checking " + LI.Label);

                // ------------------------------ here is our payload logic: ------------------------------
                OutlookResults Results = ParseTheGIFs(AllTheGifs, LI);
                bool UseHiPri = false;
                if ((int)Results.HowSevere >= m_Email_HiPri_MinSev)
                    UseHiPri = true;

                // pull previous found info from tracking.xml for comparison
                PreviousResultsInfo PRI = GetPreviousResults(LI);

                // compare previous results to see if we've already notified for these severities
                // or if the sevs have changed, email on that
                //bool HasPrevResults = CheckPreviousResults(PRI, Results, out SeverityChangedFromPrev, out HighestSevFromPrev, out DayChanged);
                ComparisonAgainPrevResults PrevComparison = CheckPreviousResults(PRI, Results, LI);
                //bool HasPrevResults = CheckPreviousResults(PRI, Results, LI, out SeverityChangedFromPrev, out HighestSevFromPrev, out DayChanged);
                bool HasPrevResults = PrevComparison != null;
                if (!HasPrevResults) PrevComparison = new ComparisonAgainPrevResults();

                bool EmailGotSent = false;
                if (!log.IsDebugEnabled)
                {
                    StringBuilder sb = new(LI.Label);
                    sb.AppendFormat(" ({0})", LI.LocationId);

                    foreach (int Day in Results.Severities.Keys)
                    {
                        SeveritiesEnum Sev = Results.Severities[Day];
                        sb.AppendFormat(", d{0}: {1}", Day + 1, Sev);
                    }
                    WriteToConsoleAndLogIt(WhichLogType.INFO, sb.ToString());
                }

                if (HasPrevResults && PrevComparison.SeverityChangedFromPrev && PrevComparison.HighestSev < LI.Min_Severity)
                {
                    WriteToConsoleAndLogIt(WhichLogType.DEBUG,
                        "Got sev changes but min sev didn't meet requirements of at least " + LI.Min_Severity.ToString());
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
                    Results.AttachedImages ??= new Dictionary<Guid, MagickImage>();
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

                WriteToConsoleAndLogIt(WhichLogType.DEBUG, string.Format(
                    "HasPrev: {0}, HowSevere: {1}, HighestSevFromPrev: {2}, MinSev: {3}, SevChangedFromPrev: {4}, MesoChanged: {5}, DayChanged: {6}",
                     HasPrevResults, (int)Results.HowSevere, PrevComparison.HighestSev, LI.Min_Severity,
                     PrevComparison.SeverityChangedFromPrev, MesoChanged, PrevComparison.DayChanged));

                if ((int)Results.HowSevere >= LI.Min_Severity)
                    m_LocationsInSevere++;

                if ((PrevComparison.DayChanged || !HasPrevResults) && (int)Results.HowSevere >= LI.Min_Severity ||
                    PrevComparison.SeverityChangedFromPrev && PrevComparison.HighestSev >= LI.Min_Severity || MesoChanged)
                {
                    string EmailSubject = "Severe Weather Forecast in " + LI.Label;
                    if (PrevComparison.SeverityChangedFromPrev)
                    {
                        WriteToConsoleAndLogIt(WhichLogType.INFO, "Sev changed from last run, sending email, l: {0}", LI.LocationId);
                        EmailSubject = string.Format("Severe Weather Forecast in {0} Has Changed!", LI.Label);
                    }

                    if (HasMesos)
                    {
                        EmailSubject += " Mesoscale Discussion Included";
                    }

                    // -------------------------------------- emails get sent here ------------------------------------
                    if (m_EmailIsEnabled)
                    {
                        EmailGotSent = SendEmail(EmailSubject, Results.EmailMessage, Results.AttachedImages, UseHiPri, LI, MesoEmail);
                        m_LocationsEmailed++;
                        if (HasMesos) m_MesosIncluded++;
                    }
                    else
                        WriteToConsoleAndLogIt(WhichLogType.INFO, "I wanted to send an email but emails are disabled");
                    //EmailGotSent = true;
                }
                else if (HasPrevResults && (int)Results.HowSevere > LI.Min_Severity)
                {
                    WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Not sending email due to previous results");
                }

                PRI ??= new PreviousResultsInfo(LI.LocationId);

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

            //m_DS_Tracking.WriteXml(TrackingFile);

            string TrackingFile = m_strCD + "Tracking.json";
            string TrackingText = System.Text.Json.JsonSerializer.Serialize(m_Tracking2, m_JsonOpts);
            try
            { System.IO.File.WriteAllText(TrackingFile, TrackingText); }
            catch (Exception ex)
            { WriteToConsoleAndLogItWithException(WhichLogType.ERROR, "Unable to write tracking file: " + TrackingFile, ex); }

            WriteToConsoleAndLogIt(WhichLogType.INFO, "Locations count: {0}, locations in severe: {1}, locations emailed: {2}, mesos included: {3}",
                 m_Locations.Count, m_LocationsInSevere, m_LocationsEmailed, m_MesosIncluded);

            return 0;
        }

        private MesoParseResult ParseMainMesoDiscussion()
        {
            m_AllMesoDetails = new Dictionary<Guid, MesoDiscussionDetails>();

            if (m_MesoHtml == null)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Meso HTML was null");
                return MesoParseResult.Had_Error;
            }

            string Html = m_MesoHtml;
            const StringComparison SC = StringComparison.CurrentCultureIgnoreCase;

            int StartingPt = Html.IndexOf("<!-- Contents below-->", SC);
            if (StartingPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Could not find starting point (0)");
                return MesoParseResult.Had_Error;
            }

            int NoMesos = Html.IndexOf("<center>No Mesoscale Discussions are currently in effect.</center>", StartingPt, SC);
            if (NoMesos > 0)
            {
                WriteToConsoleAndLogIt(WhichLogType.INFO, "No Mesoscale Discussions are currently in effect");
                return MesoParseResult.No_Meso;
            }

            StartingPt = Html.IndexOf("<map name=\"mdimgmap\">", SC);
            if (StartingPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Could not find starting point (2)");
                return MesoParseResult.Had_Error;
            }

            int EndPoint = Html.IndexOf("</map>", StartingPt + 21, SC);
            if (EndPoint < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Could not find end point");
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
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "No regex matches for foly poly coordinates");
                return MesoParseResult.Had_Error;
            }

            //WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Match count: " + Matches.Count.ToString());
            for (int i = 0; i < Matches.Count; i++)
            {
                MesoDiscussionDetails Details = new();

                Match M = Matches[i];

                if (M.Groups.Count < 9)
                {
                    WriteToConsoleAndLogIt(WhichLogType.DEBUG, $"Match {i}, group count was {M.Groups.Count}, expected >= 9");
                    return MesoParseResult.Had_Error;
                }

                Details.PolyCoords = M.Groups[4].Value;
                Details.Filename = M.Groups[8].Value;

                if (!string.IsNullOrEmpty(Details.PolyCoords))
                {
                    string[] Split = Details.PolyCoords.Split(',');
                    List<System.Drawing.Point> Points = new(Split.Length);
                    for (int t = 0; t < Split.Length; t += 2)
                    {
                        Points.Add(new System.Drawing.Point(
                            Convert.ToInt32(Split[t].Trim()),
                            Convert.ToInt32(Split[t + 1].Trim())
                            ));
                    }
                    Details.Poly = Points;
                }

                //WriteToConsoleAndLogIt(WhichLogType.DEBUG, (string.Format("{0}: Poly: {1}, link: {2}", i, Poly, Link));
                //DownloadMesoDiscussion(BaseURL, Link);

                if (Details.Poly != null && Details.Poly.Count > 0)
                    m_AllMesoDetails.Add(Details.Id, Details);
                else
                    WriteToConsoleAndLogIt(WhichLogType.INFO, string.Format("Skipping meso {0} as there was no coords", Details.Filename));
            }
            return MesoParseResult.Found_Mesos;
        }

        private bool DownloadMesoDiscussionDetails(Guid WhichOne)
        {
            MesoDiscussionDetails Details = m_AllMesoDetails[WhichOne];
            string URL = m_MesoBaseURL + Details.Filename;

            WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Downloading Meso: " + URL);

            int RetryCount = 0;
            string Html = null;

            while (RetryCount <= m_Download_MaxRetries)
            {
                Task<string> response = m_MainClient.GetStringAsync(URL);

                try
                { response.Wait(m_Download_Timeout * 1000); }
                catch (Exception ex)
                {
                    Exception ex2 = Helpers.GetSingleExceptionFromAggregateException(ex);

                    RetryCount++;
                    if (RetryCount <= m_Download_MaxRetries)
                        WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, retry {2}", URL, ex2.Message, RetryCount);
                    else
                        WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, cancelling", URL, ex2.Message);
                    continue;
                }

                if (response.IsCompletedSuccessfully)
                {
                    Html = response.Result;
                    break;
                }
                else
                {
                    if (RetryCount <= m_Download_MaxRetries)
                    {
                        if (response.IsFaulted)
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, retry {1}",
                                URL, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message, RetryCount);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Timed out, retry {0}", RetryCount);
                    }
                    else
                    {
                        if (response.IsFaulted)
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, cancelling",
                                URL, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.FATAL, "Timed out, cancelling");
                    }
                }
            } // next retry

            if (RetryCount > m_Download_MaxRetries)
                return false;

            // we got the HTML, need to parse it for the the text and image.  Will download image
            if (!ParseMesoDetails(Details, Html))
                return false;

            List<string> DownloadUrl = new()
            { m_MesoBaseURL + Details.GifName };

            List<MagickImage> Result = DownloadAllTheGIFs(DownloadUrl);
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
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, Details.Filename + ", could not find starting point (0)");
                return false;
            }

            // <img src="mcd0330.gif"
            // download gif:
            RegexStr = @"(img src=\"")(.*?)(\"")";

            Html = Html[(StartingPt + 18)..];

            Match = Regex.Match(Html, RegexStr, Opts);
            if (!Match.Success)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Unable to locate gif in html " + Details.Filename);
                return false;
            }

            Details.GifName = Match.Groups[2].Value;

            StartingPt = Html.IndexOf("<pre>", SC);

            if (StartingPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, Details.Filename + ", could not find starting point (1)");
                return false;
            }

            int EndPt = Html.IndexOf("</pre>", 0, SC);

            if (EndPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, Details.Filename + ", could not find end point");
                return false;
            }

            Html = Html.Substring(StartingPt + 5, EndPt - StartingPt - 5);
            //Html = Html.Substring(0, EndPt);
            StringBuilder sb = new();
            using (System.IO.StringReader SR = new(Html))
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
                    Details.MesoText = Details.MesoText[0..^1];
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
                    if (PRI == null || PRI.Mesos == null || !PRI.Mesos.Contains(Detail.Filename))
                        MesoChanged = true;

                    WriteToConsoleAndLogIt(WhichLogType.DEBUG, string.Format(
                        "Found applicabled meso for Id: {0}, filename: {1}, IsDownloaded: {2}",
                        LI.LocationId, Detail.Filename, Detail.IsDownloaded));
                    if (!Detail.IsDownloaded)
                    {
                        bool DownloadedIt = DownloadMesoDiscussionDetails(Id);
                        if (!DownloadedIt)
                        {
                            WriteToConsoleAndLogIt(WhichLogType.INFO, "Was unable to fully download meso discussion");
                            return null;
                        }
                    }

                    Mesos ??= new List<MesoDiscussionDetails>();
                    Mesos.Add(Detail);
                }
            }

            return Mesos;
        }

        private ComparisonAgainPrevResults CheckPreviousResults(PreviousResultsInfo PRI, OutlookResults OR, LocationsInfo LI)
        {
            ComparisonAgainPrevResults PrvResults = new();

            SeveritiesEnum HighestSevEnum = SeveritiesEnum.Nothing;
            DateTime Now = DateTime.Now;
            DateTime UseDate = new(Now.Year, Now.Month, Now.Day, 0, 0, 0, DateTimeKind.Local);
            DateTime Yesterday = UseDate.AddDays(-1);
            string[] IOrD = new string[] { "{IncreaseOrDecrease0}", "{IncreaseOrDecrease1}", "{IncreaseOrDecrease2}" };

            if (PRI == null || PRI.LastSeverities == null || !PRI.LastCheckTime.HasValue)
            {
                if (OR.EmailMessage != null)
                {
                    foreach (string Text in IOrD)
                        OR.EmailMessage = OR.EmailMessage.Replace(Text, string.Empty);
                }

                return null;
            }

            //if (PRI != null && PRI.LastSeverities != null &&
            //    (PRI.LastCheckTime.HasValue && PRI.LastCheckTime.Value.ToLocalTime() > UseDate) ||
            //    PRI != null && !PRI.LastCheckTime.HasValue)

            //bool LastCheckWasPrevDay = false;

            if (PRI.LastCheckTime.HasValue &&
                PRI.LastCheckTime.Value.ToLocalTime() < UseDate &&
                PRI.LastCheckTime.Value.ToLocalTime() > Yesterday)
            {
                PrvResults.DayChanged = true;

                // we last checked some time yesterday.  Compare the prev day 2 and 3 to see if the severity changed overnight
                // to do so, we'll shift days prev days 2 and 3 to look like today's 1 and 2
            }

            foreach (int UseDay in PRI.LastSeverities.Keys)
            {
                int UsePrevDay = UseDay;
                if (PrvResults.DayChanged)
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

                if (LastSev > CurrentSev &&
                    (int)LastSev >= LI.Min_Severity)
                {
                    // severity decreased and it was previously at or above the min severity level
                    PrvResults.SeverityChangedFromPrev = true;
                    WriteToConsoleAndLogIt(WhichLogType.INFO, "Day {0} Sev decreased from {1} to {2}", UseDay, LastSev, CurrentSev);
                    OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay],
                        " (<span style=\"background-color: #FFFF00\">decreased</span> from " + LastSev.ToString() + ")");
                }
                else if (LastSev < CurrentSev &&
                    (int)CurrentSev >= LI.Min_Severity)
                {
                    // severity increased and is at or above the min severity level
                    if (LastSev <= SeveritiesEnum.Nothing) // we don't care if it increased from nothing
                    {
                        WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Found increase from nothing, so treating it as new");
                        OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay], string.Empty);
                    }
                    else
                    {
                        WriteToConsoleAndLogIt(WhichLogType.INFO, "Day {0} Sev increased from {1} to {2}", UseDay, LastSev, CurrentSev);
                        OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay],
                            " <span style=\"background-color: #FFFF00\">(increased from " + LastSev.ToString() + ")</span>");
                    }
                    PrvResults.SeverityChangedFromPrev = true;
                }
                else
                {
                    // no change in sev, or it was below thresholds
                    if (OR.EmailMessage != null)
                        OR.EmailMessage = OR.EmailMessage.Replace(IOrD[UseDay], string.Empty);
                }
            } // next day

            PrvResults.HighestSev = (int)HighestSevEnum;
            return PrvResults;
        }

        private void AddDataForNextTime(PreviousResultsInfo PRI)
        {
            //DataView DV = new DataView(m_DS_Tracking.Tables[0]);
            //DV.RowFilter = string.Format("Location_Id = {0}", PRI.LocationId);

            string Mesos = null;
            if (PRI.Mesos != null && PRI.Mesos.Count > 0)
            {
                for (int i = 0; i < PRI.Mesos.Count; i++)
                {
                    if (i == 0) Mesos = PRI.Mesos[i];
                    else Mesos += "; " + PRI.Mesos[i];
                }
            }

            StringBuilder sb = new();
            foreach (int Day in PRI.LastSeverities.Keys)
            {
                if (sb.Length > 0) sb.Append(';');
                sb.Append((int)PRI.LastSeverities[Day]);
            }

            TrackingInfo TI = new()
            {
                Last_Check = DateTime.UtcNow,
                Last_Mesos = Mesos,
                Last_Severities = sb.ToString()
            };

            m_Tracking2.Add(PRI.LocationId, TI);
        }

        private PreviousResultsInfo GetPreviousResults(LocationsInfo LI)
        {
            //if (m_DS_Tracking.Tables[0].Rows.Count < 1) return null;
            if (m_Tracking == null || !m_Tracking.ContainsKey(LI.LocationId)) return null;

            TrackingInfo TI = m_Tracking[LI.LocationId];

            PreviousResultsInfo PRI = new()
            {
                LocationId = LI.LocationId
            };
            if (TI.Last_Check > DateTime.MinValue) PRI.LastCheckTime = TI.Last_Check;
            if (!string.IsNullOrEmpty(TI.Last_Severities))
            {
                string[] Days = TI.Last_Severities.Split(';');
                PRI.LastSeverities = new Dictionary<int, SeveritiesEnum>();
                for (int i = 0; i < Days.Length; i++)
                {
                    string TheDay = Days[i].Trim();
                    SeveritiesEnum Sev = (SeveritiesEnum)Enum.Parse(typeof(SeveritiesEnum), TheDay);
                    PRI.LastSeverities.Add(i, Sev);
                }
            }
            if (!string.IsNullOrEmpty(TI.Last_Mesos))
            {
                PRI.Mesos = new List<string>();
                string[] Temp2 = TI.Last_Mesos.Split(';');
                for (int i = 0; i < Temp2.Length; i++)
                    PRI.Mesos.Add(Temp2[i].Trim());
            }

            return PRI;
        }

        private OutlookResults ParseTheGIFs(List<MagickImage> TheGIFs, LocationsInfo LI)
        {
            OutlookResults Results = new();
            StringBuilder sb = new();
            Results.Severities = new Dictionary<int, SeveritiesEnum>();

            for (int i = 0; i < TheGIFs.Count; i++)
            {
                MagickImage Gif = TheGIFs[i];

                bool WasSevere = false;
                SevereCategoryValues SCV = new(Gif, m_Cities);

                System.Drawing.Point UsePoint = new(LI.Location.X, LI.Location.Y);
                SeveritiesEnum Severity = SCV.WhatSeverityIsThis(UsePoint, false);
                Results.Severities.Add(i, Severity);

                if (Severity > Results.HowSevere)
                    Results.HowSevere = Severity;

                int Day = i + 1;
                Guid UseGuid = Guid.Empty;

                int CheckValue = LI.Min_Severity_For_Picture ?? 2;
                if ((int)Severity >= CheckValue) // MRGL and up default
                {
                    UseGuid = Guid.NewGuid();
                    MagickImage TheCrop = SCV.GetCropOfMyArea(UsePoint);
                    Results.AttachedImages ??= new Dictionary<Guid, MagickImage>();
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

                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Day {0}, Severity: {1}",
                    Day, Severity);
                string IncreaseOrDecrease = "{IncreaseOrDecrease" + i.ToString() + "}";
                DateTime UseDate = DateTime.Now.AddDays(i);
                sb.AppendLine(string.Format("<p><a href=\"{0}\">Day {1}</a> - {2}, Severity: {3}{4}{5}</p>",
                   m_URLs[i], Day, UseDate.DayOfWeek, Severity, IncreaseOrDecrease, ImageInfo));
            }

            //if (Results.WasSevere)
            Results.EmailMessage = sb.ToString();

            return Results;
        }

        private List<MagickImage> DownloadAllTheGIFs(List<string> URLs)
        {
            List<MagickImage> Results = new();

            List<string> TheURLs = URLs;
            //if (m_GetExtendedOutlook)
            //{
            //    for (int i = 4; i <= 8; i++)
            //    {
            //        string Url = m_URL_Extended + string.Format("/day{0}prob.gif", i);

            //        TheURLs.Add(m_URL_Extended);
            //    }
            //}

            for (int i = 0; i < TheURLs.Count; i++)
            {
                string URL = TheURLs[i];

                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Downloading: " + URL);
                int RetryCount = 0;
                while (RetryCount <= m_Download_MaxRetries)
                {
                    Task<byte[]> response = m_MainClient.GetByteArrayAsync(URL);

                    try
                    { response.Wait(m_Download_Timeout * 1000); }
                    catch (Exception ex)
                    {
                        Exception ex2 = Helpers.GetSingleExceptionFromAggregateException(ex);

                        RetryCount++;
                        if (RetryCount <= m_Download_MaxRetries)
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, retry {2}", URL, ex2.Message, RetryCount);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, cancelling", URL, ex2.Message);
                        continue;
                    }

                    if (response.IsCompletedSuccessfully)
                    {
                        try
                        {
                            MagickImage img = new(response.Result);
                            Results.Add(img);
                        }
                        catch (Exception ex)
                        {
                            WriteToConsoleAndLogItWithException(WhichLogType.ERROR, "Unable to convert to image " + URL, ex);
                            return null;
                        }
                        break;
                    }
                    else
                    {
                        RetryCount++;

                        if (RetryCount <= m_Download_MaxRetries)
                        {
                            if (response.IsFaulted)
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, retry {1}",
                                    URL, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message, RetryCount);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Timed out, retry {0}", RetryCount);
                        }
                        else
                        {
                            if (response.IsFaulted)
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, cancelling",
                                    URL, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.FATAL, "Timed out, cancelling");
                        }
                    }
                } // next retry

                if (RetryCount > m_Download_MaxRetries)
                    return null;
            } // next URL

            return Results;
        }

        private static List<MailboxAddress> ParseEmailRecipients(string Input)
        {
            if (string.IsNullOrEmpty(Input)) return null;

            string[] ToPeople = Input.Split(';');
            List<MailboxAddress> Addresses = new();

            for (int i = 0; i < ToPeople.Length; i++)
            {
                string To = ToPeople[i].Trim();
                Match M = Regex.Match(To, @"\(.*\)");
                if (M.Success)
                {
                    string Displayname = M.Value[1..^1];
                    string Email = To[..M.Index].Trim();

                    MailboxAddress MA = new(Displayname, Email);
                    Addresses.Add(MA);
                }
                else
                    Addresses.Add(new MailboxAddress(null, To));
            }

            if (Addresses.Count > 0) return Addresses;
            else return null;
        }

        private bool SendEmail(string Subject, string Body, Dictionary<Guid, MagickImage> Images, bool UseHighPri, LocationsInfo LI, string Meso)
        {
            using SmtpClient client = new();

            MimeMessage message = new();
            message.From.Add(new MailboxAddress(null, m_EmailFrom));
            if (UseHighPri) message.Priority = MessagePriority.Urgent;

#if DEBUG
            //if (!string.IsNullOrEmpty(m_EmailTestAddress))
            //    Recips = m_EmailTestAddress;
#endif

            // optionally email address can have a display name.  Syntax:
            // email@address.com (Display Name)
            List<MailboxAddress> ToPeople = ParseEmailRecipients(LI.Email_Recipients);
            if (ToPeople != null)
            {
                for (int i = 0; i < ToPeople.Count; i++)
                    message.To.Add(ToPeople[i]);
            }

            List<MailboxAddress> BCCPeople = ParseEmailRecipients(LI.Email_Recipients_BCC);
            if (BCCPeople != null)
            {
                for (int i = 0; i < BCCPeople.Count; i++)
                    message.Bcc.Add(BCCPeople[i]);
            }

            string Extra = string.Empty;
            if (Images != null && Images.Count > 0)
            {
                Guid g = Guid.NewGuid();

                Body += string.Format("\r\n<p><img src=cid:{0} ></p>", g);
                Images.Add(g, m_Legend);
            }

            if (!string.IsNullOrEmpty(m_EmailReplyTo))
                message.ReplyTo.Add(new MailboxAddress(null, m_EmailReplyTo));


            string Meso2 = Meso ?? string.Empty;
            string UseBody = "<html><body>\r\n" + Body + "\r\n" + Meso2 + "</body></html>";

            message.Subject = Subject;
            BodyBuilder bb = new();

            AddImagesToEmail(bb, Images);

            bb.HtmlBody = UseBody;
            message.Body = bb.ToMessageBody();

#if DEBUG
            System.IO.StreamWriter SW = new(@"c:\temp\test.html", false);
            SW.Write(UseBody);
            SW.Close();
#endif

            bool ItGotSent = false;
            try
            {
                client.Connect(m_EmailHost, m_EmailPort, m_UseSSL);
                client.AuthenticationMechanisms.Remove("XOAUTH2");
                client.Authenticate(m_EmailUser, m_EmailPassword);

                client.Send(message);
                client.Disconnect(true);
                ItGotSent = true;
            }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.ERROR, "Error sending email:", ex);
            }

            if (ItGotSent)
                WriteToConsoleAndLogIt(WhichLogType.INFO, string.Format("{0} ({1}), Email got sent", LI.Label, LI.LocationId));

            return ItGotSent;
        }

        private static void AddImagesToEmail(BodyBuilder bb, Dictionary<Guid, MagickImage> images)
        {
            if (images == null) return;

            foreach (Guid g in images.Keys)
            {
                MagickImage bm = images[g];

                using System.IO.MemoryStream ms = new();
                bm.Write(ms, MagickFormat.Jpeg);
                ms.Seek(0, System.IO.SeekOrigin.Begin);
                MimeEntity attachedImage = bb.LinkedResources.Add(g.ToString(), ms);
                attachedImage.ContentId = g.ToString();
            }
        }

        private string ExtractSummaryFromSource(int WhichDay)
        {
            string Source = m_SourceCodes[WhichDay];

            const StringComparison SC = StringComparison.CurrentCultureIgnoreCase;
            int Start = Source.IndexOf("...SUMMARY...", SC);
            if (Start < 0) return null;

            StringBuilder sb = new();
            string Source2 = Source[Start..];
            System.IO.StringReader SR = new(Source2);
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

            List<string> Results = new();
            List<string> TheURLs = m_URLs;

            for (int i = 0; i < TheURLs.Count; i++)
            {
                string URL = TheURLs[i];

                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Downloading: " + URL);
                int RetryCount = 0;
                while (RetryCount <= m_Download_MaxRetries)
                {
                    Task<string> response = m_MainClient.GetStringAsync(URL);

                    try
                    { response.Wait(m_Download_Timeout * 1000); }
                    catch (Exception ex)
                    {
                        Exception ex2 = Helpers.GetSingleExceptionFromAggregateException(ex);

                        RetryCount++;
                        if (RetryCount <= m_Download_MaxRetries)
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, retry {2}", URL, ex2.Message, RetryCount);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, cancelling", URL, ex2.Message);
                        continue;
                    }

                    if (response.IsCompletedSuccessfully)
                    {
                        string Source = response.Result;
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
                            string Value = m.Value[2..^2];
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
                        RetryCount++;

                        if (RetryCount <= m_Download_MaxRetries)
                        {
                            if (response.IsFaulted)
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, retry {1}",
                                    URL, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message, RetryCount);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Timed out, retry {0}", RetryCount);
                        }
                        else
                        {
                            if (response.IsFaulted)
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, cancelling",
                                    URL, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.FATAL, "Timed out, cancelling");
                        }
                    }
                } // next retry

                if (RetryCount > m_Download_MaxRetries)
                    return null;
            }

            return Results;
        }

        private void WriteToConsoleAndLogIt(WhichLogType Logtype, string Text)
        {
            switch (Logtype)
            {
                case WhichLogType.DEBUG:
                    log.Debug(Text);
                    if (log.IsDebugEnabled) Console.WriteLine(Text);
                    break;
                case WhichLogType.ERROR:
                    log.Error(Text);
                    if (log.IsErrorEnabled) Console.WriteLine(Text);
                    break;
                case WhichLogType.FATAL:
                    log.Fatal(Text);
                    if (log.IsFatalEnabled) Console.WriteLine(Text);
                    break;
                case WhichLogType.INFO:
                    log.Info(Text);
                    if (log.IsInfoEnabled) Console.WriteLine(Text);
                    break;
                case WhichLogType.WARN:
                    log.Warn(Text);
                    if (log.IsWarnEnabled) Console.WriteLine(Text);
                    break;
            }
        }

        private void WriteToConsoleAndLogIt(WhichLogType Logtype, string Text, params object[] args)
        {
            switch (Logtype)
            {
                case WhichLogType.DEBUG:
                    log.DebugFormat(Text, args);
                    if (log.IsDebugEnabled) Console.WriteLine(Text, args);
                    break;
                case WhichLogType.ERROR:
                    log.ErrorFormat(Text, args);
                    if (log.IsErrorEnabled) Console.WriteLine(Text, args);
                    break;
                case WhichLogType.FATAL:
                    log.FatalFormat(Text, args);
                    if (log.IsFatalEnabled) Console.WriteLine(Text, args);
                    break;
                case WhichLogType.INFO:
                    log.InfoFormat(Text, args);
                    if (log.IsInfoEnabled) Console.WriteLine(Text, args);
                    break;
                case WhichLogType.WARN:
                    log.WarnFormat(Text, args);
                    if (log.IsWarnEnabled) Console.WriteLine(Text, args);
                    break;
            }
        }

        private void WriteToConsoleAndLogItWithException(WhichLogType Logtype, string Text, Exception ex)
        {
            switch (Logtype)
            {
                case WhichLogType.DEBUG:
                    log.Debug(Text, ex);
                    if (log.IsDebugEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.ERROR:
                    log.Error(Text, ex);
                    if (log.IsErrorEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.FATAL:
                    log.Fatal(Text, ex);
                    if (log.IsFatalEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.INFO:
                    log.Info(Text, ex);
                    if (log.IsInfoEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.WARN:
                    log.Warn(Text, ex);
                    if (log.IsWarnEnabled)
                    {
                        Console.WriteLine(Text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
            }
        }
    }
}
