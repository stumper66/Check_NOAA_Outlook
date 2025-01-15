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
    public partial class MainProgram
    {
        public MainProgram()
        {
            m_strCD = Environment.CurrentDirectory;
            if (!m_strCD.EndsWith('\\')) m_strCD += "\\";

            m_Locations = [];

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
        private static readonly Regex RegexTo = new(@"\(.*\)");
        private static readonly Regex RegexTitle = new(".*title=.*\"Categorical Outlook\"");
        private static readonly Regex RegexOnClick = new(@"\((.*?)\)");

        private bool LoadSettings()
        {
            string appSettingsFile = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

            if (!System.IO.File.Exists(appSettingsFile))
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "File not found: " + appSettingsFile);
                return false;
            }

            string appSettingsText;
            try
            { appSettingsText = System.IO.File.ReadAllText(appSettingsFile); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to read " + appSettingsFile, ex);
                return false;
            }

            Dictionary<string, object> appSettings;
            try
            {
                appSettings = new Dictionary<string, object>(
                    System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(
                    appSettingsText, m_JsonOpts), StringComparer.CurrentCultureIgnoreCase);
            }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to deserialize config file " + appSettingsFile, ex);
                return false;
            }

            string temp = (string)GetAppSetting("LoggingLevel", "INFO", appSettings);
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
            temp = (string)GetAppSetting("LogFile", null, appSettings);
            if (!string.IsNullOrEmpty(temp)) this.log.LogFile = Environment.ExpandEnvironmentVariables(temp);

            AppDomain.CurrentDomain.SetData("LogDir", System.IO.Path.GetDirectoryName(this.log.LogFile));

            int maxLogSizeMB = Convert.ToInt32(GetAppSetting("LogMaxFileSizeMB", 0, appSettings));
            int maxLogRotations = Convert.ToInt32(GetAppSetting("LogFilesToKeep", 0, appSettings));

            if (maxLogSizeMB > 0)
            {
                bool didRotate = this.log.CheckIfRotationNeeded(maxLogRotations, maxLogSizeMB, out string RotationMsg);
                if (didRotate)
                    log.Info("Log rotation: " + RotationMsg);
                else if (RotationMsg != null)
                    log.Debug("Log rotation: " + RotationMsg);
            }

            string versionStr = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            log.InfoFormat("------------------------------ Program starting.  Version {0} -----------------------------------------", versionStr);

            m_EmailHost = (string)GetAppSetting("Email_Host", "localhost", appSettings);
            m_EmailUser = (string)GetAppSetting("Email_Username", null, appSettings);
            m_EmailPassword = (string)GetAppSetting("Email_Password", null, appSettings);
            m_EmailPort = Convert.ToInt32(GetAppSetting("Email_Port", 587, appSettings));
            m_UseSSL = Convert.ToBoolean(GetAppSetting("Email_UseSSL", false, appSettings));
            m_EmailFrom = (string)GetAppSetting("Email_From", null, appSettings);
            m_EmailReplyTo = (string)GetAppSetting("Email_ReplyTo", null, appSettings);
            m_Email_HiPri_MinSev = Convert.ToInt32(GetAppSetting("Email_HighPri_Min_Criteria", "2", appSettings));
            m_EmailIsEnabled = Convert.ToBoolean(GetAppSetting("Email_IsEnabled", "true", appSettings));
            m_Download_MaxRetries = Convert.ToInt32(GetAppSetting("Download_Max_Retries", "3", appSettings));
            m_Download_Timeout = Convert.ToInt32(GetAppSetting("Download_Timeout_Seconds", "10", appSettings));
            m_NOAA_URL = (string)GetAppSetting("NOAA_Main_URL", "http://www.spc.noaa.gov/products/outlook/", appSettings);
            m_MesoBaseURL = (string)GetAppSetting("NOAA_Meso_URL", "http://www.spc.noaa.gov/products/outlook/", appSettings);
            //m_EmailTestAddress = (string)GetAppSetting("Email_Test_Address", null, AppSettings);
            m_IncludeMeso = Convert.ToBoolean(GetAppSetting("IncludeMesoDiscussions", "false", appSettings));

            if (!m_NOAA_URL.EndsWith('/')) m_NOAA_URL += "/";

            m_URLs =
            [
                m_NOAA_URL + "day1otlk.html",
                m_NOAA_URL + "day2otlk.html",
                m_NOAA_URL + "day3otlk.html"
            ];
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
            string locationsFile = m_strCD + "Locations.json";
            string trackingFile = m_strCD + "Tracking.json";

            if (!System.IO.File.Exists(locationsFile))
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "File not found: " + locationsFile);
                return false;
            }

            string fileContents;
            try
            { fileContents = System.IO.File.ReadAllText(locationsFile); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to read " + locationsFile, ex);
                return false;
            }

            Dictionary<string, LocationsInfo> locationsDictionary;
            try
            { locationsDictionary = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, LocationsInfo>>(fileContents, m_JsonOpts); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to deserialize " + locationsFile, ex);
                return false;
            }

            if (locationsDictionary == null || locationsDictionary.Count == 0)
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "No locations were defined");
                return false;
            }

            foreach (string locationId in locationsDictionary.Keys)
            {
                LocationsInfo li = locationsDictionary[locationId];
                li.locationId = locationId;

                //LI.LocationId = (int)Row["Id"];
                string loc = li.Coordinates;
                string[] points = loc.Split(',');
                if (points.Length != 2)
                {
                    WriteToConsoleAndLogIt(WhichLogType.FATAL, "Invalid coordinates: " + loc);
                    return false;
                }

                li.location = new System.Drawing.Point(
                    Convert.ToInt32(points[0]),
                    Convert.ToInt32(points[1])
                    );

                if (!string.IsNullOrWhiteSpace(li.MesoMap_Coordinates))
                {
                    points = li.MesoMap_Coordinates.Split(',');
                    if (points.Length == 2)
                    {
                        li.mesoMapLocation = new System.Drawing.Point(
                            Convert.ToInt32(points[0]),
                            Convert.ToInt32(points[1])
                            );
                    }
                }

                m_Locations.Add(li);
            }

            if (!System.IO.File.Exists(trackingFile))
                return true;

            try
            { fileContents = System.IO.File.ReadAllText(trackingFile); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to read " + trackingFile, ex);
                return false;
            }

            try
            { m_Tracking = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, TrackingInfo>>(fileContents, m_JsonOpts); }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.FATAL, "Unable to deserialize " + trackingFile, ex);
                return false;
            }

            return true;
        }

        private static object GetAppSetting(string propertyName, object defaultValue, Dictionary<string, object> settingsDictionary)
        {
            if (!settingsDictionary.TryGetValue(propertyName, out object value))
                return defaultValue;

            object temp = value;
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

            return value;
        }

        public int RunMain()
        {
            log = new Logging("Check_NOAA_Outlook.log");

            int result = RunMain2();
            log.Dispose();

            return result;
        }

        private int RunMain2()
        {
            if (!LoadSettings()) return 10;

            m_MainClient = new System.Net.Http.HttpClient();
            //WriteToConsoleAndLogIt(WhichLogType.INFO, string.Format("Starting program, locations: {0}", m_Locations.Count));

            // step 1, download the 3 days HTML and pull the image URLs from them
            List<string> imageURLs = GetAllImageURLs();
            if (imageURLs == null)
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "ImageURLs was null");
                return 3;
            }

            if (m_IncludeMeso)
            {
                MesoParseResult mesoResult = ParseMainMesoDiscussion();
                if (mesoResult != MesoParseResult.Found_Mesos)
                {
                    if (mesoResult == MesoParseResult.Had_Error) WriteToConsoleAndLogIt(WhichLogType.ERROR, "Had Error parsing main meso discussion");
                    else WriteToConsoleAndLogIt(WhichLogType.INFO, "No meso discussions present");
                    m_IncludeMeso = false;
                }
                else
                    WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Mesos found: " + m_AllMesoDetails.Count.ToString());
            }

            // step 2, download the GIFs for each of the outlook days
            List<MagickImage> allTheGifs = DownloadAllTheGIFs(imageURLs);
            if (allTheGifs == null)
            {
                WriteToConsoleAndLogIt(WhichLogType.FATAL, "AllTheGIFs was null");
                return 4;
            }

            m_Tracking2 = [];
            // step 3, loop thru each defined location and compare the coordinates from each to the 
            // colors on the downloaded GIFs
            for (int i = 0; i < m_Locations.Count; i++)
            {
                LocationsInfo li = m_Locations[i];
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Checking " + li.Label);

                // ------------------------------ here is our payload logic: ------------------------------
                OutlookResults results = ParseTheGIFs(allTheGifs, li);
                bool useHiPri = false;
                if ((int)results.howSevere >= m_Email_HiPri_MinSev)
                    useHiPri = true;

                // pull previous found info from tracking.xml for comparison
                PreviousResultsInfo pri = GetPreviousResults(li);

                // compare previous results to see if we've already notified for these severities
                // or if the sevs have changed, email on that
                //bool HasPrevResults = CheckPreviousResults(PRI, Results, out SeverityChangedFromPrev, out HighestSevFromPrev, out DayChanged);
                ComparisonAgainPrevResults prevComparison = CheckPreviousResults(pri, results, li);
                //bool HasPrevResults = CheckPreviousResults(PRI, Results, LI, out SeverityChangedFromPrev, out HighestSevFromPrev, out DayChanged);
                bool hasPrevResults = prevComparison != null;
                if (!hasPrevResults) prevComparison = new ComparisonAgainPrevResults();

                bool emailGotSent = false;
                if (!log.IsDebugEnabled)
                {
                    StringBuilder sb = new(li.Label);
                    sb.AppendFormat(" ({0})", li.locationId);

                    foreach (int day in results.severities.Keys)
                    {
                        SeveritiesEnum sev = results.severities[day];
                        sb.AppendFormat(", d{0}: {1}", day + 1, sev);
                    }
                    WriteToConsoleAndLogIt(WhichLogType.INFO, sb.ToString());
                }

                if (hasPrevResults && prevComparison.severityChangedFromPrev && prevComparison.highestSev < li.Min_Severity)
                {
                    WriteToConsoleAndLogIt(WhichLogType.DEBUG,
                        "Got sev changes but min sev didn't meet requirements of at least " + li.Min_Severity.ToString());
                }

#if DEBUG
                //SeverityChangedFromPrev = true;
#endif

                List<MesoDiscussionDetails> applicableMesos = null;
                bool mesoChanged = false;
                if (m_IncludeMeso) applicableMesos = GetMesoDiscussionForLocation(li, pri, out mesoChanged);
                bool hasMesos = applicableMesos != null && applicableMesos.Count > 0;
                string mesoEmail = null;

                if (hasMesos)
                {
                    results.attachedImages ??= [];
                    for (int t = 0; t < applicableMesos.Count; t++)
                    {
                        MesoDiscussionDetails detail = applicableMesos[t];
                        results.attachedImages.Add(detail.id, detail.mesoPicture);
                        if (t == 0) mesoEmail = "\r\n<p><strong>Mesoscale Discussion Below:</strong></p>";
                        mesoEmail += string.Format("\r\n<p><img src=cid:{0} ></p>\r\n<p>{1}</p>", detail.id, detail.mesoText.Replace(
                            "\r\n\r\n", "</p>\r\n\r\n<p>")); //.Replace("\r\n", "<br>\r\n"));
                    }
                    useHiPri = true;
                }

                WriteToConsoleAndLogIt(WhichLogType.DEBUG, string.Format(
                    "HasPrev: {0}, HowSevere: {1}, HighestSevFromPrev: {2}, MinSev: {3}, SevChangedFromPrev: {4}, MesoChanged: {5}, DayChanged: {6}",
                     hasPrevResults, (int)results.howSevere, prevComparison.highestSev, li.Min_Severity,
                     prevComparison.severityChangedFromPrev, mesoChanged, prevComparison.dayChanged));

                if ((int)results.howSevere >= li.Min_Severity)
                    m_LocationsInSevere++;

                if ((prevComparison.dayChanged || !hasPrevResults) && (int)results.howSevere >= li.Min_Severity ||
                    prevComparison.severityChangedFromPrev && prevComparison.highestSev >= li.Min_Severity || mesoChanged)
                {
                    string emailSubject = "Severe Weather Forecast in " + li.Label;
                    if (prevComparison.severityChangedFromPrev)
                    {
                        WriteToConsoleAndLogIt(WhichLogType.INFO, "Sev changed from last run, sending email, l: {0}", li.locationId);
                        emailSubject = string.Format("Severe Weather Forecast in {0} Has Changed!", li.Label);
                    }

                    if (hasMesos)
                    {
                        emailSubject += " Mesoscale Discussion Included";
                    }

                    // -------------------------------------- emails get sent here ------------------------------------
                    if (m_EmailIsEnabled)
                    {
                        emailGotSent = SendEmail(emailSubject, results.emailMessage, results.attachedImages, useHiPri, li, mesoEmail);
                        m_LocationsEmailed++;
                        if (hasMesos) m_MesosIncluded++;
                    }
                    else
                        WriteToConsoleAndLogIt(WhichLogType.INFO, "I wanted to send an email but emails are disabled");
                    //EmailGotSent = true;
                }
                else if (hasPrevResults && (int)results.howSevere > li.Min_Severity)
                {
                    WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Not sending email due to previous results");
                }

                pri ??= new PreviousResultsInfo(li.locationId);

                pri.lastSeverities = results.severities;
                if (hasMesos)
                {
                    pri.mesos = [];
                    for (int t = 0; t < applicableMesos.Count; t++)
                        pri.mesos.Add(applicableMesos[t].filename);
                }
                else
                    pri.mesos = null;

                AddDataForNextTime(pri);
            } // next location

            //m_DS_Tracking.WriteXml(TrackingFile);

            string trackingFile = m_strCD + "Tracking.json";
            string trackingText = System.Text.Json.JsonSerializer.Serialize(m_Tracking2, m_JsonOpts);
            try
            { System.IO.File.WriteAllText(trackingFile, trackingText); }
            catch (Exception ex)
            { WriteToConsoleAndLogItWithException(WhichLogType.ERROR, "Unable to write tracking file: " + trackingFile, ex); }

            WriteToConsoleAndLogIt(WhichLogType.INFO, "Locations count: {0}, locations in severe: {1}, locations emailed: {2}, mesos included: {3}",
                 m_Locations.Count, m_LocationsInSevere, m_LocationsEmailed, m_MesosIncluded);

            return 0;
        }

        private MesoParseResult ParseMainMesoDiscussion()
        {
            m_AllMesoDetails = [];

            if (m_MesoHtml == null)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Meso HTML was null");
                return MesoParseResult.Had_Error;
            }

            string html = m_MesoHtml;
            const StringComparison sc = StringComparison.CurrentCultureIgnoreCase;

            int startingPt = html.IndexOf("<!-- Contents below-->", sc);
            if (startingPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Could not find starting point (0)");
                return MesoParseResult.Had_Error;
            }

            int noMesos = html.IndexOf("<center>No Mesoscale Discussions are currently in effect.</center>", startingPt, sc);
            if (noMesos > 0)
            {
                WriteToConsoleAndLogIt(WhichLogType.INFO, "No Mesoscale Discussions are currently in effect");
                return MesoParseResult.No_Meso;
            }

            startingPt = html.IndexOf("<map name=\"mdimgmap\">", sc);
            if (startingPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Could not find starting point (2)");
                return MesoParseResult.Had_Error;
            }

            int endPoint = html.IndexOf("</map>", startingPt + 21, sc);
            if (endPoint < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Could not find end point");
                return MesoParseResult.Had_Error;
            }

            html = html.Substring(startingPt + 21, endPoint - startingPt - 21);
            /*
             <map name="mdimgmap">
 <area shape="poly" coords="463,122,475,125,494,128,493,138,483,142,462,141,453,130,463,122,463,122" href="https://www.spc.noaa.gov/products/md/md0302.html" title="Mesoscale Discussion # 302">
 <area shape="poly" coords="398,142,406,135,419,132,427,131,439,129,451,123,462,126,454,144,441,149,415,152,388,144,398,142,398,142" href="https://www.spc.noaa.gov/products/md/md0301.html" title="Mesoscale Discussion # 301">
 </map>
            */

            const string regexStr = @"(\""poly\"")(.*?)(coords=\"")(.*?)(\"")(.*?)(href=\"")(.*?)(\"")";
            RegexOptions opts = RegexOptions.IgnoreCase | RegexOptions.Multiline;
            MatchCollection matches = Regex.Matches(html, regexStr, opts);

            if (matches == null || matches.Count == 0)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "No regex matches for foly poly coordinates");
                return MesoParseResult.Had_Error;
            }

            //WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Match count: " + Matches.Count.ToString());
            for (int i = 0; i < matches.Count; i++)
            {
                MesoDiscussionDetails details = new();

                Match m = matches[i];

                if (m.Groups.Count < 9)
                {
                    WriteToConsoleAndLogIt(WhichLogType.DEBUG, $"Match {i}, group count was {m.Groups.Count}, expected >= 9");
                    return MesoParseResult.Had_Error;
                }

                details.polyCoords = m.Groups[4].Value;
                details.filename = m.Groups[8].Value;

                if (!string.IsNullOrEmpty(details.polyCoords))
                {
                    string[] split = details.polyCoords.Split(',');
                    List<System.Drawing.Point> points = new(split.Length);
                    for (int t = 0; t < split.Length; t += 2)
                    {
                        points.Add(new System.Drawing.Point(
                            Convert.ToInt32(split[t].Trim()),
                            Convert.ToInt32(split[t + 1].Trim())
                            ));
                    }
                    details.poly = points;
                }

                //WriteToConsoleAndLogIt(WhichLogType.DEBUG, (string.Format("{0}: Poly: {1}, link: {2}", i, Poly, Link));
                //DownloadMesoDiscussion(BaseURL, Link);

                if (details.poly != null && details.poly.Count > 0)
                    m_AllMesoDetails.Add(details.id, details);
                else
                    WriteToConsoleAndLogIt(WhichLogType.INFO, string.Format("Skipping meso {0} as there was no coords", details.filename));
            }
            return MesoParseResult.Found_Mesos;
        }

        private bool DownloadMesoDiscussionDetails(Guid whichOne)
        {
            MesoDiscussionDetails details = m_AllMesoDetails[whichOne];
            string url = m_MesoBaseURL + details.filename;

            WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Downloading Meso: " + url);

            int retryCount = 0;
            string html = null;

            while (retryCount <= m_Download_MaxRetries)
            {
                Task<string> response = m_MainClient.GetStringAsync(url);

                try
                { response.Wait(m_Download_Timeout * 1000); }
                catch (Exception ex)
                {
                    Exception ex2 = Helpers.GetSingleExceptionFromAggregateException(ex);

                    retryCount++;
                    if (retryCount <= m_Download_MaxRetries)
                        WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, retry {2}", url, ex2.Message, retryCount);
                    else
                        WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, cancelling", url, ex2.Message);
                    continue;
                }

                if (response.IsCompletedSuccessfully)
                {
                    html = response.Result;
                    break;
                }
                else
                {
                    if (retryCount <= m_Download_MaxRetries)
                    {
                        if (response.IsFaulted)
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, retry {1}",
                                url, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message, retryCount);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Timed out, retry {0}", retryCount);
                    }
                    else
                    {
                        if (response.IsFaulted)
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, cancelling",
                                url, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.FATAL, "Timed out, cancelling");
                    }
                }
            } // next retry

            if (retryCount > m_Download_MaxRetries)
                return false;

            // we got the HTML, need to parse it for the the text and image.  Will download image
            if (!ParseMesoDetails(details, html))
                return false;

            List<string> downloadUrl = [m_MesoBaseURL + details.gifName];

            List<MagickImage> result = DownloadAllTheGIFs(downloadUrl);
            if (result == null || result.Count == 0)
                return false;

            details.mesoPicture = result[0];
            details.isDownloaded = true;

            return true;
        }

        private bool ParseMesoDetails(MesoDiscussionDetails details, string mesoHtml)
        {
            string html = mesoHtml;
            string regexStr = @"(<title>)(.*?)(</title>)";
            RegexOptions opts = RegexOptions.IgnoreCase | RegexOptions.Multiline;
            Match match = Regex.Match(html, regexStr, opts);

            details.title = details.filename; // if we can't find the title we'll use the filename
            if (match.Success && match.Groups.Count > 2)
                details.title = match.Groups[2].Value.Trim();

            const StringComparison sc = StringComparison.CurrentCultureIgnoreCase;
            int startingPt = html.IndexOf("a name=\"contents\"", sc);

            if (startingPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, details.filename + ", could not find starting point (0)");
                return false;
            }

            // <img src="mcd0330.gif"
            // download gif:
            regexStr = @"(img src=\"")(.*?)(\"")";

            html = html[(startingPt + 18)..];

            match = Regex.Match(html, regexStr, opts);
            if (!match.Success)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Unable to locate gif in html " + details.filename);
                return false;
            }

            details.gifName = match.Groups[2].Value;

            startingPt = html.IndexOf("<pre>", sc);

            if (startingPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, details.filename + ", could not find starting point (1)");
                return false;
            }

            int endPt = html.IndexOf("</pre>", 0, sc);

            if (endPt < 1)
            {
                WriteToConsoleAndLogIt(WhichLogType.DEBUG, details.filename + ", could not find end point");
                return false;
            }

            html = html.Substring(startingPt + 5, endPt - startingPt - 5);
            //Html = Html.Substring(0, EndPt);
            StringBuilder sb = new();
            using (System.IO.StringReader sr = new(html))
            {
                bool foundStart = false;
                while (sr.Peek() >= 0)
                {
                    string line = sr.ReadLine();
                    if (line == null) continue;
                    line = line.Trim();

                    if (!foundStart)
                    {
                        if (line.Length == 0) continue;
                        foundStart = true;
                    }

                    sb.AppendLine(line);
                }
            }

            details.mesoText = sb.ToString();

            // remove trailing carriage returns, up to 5
            for (int i = 0; i < 5; i++)
            {
                if (details.mesoText.EndsWith('\n'))
                    details.mesoText = details.mesoText[0..^1];
                else
                    break;
            }

            return true;
        }

        private List<MesoDiscussionDetails> GetMesoDiscussionForLocation(LocationsInfo li, PreviousResultsInfo pri, out bool mesoChanged)
        {
            mesoChanged = false; // only true if there are mesos present now that wasn't before

            if (m_AllMesoDetails.Count == 0) return null;
            if (li.mesoMapLocation.IsEmpty) return null;

            List<MesoDiscussionDetails> mesos = null;

            foreach (Guid id in m_AllMesoDetails.Keys)
            {
                MesoDiscussionDetails detail = m_AllMesoDetails[id];
                if (detail.IsPointInPolygon(li.mesoMapLocation))
                {
                    // found a meso that is for this location

                    // first see if we've previously seen this filename
                    if (pri == null || pri.mesos == null || !pri.mesos.Contains(detail.filename))
                        mesoChanged = true;

                    WriteToConsoleAndLogIt(WhichLogType.DEBUG, string.Format(
                        "Found applicabled meso for Id: {0}, filename: {1}, IsDownloaded: {2}",
                        li.locationId, detail.filename, detail.isDownloaded));
                    if (!detail.isDownloaded)
                    {
                        bool downloadedIt = DownloadMesoDiscussionDetails(id);
                        if (!downloadedIt)
                        {
                            WriteToConsoleAndLogIt(WhichLogType.INFO, "Was unable to fully download meso discussion");
                            return null;
                        }
                    }

                    mesos ??= [];
                    mesos.Add(detail);
                }
            }

            return mesos;
        }

        private ComparisonAgainPrevResults CheckPreviousResults(PreviousResultsInfo pri, OutlookResults or, LocationsInfo li)
        {
            ComparisonAgainPrevResults prvResults = new();

            SeveritiesEnum highestSevEnum = SeveritiesEnum.Nothing;
            DateTime now = DateTime.Now;
            DateTime useDate = new(now.Year, now.Month, now.Day, 0, 0, 0, DateTimeKind.Local);
            DateTime yesterday = useDate.AddDays(-1);
            string[] iOrD = ["{IncreaseOrDecrease0}", "{IncreaseOrDecrease1}", "{IncreaseOrDecrease2}"];

            if (pri == null || pri.lastSeverities == null || !pri.lastCheckTime.HasValue)
            {
                if (or.emailMessage != null)
                {
                    foreach (string text in iOrD)
                        or.emailMessage = or.emailMessage.Replace(text, string.Empty);
                }

                return null;
            }

            //if (PRI != null && PRI.LastSeverities != null &&
            //    (PRI.LastCheckTime.HasValue && PRI.LastCheckTime.Value.ToLocalTime() > UseDate) ||
            //    PRI != null && !PRI.LastCheckTime.HasValue)

            //bool LastCheckWasPrevDay = false;

            if (pri.lastCheckTime.HasValue &&
                pri.lastCheckTime.Value.ToLocalTime() < useDate &&
                pri.lastCheckTime.Value.ToLocalTime() > yesterday)
            {
                prvResults.dayChanged = true;

                // we last checked some time yesterday.  Compare the prev day 2 and 3 to see if the severity changed overnight
                // to do so, we'll shift days prev days 2 and 3 to look like today's 1 and 2
            }

            foreach (int useDay in pri.lastSeverities.Keys)
            {
                int usePrevDay = useDay;
                if (prvResults.dayChanged)
                {
                    if (pri.lastSeverities.ContainsKey(usePrevDay + 1))
                        usePrevDay++;
                    else
                    {
                        if (or.emailMessage != null)
                            or.emailMessage = or.emailMessage.Replace(iOrD[useDay], string.Empty);
                        continue;
                    }
                }

                SeveritiesEnum lastSev = pri.lastSeverities[usePrevDay];
                SeveritiesEnum currentSev = or.severities[useDay];

                if (lastSev > highestSevEnum)
                    highestSevEnum = lastSev;
                if (currentSev > highestSevEnum)
                    highestSevEnum = currentSev;

                if (lastSev > currentSev &&
                    (int)lastSev >= li.Min_Severity)
                {
                    // severity decreased and it was previously at or above the min severity level
                    prvResults.severityChangedFromPrev = true;
                    WriteToConsoleAndLogIt(WhichLogType.INFO, "Day {0} Sev decreased from {1} to {2}", useDay, lastSev, currentSev);
                    or.emailMessage = or.emailMessage.Replace(iOrD[useDay],
                        " (<span style=\"background-color: #FFFF00\">decreased</span> from " + lastSev.ToString() + ")");
                }
                else if (lastSev < currentSev &&
                    (int)currentSev >= li.Min_Severity)
                {
                    // severity increased and is at or above the min severity level
                    if (lastSev <= SeveritiesEnum.Nothing) // we don't care if it increased from nothing
                    {
                        WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Found increase from nothing, so treating it as new");
                        or.emailMessage = or.emailMessage.Replace(iOrD[useDay], string.Empty);
                    }
                    else
                    {
                        WriteToConsoleAndLogIt(WhichLogType.INFO, "Day {0} Sev increased from {1} to {2}", useDay, lastSev, currentSev);
                        or.emailMessage = or.emailMessage.Replace(iOrD[useDay],
                            " <span style=\"background-color: #FFFF00\">(increased from " + lastSev.ToString() + ")</span>");
                    }
                    prvResults.severityChangedFromPrev = true;
                }
                else
                {
                    // no change in sev, or it was below thresholds
                    if (or.emailMessage != null)
                        or.emailMessage = or.emailMessage.Replace(iOrD[useDay], string.Empty);
                }
            } // next day

            prvResults.highestSev = (int)highestSevEnum;
            return prvResults;
        }

        private void AddDataForNextTime(PreviousResultsInfo pri)
        {
            //DataView DV = new DataView(m_DS_Tracking.Tables[0]);
            //DV.RowFilter = string.Format("Location_Id = {0}", PRI.LocationId);

            string mesos = null;
            if (pri.mesos != null && pri.mesos.Count > 0)
            {
                for (int i = 0; i < pri.mesos.Count; i++)
                {
                    if (i == 0) mesos = pri.mesos[i];
                    else mesos += "; " + pri.mesos[i];
                }
            }

            StringBuilder sb = new();
            foreach (int day in pri.lastSeverities.Keys)
            {
                if (sb.Length > 0) sb.Append(';');
                sb.Append((int)pri.lastSeverities[day]);
            }

            TrackingInfo ti = new()
            {
                Last_Check = DateTime.UtcNow,
                Last_Mesos = mesos,
                Last_Severities = sb.ToString()
            };

            m_Tracking2.Add(pri.locationId, ti);
        }

        private PreviousResultsInfo GetPreviousResults(LocationsInfo li)
        {
            //if (m_DS_Tracking.Tables[0].Rows.Count < 1) return null;
            if (m_Tracking == null || !m_Tracking.TryGetValue(li.locationId, out TrackingInfo value)) return null;

            TrackingInfo ti = value;

            PreviousResultsInfo pri = new()
            {
                locationId = li.locationId
            };
            if (ti.Last_Check > DateTime.MinValue) pri.lastCheckTime = ti.Last_Check;
            if (!string.IsNullOrEmpty(ti.Last_Severities))
            {
                string[] days = ti.Last_Severities.Split(';');
                pri.lastSeverities = [];
                for (int i = 0; i < days.Length; i++)
                {
                    string theDay = days[i].Trim();
                    SeveritiesEnum sev = Enum.Parse<SeveritiesEnum>(theDay);
                    pri.lastSeverities.Add(i, sev);
                }
            }
            if (!string.IsNullOrEmpty(ti.Last_Mesos))
            {
                pri.mesos = [];
                string[] temp2 = ti.Last_Mesos.Split(';');
                for (int i = 0; i < temp2.Length; i++)
                    pri.mesos.Add(temp2[i].Trim());
            }

            return pri;
        }

        private OutlookResults ParseTheGIFs(List<MagickImage> theGIFs, LocationsInfo li)
        {
            OutlookResults results = new();
            StringBuilder sb = new();
            results.severities = [];

            for (int i = 0; i < theGIFs.Count; i++)
            {
                MagickImage gif = theGIFs[i];

                bool wasSevere = false;
                SevereCategoryValues scv = new(gif, m_Cities);

                System.Drawing.Point usePoint = new(li.location.X, li.location.Y);
                SeveritiesEnum severity = scv.WhatSeverityIsThis(usePoint, false);
                results.severities.Add(i, severity);

                if (severity > results.howSevere)
                    results.howSevere = severity;

                int day = i + 1;
                Guid useGuid = Guid.Empty;

                int checkValue = li.Min_Severity_For_Picture ?? 2;
                if ((int)severity >= checkValue) // MRGL and up default
                {
                    useGuid = Guid.NewGuid();
                    MagickImage theCrop = scv.GetCropOfMyArea(usePoint);
                    results.attachedImages ??= [];
                    results.attachedImages.Add(useGuid, theCrop);
                    results.wasSevere = true;
                    wasSevere = true;
                }

                // <img src=""cid:{0}"" />

                string summaryText = null;
                if (wasSevere)
                {
                    results.wasSevere = true;
                    summaryText = ExtractSummaryFromSource(i);
                }

                string imageInfo = wasSevere ?
                    string.Format("<br><b>Summary:</b> {2}<br>\r\n<a href=\"{0}\"><img src=cid:{1} ></a>",
                    m_URLs[i], useGuid, summaryText) :
                    string.Empty;

                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Day {0}, Severity: {1}",
                    day, severity);
                string increaseOrDecrease = "{IncreaseOrDecrease" + i.ToString() + "}";
                DateTime useDate = DateTime.Now.AddDays(i);
                sb.AppendLine(string.Format("<p><a href=\"{0}\">Day {1}</a> - {2}, Severity: {3}{4}{5}</p>",
                   m_URLs[i], day, useDate.DayOfWeek, severity, increaseOrDecrease, imageInfo));
            }

            //if (Results.WasSevere)
            results.emailMessage = sb.ToString();

            return results;
        }

        private List<MagickImage> DownloadAllTheGIFs(List<string> urls)
        {
            List<MagickImage> results = [];

            List<string> theURLs = urls;

            for (int i = 0; i < theURLs.Count; i++)
            {
                string url = theURLs[i];

                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Downloading: " + url);
                int retryCount = 0;
                while (retryCount <= m_Download_MaxRetries)
                {
                    Task<byte[]> response = m_MainClient.GetByteArrayAsync(url);

                    try
                    { response.Wait(m_Download_Timeout * 1000); }
                    catch (Exception ex)
                    {
                        Exception ex2 = Helpers.GetSingleExceptionFromAggregateException(ex);

                        retryCount++;
                        if (retryCount <= m_Download_MaxRetries)
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, retry {2}", url, ex2.Message, retryCount);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, cancelling", url, ex2.Message);
                        continue;
                    }

                    if (response.IsCompletedSuccessfully)
                    {
                        try
                        {
                            MagickImage img = new(response.Result);
                            results.Add(img);
                        }
                        catch (Exception ex)
                        {
                            WriteToConsoleAndLogItWithException(WhichLogType.ERROR, "Unable to convert to image " + url, ex);
                            return null;
                        }
                        break;
                    }
                    else
                    {
                        retryCount++;

                        if (retryCount <= m_Download_MaxRetries)
                        {
                            if (response.IsFaulted)
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, retry {1}",
                                    url, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message, retryCount);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Timed out, retry {0}", retryCount);
                        }
                        else
                        {
                            if (response.IsFaulted)
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, cancelling",
                                    url, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.FATAL, "Timed out, cancelling");
                        }
                    }
                } // next retry

                if (retryCount > m_Download_MaxRetries)
                    return null;
            } // next URL

            return results;
        }

        private static List<MailboxAddress> ParseEmailRecipients(string input)
        {
            if (string.IsNullOrEmpty(input)) return null;

            string[] toPeople = input.Split(';');
            List<MailboxAddress> addresses = [];

            for (int i = 0; i < toPeople.Length; i++)
            {
                string to = toPeople[i].Trim();
                Match m = RegexTo.Match(to);
                if (m.Success)
                {
                    string displayname = m.Value[1..^1];
                    string email = to[..m.Index].Trim();

                    MailboxAddress ma = new(displayname, email);
                    addresses.Add(ma);
                }
                else
                    addresses.Add(new MailboxAddress(null, to));
            }

            if (addresses.Count > 0) return addresses;
            else return null;
        }

        private bool SendEmail(string subject, string body, Dictionary<Guid, MagickImage> images, bool useHighPri, LocationsInfo li, string meso)
        {
            using SmtpClient client = new();

            MimeMessage message = new();
            message.From.Add(new MailboxAddress(null, m_EmailFrom));
            if (useHighPri) message.Priority = MessagePriority.Urgent;

#if DEBUG
            //if (!string.IsNullOrEmpty(m_EmailTestAddress))
            //    Recips = m_EmailTestAddress;
#endif

            // optionally email address can have a display name.  Syntax:
            // email@address.com (Display Name)
            List<MailboxAddress> toPeople = ParseEmailRecipients(li.Email_Recipients);
            if (toPeople != null)
            {
                for (int i = 0; i < toPeople.Count; i++)
                    message.To.Add(toPeople[i]);
            }

            List<MailboxAddress> bccPeople = ParseEmailRecipients(li.Email_Recipients_BCC);
            if (bccPeople != null)
            {
                for (int i = 0; i < bccPeople.Count; i++)
                    message.Bcc.Add(bccPeople[i]);
            }

            string extra = string.Empty;
            if (images != null && images.Count > 0)
            {
                Guid g = Guid.NewGuid();

                body += string.Format("\r\n<p><img src=cid:{0} ></p>", g);
                images.Add(g, m_Legend);
            }

            if (!string.IsNullOrEmpty(m_EmailReplyTo))
                message.ReplyTo.Add(new MailboxAddress(null, m_EmailReplyTo));


            string meso2 = meso ?? string.Empty;
            string useBody = "<html><body>\r\n" + body + "\r\n" + meso2 + "</body></html>";

            message.Subject = subject;
            BodyBuilder bb = new();

            AddImagesToEmail(bb, images);

            bb.HtmlBody = useBody;
            message.Body = bb.ToMessageBody();

#if DEBUG
            System.IO.StreamWriter sw = new(@"c:\temp\test.html", false);
            sw.Write(useBody);
            sw.Close();
#endif

            bool itGotSent = false;
            try
            {
                client.Connect(m_EmailHost, m_EmailPort, m_UseSSL);
                client.AuthenticationMechanisms.Remove("XOAUTH2");
                client.Authenticate(m_EmailUser, m_EmailPassword);

                client.Send(message);
                client.Disconnect(true);
                itGotSent = true;
            }
            catch (Exception ex)
            {
                WriteToConsoleAndLogItWithException(WhichLogType.ERROR, "Error sending email:", ex);
            }

            if (itGotSent)
                WriteToConsoleAndLogIt(WhichLogType.INFO, string.Format("{0} ({1}), Email got sent", li.Label, li.locationId));

            return itGotSent;
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

        private string ExtractSummaryFromSource(int whichDay)
        {
            string source = m_SourceCodes[whichDay];

            const StringComparison sc = StringComparison.CurrentCultureIgnoreCase;
            int start = source.IndexOf("...SUMMARY...", sc);
            if (start < 0) return null;

            StringBuilder sb = new();
            string source2 = source[start..];
            System.IO.StringReader sr = new(source2);
            sr.ReadLine();
            while (sr.Peek() > 0)
            {
                string Line = sr.ReadLine();
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
            string url_Gif = m_NOAA_URL + "day{0}{1}.gif";

            List<string> results = [];
            List<string> theURLs = m_URLs;

            for (int i = 0; i < theURLs.Count; i++)
            {
                string url = theURLs[i];

                WriteToConsoleAndLogIt(WhichLogType.DEBUG, "Downloading: " + url);
                int retryCount = 0;
                while (retryCount <= m_Download_MaxRetries)
                {
                    Task<string> response = m_MainClient.GetStringAsync(url);

                    try
                    { response.Wait(m_Download_Timeout * 1000); }
                    catch (Exception ex)
                    {
                        Exception ex2 = Helpers.GetSingleExceptionFromAggregateException(ex);

                        retryCount++;
                        if (retryCount <= m_Download_MaxRetries)
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, retry {2}", url, ex2.Message, retryCount);
                        else
                            WriteToConsoleAndLogIt(WhichLogType.WARN, "Error downloading {0}, {1}, cancelling", url, ex2.Message);
                        continue;
                    }

                    if (response.IsCompletedSuccessfully)
                    {
                        string source = response.Result;
                        if (i < 3)
                        {
                            // day 1-3 outlook
                            m_SourceCodes[i] = source;

                            // a title="Categorical Outlook"
                            Match m = RegexTitle.Match(source);
                            if (!m.Success) return null;

                            // <td OnClick="show_tab('otlk_0100')" OnMouseOver="show_tab('otlk_0100')"><a title="Categorical Outlook"
                            m = RegexOnClick.Match(m.Value);
                            if (!m.Success) return null;

                            //('otlk_0100')
                            string value = m.Value[2..^2];
                            // otlk_0100

                            int useDay = i + 1;
                            string useURL = string.Format(url_Gif, useDay, value);
                            results.Add(useURL);
                        }
                        else
                        {
                            // meso discussion
                            m_MesoHtml = source;
                        }

                        break;
                    }
                    else
                    {
                        retryCount++;

                        if (retryCount <= m_Download_MaxRetries)
                        {
                            if (response.IsFaulted)
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, retry {1}",
                                    url, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message, retryCount);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Timed out, retry {0}", retryCount);
                        }
                        else
                        {
                            if (response.IsFaulted)
                                WriteToConsoleAndLogIt(WhichLogType.WARN, "Got exception downloading {0}, {1}, cancelling",
                                    url, Helpers.GetSingleExceptionFromAggregateException(response.Exception).Message);
                            else
                                WriteToConsoleAndLogIt(WhichLogType.FATAL, "Timed out, cancelling");
                        }
                    }
                } // next retry

                if (retryCount > m_Download_MaxRetries)
                    return null;
            }

            return results;
        }

        private void WriteToConsoleAndLogIt(WhichLogType logtype, string text)
        {
            switch (logtype)
            {
                case WhichLogType.DEBUG:
                    log.Debug(text);
                    if (log.IsDebugEnabled) Console.WriteLine(text);
                    break;
                case WhichLogType.ERROR:
                    log.Error(text);
                    if (log.IsErrorEnabled) Console.WriteLine(text);
                    break;
                case WhichLogType.FATAL:
                    log.Fatal(text);
                    if (log.IsFatalEnabled) Console.WriteLine(text);
                    break;
                case WhichLogType.INFO:
                    log.Info(text);
                    if (log.IsInfoEnabled) Console.WriteLine(text);
                    break;
                case WhichLogType.WARN:
                    log.Warn(text);
                    if (log.IsWarnEnabled) Console.WriteLine(text);
                    break;
            }
        }

        private void WriteToConsoleAndLogIt(WhichLogType logtype, string text, params object[] args)
        {
            switch (logtype)
            {
                case WhichLogType.DEBUG:
                    log.DebugFormat(text, args);
                    if (log.IsDebugEnabled) Console.WriteLine(text, args);
                    break;
                case WhichLogType.ERROR:
                    log.ErrorFormat(text, args);
                    if (log.IsErrorEnabled) Console.WriteLine(text, args);
                    break;
                case WhichLogType.FATAL:
                    log.FatalFormat(text, args);
                    if (log.IsFatalEnabled) Console.WriteLine(text, args);
                    break;
                case WhichLogType.INFO:
                    log.InfoFormat(text, args);
                    if (log.IsInfoEnabled) Console.WriteLine(text, args);
                    break;
                case WhichLogType.WARN:
                    log.WarnFormat(text, args);
                    if (log.IsWarnEnabled) Console.WriteLine(text, args);
                    break;
            }
        }

        private void WriteToConsoleAndLogItWithException(WhichLogType logtype, string text, Exception ex)
        {
            switch (logtype)
            {
                case WhichLogType.DEBUG:
                    log.Debug(text, ex);
                    if (log.IsDebugEnabled)
                    {
                        Console.WriteLine(text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.ERROR:
                    log.Error(text, ex);
                    if (log.IsErrorEnabled)
                    {
                        Console.WriteLine(text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.FATAL:
                    log.Fatal(text, ex);
                    if (log.IsFatalEnabled)
                    {
                        Console.WriteLine(text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.INFO:
                    log.Info(text, ex);
                    if (log.IsInfoEnabled)
                    {
                        Console.WriteLine(text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
                case WhichLogType.WARN:
                    log.Warn(text, ex);
                    if (log.IsWarnEnabled)
                    {
                        Console.WriteLine(text);
                        Console.WriteLine("Exception: " + ex.Message);
                    }
                    break;
            }
        }
    }
}
