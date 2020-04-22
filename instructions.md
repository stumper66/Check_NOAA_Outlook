# How to use Check_NOAA_Outlook

There are 6 files required to be present:

Mode  | LastWriteDate | Time   | Length | Name   
---   | ---           | ---    | ---    | ---
-a--- | 04/20/2020    | 08:06  | 37376  | Check_NOAA_Outlook.exe
-a--- | 04/11/2020    | 17:58  | 1906   | Check_NOAA_Outlook.exe.config
-a--- | 02/18/2017    | 22:50  | 8039   | cities.gif
-a--- | 02/18/2017    | 22:53  | 13968  | legend.jpg
-a--- | 04/20/2020    | 08:08  | 3734   | Locations.xml
-a--- | 03/08/2017    | 19:26  | 276480 | log4net.dll

First edit Check_NOAA_Outlook.exe.config and put in the proper values which is mostly email config.  Be sure to specify a proper SMTP host.

![appconfig](/images/appconfig.PNG)

Next you'll need to edit locations.xml and fill in the following:

* Id.  Must be a unique number.  Start from 0
* Coordinates.  The location on the map for the desired location.  See below for more information.
* Email_Recipients.  Enter a semi-colon separated list of email recipients.  For display name put it in parenthesis (optional).  This field is optional if you have BCC recipients below.
* Email_Recipients_BCC.  Optional: Enter a semi-colon separated list of email bcc recipients.
* Min_Severity.  The minimum severity to issue an email, based upon NOAA categories.  0 = nothing, 1 = TSTM, 2, = MRGL, 3 = SLGT, 4 = ENH, 5 = MDT, 6 = HIGH
* Min_Severity_For_Picture.  (Optional) the minimum severity that the picture will be included in the email.
* MesoMap_Coordinates.  (Optional) to receive mesoscale discussions, enter the coordinates from the validmd.png image (more info below).

Once the above files are configured, simply execute Check_NOAA_Outlook.exe as often as desired.  I recommend using a scheduled task to have it run every 5 minutes from 7am to 11pm.

## Locations.xml example
![locations.xml](/images/locationsxml.PNG)

## Running The Program
Simply execute the executable, ideally on a repeating schedule.
![sample1](/images/sample1.PNG)

A log file is produced with the configured logging level in the app.config.
![logsample](/images/logsample.PNG)

## Maps and coordinates

In order for it to know where the alert locations are, you must determine the x and y coordinates based upon the NOAA graphic.  This is the graphic found in any day x webpage, such as: <https://www.spc.noaa.gov/products/outlook/day1otlk.html>

The file is saved in the project as 'map with cities.gif'.  Open this image with an editor (I like to use Paint.NET).  Find your location in the map and note the coordinates for that picture, then use it for locations.xml

You'll need to use a separate set of coordinates for mesoscale discussions.  This image in the project as 'validmd.png'.  It is from: <https://www.spc.noaa.gov/products/md/>

Once you have the coordinates from validmd.png, use them for 'MesoMap_Coordinates' in locations.xml

[Day1]:https://www.spc.noaa.gov/products/outlook/day1otlk.html ("https://www.spc.noaa.gov/products/outlook/day1otlk.html")
[Meso]:https://www.spc.noaa.gov/products/md/   "https://www.spc.noaa.gov/products/md/"
