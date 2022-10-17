# Requirements
* A supported version of Windows, such as Server 2019 or Windows 10.
* [Microsoft.Net Core 3.1](https://dotnet.microsoft.com/download/dotnet-core/3.1)
* Internet connectivity for downloading NOAA outlooks and sending email.

# How to use Check_NOAA_Outlook

There are 11 files and 1 directory required to be present:

Mode  | LastWriteDate | Time   | Length | Name   
---   | ---           | ---    | ---    | ---
d---- |        07/15/2020 |  23:23 |             | runtimes
-a--- |        08/12/2020 |  13:11 |          643| appsettings.json
-a--- |        08/12/2020 |  12:44 |         2492| Check_NOAA_Outlook.deps.json
-a--- |        08/12/2020 |  12:46 |        47616| Check_NOAA_Outlook.dll
-a--- |        08/12/2020 |  12:46 |       169984| Check_NOAA_Outlook.exe
-a--- |        08/12/2020 |  12:44 |          240| Check_NOAA_Outlook.runtimeconfig.dev.json
-a--- |        08/12/2020 |  12:44 |          154| Check_NOAA_Outlook.runtimeconfig.json
-a--- |        02/18/2017 |  22:50 |         8039| cities.gif
-a--- |        02/18/2017 |  22:53 |        13968| legend.jpg
-a--- |        08/12/2020 |  12:53 |         1147| Locations.json
-a--- |        07/26/2020 |  14:46 |      1408208| Magick.NET.Core.dll
-a--- |        07/26/2020 |  14:50 |       528592| Magick.NET-Q8-AnyCPU.dll

First edit appsettings.json and put in the proper values which is mostly email config.  Be sure to specify a proper SMTP host.

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

## Locations.json example
![locations.PNG](/images/locations.PNG)

## Running The Program
Simply execute the executable, ideally on a repeating schedule.
![sample1](/images/sample1.PNG)

A log file is produced with the configured logging level in the app.config.
![logsample](/images/logsample.PNG)

## Maps and coordinates

In order for it to know where the alert locations are, you must determine the x and y coordinates based upon the NOAA graphic.  This is the graphic found in any day x webpage, such as: <https://www.spc.noaa.gov/products/outlook/day1otlk.html>

The file is saved in the project as 'map with cities.gif'.  Open this image with an editor (I like to use Paint.NET).  Find your location in the map and note the coordinates for that picture, then use it for locations.xml

You'll need to use a separate set of coordinates for mesoscale discussions.  This image in the project as 'validmd.png'.  It is from: <https://www.spc.noaa.gov/products/md/>

Once you have the coordinates from validmd.png, use them for 'MesoMap_Coordinates' in locations.json

[Day1]:https://www.spc.noaa.gov/products/outlook/day1otlk.html ("https://www.spc.noaa.gov/products/outlook/day1otlk.html")
[Meso]:https://www.spc.noaa.gov/products/md/   "https://www.spc.noaa.gov/products/md/"
