## Binaries

Requirements:
* A supported version of Windows, such as Server 2019 or Windows 10.
* [Microsoft .Net Core 3.1](https://dotnet.microsoft.com/download/dotnet-core/3.1)

Name | Version | Notes | Size | Date
---  | ---     | ---   | ---  | ---
[Check NOAA Outlook](https://github.com/stumper66/Check_NOAA_Outlook/releases/download/v1.1/Check_NOAA_Outlook-v1.1.zip) | 1.1 | Windows | 12,579 KB | 08/12/2020

## Instructions

Click [here for instructions](/instructions.md)

## Check NOAA Outlook
Checks the [NOAA outlook] and emails the defined recipients if they are in a severe risk area.  Includes any mesoscale discussions as well

The NOAA outlook shows the severe weather outlook for the next 3 days and for the following 4 days occasionally if the forecast shows something long-term.

For days 1-3 you then click on that day to see more detailed information which includes a map, summary and details information.
This program checks days 1-3 and if the defined location(s) fall into a severe category or optionally just thunderstorms it will generate a crop of the area and parses the summary and emails it.

It keeps track of any previous runs and will email of any of the severity categories change.

NOAA outlook 1-3 day sample:
![NOAA Outlook sample](/images/noaa.PNG)

[NOAA outlook]: https://www.spc.noaa.gov/products/outlook/

## Email Examples
Here are examples of what the emails looks like.
![Email 1](/images/email1.PNG)
![Email 2](/images/email2.PNG)

If a mescoscale discussion is detected in your target area, it will be included in the bottom of the email and the subject will say "Mesoscale Discussion Included".
![Email 3](/images/email3.PNG)
