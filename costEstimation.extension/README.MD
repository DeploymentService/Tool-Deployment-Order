# COST ESTIMATION APP

The Cost Estimation App is a module that is intended to get elements within a Revit Project that will be counted in order to get the BOM/BOQ.

This script creates a table containing the data we are interested in for each element.

<br>
<h3><strong>
The script filters the elements by two layers, which are the following:
</strong></h3>

<ol>
<li> Only the elements that have a value in the Category Type Parameter (it is a custom shared parameter, not the Built-In one) will be counted, if that parameter is null in a certain type, that type will not be counted.
<li> Only the elements that are present in a schedule will be counted. If the element is present in more that one schedule (repeated) it will be counted just one time.
</ol>

<br>
<h2>OAUTH</h2>

The credentials.json file can be found in the <strong>folder 'credentials'</strong> within the pushbutton folder. This file must be put in:
> %APPDATA%\Autodesk\Revit\Addins\GoogleAPI\

When the script is run for the first time it will create a <i>token.json</i> file inside that same folder. Then the script is going to take you to your browser in order to get your permissions to get access to Google Drive and Google Sheets, make sure log in with your <strong>@ra.house</strong> account to give these accesses.

Once this is done, the script will prompt you to specify a name for the spreadsheet to be exported.

For now, all the spreadsheets are getting exported to this folder in Google Drive <a>https://drive.google.com/drive/u/1/folders/1ormo29Yt6Pk83MKB5zRJaYVq5UaJ7IOw</a>, then, you can take and copy/move the recently exported schedule to the folder of your preference.

<br>
<h2>COMMENTS</h2>
<ul>
<li>At this moment because of technical limitations the script only works when it is run once. If you want to run the script once more please close and open Revit again.
</ul>