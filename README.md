# Google Drive Time Manager

The main purpose of this is to set the changed time for Google Drive folders such that their "modified" time as 
shown on the site reflects the latest time a file within that folder was changed, looking recursively at all the files.
If a folder is empty, its time will be set to January 1st 1970, 12:00 UTC.

The default behavior for a folder's "modified" time is to reflect when the user created or last changed the folder 
itself, while files or folders in it have no impact on it.
  
A secondary purpose is to serve as an example of using *Google Apps Script* to implement custom functionality in a
spreadsheet.

## Example
We have these files and folders:
* `/dir1/`, changed at 07:00
* `/dir1/dir2/`, changed at 08:00
* `/dir1/dir2/file1`, changed at 09:00
* `/dir1/file2`, changed at 10:00
* `/dir1/dir3/`, changed at 11:00

(To keep things shorter, all times are supposed to be in the same day, so the date is not shown.)
   
After running the script, we'll have these (changes in **bold**):
* `/dir1/`, changed at **10:00**
* `/dir1/dir2/`, changed at **09:00**
* `/dir1/dir2/file1`, changed at 09:00
* `/dir1/file2`, changed at 10:00
* `/dir1/dir3/`, changed at **1970-01-01 12:00 UTC**

## Limitations
* The script wasn't tested with shared drives, and it's unclear what to expect in such a scenario.
* Shortcuts are currently ignored, as their change time seems best ignored.
* For folders that have other owners we don't try to set the date. (It doesn't seem to work, anyway.)

## Installation
* Go to [Google Drive](https://drive.google.com/drive/) and create a new spreadsheet. Let's say you call it *Drive
  Time Manager* and you put it in the root of your *Drive*, but you can give it any name and put it in some subfolder,
  just make sure you remember where you put it and how you called it.
* With the new spreadsheet open, open the *Extensions* menu and click on *App Script*.
* In the new browser window/tab that opens, give a meaningful name to the project, like *Drive Time Manager* and then 
copy the content of the file *google-drive-set-time.js*, which you can get from [here](https://raw.githubusercontent.com/mciobanu/GDriveTimeManager/main/google-drive-set-time.js).
* Click on the *Save* button, then close both the *App Script* and your new spreadsheet's tab.
* Reopen the spreadsheet. It should be changed in this way:
  * It should have a sheet called "Folders", with 3 cells having text in them, and the background color set for some cells 
  * It should have a custom menu called *Modification times*, as the last menu entry

## Usage
You can run the script for the whole drive or just for some folders. In *Column A* you can specify folders by name
or by ID, under the automatically created cells, and you can add more rows if you need them.

A reason to give IDs is when you have multiple folders with the same name. Then, when you try to run the script,
error messages will be created in *Column B* for such folders, and you can see the full path and the ID there and 
choose the one you want.

If you don't enter any name or ID, the script runs on the whole drive.

To run the script, use the newly added custom menu (*Set time for specified folders*, under *Modification times*).

While running, the script logs what it does in the spreadsheet. Feel free to delete the log or parts of it after that.

**Do not** change the content of the "label" cells, which got created automatically, as the script uses them to 
determine where names and IDs begin and end.

## Note for developers
The dependency on `@types/google-apps-script` is just to help auto-completion, when editing files locally.

## Future work
* Perhaps add some options, like whether to ignore the shortcuts or not
* Maybe add a "dry-run" mode.
* Add support for changing the date for files.
