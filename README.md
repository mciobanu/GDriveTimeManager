# Google Drive™ Time Manager

The main purpose of this is to set the modified time for a Google Drive™ folder (or more) to 
reflect the latest time a file within that folder was changed, looking at all the files directly in that 
folder as well as recursively in its sub-folders.
If a folder is empty, its time will be set to January 1st 1970, 12:00 UTC.

The default behavior for a folder's "modified" time is to reflect when the user created or last changed the folder
itself, while files or folders in it have no impact on the folder's "modified" time.

The modified times can also be set for files.

Another purpose is to serve as an example of using *Google Apps Script™* to implement custom functionality in a
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
            
Here is how this example looks like when setting folder times:
![set dir1 folder times](https://github.com/mciobanu/GDriveTimeManager/blob/main/publish/set-time-example.png?raw=true)

## Limitations
* The script wasn't tested with shared drives, and it's unclear what to expect in such a scenario.
* Shortcuts are currently skipped, as their change time seems best ignored.
* For folders that have other owners, we don't try to set the date. (It doesn't seem to work, anyway.)

## Installation
* Go to [Google Drive](https://drive.google.com/drive/) and create a new spreadsheet. Let's say you call it *Drive
  Time Manager* and you put it in the root of your *Drive*, but you can give it any name and put it in some subfolder,
  just make sure you remember where you put it and how you called it.
* With the new spreadsheet open, open the *Extensions* menu and click on *App Script*.
* In the new browser window/tab that opens, give a meaningful name to the project, like *Drive Time Manager*, and then
  replace the autogenerated code with the content of the file *google-drive-set-time.js*, which you can get
  from [here](https://raw.githubusercontent.com/mciobanu/GDriveTimeManager/main/google-drive-set-time.js). Save (*Control+S*, or the *Save* button).
* On the left, click on the *Project settings* and enable *Show "appsscript.json" manifest file in editor*
* Click on the "<>" button on the left to go back to the editor, select the file *appsscript.json*, copy your
  *timeZone* value to clipboard, overwrite it with the content from [here](https://raw.githubusercontent.com/mciobanu/GDriveTimeManager/main/appscript.json),
  then restore your *timeZone* and save.
* Close both the *App Script* tab and your new spreadsheet's tab.
* Reopen the spreadsheet. It should be changed in this way:
  * It should have a sheet called *Folders* and one called *Files*, with the background color set
    for some cells. Some of these also have text, which shouldn't be changed. **Note**: Creating these sheets may 
    take tens of seconds, so please be patient and let it finish. In the end, it should look something like this:
    ![fresh](https://github.com/mciobanu/GDriveTimeManager/blob/main/publish/empty.png?raw=true)
  * It should have a custom menu called *Modification times*, as the last menu entry

**Note**: A *Google Workspace Marketplace™* add-on is in development, which uses the same code but is somehow easier to
install. A link will be provided here when that becomes available.
<!--- ttt0 Update what is created automatically if deferring what's not needed --->

## Usage
You can run the script for the whole drive or just for some folders or files. In *Column A* of the respective sheet,
you can specify folders or files by name or by ID, under the automatically created cells. You can add more rows
if you need them.

A reason to use IDs is when you have multiple folders or files with the same name. Then, when you try to run the script,
error messages will be created in *Column B* for such folders or files, and you can see the full path and the ID there
and choose the one you want.

If you don't enter any name or ID, the script runs on the whole drive.

To run the script, use the newly added custom menu entries, described below.
Note that the first time you run it, you need to give it access to your drive, which requires choosing a *gMail*
account and clicking on the *Advanced* link to enable access rights, as it is not verified by Google. (It would be
the same for whatever scripts you create that access *Drive*.)

**Note:** While the other permissions should look reasonable given the add-on's functionality, the request to *See your 
primary Google Account email address* is due to what seems to be a bug in determining if the current user is the owner
of a file or folder. There is a flag for this, but it stopped working during development. This is described in the
comment of the function `getOwnedByMe()`. (In turn, this is needed because it seems that you cannot change the time
for a file or folder you don't own.)

**Note**: The request for granting permissions can come at a random moment, stopping the script from setting up the
sheets and leaving them in an inconsistent state. Should that happen, you can rename the affected sheet and start
again, for example by closing the spreadsheet and reopening.

While running, the script logs what it does in the spreadsheet. Feel free to delete the log or parts of it after that.

**Do not** change the content of the "label" cells, which got created automatically, as the script uses them to
determine where names and IDs begin and end.

If you rename the *Folders* or *Files* sheets, they will be recreated. This way it's easier to start over or keep
a snapshot of your data at some point in time. This can also help with you changing the wrong cells or potential
bugs in the script.

### Menu entries
                  
![Menu entries](https://github.com/mciobanu/GDriveTimeManager/blob/main/publish//menu-manual.png?raw=true)
* *Validate folder data*: Takes names and IDs in *Column A* and makes sure they are valid (they may not exist or a
  name may refer to multiple folders). The results are put in *Column B*. The main reason to use IDs is that you might
  have multiple folders with the same name and don't want to process them all. The IDs can be copied from *Column B*,
  after getting the error of having multiple folders with the same name.
  * **Note**: If no names or IDs are entered, the assumption is that we process the whole drive.
* *Set time for specified folders*: Goes through the folders given by name or ID in colum *A* and sets the change time
  as the most recent time of the files inside that folder (directly or somewhere in the folder tree). Validation is
  run before it, and the script doesn't proceed if there are any errors.
* *List content of specified  folders*: Generates the list with the files in the given folders and puts it in the log area.
* *Validate file data*: Like *Validate folder data* but for files: names and IDs must exist and be unique. Unlike for
  folders, you need to specify some files. (If a reason becomes apparent for setting the change time to the same value
  for all the files, this might get implemented.) The results are now in *Column C*, as *Column B* is used to specify
  the new date for a particular file. The formats you can use for dates depend on your locale settings. Another thing
  that is checked here is that there is a 1:1 correspondence between dates and names or IDs.
* *Set time for specified files*: Calls the validation and, if all is fine, sets the times as instructed.

## Notes for developers
* The dependency on `@types/google-apps-script` is just to help auto-completion, when editing files locally.
* Currently, everything is in one file, which is easy to install but rather unwieldy to work with. This can be changed,
  if it seems useful

## Future work
* Perhaps add some options, like whether to ignore the shortcuts or not. These could be put in a dedicated sheet.
* Maybe add a "dry-run" mode.

## Trademark notices
Registered trademarks and service marks are the property of their respective owners.

Google Drive™, Google Workspace Marketplace™, and Google Apps Script™ are trademarks of Google LLC.

[//]: # (ttt0 replace all references to "script" with "add-on", here and in HTML)
