/*

Copyright (c) 2023 Marian Ciobanu

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

 */

const DEFAULT_SOURCE_HEIGHT = 5;

const FOLDERS_SHEET_NAME = 'Folders';
const FILES_SHEET_NAME = 'Files';


const SMALLEST_TIME = '1970-01-01T12:00:00.000Z'; //ttt3 Review if something else would be better. (Hour
// is set at noon, so most timezones will see it as January 1st)

//const ROOT_ID = 'root';

// noinspection JSUnusedGlobalSymbols
function tst02() {
    // const a = SOURCE_LIST;
    // const b = a;
    // const activeSheet = SpreadsheetApp.getActiveSheet();
    // // sheet = ss.getSheets()[0];
    // const range = activeSheet.getRange(11, 1);
    // range.setValue('abc');
    // const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    // if (sheets.length !== 3) {
    //     SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    // }
    //getFoldersSheet().activate();
    // SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].activate();
    // const aa = getFilesSheet().getLastRow();
    // return aa;

    //logS('msg1');
    //logS('msg2');

    //const query = `"1wKMfIBhstKUVf4yRYQlgXUuOlsiI6OZ8" in parents and trashed = false`;
    const query = `title = 'testmtime03' and trashed = false`;
    let pageToken = null;

    const items = Drive.Files.list({
        q: query,
        maxResults: 100,
        pageToken,
    });

    console.log(items.items ? items.items.length : 'no results');
}


/*
{
    namesBegin: 3,
    namesEnd: 12,
    idsBegin: 12,
    idsEnd: 18,
    logsBegin: 18,
    logsEnd: 40,
}
*/

/**
 * Where the ranges of names, IDs, and logs begin and end
 *
 * @typedef {Object} RangeInfo
 * @property {number} namesBegin
 * @property {number} namesEnd
 * @property {number} idsBegin
 * @property {number} idsEnd
 * @property {number} logsBegin
 * @property {number} logsEnd
 */


/**
 * Drive info about a particular ID. If it has multiple parents, only one path through one of them is included
 *
 * @typedef {Object} IdInfo
 *
 * @property {string} id
 * @property {string} path
 * @property {boolean} multiplePaths
 * @property {string} modifiedDate
 * @property {boolean} ownedByMe
 */

/**
 * Information about a folder or file (based on either name or ID, with an optional date; all come from user input).
 * Corresponds to a row in the spreadsheet (one of name, id, name+date, id+date).
 *
 * @typedef {Object} InputInfo
 *
 * @property {IdInfo[]} idInfos for IDs this array has at most 1 element
 * @property {string[]} errors
 * @property {string} [date]
 */


/*

IdInfo: {
    id: "ke8436",
    path: "/kf84/asf",
    multiplePaths: true, // optional, for when there are multiple parents
    modifiedDate: "2000-01-01T10:00:00.000Z",
    ownedByMe: true,
}

InputInfo:
{
    idInfos: [
        { // IdInfo
            id: "hd73hb",
            path: "/de/fd/gth",
            multiplePaths: false,
            modifiedDate: "2000-01-01T10:00:00.000Z",
            ownedByMe: true,
        },
        {
            id: "ke8436",
            path: "/kf84/asf",
            multiplePaths: false,
            modifiedDate: "2000-01-01T10:00:00.000Z",
            ownedByMe: true,
        },
    ],
    errors: [
        "Found multiple folders with name 'name1'",
    ],
    date: "21/4/2015 10:11",
}
*/

/**
 * @typedef {Object} NameAndIdValidationInfo  //ttt0: Rename "Folder"
 *
 * @property {RangeInfo} rangeInfo
 * @property {InputInfo[]} inputNameInfos
 * @property {InputInfo[]} inputIdInfos
 * @property {Map<string, IdInfo>} idInfosMap
 * @property {boolean} hasErrors
*/


// noinspection JSUnusedGlobalSymbols
/**
 * Run automatically when the corresponding spreadsheet is opened
 */
function onOpen() {
    // https://developers.google.com/apps-script/guides/menus
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Modification times')
        .addItem(`Set time for specified folders`, 'menuSetTimesFolders')
        //.addItem(`List content of specified  folders`, 'menuListFolders')
        //.addItem(`Change time for file ${SOURCE_LIST}`, 'menuSetTimesFile') //ttt0 implement
        .addToUi();

    setupSheets();
}
//ttt3 Review other triggers: https://developers.google.com/apps-script/guides/triggers


const LOG_START = 'Log (don\'t change this cell)';

const NAMES_BG = '#efe';
const IDS_BG = '#eef';
const LOG_BG = '#ffc';
const ERROR_BG = '#fbb';


class DriveObjectProcessor {

    /**
     * @param {string} sheetName
     * @param {string} nameLabelStart
     * @param {string} idLabelStart
     * @param {boolean} expectFolders whether we want folders or files; (enum would be nicer, but JS doesn't have them)
     */
    constructor(sheetName, nameLabelStart, idLabelStart, expectFolders) {
        this.sheetName = sheetName;
        this.nameLabelStart = nameLabelStart;
        this.idLabelStart = idLabelStart;
        this.logLabelStart = LOG_START;
        this.rangeNotFoundErr = 'Couldn\'t find section delimiters. If a manual fix is not obvious,'
            + ` delete or rename the "${this.sheetName}" sheet and then reopen the spreadsheet`; //ttt1 perhaps
        // make most members private, but not sure it's woth it, and V8 doesn't seem to support it

        this.expectFolders = expectFolders;
        this.objectLabel = expectFolders ? 'Folder' : 'File';
        this.objectLabelLc = expectFolders ? 'folder' : 'file';
        this.reverseObjectLabelLc = expectFolders ? 'file' : 'folder';
    }

    /**
     * @returns {GoogleAppsScript.Spreadsheet.Sheet}
     */
    getSheet() {
        //return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
        for (let i = 0; i < 2; i += 1) {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetName);
            if (sheet) {
                return sheet;
            }
            setupSheets(); //ttt1: Review this
        }
        return null; // this should never be reached
    }

    /**
     * @param {SpreadsheetApp.Sheet} sheet
     */
    setupSheet(sheet) {
        this.addLabelsIfEmptySheet(sheet);
        this.applyColorToSheet(sheet);
    }

    /**
     * If the sheet is empty, add the section start labels
     *
     * @param {SpreadsheetApp.Sheet} sheet
     */
    addLabelsIfEmptySheet(sheet) {

        if (sheet.getLastRow() !== 0) {
            return;
        }

        const BOLD_FONT = 'bold';
        let crtLine = 1;

        function addLabel(label, increment) {
            const range = sheet.getRange(crtLine, 1);
            range.setValue(label);
            range.setFontWeight(BOLD_FONT);
            //range.protect().removeEditor('XYZ@gmail.com'); // It's what we want, but doesn't work
            range.protect().setWarningOnly(true); // We really want nobody being able to edit, but it can't be done, so we use warnings
            crtLine += increment;
        }

        addLabel(this.nameLabelStart, DEFAULT_SOURCE_HEIGHT);
        addLabel(this.idLabelStart, DEFAULT_SOURCE_HEIGHT);
        addLabel(this.logLabelStart, DEFAULT_SOURCE_HEIGHT);

        // Use some educated guesses for the column widths ...
        sheet.autoResizeColumn(1);
        const w = sheet.getColumnWidth(1);
        sheet.setColumnWidth(1, w * 1.2);
        sheet.setColumnWidth(2, w * 2.4);
    }

    /**
     * Sets the background for the first column, so names, IDs, and logs each have their own color.
     * Throws if (some of) the section starts are not found or are not in their proper order.
     *
     * @param {SpreadsheetApp.Sheet} sheet
     * @returns {boolean} true iff the range is valid
     */
    applyColorToSheet(sheet) {
        const rangeInfo = this.getRangeInfo(sheet);
        if (!rangeInfo) {
            return false;
        }
        sheet.getRange(rangeInfo.namesBegin, 1, rangeInfo.namesEnd - rangeInfo.namesBegin, 2).setBackground(NAMES_BG);
        sheet.getRange(rangeInfo.idsBegin, 1, rangeInfo.idsEnd - rangeInfo.idsBegin, 2).setBackground(IDS_BG);
        sheet.getRange(rangeInfo.logsBegin, 1, rangeInfo.logsEnd - rangeInfo.logsBegin, 2).setBackground(LOG_BG);
        return true;
    }


    /**
     * Reads the names and IDs and generates errors, if necessary
     * @param {SpreadsheetApp.Sheet} sheet
     *
     * @returns {NameAndIdValidationInfo|null} A null is returned iff it couldn't find the ranges
     */
    validateNamesAndIds(sheet) {
        sheet.activate();
        const rangeInfo = this.getRangeInfo(sheet);
        if (!rangeInfo) {
            return null;
        }

        /** @type {Map<string, IdInfo>} */
        const idInfosMap = new Map();
        const names = DriveObjectProcessor.getColumnData(sheet, 1, rangeInfo.namesBegin + 1, rangeInfo.namesEnd);
        const inputNameInfos = this.validateNames(sheet, names, idInfosMap);
        const inputIds = DriveObjectProcessor.getColumnData(sheet, 1, rangeInfo.idsBegin + 1, rangeInfo.idsEnd);
        const inputIdInfos = this.validateIds(sheet, inputIds, idInfosMap);
        const nameErrors = inputNameInfos.filter((val) => val.errors.length);
        const idErrors = inputIdInfos.filter((val) => val.errors.length);
        const hasErrors = nameErrors.length > 0 || idErrors.length > 0;

        return {
            rangeInfo,
            inputNameInfos,
            inputIdInfos,
            hasErrors,
            idInfosMap,
        };
    }


    /**
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {number} column
     * @param {number} rowStart
     * @param {number} rowEnd exclusive
     * @returns {string[]} data in a column, between 2 rows, as an array.
     */
    static getColumnData(sheet, column, rowStart, rowEnd) {
        const rows = sheet.getRange(rowStart, column, rowEnd - rowStart).getValues();
        /** @type string[] */
        const res = [];
        for (let i = 0; i < rows.length; i += 1) {
            res.push(rows[i][0]);
        }
        return res;
    }


    /**
     * Checks the first column for section delimiters and data and returns the ranges for names, IDs, and logs.
     * To be used for coloring or log clearing.
     * If the delimiters are not found in the expected order, returns null.
     *
     * The ends are exclusive, and, for now, coincide with the beginning of the next section. This might change, though.
     *
     * @param {SpreadsheetApp.Sheet} sheet
     * @returns {(RangeInfo|null)} null if the range is invalid
     */
    getRangeInfo(sheet) {
        const lastRow = sheet.getLastRow();
        const rows = sheet.getRange(1, 1, lastRow).getValues();
        const expected = [this.nameLabelStart, this.idLabelStart, this.logLabelStart];
        let expectedIndex = 0;
        const found = [];
        for (let i = 0; i < rows.length; i += 1) {
            if (rows[i][0] === expected[expectedIndex]) {
                found.push(i + 1); // spreadsheet indexes start at 1
                expectedIndex++;
                if (expectedIndex === expected.length) {
                    break;
                }
            }
        }
        if (expectedIndex !== expected.length) {
            this.showMessage(sheet, this.rangeNotFoundErr);
            return null;
        }
        return {
            namesBegin: found[0],
            namesEnd: found[1],
            idsBegin: found[1],
            idsEnd: found[2],
            logsBegin: found[2],
            logsEnd: lastRow + 1, // "+1" to account for exclusivity of the end
        };
    }


    /**
     * Returns an array the size of up to the "names" section without the title with the type InputInfo.
     * For empty names, the corresponding entry is undefined (or missing, at the end of the array).
     * For non-empty names, the corresponding entry is ideally one entry, when we have one path; otherwise it's an error.
     * Not clear how to deal with multipaths. We should probably cache them, and maybe generate a warning.
     *
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
     * @param {string[]} names - array of strings
     * @param {Map<string, IdInfo>} idInfosMap - map with ID as key and a IdInfo as the value; an InputInfo with
     *      multiple entries will also have multiple entries in the map; the map is used here to figure out
     *      which folders or files  are already defined and then as the input for the actual processing
     * @returns {InputInfo[]}
     */
    validateNames(sheet, names, idInfosMap) {
        /** @type InputInfo[] */
        const res = [];
        for (let i = 0; i < names.length; i += 1) {
            const name = names[i];
            if (!name) {
                continue;
            }
            /** @type IdInfo[] */
            const idInfos = [];
            /** @type string[] */
            const errors = [];

            const query = `title = '${name}' and trashed = false`;
            let pageToken = null;

            do {
                try {
                    const items = Drive.Files.list({
                        q: query,
                        maxResults: 100,
                        pageToken,
                    });

                    if (!items.items || items.items.length === 0) {
                        break;
                    }
                    for (let i = 0; i < items.items.length; i++) {
                        /** @type IdInfo */
                        let idInfo;
                        const item = items.items[i];
                        if ((this.expectFolders && item.mimeType === FOLDER_MIME) || (!this.expectFolders && item.mimeType !== SHORTCUT_MIME)) {
                            idInfo = getIdInfo(item.id);
                            idInfos.push(idInfo);
                            const existing = idInfosMap.get(idInfo.id);
                            if (existing) {
                                errors.push(`${this.objectLabel} with the ID ${idInfo.id} already added`);
                            } else {
                                idInfosMap.set(idInfo.id, idInfo);
                            }
                        } else if (item.mimeType !== SHORTCUT_MIME) {
                            // This is not an error, but we'd still like to log something
                            logS(sheet, `Expected a ${this.objectLabelLc} but got a ${this.reverseObjectLabelLc} for ${item.title} [${item.id}]`);
                        } else {
                            logS(sheet, `Ignoring shortcut ${item.title} [${item.id}]`);
                        }
                    }
                    pageToken = items.nextPageToken;
                } catch (err) {
                    errors.push(`Failed to process ${this.objectLabelLc} '${name}'. ${err}`);
                    //ttt2 improve
                }
            } while (pageToken);

            const cnt = idInfos.length;
            if (!cnt) {
                errors.push(`No ${this.objectLabelLc} found with the name '${name}'`);
            } else if (cnt > 1) {
                errors.push(`Found ${cnt} ${this.objectLabelLc}s with the name '${name}'`);
            }
            res[i] = {
                idInfos,
                errors,
            };
        }
        return res;
    }


    /**
     * Returns an array the size of up to the "names" section without the title with the type InputInfo
     * For empty IDs, the corresponding entry is unedfined (or missing, at the end of the array).
     *
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
     * @param {string[]} ids
     * @param {Map<string, IdInfo>} idInfosMap - map with ID as key and an IdInfo as the value; an InputInfo with
     *      multiple entries will also have multiple entries in the map; the map is used here to figure out
     *      which folders or files are already defined and then as the input for the actual processing
     * @returns {InputInfo[]}
     */
    validateIds(sheet, ids, idInfosMap) {
        /** @type InputInfo[] */
        const res = [];
        for (let i = 0; i < ids.length; i += 1) {
            const id = ids[i];
            if (!id) {
                continue;
            }
            /** @type string[] */
            const errors = [];
            /** @type IdInfo */
            let idInfo;
            try {
                const item = Drive.Files.get(id);

                if ((this.expectFolders && item.mimeType === FOLDER_MIME) || (!this.expectFolders && item.mimeType !== SHORTCUT_MIME)) {  //ttt1: duplicate code with names
                    idInfo = getIdInfo(item.id);
                    const existing = idInfosMap.get(idInfo.id);
                    if (existing) {
                        errors.push(`${this.objectLabel} with the ID ${idInfo.id} already added`);
                    } else {
                        idInfosMap.set(idInfo.id, idInfo);
                    }
                } else if (item.mimeType !== SHORTCUT_MIME) {
                    // This is not an error, but we'd still like to log something
                    logS(sheet, `Expected a ${this.objectLabelLc} but got a ${this.reverseObjectLabelLc} for ${item.title} [${item.id}]`);
                } else {
                    logS(sheet, `Ignoring shortcut ${item.title} [${item.id}]`);
                }
            } catch (err) {
                errors.push(`${this.objectLabel} with ID ${id} not found`);
            }
            /** @type IdInfo[] */
            const idInfos = [];
            if (idInfo) {
                idInfos.push(idInfo)
            }
            res[i] = {
                idInfos,
                errors,
            };
        }
        return res;
    }


    /**
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {RangeInfo} rangeInfo
     * @param {InputInfo[]} inputNameInfos
     * @param {InputInfo[]} inputIdInfos
     *
     * @returns {boolean} true iff the range is valid
     */
    updateUiAfterValidation(sheet, rangeInfo, inputNameInfos, inputIdInfos) {
        this.clearErrors(sheet);
        if (!this.applyColorToSheet(sheet)) {
            return false;
        }

        for (let i = 0; i < inputNameInfos.length; i += 1) {
            /** @type InputInfo */
            const inputNameInfo = inputNameInfos[i];
            if (!inputNameInfo) {
                continue;
            }
            /** @type string[] */
            let lines = [];
            for (const idInfo of inputNameInfo.idInfos) {
                lines.push(`${idInfo.path}${idInfo.multiplePaths ? ' (and others)' : ''} [${idInfo.id}]`);
            }
            const cellRange = sheet.getRange(rangeInfo.namesBegin + 1 + i, 2);
            if (inputNameInfo.errors.length) {
                lines.push(...inputNameInfo.errors);
                cellRange.setBackground(ERROR_BG);
            }
            cellRange.setValue(lines.join('\n'));
        }

        for (let i = 0; i < inputIdInfos.length; i += 1) {
            /** @type InputInfo */
            const inputIdInfo = inputIdInfos[i];
            if (!inputIdInfo) {
                continue;
            }
            /** @type string[] */
            let lines = [];
            const idInfo = inputIdInfo.idInfos[0]; //!!! It's OK if the array is empty, as [0] will return
            // undefined. (For IDs, we have an entry if the ID is correct and none if there's an error.)
            if (idInfo) {
                lines.push(`${idInfo.path}${idInfo.multiplePaths ? ' (and others)' : ''}  [${idInfo.id}]`); //ttt1 duplicate code "(and others)"
            }
            const cellRange = sheet.getRange(rangeInfo.idsBegin + 1 + i, 2);
            if (inputIdInfo.errors.length) {
                lines.push(...inputIdInfo.errors);
                cellRange.setBackground(ERROR_BG);
            }
            cellRange.setValue(lines.join('\n'));
        }

        //sheet.autoResizeColumns(1, 2); //!!! Not right, as it include the logs
        return true;
    }


    /**
     * Erases the second column, which is supposed to be used for errors.
     *
     * @param {SpreadsheetApp.Sheet} sheet
     */
    clearErrors(sheet) {
        sheet.getRange(1, 2, sheet.getLastRow()).clearContent();
    }


    /**
     * Show a popup message. If that fails, logs to the log section in
     * the corresponding sheet. If that also fails, logs to the console.
     *
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {string} message
     */
    showMessage(sheet, message) {
        try {
            SpreadsheetApp.getUi().alert(message);
        } catch {
            logS(sheet, message);
        }
    }
}


const FOLDER_NAME_START = 'Folder names, one per cell (don\'t change this cell)';
const FOLDER_ID_START = 'Folder IDs, one per cell (don\'t change this cell)';

class DriveFolderProcessor extends DriveObjectProcessor {
    constructor() {
        super(FOLDERS_SHEET_NAME,
            FOLDER_NAME_START,
            FOLDER_ID_START,
            true);
    }

    /**
     * @param {boolean} showConfirmation
     * @returns {boolean} true iff all was OK (the range is valid and the user confirmed it's OK to proceed, then we made the updates)
     */
    setTimes(showConfirmation) {
        const sheet = this.getSheet();

        //@param {SpreadsheetApp.Sheet} sheet
        //@param {IdInfo[]} idInfos
        let idInfos = this.getProcessedInputData(sheet);
        if (!idInfos) {
            return false;
        }

        // if (!idInfos.length) {
        //     this.showMessage('No files were specified, but at least one is needed');
        //     return false;
        // }

        if (showConfirmation && !showConfirmYesNoBox(`Really set the dates for ${idInfos.length ? 'the specified' : 'all the'} folders?`)) {
            return false;
        }

        if (!idInfos.length) {
            const rootFolder = DriveApp.getRootFolder(); // ttt2 This probably needs to change for shared drives
            idInfos.push({
                id: rootFolder.getId(),
                path: '',
                multiplePaths: false,
                modifiedDate: SMALLEST_TIME, // Not right, but it will be ignored
                ownedByMe: true, // doesn't really matter
            });
            logS(sheet, 'Processing all the folders, as no folder names or IDs were specified');
        }

        logS(sheet, '------------------ Starting update ------------------');
        const timeSetter = new TimeSetter();
        for (const idInfo of idInfos) {
            timeSetter.processFolder(sheet, idInfo);
        }
        logS(sheet, '------------------ Update finished ------------------');
        return true;
    }

    /**
     * Gathers and validates the data. Mainly makes sure that files / folders exist and there are no duplicates.
     *
     * @param {SpreadsheetApp.Sheet} sheet
     * @returns {(IdInfo[]|null)} an array (which might be empty) with an IdInfo for each user input, if all is OK; null, if there are errors
     */
    getProcessedInputData(sheet) {

        const vld = this.validateNamesAndIds(sheet);
        if (!vld) {
            return null;
        }

        //!!! We want to update the UI here, and not after the confirmation, as it's important to see what the IDs point
        // to before confirmation. //ttt1 However, the UI doesn't change until after the confirmation.
        if (!this.updateUiAfterValidation(sheet, vld.rangeInfo, vld.inputNameInfos, vld.inputIdInfos)) {
            // Range couldn't be computed, even though a few lines above it could. Perhaps the user deleted a label.
            return null;
        }

        if (vld.hasErrors) {
            this.showMessage(sheet, 'Found some errors, which need to be resolved before proceeding with setting the times');
            return null;
        }

        /** @type {IdInfo[]} */
        const idInfosArr = Array.from(vld.idInfosMap.values());
        return idInfosArr;
    }

}
// const FILE_NAME_START = 'File names, one per cell (don\'t change this cell)';
// const FILE_ID_START = 'File IDs, one per cell (don\'t change this cell)';


const driveFolderProcessor = new DriveFolderProcessor();


/**
 * @returns {SpreadsheetApp.Sheet}
 */
function getFilesSheet() {
    //return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FILES_SHEET_NAME);
}


/**
 * Called when opening the document, to see if the sheets exist and have the right content and tell the user if not.
 */
function setupSheets() {
    if (!driveFolderProcessor.getSheet()) {
        // If there's a single sheet, it's probably a new install, so we can rename it.
        const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
        if (sheets.length === 1) {
            if (sheets[0].getLastRow() === 0) {
                // There's no data. Just rename // ttt2 There might be formatting
                sheets[0].setName(FOLDERS_SHEET_NAME);
            }
        }
    }

    if (!driveFolderProcessor.getSheet()) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        sheet.setName(FOLDERS_SHEET_NAME);
    }
    /*if (!getFilesSheet()) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        sheet.setName(FILES_SHEET_NAME);
    }*/

    const folderSheet = driveFolderProcessor.getSheet();
    folderSheet.activate(); //ttt3 This doesn't work when running the script in the editor, but
    // works when starting from the Sheet menu. At least it doesn't crash

    driveFolderProcessor.setupSheet(folderSheet);
}


/**
 *
 * @returns {boolean} true iff all was OK (the range is valid and the user confirmed it's OK to proceed, then we made the updates)
 */
function menuSetTimesFolders() {
    return driveFolderProcessor.setTimes(true);
}

// noinspection JSUnusedGlobalSymbols
/**
 * For debugging, to be called from the Google Apps Script web IDE, where a UI is not accessible.
 *
 * @returns {boolean}
 */
function setTimesFoldersDebug() {
    return driveFolderProcessor.setTimes(false);
}


const FOLDER_MIME = 'application/vnd.google-apps.folder';
const SHORTCUT_MIME = 'application/vnd.google-apps.shortcut';

class TimeSetter {

    constructor() {
        /** @type {Map<string, string>} */
        this.processed = new Map();
    }

    /**
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {IdInfo} idInfo
     * @returns {string} when the folder was last modified, in the format '2000-01-01T10:00:00.000Z'
     */
    processFolder(sheet, idInfo) {
        const existing = this.processed.get(idInfo.id);
        if (existing) {
            return existing;
        }

        const query = `"${idInfo.id}" in parents and trashed = false`;
        let pageToken = null;
        let res = SMALLEST_TIME;

        do {
            try {
                const items = Drive.Files.list({
                    q: query,
                    maxResults: 100,
                    pageToken,
                });

                if (!items.items || items.items.length === 0) {
                    //console.log('No folders found.');
                    break;
                }
                for (let i = 0; i < items.items.length; i++) {
                    const item = items.items[i];
                    let newTime = res;
                    if (item.mimeType === FOLDER_MIME) {
                        /** @type {IdInfo}} */
                        const childIdInfo = {
                            id: item.id,
                            modifiedDate: item.modifiedDate,
                            multiplePaths: false,  // not correct, but it doesn't matter; it's just to have something
                            path: `${idInfo.path}/${item.title}`,
                            ownedByMe: getOwnedByMe(item),   //ttt1: See why there's no warning here, as getOwnedByMe()
                            // may return undefined, while the field is just boolean
                        };
                        newTime = this.processFolder(sheet, childIdInfo);
                    } else {
                        if (item.mimeType !== SHORTCUT_MIME) {
                            newTime = item.modifiedDate;
                        } else {
                            logS(sheet, `Ignoring date of shortcut ${idInfo.path}/${item.title}`);
                        }
                    }
                    if (newTime > res) {
                        res = newTime;
                    }
                }
                pageToken = items.nextPageToken;
            } catch (err) {
                const msg = `Failed to process folder '${idInfo.path}' [${idInfo.id}]. ${err}`; //ttt2 We might want
                // ${err.message}, but that might not always exist, and then we get "undefined". This would work, but not
                // sure what value it provides: ${err.message || err}. If the exception being thrown inherits Error (as
                // all exceptions are supposed to), then err.message exists. But some code might throw arbitrary expressions
                logS(sheet, msg);
                //ttt2 improve
            }
        } while (pageToken);

        if (res !== idInfo.modifiedDate) {
            if (idInfo.path) {
                // We are not dealing here with the root, which cannot be updated (and for which you couldn't easily see the date anyway)
                try {
                    if (idInfo.ownedByMe) {
                        logS(sheet, `Setting time to ${res} for ${idInfo.path}. It was ${idInfo.modifiedDate}`);
                        updateModifiedTime(idInfo.id, res);
                    } else {
                        logS(sheet, `Not updating ${idInfo.path}, which has a different owner`);
                    }
                } catch (err) {
                    const msg = `Failed to update time for folder '${idInfo.path}' [${idInfo.id}]. ${err}`;
                    logS(sheet, msg);
                    //ttt2 improve
                }
            }
        } else {
            logS(sheet, `Time ${res} is already correct for ${idInfo.path}`);
        }
        this.processed.set(idInfo.id, res);
        return res;
    }
}


/**
 * @param {string} id
 * @param {string} itemTime format: '2020-05-05T10:00:00.000Z'
 */
function updateModifiedTime(id, itemTime) {
    const body = {modifiedDate: itemTime}; // type File: https://developers.google.com/drive/api/reference/rest/v2/files#File
    const blob = null;
    const optionalArgs = {setModifiedDate: true}; // https://developers.google.com/drive/api/reference/rest/v2/files/update#query-parameters
    Drive.Files.update(body, id, blob, optionalArgs); // This fails silently for non-owners, but we check for that
    // before calling updateModifiedTime()
}



/**
 * @param {string} id ID of a file or a folder
 * @returns {IdInfo}
 */
function getIdInfo(id) {
    const item = Drive.Files.get(id)
    let parents = item.parents;
    let parentCnt = parents.length;
    if (parentCnt === 0) {
        // It's a (usually "the") root
        return {
            id,
            path: '',
            multiplePaths: false,
            modifiedDate: item.modifiedDate,
            ownedByMe: item.ownedByMe, // doesn't matter
        };
    }
    let parent = parents[0];
    let parentInfo = getIdInfo(parent.id);
    return {
        id,
        path: `${parentInfo.path}/${item.title}`,
        multiplePaths: parentInfo.multiplePaths || (parentCnt > 1),
        modifiedDate: item.modifiedDate,
        ownedByMe: item.ownedByMe,
    };
}


const USER_EMAIL = Session.getActiveUser().getEmail();

/**
 * For whatever reason the flag ownedByMe stopped working on 2023.11.10. After reverting the code to the one when
 * the feature was introduced and really tested, and seeing that it was the same, the conclusion is that the issue
 * wasn't introduced by some bug, but comes from Drive. This is the corresponding workaround.
 *
 * @param {GoogleAppsScript.Drive.Schema.File} file
 * @returns {(boolean|undefined)}
 */
function getOwnedByMe(file) {
    if (file.ownedByMe !== undefined) {
        return file.ownedByMe;
    }
    if (file.owners) {
        for (const owner of file.owners) {
            if (owner.emailAddress === USER_EMAIL) {
                return true;
            }
        }
        return false;
    }
    return undefined;
}

/**
 * Logs a message in the given sheet. If that fails, logs to the console.
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {string} message
 */
function logS(sheet, message) {
    try {
        const row = sheet.getLastRow() + 1;
        const range = sheet.getRange(row, 1);
        range.setValue(`${formatDate(new Date())} ${message}`);
        sheet.getRange(row, 1, 1, 2).setBackground(LOG_BG);
    } catch (err) {
        console.log(`Failed to log in UI "${message}". "${err}"`);
    }
}


/**
 * @param {string} message
 * @returns {boolean} whether the user chose "Yes"
 */
function showConfirmYesNoBox(message) {
    //let choice = Browser.msgBox(message, Browser.Buttons.YES_NO);
    //return choice === 'yes';
    const ui = SpreadsheetApp.getUi();
    const choice = ui.alert(message, ui.ButtonSet.YES_NO);
    return choice === ui.Button.YES;
}

/**
 * Formats a date using HH:mm:ss.SSS (SSS is for milliseconds)
 * @param {Date} date
 * @returns {string}
 */
function formatDate(date) {
    const main = date.toLocaleDateString('ro-RO', {
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
    }).substring(4);
    const millis = String(date.getMilliseconds()).padStart(3, '0');
    return `${main}.${millis}`;
}

//ttt1 Perhaps have a "dry-run"

//ttt0: Rename sheet, go to menu. Starts creating and then there's an error. It should exit immediately, or ask for
// confirmation to create, and, if so, work OK
