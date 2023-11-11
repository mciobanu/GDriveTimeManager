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

const PLAIN_TEXT_FMT = '@STRING@';
const LIST_DATETIME_FMT = 'yyyy-MM-dd hh:mm';

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

/**
 * @typedef {Function} SimpleLogger
 * @param {string} message
 */

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
 * @typedef {Object} ValidationInfo
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
        .addItem(`Validate folder data`, 'menuValidateFolders')
        .addItem(`Set time for specified folders`, 'menuSetTimesFolders')
        .addItem(`List content of specified  folders`, 'menuListFolders')
        .addSeparator()
        .addItem(`Validate file data`, 'menuValidateFiles')
        .addItem(`Set time for specified files`, 'menuSetTimesFiles')
        .addToUi();

    setupSheets();
}
//ttt3 Review other triggers: https://developers.google.com/apps-script/guides/triggers


const LOG_START = 'Log (don\'t change this cell)';

const NAMES_BG = '#e0ffe0';
const NAMES_OUT_BG = '#f0fff0';
const IDS_BG = '#e0e0ff';
const IDS_OUT_BG = '#f0f0ff';
const LOG_BG = '#ffc';
const ERROR_BG = '#fbb';


class DriveObjectProcessor {

    /**
     * @param {string} sheetName
     * @param {string} nameLabelStart
     * @param {string} idLabelStart
     * @param {string[]} columnLabels
     * @param {boolean} expectFolders whether we want folders or files; (enum would be nicer, but JS doesn't have them)  //ttt0: find better name
     */
    constructor(sheetName, nameLabelStart, idLabelStart, columnLabels, expectFolders) {
        this.sheetName = sheetName;
        this.nameLabelStart = nameLabelStart;
        this.idLabelStart = idLabelStart;
        this.columnLabels = columnLabels;
        this.logLabelStart = LOG_START;
        this.rangeNotFoundErr = 'Couldn\'t find section delimiters. If a manual fix is not obvious,'
            + ` delete or rename the "${this.sheetName}" sheet and then reopen the spreadsheet`; //ttt1 perhaps
        // make most members private, but not sure it's worth it, and V8 doesn't seem to support it

        this.expectFolders = expectFolders;
        this.objectLabel = expectFolders ? 'Folder' : 'File';  //ttt1: Review idea of computing these here vs. passing
        // them as params. Adds coupling but cuts param count.
        this.objectLabelLc = expectFolders ? 'folder' : 'file';
        this.reverseObjectLabelLc = expectFolders ? 'file' : 'folder';
        this.outputColumn = expectFolders ? 2 : 3;
        this.dateColumn = expectFolders ? 0 : 2;
    }

    /**
     * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet associated with the object. It creates it if it doesn't
     * exist
     */
    getSheet() {
        //return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetName);
        if (!sheet) {
            sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
            sheet.setName(this.sheetName);
            this.setupSheet();
        }
        return sheet;
    }

    sheetExists() {
        return !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetName);
    }

    /**
     * Adds labels (only if sheet is empty) and sets background colors
     */
    setupSheet() {
        /** @type SpreadsheetApp.Sheet */
        const sheet = this.getSheet();
        this.addLabelsIfEmptySheet(sheet);
        if (!this.applyColorToSheet(sheet)) {
            throw new Error('applyColorToSheet() failed');  //ttt1: see how to improve; we want getSheet() to not return non-null until all is set up
        }
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

        /**
         *
         * @param {string[]} labels
         * @param {number} increment
         */
        function addLabel(labels, increment) {
            for (let i = 0; i < labels.length; i += 1) {
                const range = sheet.getRange(crtLine, i + 1);
                const label = labels[i];
                range.setValue(label);
                range.setFontWeight(BOLD_FONT);
                //range.protect().removeEditor('XYZ@gmail.com'); // It's what we want, but doesn't work
                range.protect().setWarningOnly(true); // We really want nobody being able to edit, but it can't be done, so we use warnings
            }
            crtLine += increment;
        }

        addLabel([this.nameLabelStart, ...this.columnLabels], DEFAULT_SOURCE_HEIGHT);
        addLabel([this.idLabelStart, ...this.columnLabels], DEFAULT_SOURCE_HEIGHT);
        addLabel([this.logLabelStart], DEFAULT_SOURCE_HEIGHT);

        // Use some educated guesses for the column widths ...
        sheet.autoResizeColumn(1);
        const w = sheet.getColumnWidth(1);
        sheet.setColumnWidth(1, w * 1.2);
        if (this.dateColumn) {
            sheet.setColumnWidth(this.dateColumn, w * 0.7);
        }
        sheet.setColumnWidth(this.outputColumn, w * 2.4);
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
        sheet.getRange(rangeInfo.namesBegin, 1, rangeInfo.namesEnd - rangeInfo.namesBegin, this.outputColumn - 1)
            .setBackground(NAMES_BG);
        sheet.getRange(rangeInfo.namesBegin, this.outputColumn, rangeInfo.namesEnd - rangeInfo.namesBegin, 1)
            .setBackground(NAMES_OUT_BG)
            //.protect().setWarningOnly(true)    //!!! It's nice to write-protect, but then there are warnings when inserting. ttt2: See if possible to improve
        ;
        sheet.getRange(rangeInfo.idsBegin, 1, rangeInfo.idsEnd - rangeInfo.idsBegin, this.outputColumn - 1)
            .setBackground(IDS_BG);
        sheet.getRange(rangeInfo.idsBegin, this.outputColumn, rangeInfo.idsEnd - rangeInfo.idsBegin, 1)
            .setBackground(IDS_OUT_BG)
            //.protect().setWarningOnly(true)
        ;
        sheet.getRange(rangeInfo.logsBegin, 1, rangeInfo.logsEnd - rangeInfo.logsBegin, this.outputColumn)
            .setBackground(LOG_BG);

        if (this.dateColumn) {
            sheet.getRange(rangeInfo.namesBegin, this.dateColumn, rangeInfo.idsEnd - rangeInfo.namesBegin, this.dateColumn)
                .setNumberFormat(PLAIN_TEXT_FMT);
        }
        return true;
    }


    /**
     * Reads the names and IDs (and dates, for files) and generates errors, if necessary
     * @param {SpreadsheetApp.Sheet} sheet
     *
     * @returns {ValidationInfo|null} All the data necessary to set the times. A null is returned iff it couldn't find the ranges
     */
    getValidationInfo(sheet) {
        sheet.activate(); //ttt1: This causes cell A1 to become selected. Perhaps get what's current first, then activate, then set current
        const rangeInfo = this.getRangeInfo(sheet);
        if (!rangeInfo) {
            return null;
        }

        /** @type {Map<string, IdInfo>} */
        const idInfosMap = new Map();
        const inputNames = DriveObjectProcessor.getColumnData(sheet, 1, rangeInfo.namesBegin + 1, rangeInfo.namesEnd);
        const inputNameInfos = this.validateNames(sheet, inputNames, idInfosMap);
        const inputIds = DriveObjectProcessor.getColumnData(sheet, 1, rangeInfo.idsBegin + 1, rangeInfo.idsEnd);
        const inputIdInfos = this.validateIds(sheet, inputIds, idInfosMap);
        if (this.dateColumn) {
            const inputNameDates = DriveObjectProcessor.getColumnData(sheet, this.dateColumn, rangeInfo.namesBegin + 1, rangeInfo.namesEnd);
            const inputIdDates = DriveObjectProcessor.getColumnData(sheet, this.dateColumn, rangeInfo.idsBegin + 1, rangeInfo.idsEnd);
            this.validateTimes(sheet, inputNameInfos, inputNameDates);
            this.validateTimes(sheet, inputIdInfos, inputIdDates);
        }
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
     * Changes inputInfos by setting the date field. When there's a mismatch between entries, adds an error, and
     * also adds an element to inputInfos if just the date is present.
     *
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {InputInfo[]} inputInfos
     * @param {string[]} inputDates
     */
    validateTimes(sheet, inputInfos, inputDates) {
        for (let i = 0; i < inputDates.length; i++) {
            /** @type {InputInfo} */
            let inputInfo = inputInfos[i];
            const dateStr = inputDates[i];
            if (dateStr) {
                if (!inputInfo) {
                    inputInfo = {
                        idInfos: [],
                        errors: [`Missing ${this.objectLabelLc} field`],
                    }
                    inputInfos[i] = inputInfo;
                }
                try {
                    const date = new Date(dateStr);
                    inputInfo.date = date.toISOString();
                } catch {
                    inputInfo.errors.push('Error parsing date field');
                }
                continue;
            }
            if (!inputInfo) {
                continue;
            }
            inputInfo.errors.push('Missing time field');
        }
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
                        if ((this.expectFolders && item.mimeType === FOLDER_MIME) || (!this.expectFolders && item.mimeType !== SHORTCUT_MIME && item.mimeType !== FOLDER_MIME)) {
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
                            this.log(sheet, `Expected a ${this.objectLabelLc} but got a ${this.reverseObjectLabelLc} for ${item.title} [${item.id}]`);
                        } else {
                            this.log(sheet, `Ignoring shortcut ${item.title} [${item.id}]`);
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

                if ((this.expectFolders && item.mimeType === FOLDER_MIME) || (!this.expectFolders && item.mimeType !== SHORTCUT_MIME && item.mimeType !== FOLDER_MIME)) {  //ttt1: duplicate code with names
                    idInfo = getIdInfo(item.id);
                    const existing = idInfosMap.get(idInfo.id);
                    if (existing) {
                        errors.push(`${this.objectLabel} with the ID ${idInfo.id} already added`);
                    } else {
                        idInfosMap.set(idInfo.id, idInfo);
                    }
                } else if (item.mimeType !== SHORTCUT_MIME) {
                    // This is not an error, but we'd still like to log something
                    this.log(sheet, `Expected a ${this.objectLabelLc} but got a ${this.reverseObjectLabelLc} for ${item.title} [${item.id}]`);
                } else {
                    this.log(sheet, `Ignoring shortcut ${item.title} [${item.id}]`);
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

        this.updateUiHlp(sheet, inputNameInfos, rangeInfo.namesBegin);
        this.updateUiHlp(sheet, inputIdInfos, rangeInfo.idsBegin);

        //sheet.autoResizeColumns(1, 2); //!!! Not right, as it include the logs
        return true;
    }

    /**
     *
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {InputInfo[]} inputInfos
     * @param firstRow
     */
    updateUiHlp(sheet, inputInfos, firstRow) {
        for (let i = 0; i < inputInfos.length; i += 1) {
            /** @type InputInfo */
            const inputInfo = inputInfos[i];
            if (!inputInfo) {
                continue;
            }
            /** @type string[] */
            let lines = [];
            for (const idInfo of inputInfo.idInfos) {
                lines.push(`${idInfo.path}${idInfo.multiplePaths ? ' (and others)' : ''} [${idInfo.id}]`);
            }
            if (inputInfo.date) {
                lines.push(isoDateToShort(inputInfo.date));
            }
            const cellRange = sheet.getRange(firstRow + 1 + i, this.outputColumn);
            if (inputInfo.errors.length) {
                lines.push(...inputInfo.errors);
                cellRange.setBackground(ERROR_BG);
            }
            cellRange.setValue(lines.join('\n'));
        }
    }


    /**
     * Updates the output column to reflect folders / IDs / dates / errors
     */
    validateInput() {
        const sheet = this.getSheet();
        this.getProcessedInputData(sheet);
    }

    /**
     * Gathers and validates the data, updating the sheet in the process. Mainly makes sure that files / folders
     * exist and there are no duplicates. Also, for files, it checks that the times match and can be parsed. //ttt1 coupling
     *
     * @param {SpreadsheetApp.Sheet} sheet
     * @returns {(InputInfo[]|null)} an array (which might be empty) with an IdInfo for each user input, if all is OK; null, if there are errors
     */
    getProcessedInputData(sheet) {

        const vld = this.getValidationInfo(sheet);
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

        /** @type {InputInfo[]} */
        const tmp = [...vld.inputNameInfos];
        tmp.push(...vld.inputIdInfos);
        const res = tmp.filter((x) => !!x);
        return res;
    }


    /**
     * Erases the second column, which is supposed to be used for errors.
     *
     * @param {SpreadsheetApp.Sheet} sheet
     */
    clearErrors(sheet) {
        const rangeInfo = this.getRangeInfo(sheet);
        if (!rangeInfo) {
            return;
        }
        sheet.getRange(rangeInfo.namesBegin + 1, this.outputColumn, rangeInfo.namesEnd - rangeInfo.namesBegin - 1).clearContent();
        sheet.getRange(rangeInfo.idsBegin + 1, this.outputColumn, rangeInfo.idsEnd - rangeInfo.idsBegin - 1).clearContent();
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
            this.log(sheet, message);
        }
    }

    /**
     * Logs a message in the given sheet. If that fails, logs to the console.
     *
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {string} message
     */
    log(sheet, message) {
        try {
            const row = sheet.getLastRow() + 1;
            const range = sheet.getRange(row, 1);
            range.setValue(`${formatLogDate(new Date())} ${message}`);
            sheet.getRange(row, 1, 1, this.outputColumn).setBackground(LOG_BG);
        } catch (err) {
            console.log(`Failed to log in UI "${message}". "${err}"`);
        }
    }
}

const OUTPUT_COLUMN_LABEL = 'Interpreted input data (changes to this column get overwritten)'

const FOLDER_NAME_START = 'Folder names, one per cell (don\'t change this cell)';
const FOLDER_ID_START = 'Folder IDs, one per cell (don\'t change this cell)';
const FOLDER_COLUMN_LABELS = [OUTPUT_COLUMN_LABEL];

class DriveFolderProcessor extends DriveObjectProcessor {
    constructor() {
        super(FOLDERS_SHEET_NAME,
            FOLDER_NAME_START,
            FOLDER_ID_START,
            FOLDER_COLUMN_LABELS,
            true);
    }

    getSheet() {
        //return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetName);
        if (!sheet) {
            // If there's a single sheet, it's probably a new install, so we can rename it.
            const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
            if (sheets.length === 1) {
                if (sheets[0].getLastRow() === 0) {
                    // There's no data. Just rename // ttt2 There might be formatting
                    sheets[0].setName(this.sheetName);
                    this.setupSheet();
                }
            }
        }
        return super.getSheet();
    }


    /**
     * @param {boolean} showConfirmation
     * @returns {boolean} true iff all was OK (the range is valid and the user confirmed it's OK to proceed, then we made the updates)
     */
    setTimes(showConfirmation) {
        const sheet = this.getSheet();

        const inputInfos = this.getInputInfos(sheet, (size) => {
            if (!showConfirmation) {
                return '';
            }
            return `Really set the dates for ${size ? 'the specified' : 'all the'} folders?`;
        });
        if (!inputInfos) {
            return false;
        }

        this.log(sheet, '------------------ Starting update ------------------');
        const timeSetter = new TimeSetter();
        for (const inputInfo of inputInfos) {
            timeSetter.processFolder(inputInfo.idInfos[0], (message => this.log(sheet, message)));
        }
        this.log(sheet, '------------------ Update finished ------------------');
        return true;
    }


    /**
     * @param {boolean} showConfirmation
     */
    listFiles(showConfirmation) {
        const sheet = this.getSheet();

        const inputInfos = this.getInputInfos(sheet, (size) => {
            if (!showConfirmation) {
                return '';
            }
            return `List the files in ${size ? 'the specified' : 'all the'} folders?`;
        });
        if (!inputInfos) {
            return;
        }

        this.log(sheet, '------------------ Starting listing ------------------');
        /**
         * @typedef ListInfo
         *
         * @property {string} name
         * @property {string} id
         * @property {string} path
         * @property {string} time
         * @property {string} size
         * @property {string} mime
         */

        // /** @type {Map<string, ListInfo>} */
        // const filesMap = new Map();  //ttt1: Not sure how to approach the issue of multiple paths to the same file
        /** @type {ListInfo[]} */
        const files= [];
        /** @type {Set<string>} */
        const exploredFolders = new Set();

        for (const inputInfo of inputInfos) {
            this.listFolder(sheet, inputInfo.idInfos[0], files, exploredFolders); // When there are no errors, there is exactly 1 ID per entry
        }
        this.log(sheet, '------------------------------------');
        files.sort((l1, l2) => (l1.path + l1.name).localeCompare(l2.path + l2.name));
        for (const file of files) {
            const row = sheet.getLastRow() + 1;
            const arr = [file.path, file.name, file.time, file.size, file.id, file.mime];
            for (let i = 0; i < arr.length; i += 1) {
                const range = sheet.getRange(row, i + 1);
                let val = arr[i];
                if (i === 2) {
                    val = new Date(arr[i]);  //ttt1: See what to do about this. It's supposed to convert to
                    // a sensible string, based on the spreadsheet and the browser settings, but it shows the date and
                    // no time
                }
                range
                    .setValue(val)
                    .setBackground(LOG_BG);
                if (i === 2) {
                    range.setNumberFormat(LIST_DATETIME_FMT);
                } else if (i !== 3) { //ttt2: hardcoded
                    range.setNumberFormat(PLAIN_TEXT_FMT);
                }
            }
        }
        this.log(sheet, '------------------ Listing finished ------------------');
    }


    /**
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {IdInfo} idInfo
     * @param {ListInfo[]} files
     * @param {Set<string>} exploredFolders
     */
    listFolder(sheet, idInfo, files, exploredFolders) {
        if (exploredFolders.has(idInfo.id)) {
            this.log(sheet, `Already processed ${idInfo.path} [${idInfo.id}]`)
            return;
        }
        exploredFolders.add(idInfo.id)

        const query = `"${idInfo.id}" in parents and trashed = false`;
        let pageToken = null;

        do {
            try {
                const items = Drive.Files.list({  //ttt0: duplicate code for going through children and do something
                    q: query,
                    maxResults: 100,
                    pageToken,
                });

                if (!items.items || items.items.length === 0) {
                    break;
                }
                for (let i = 0; i < items.items.length; i++) {
                    const item = items.items[i];
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
                        this.listFolder(sheet, childIdInfo, files, exploredFolders);
                    } else {
                        if (item.mimeType !== SHORTCUT_MIME) {
                            files.push({
                                id: item.id,
                                name: item.title,
                                path: idInfo.path,
                                size: item.fileSize,  //ttt2: see why is this a string
                                time: item.modifiedDate,
                                mime: item.mimeType,
                            })
                        } else {
                            this.log(sheet, `Ignoring shortcut ${idInfo.path}/${item.title}`);
                        }
                    }
                }
                pageToken = items.nextPageToken;
            } catch (err) {
                const msg = `Failed to process folder '${idInfo.path}' [${idInfo.id}]. ${err}`; //ttt2 We might want
                // ${err.message}, but that might not always exist, and then we get "undefined". This would work, but not
                // sure what value it provides: ${err.message || err}. If the exception being thrown inherits Error (as
                // all exceptions are supposed to), then err.message exists. But some code might throw arbitrary expressions
                this.log(sheet, msg); //ttt0: add "this." before " log"
                //ttt2 improve
            }
        } while (pageToken);
    }


    /**
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {function(number): string} confirmationMessageGetter
     * @returns {(InputInfo[]|null)} an array (which might be empty) with an IdInfo for each user input, if all is OK; null, if there are errors
     */
    getInputInfos(sheet, confirmationMessageGetter) {
        let inputInfos = this.getProcessedInputData(sheet);
        if (!inputInfos) {
            return null;
        }

        const confirmationMessage = confirmationMessageGetter(inputInfos.length);
        if (confirmationMessage && !showConfirmYesNoBox(confirmationMessage)) {
            return null;
        }

        if (!inputInfos.length) {
            const rootFolder = DriveApp.getRootFolder(); // ttt2 This probably needs to change for shared drives
            inputInfos.push({
                idInfos: [{
                    id: rootFolder.getId(),
                    path: '',
                    multiplePaths: false,
                    modifiedDate: SMALLEST_TIME, // Not right, but it will be ignored
                    ownedByMe: true, // doesn't really matter
                }],
                errors: [],
            });
            this.log(sheet, 'Processing all the folders, as no folder names or IDs were specified');
        }
        return inputInfos;
    }
}



const FILE_NAME_START = 'File names, one per cell (don\'t change this cell)';
const FILE_ID_START = 'File IDs, one per cell (don\'t change this cell)';

const FILE_COLUMN_LABELS = ['New date', OUTPUT_COLUMN_LABEL];

class DriveFileProcessor extends DriveObjectProcessor {
    constructor() {
        super(FILES_SHEET_NAME,
            FILE_NAME_START,
            FILE_ID_START,
            FILE_COLUMN_LABELS,
            false);
    }

    /**
     * @param {boolean} showConfirmation
     * @returns {boolean} true iff all was OK (the range is valid and the user confirmed it's OK to proceed, then we made the updates)
     */
    setTimes(showConfirmation) {
        const sheet = this.getSheet();

        let inputInfos = this.getProcessedInputData(sheet);
        if (!inputInfos) {
            return false;
        }

        if (!inputInfos.length) {
            this.showMessage(sheet, 'No files were specified, but at least one is needed');
            return false;
        }

        if (showConfirmation && !showConfirmYesNoBox(`Really set the dates for the specified files?`)) {
            return false;
        }

        this.log(sheet, '------------------ Starting update ------------------');
        const timeSetter = new TimeSetter();
        for (const inputInfo of inputInfos) {
            timeSetter.processFile(inputInfo, (message => this.log(sheet, message)));
        }
        this.log(sheet, '------------------ Update finished ------------------');
        return true;
    }
}

const driveFolderProcessor = new DriveFolderProcessor();
const driveFileProcessor = new DriveFileProcessor();



/**
 * Called when opening the document, to see if the sheets exist and have the right content and tell the user if not.
 */
function setupSheets() {
    const folderSheet = driveFolderProcessor.getSheet();
    const fileSheetExists = driveFileProcessor.sheetExists();
    if (!fileSheetExists) {
        driveFileProcessor.getSheet();
    }

    driveFolderProcessor.setupSheet();
    driveFileProcessor.setupSheet();

    if (!fileSheetExists) {
        // After the sheets have been created, we want to leave the active one as the user set it. At creation,
        // we want to activate folders, as it's what the user probably wants.
        folderSheet.activate(); //ttt3 This doesn't work when running the script in the editor, but
        // works when starting from the Sheet menu. At least it doesn't crash
    }
}


function menuValidateFolders() {
    return driveFolderProcessor.validateInput();
}

/**
 * @returns {boolean} true iff all was OK (the range is valid and the user confirmed it's OK to proceed, then we made the updates)
 */
function menuSetTimesFolders() {
    return driveFolderProcessor.setTimes(true);
}

function menuListFolders() {
    return driveFolderProcessor.listFiles(true);
}


function menuValidateFiles() {
    return driveFileProcessor.validateInput();
}

/**
 * @returns {boolean} true iff all was OK (the range is valid and the user confirmed it's OK to proceed, then we made the updates)
 */
function menuSetTimesFiles() {
    return driveFileProcessor.setTimes(true);
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

// noinspection JSUnusedGlobalSymbols
/**
 * For debugging, to be called from the Google Apps Script web IDE, where a UI is not accessible.
 *
 * @returns {boolean}
 */
function setTimesFilesDebug() {
    return driveFileProcessor.setTimes(false);
}


const FOLDER_MIME = 'application/vnd.google-apps.folder';
const SHORTCUT_MIME = 'application/vnd.google-apps.shortcut';

class TimeSetter {

    constructor() {
        /** @type {Map<string, string>} */
        this.processed = new Map();
    }

    /**
     * @param {IdInfo} idInfo
     * @param {SimpleLogger} log
     * @returns {string} when the folder was last modified, in the format '2000-01-01T10:00:00.000Z'
     */
    processFolder(idInfo, log) {
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
                        newTime = this.processFolder(childIdInfo, log);
                    } else {
                        if (item.mimeType !== SHORTCUT_MIME) {
                            newTime = item.modifiedDate;
                        } else {
                            log(`Ignoring date of shortcut ${idInfo.path}/${item.title}`);
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
                log(msg);
                //ttt2 improve
            }
        } while (pageToken);

        if (res !== idInfo.modifiedDate) {
            if (idInfo.path) {
                // We are not dealing here with the root, which cannot be updated (and for which you couldn't easily see the date anyway)
                try {
                    if (idInfo.ownedByMe) {
                        log(`Setting time to ${res} for ${idInfo.path}. It was ${idInfo.modifiedDate}`);
                        updateModifiedTime(idInfo.id, res);
                    } else {
                        log(`Not updating ${idInfo.path}, which has a different owner`);
                    }
                } catch (err) {
                    const msg = `Failed to update time for folder '${idInfo.path}' [${idInfo.id}]. ${err}`;
                    log(msg);
                    //ttt2 improve
                }
            }
        } else {
            log(`Time ${res} is already correct for ${idInfo.path}`);
        }
        this.processed.set(idInfo.id, res);
        return res;
    }

    /**
     * @param {InputInfo} inputInfo
     * @param {SimpleLogger} log
     * @returns {string} when the folder was last modified, in the format '2000-01-01T10:00:00.000Z'
     */
    processFile(inputInfo, log) {
        const idInfo = inputInfo.idInfos[0];
        log(`Setting time to ${inputInfo.date} for ${idInfo.path}`);
        updateModifiedTime(idInfo.id, inputInfo.date);
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
function formatLogDate(date) {
    const main = date.toLocaleDateString('ro-RO', {
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
    }).substring(4);
    const millis = String(date.getMilliseconds()).padStart(3, '0');
    return `${main}.${millis}`;
}

// /**
//  * Formats a date using YYYY-MM-DD HH:mm
//  * @param {Date} date
//  * @returns {string}
//  */
// function formatShortDate(date) {
//     return `${date.toISOString().replace('T', ' ').replace(/\..*/, '')} UTC`;
// }

/**
 * Converts an ISO date representation to YYYY-MM-DD HH:mm
 * @param {string} isoDate
 * @returns {string}
 */
function isoDateToShort(isoDate) {
    return `${isoDate
        .replace('T', ' ')
        .replace(/\..*/, '')
        .replace(/:00$/, '')} UTC`;
}

//ttt1 Perhaps have a "dry-run", possibly enabled via a "Settings" sheet
