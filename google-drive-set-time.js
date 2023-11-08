const DEFAULT_SOURCE_HEIGHT = 5;

const FOLDERS_SHEET_NAME = 'Folders';
const FILES_SHEET_NAME = 'Files';

const FOLDER_NAME_START = 'Folder names, one per cell (don\'t change this cell)';
const FOLDER_ID_START = 'Folder IDs, one per cell (don\'t change this cell)';
const LOG_START = 'Log (don\'t change this cell)';

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

    logF('msg1');
    logF('msg2');
}

/**
 * @returns {SpreadsheetApp.Sheet}
 */
function getFoldersSheet() {
    //return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOLDERS_SHEET_NAME);
}

/**
 * @returns {SpreadsheetApp.Sheet}
 */
function getFilesSheet() {
    //return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FILES_SHEET_NAME);
}


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
//ttt1 Review other triggers: https://developers.google.com/apps-script/guides/triggers

/**
 * Called when opening the document, to see if the sheets exist and have the right content and tell the user if not.
 */
function setupSheets() {
    if (!getFoldersSheet()) {
        // If there's a single sheet, it's probably a new install, so we can rename it.
        const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
        if (sheets.length === 1) {
            if (sheets[0].getLastRow() === 0) {
                // There's no data. Just rename // ttt2 There might be formatting
                sheets[0].setName(FOLDERS_SHEET_NAME);
            }
        }
    }

    if (!getFoldersSheet()) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        sheet.setName(FOLDERS_SHEET_NAME);
    }
    if (!getFilesSheet()) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
        sheet.setName(FILES_SHEET_NAME);
    }

    getFoldersSheet().activate(); //ttt1 This doesn't work when running the script in the editor, but
    // works when starting from the Sheet menu. At least it doesn't crash

    setupFoldersSheet();
}

function setupFoldersSheet() {
    addLabelsIfEmptyFoldersSheet();
    applyColorToFoldersSheet();
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
 * @typedef {Object} FolderRangeInfo
 * @property {number} namesBegin
 * @property {number} namesEnd
 * @property {number} idsBegin
 * @property {number} idsEnd
 * @property {number} logsBegin
 * @property {number} logsEnd
 */

/**
 * Checks the first column for section delimiters and data and returns the ranges for names, IDs, and logs.
 * To be used for coloring or log clearing.
 * If the delimiters are not found in the expected order, returns null.
 *
 * The ends are exclusive, and, for now, coincide with the beginning of the next section. This might change, though.
 *
 * @returns {(FolderRangeInfo|null)} null if the range is invalid
 */
function getFolderRangeInfo() {
    const foldersSheet = getFoldersSheet();
    const lastRow = foldersSheet.getLastRow();
    const rows = foldersSheet.getRange(1, 1, lastRow).getValues();
    const expected = [FOLDER_NAME_START, FOLDER_ID_START, LOG_START];
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
 * If the "Folders" sheet is empty, add the section start labels
 */
function addLabelsIfEmptyFoldersSheet() {
    const foldersSheet = getFoldersSheet();

    if (foldersSheet.getLastRow() !== 0) {
        return;
    }

    const BOLD_FONT = 'bold';
    let crtLine = 1;

    function addLabel(label, increment) {
        const range = foldersSheet.getRange(crtLine, 1);
        range.setValue(label);
        range.setFontWeight(BOLD_FONT);
        //range.protect().removeEditor('XYZ@gmail.com'); // It's what we want, but doesn't work
        range.protect().setWarningOnly(true); // We really want nobody being able to edit, but it can't be done, so we use warnings
        crtLine += increment;
    }

    addLabel(FOLDER_NAME_START, DEFAULT_SOURCE_HEIGHT);
    addLabel(FOLDER_ID_START, DEFAULT_SOURCE_HEIGHT);
    addLabel(LOG_START, DEFAULT_SOURCE_HEIGHT);

    // Use some educated guesses for the column widths ...
    foldersSheet.autoResizeColumn(1);
    const w = foldersSheet.getColumnWidth(1);
    foldersSheet.setColumnWidth(1, w * 1.2);
    foldersSheet.setColumnWidth(2, w * 2.4);
}

const NAMES_BG = '#efe';
const IDS_BG = '#eef';
const LOG_BG = '#ffc';
const ERROR_BG = '#fbb';

const RANGE_NOT_FOUND_ERR = 'Couldn\'t find section delimiters. If a manual fix is not obvious,'
    + ' delete or rename the "Folders" sheet and then reopen the spreadsheet';

/**
 * Sets the background for the first column, so names, IDs, and logs each have their own color.
 * Throws if (some of) the section starts are not found or are not in their proper order.
 */
function applyColorToFoldersSheet() {
    const foldersSheet = getFoldersSheet();
    const rangeInfo = getFolderRangeInfo();
    if (!rangeInfo) {
        showFolderMessage(RANGE_NOT_FOUND_ERR); //ttt0 Make sure this is only shown once
        return;
    }
    foldersSheet.getRange(rangeInfo.namesBegin, 1, rangeInfo.namesEnd - rangeInfo.namesBegin, 2).setBackground(NAMES_BG);
    foldersSheet.getRange(rangeInfo.idsBegin, 1, rangeInfo.idsEnd - rangeInfo.idsBegin, 2).setBackground(IDS_BG);
    foldersSheet.getRange(rangeInfo.logsBegin, 1, rangeInfo.logsEnd - rangeInfo.logsBegin, 2).setBackground(LOG_BG);
}

/**
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {number} column
 * @param {number} rowStart
 * @param {number} rowEnd exclusive
 * @returns {string[]} data in a column, between 2 rows, as an array.
 */
function getColumnData(sheet, column, rowStart, rowEnd) {
    const rows = sheet.getRange(rowStart, column, rowEnd - rowStart).getValues();
    /** @type string[] */
    const res = [];
    for (let i = 0; i < rows.length; i += 1) {
        res.push(rows[i][0]);
    }
    return res;
}

const FOLDER_MIME = 'application/vnd.google-apps.folder';
const SHORTCUT_MIME = 'application/vnd.google-apps.shortcut';

function menuSetTimesFolders() {
    //ttt0 confirmation
    const foldersSheet = getFoldersSheet();
    foldersSheet.activate();
    const rangeInfo = getFolderRangeInfo();
    if (!rangeInfo) {
        showFolderMessage(RANGE_NOT_FOUND_ERR); //ttt0 Make sure this is only shown once
        return;
    }
//foldersSheet.getRange()
    /** @type {Map<string, IdInfo>} */
    const idInfosMap = new Map();
    const names = getColumnData(foldersSheet, 1, rangeInfo.namesBegin + 1, rangeInfo.namesEnd);
    const inputNameInfos = validateFolderNames(names, idInfosMap);
    const inputIds = getColumnData(foldersSheet, 1, rangeInfo.idsBegin + 1, rangeInfo.idsEnd);
    const inputIdInfos = validateFolderIds(inputIds, idInfosMap);
    updateFolderUiAfterValidation(rangeInfo, inputNameInfos, inputIdInfos);
    const nameErrors = inputNameInfos.filter((val) => val.errors.length);
    const idErrors = inputIdInfos.filter((val) => val.errors.length);
    if (nameErrors.length || idErrors.length) {
        showFolderMessage('Found some errors, which need to be resolved before proceeding with setting the times');
        return;
    }

    logF('------------------ Starting update ------------------');
    const idInfosArr = Array.from(idInfosMap.values());
    setTimes(idInfosArr);
    logF('------------------ Update finished ------------------');
}

const SMALLEST_TIME = '1970-01-01T12:00:00.000Z'; //ttt2 Review if something else would be better. (Hour
// is set at noon, so most timezones will see it as January 1st)

class TimeSetter {
    constructor() {
        /** @type {Map<string, string>} */
        this.processed = new Map();
    }

    /**
     * @param {IdInfo} idInfo
     * @returns {string} when the folder was last modified, in the format '2000-01-01T10:00:00.000Z'
     */
    process(idInfo) {
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
                        newTime = this.process({
                            id: item.id,
                            modifiedDate: item.modifiedDate,
                            multiplePaths: false,  // not correct, but it doesn't matter; it's just to have something
                            path: `${idInfo.path}/${item.title}`,
                        });
                    } else {
                        if (item.mimeType !== SHORTCUT_MIME) {
                            //ttt1 Perhaps log that link was skipped, or have an option not to skip it
                            newTime = item.modifiedDate;
                        }
                    }
                    if (newTime > res) {
                        res = newTime;
                    }
                }
                pageToken = items.nextPageToken;
            } catch (err) {
                const msg = `Failed to process folder '${idInfo.path}' [${idInfo.id}]. Error: ${err.message}`; //ttt1 Review the use of ".message", here and elsewhere. Might not exist
                logF(msg);
                //ttt2 improve
            }
        } while (pageToken);

        if (res !== idInfo.modifiedDate) {
            if (idInfo.path) {
                // We are not dealing here with the root, which cannot be updated (and you couldn't see the date anyway)
                try {
                    logF(`Setting time to ${res} for ${idInfo.path}. It was ${idInfo.modifiedDate}`);
                    updateModifiedTime(idInfo.id, res);
                } catch (err) {
                    const msg = `Failed to update time for folder '${idInfo.path}' [${idInfo.id}]. Error: ${err.message}`;
                    logF(msg);
                    //ttt2 improve
                }
            }
        } else {
            logF(`Time ${res} is already correct for ${idInfo.path}`);
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
    Drive.Files.update(body, id, blob, optionalArgs); //ttt1 This fails silently for non-owners, so perhaps don't call it
}


/**
 * @param {IdInfo[]} idInfos
 */
function setTimes(idInfos) {
    if (!idInfos.length) {
        const rootFolder = DriveApp.getRootFolder();
        idInfos.push({
            id: rootFolder.getId(),
            path: '',
            multiplePaths: false,
            modifiedDate: SMALLEST_TIME, // Not right, but it will be ignored
        });
        logF('Processing all the files, as no folder names or IDs were specified');
    }

    const timeSetter = new TimeSetter();
    for (const idInfo of idInfos) {
        timeSetter.process(idInfo);
    }
}


/**
 * IdInfo descr
 * @typedef {Object} IdInfo
 * @property {string} id
 * @property {string} path
 * @property {boolean} multiplePaths
 * @property {string} modifiedDate
 */

/**
 * InputNameInfo descr
 * @typedef {Object} InputNameInfo
 * @property {IdInfo[]} folders
 * @property {string[]} errors
 */

/**
 * InputIdInfo descr
 * @typedef {Object} InputIdInfo
 * @property {IdInfo} folder
 * @property {string[]} errors
 */

/*

IdInfo: {
    id: "ke8436",
    //name: "name1",
    path: "/kf84/asf",
    multiplePaths: true, // optional, for when there are multiple parents
    //cnt: 3, // just for errors, to be able to tell how many times a folder has been added already //ttt1 See if we want this
    modifiedDate: "2000-01-01T10:00:00.000Z",
}

InputNameInfo:
{
    folders: [
        { // IdInfo
            id: "hd73hb",
            //name: "name1",
            path: "/de/fd/gth",
            modifiedDate: "2000-01-01T10:00:00.000Z",
        },
        {
            id: "ke8436",
            //name: "name1",
            path: "/kf84/asf",
            multiplePaths: true,
            modifiedDate: "2000-01-01T10:00:00.000Z",
        },
    ],
    errors: [
        "Found multiple folders with name 'name1'",
    ]
}

InputIdInfo:
{
    folder: { // IdInfo
        id: "hd73hb",
        name: "name1",
        path: "/de/fd/gth",
    },
    errors: [
        "ID 'hd73hb' already found",
    ]
}
*/


/**
 * Returns an array the size of up to the "names" section without the title with the type InputNameInfo.
 * For empty names, the corresponding entry is undefined (or missing, at the end of the array).
 * For non-empty names, the corresponding entry is like in the comment above. We want exactly one entry, otherwise it's an error.
 * Not clear how to deal with multipaths. We should probably cache them, and maybe generate a warning. //ttt0 cache
 *
 * @param {string[]} names - array of strings
 * @param {Map<string, IdInfo>} idInfos - map with ID as key and a IdInfo as the value; a FolderInfo with multiple folders
 *      entries will also have multiple entries in the map; the map is used here to figure out which folders are
 *      already defined and then as the input for the actual processing
 * @returns {InputNameInfo[]}
 */
function validateFolderNames(names, idInfos) {
    /** @type InputNameInfo[] */
    const res = [];
    for (let i = 0; i < names.length; i += 1) {
        const name = names[i];
        if (!name) {
            continue;
        }
        const folders = [];
        const errors = [];
        const driveFolders = DriveApp.getFoldersByName(name);
        while (driveFolders.hasNext()) {
            const driveFolder = driveFolders.next();
            const idInfo = getIdInfo(driveFolder.getId());
            folders.push(idInfo);
            const existing = idInfos.get(idInfo.id);
            if (existing) {
                errors.push(`Folder with the ID ${idInfo.id} already added`);
            } else {
                idInfos.set(idInfo.id, idInfo);
            }
            // driveFolders.getContinuationToken(...) // Looks like something to get more
            // entries, but we don't really care about this. "More than 1" is good enough
        }
        const cnt = folders.length;
        if (!cnt) {
            errors.push(`No folder found with the name '${name}'`);
        } else if (cnt > 1) {
            errors.push(`Found ${cnt} folders with the name '${name}'`);
        }
        res[i] = {
            folders,
            errors,
        };
    }
    return res;
}


/**
 * Returns an array the size of up to the "names" section without the title with the type InputIdInfo
 * For empty IDs, the corresponding entry is unedfined (or missing, at the end of the array).
 *
 * @param {string[]} ids
 * @param {Map<string, IdInfo>} idInfos - map with ID as key and a IdInfo as the value; a FolderInfo with multiple folders
 *      entries will also have multiple entries in the map; the map is used here to figure out which folders are
 *      already defined and then as the input for the actual processing
 * @returns {InputIdInfo[]}
 */
function validateFolderIds(ids, idInfos) {
    /** @type InputIdInfo[] */
    const res = [];
    for (let i = 0; i < ids.length; i += 1) {
        const id = ids[i];
        if (!id) {
            continue;
        }
        const errors = [];
        let folder;
        try {
            const idInfo = getIdInfo(id);
            const existing = idInfos.get(idInfo.id);
            if (existing) {
                errors.push(`Folder with the ID ${idInfo.id} already added`);
            } else {
                idInfos.set(idInfo.id, idInfo);
            }
            folder = idInfo;
        } catch (err) {
            errors.push(`Folder with ID ${id} not found`);
        }
        res[i] = {
            folder,
            errors,
        }
    }
    return res;
}



/**
 * @param {string} id ID of a file or a folder
 * @returns {IdInfo}
 */
function getIdInfo(id) {
    const objectInfo = Drive.Files.get(id)
    let parents = objectInfo.parents;
    let parentCnt = parents.length;
    if (parentCnt === 0) {
        return {
            id,
            path: '',
            multiplePaths: false,
            modifiedDate: objectInfo.modifiedDate,
        };
    }
    let parent = parents[0];
    let parentInfo = getIdInfo(parent.id);
    return {
        id,
        path: `${parentInfo.path}/${objectInfo.title}`,
        multiplePaths: parentInfo.multiplePaths || (parentCnt > 1),
        modifiedDate: objectInfo.modifiedDate,
    };
}


/**
 * @param {FolderRangeInfo} rangeInfo
 * @param {InputNameInfo[]} inputNameInfos
 * @param {InputIdInfo[]} inputIdInfos
 */
function updateFolderUiAfterValidation(rangeInfo, inputNameInfos, inputIdInfos) {
    clearErrors();
    applyColorToFoldersSheet();
    const foldersSheet = getFoldersSheet();

    for (let i = 0; i < inputNameInfos.length; i += 1) {
        const folderNameInfo = inputNameInfos[i];
        if (!folderNameInfo) {
            continue;
        }
        let lines = [];
        for (const folder of folderNameInfo.folders) {
            lines.push(`${folder.path}${folder.multiplePaths ? ' (and others)' : ''} [${folder.id}]`);
        }
        const cellRange = foldersSheet.getRange(rangeInfo.namesBegin + 1 + i, 2);
        if (folderNameInfo.errors.length) {
            lines.push(...folderNameInfo.errors);
            cellRange.setBackground(ERROR_BG);
        }
        cellRange.setValue(lines.join('\n'));
    }

    for (let i = 0; i < inputIdInfos.length; i += 1) {
        const folderIdInfo = inputIdInfos[i];
        if (!folderIdInfo) {
            continue;
        }
        let lines = [];
        if (folderIdInfo.folder) {
            lines.push(`${folderIdInfo.folder.path}${folderIdInfo.folder.multiplePaths ? ' (and others)' : ''}  [${folderIdInfo.folder.id}]`); //ttt1 duplicate code "(and others)"
        }
        const cellRange = foldersSheet.getRange(rangeInfo.idsBegin + 1 + i, 2);
        if (folderIdInfo.errors.length) {
            lines.push(...folderIdInfo.errors);
            cellRange.setBackground(ERROR_BG);
        }
        cellRange.setValue(lines.join('\n'));
    }

    //foldersSheet.autoResizeColumns(1, 2); //!!! Not right, as it include the logs
}

/**
 * Erases the second column, which is supposed to be used for errors.
 */
function clearErrors() {
    const foldersSheet = getFoldersSheet();
    foldersSheet.getRange(1, 2, foldersSheet.getLastRow()).clearContent();
}


/**
 * Logs a message in the "Folders" sheet. If that fails, logs to the console.
 * @param {string} message
 */
function logF(message) {
    try {
        const foldersSheet = getFoldersSheet();
        const row = foldersSheet.getLastRow() + 1;
        const range = foldersSheet.getRange(row, 1);
        range.setValue(`${formatDate(new Date())} ${message}`);
        foldersSheet.getRange(row, 1, 1, 2).setBackground(LOG_BG);
    } catch (err) {
        console.log(`Failed to log in UI "${message}". Reason: "${err.message}"`);
    }
}


/**
 * Show a popup message in the "Folders" sheet. If that fails, logs to the log section in
 * the "Folders" sheet. If that also fails, logs to the console.
 *
 * @param {string} message
 */
function showFolderMessage(message) {
    try {
        SpreadsheetApp.getUi().alert(message);
    } catch {
        logF(message);
    }
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
