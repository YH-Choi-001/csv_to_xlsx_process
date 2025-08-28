/**
 * @file csv_to_xlsx_process.ts
 * @brief Applies formatting to XLSX files from merged nessus CSV file.
 * @author David Choi <david.choi@pentastic.hk>
 * @version 2025.07.11
 */

/**
 * @brief The entry point of this script.
 * @param workbook The .xlsx file running this script.
 */
function main(workbook: ExcelScript.Workbook) {
	// retrieve the worksheet
	const selectedSheet = workbook.getActiveWorksheet();

	const FILENAME_WITH_EXTENSION = workbook.getName();
	const FILENAME_EXTENSION_SEPARATOR_INDEX = FILENAME_WITH_EXTENSION.lastIndexOf(".");
	const FILENAME =
		FILENAME_EXTENSION_SEPARATOR_INDEX < 0 ?
			FILENAME_WITH_EXTENSION :
			FILENAME_WITH_EXTENSION.substring(0, FILENAME_WITH_EXTENSION.lastIndexOf("."));
	
	if (!FILENAME_WITH_EXTENSION.endsWith(".xlsx")) {
		throw new Error(
			"You can only run this script in .xlsx files.\n" +
			"Save this file as a .xlsx file first.");
	}

	// hide some specific columns
	const COLUMNS_TOBE_HIDDEN: string = "A:C";
	hideColumns(selectedSheet, COLUMNS_TOBE_HIDDEN);

	// apply filter to some specific columns
	const COLUMNS_TOBE_FILTERED: string = "D:O";
	filterColumns(
		selectedSheet,
		COLUMNS_TOBE_FILTERED);

	// sort in ascending order by column D
	// 0 for column A, 1 for column B, 2 for column C, ...
	const RISK_COLUMN = 3; // column D
	sort(selectedSheet, RISK_COLUMN);

	// remove duplicates
	//               considered columns:   D  E  F  G  H  I  J   K   L   M,  N,  O
	const DUPLICATES_CONSIDERED_COLUMNS = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14];
	removeDuplicates(selectedSheet, DUPLICATES_CONSIDERED_COLUMNS);

	// set height of occupied rows to UNIFIED_ROW_HEIGHT
	const UNIFIED_ROW_HEIGHT = 14;
	setRowHeight(selectedSheet, UNIFIED_ROW_HEIGHT);

	// hide unwanted risk levels
	const RISK_LEVELS_TOBE_SHOWN = [
		"Critical",
		"High",
		"Low",
		"Medium"
	];
	onlyShowFiltered(selectedSheet, RISK_COLUMN, RISK_LEVELS_TOBE_SHOWN);

	// autofit specific columns
	const DEVICE_COLUMN:   string = "E";
	const HOST_COLUMN:     string = "F";
	const PORT_COLUMN:     string = "I";
	const NAME_COLUMN:     string = "J";
	const SOLUTION_COLUMN: string = "M";
	const COLUMNS_TOBE_AUTOFITTED = [
		DEVICE_COLUMN,
		HOST_COLUMN,
		PORT_COLUMN,
		NAME_COLUMN,
		SOLUTION_COLUMN];
	autofitColumns(selectedSheet, COLUMNS_TOBE_AUTOFITTED);

	// No need to insert device name in this script,
	// because we let the merge_csv.py script does it.
	// // insert column left of host as device name
	// const deviceName = FILENAME;
	// insertColumn(selectedSheet, HOST_COLUMN, "Device", deviceName);
}

/**
 * @brief Hides column(s).
 * @param worksheet The worksheet to deal with.
 * @param columnsTobeHidden The column(s) to be hidden, e.g. for columns A to C, use "A:C"; for column G only, use "G:G".
 */
function hideColumns(worksheet: ExcelScript.Worksheet, columnsTobeHidden: string) {
	worksheet.getRange(columnsTobeHidden).setColumnHidden(true);
	console.log(`Columns ${columnsTobeHidden} are hidden.`);
}

/**
 * @brief Filters column(s).
 * @param worksheet The worksheet to deal with.
 * @param columnTobeFiltered The column(s) to be filtered, e.g. for columns A to C, use "A:C"; for column G only, use "G:G".
 */
function filterColumns(worksheet: ExcelScript.Worksheet, columnsTobeFiltered: string) {
	worksheet.getAutoFilter().apply(worksheet.getRange(columnsTobeFiltered));
	console.log(`Filters applied to columns ${columnsTobeFiltered}.`);
}

/**
 * @brief Sorts a table by one of its column, in ascending order.
 * @param worksheet The worksheet to deal with.
 * @param columnIndex The index of the column used to sort the table, e.g. 0 for column A, 1 for column B, 2 for column C, ...
 */
function sort(worksheet: ExcelScript.Worksheet, columnIndex: number) {
	if (columnIndex < 0) {
		throw new RangeError(`columnIndex = ${columnIndex} cannot be less than 0.`);
	}
	const filterableRange = worksheet.getAutoFilter().getRange();
	filterableRange.getSort()
		.apply(
			[{
				key: columnIndex - filterableRange.getColumnIndex(), // The index of the column in the range used to sort.
				ascending: true
			}],
			false, // Match case: false.
			true); // Treat first row as a header row: true.
	
	console.log(`Table sorted by column ${columnIndex}.`);
}

/**
 * @brief Only show rows that are filtered.
 * @param columnIndex The index of the column used for filtering.
 * @param valuesTobeShown The array of values to be shown.
 */
function onlyShowFiltered(worksheet: ExcelScript.Worksheet, columnIndex: number, valuesTobeShown: string[]) {
	const autoFilter = worksheet.getAutoFilter();
	const range = autoFilter.getRange();
	autoFilter.apply(
		range,
		// columnIndex - range.getColumnIndex(),
		0,
		{
			filterOn: ExcelScript.FilterOn.values,
			values: valuesTobeShown
		}
	);
	console.log(`Table only show filtered at column ${columnIndex} with values ${valuesTobeShown}.`);
}

/**
 * @brief Removes duplicate rows in a table.
 * @param worksheet The worksheet to deal with.
 * @param consideredColumns The columns to be considered when comparing between rows, e.g. for columns A, C, D, F, use [0, 2, 3, 5]
 */
function removeDuplicates(worksheet: ExcelScript.Worksheet, consideredColumns: number[]) {
	const usedRange: ExcelScript.Range = worksheet.getUsedRange();
	for (let i = 0; i < consideredColumns.length; i++) {
		consideredColumns[i] -= usedRange.getColumnIndex(); // remove the offset of unoccupied columns
	}
	usedRange.removeDuplicates(
		consideredColumns,
		true); // The range contains header: true.
	
	console.log("Duplicates removed.");
}

/**
 * @brief Sets the row height of all occupied cells.
 * @param worksheet The worksheet to deal with.
 * @param rowHeight The desired row height to be set.
 */
function setRowHeight(worksheet: ExcelScript.Worksheet, rowHeight: number) {
	if (rowHeight <= 0) {
		throw new RangeError(`rowHeight = ${rowHeight} cannot be less than or equal to 0.`);
	}
	const usedRange: ExcelScript.Range = worksheet.getUsedRange();
	usedRange.getFormat().setRowHeight(rowHeight);
	console.log(`Height of occupied cells is set to ${rowHeight}.`);
}

/**
 * @brief Applies autofit to some specific columns in the worksheet.
 * @param worksheet The worksheet to deal with.
 * @param columnsTobeAutofitted The columns to be autofitted, e.g. for columns C, G, Q, T, use ["C", "G", "Q", "T"].
 */
function autofitColumns(worksheet: ExcelScript.Worksheet, columnsTobeAutofitted: string[]) {
	// apply autofit to specific columns
	for (const columnTobeAutofitted of columnsTobeAutofitted) {
		autofitColumn(worksheet, columnTobeAutofitted);
	}
	console.log(`Autofit is applied to columns ${columnsTobeAutofitted}.`);
}

/**
 * @brief Applies autofit to a single column in a worksheet.
 * @param worksheet The worksheet to deal with.
 * @param column The column to apply autofit, e.g. for column E, use "E".
 */
function autofitColumn(worksheet: ExcelScript.Worksheet, column: string) {
	const columnRange: string = `${column}:${column}`;
	worksheet.getRange(columnRange).getFormat().autofitColumns();
	console.log(`Autofit is applied to column ${column}.`);
}

function insertColumn(sheet: ExcelScript.Worksheet, insertBeforeThisColumn: string, newColumnName: string, newColumnFields: string) {
	const beforeColumnRangeString: string =
		`${insertBeforeThisColumn}:${insertBeforeThisColumn}`;
	const beforeColumnRange = sheet.getRange(beforeColumnRangeString);

	// Insert the column, shifting existing columns to the right.
	beforeColumnRange.insert(ExcelScript.InsertShiftDirection.right);

	const newColumnRange = sheet.getRange(beforeColumnRangeString);
	const usedRange = sheet.getUsedRange();
	const firstRow = 1;
	const lastRow = usedRange.getRowIndex() + usedRange.getRowCount() - 1;
	const newColumnIndex = newColumnRange.getColumnIndex();
	sheet.getCell(0, newColumnIndex).setValue(newColumnName);
	for (let i = firstRow; i <= lastRow; i++) {
		sheet.getCell(i, newColumnIndex).setValue(newColumnFields);
	}

	console.log(`Column "${newColumnName}" inserted at Column ${insertBeforeThisColumn} populated with "${newColumnFields}"`);
}