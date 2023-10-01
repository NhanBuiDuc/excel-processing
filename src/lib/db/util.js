import { read, utils } from 'xlsx';
import xlsx from 'xlsx';

import fs from 'fs';
import unorm from 'unorm';
import { attributeMappings, attributeOrder } from './keyword_mapping.js';
import {
	insertStudent,
	selectIdFromStudentData,
	getStudentsFromClassroom,
	deleteChainAttendanceEventAttendanceStudentParent
} from './student.js';
import { DateTime } from 'luxon';
import { insertParent } from './parent.js';
import xlsxPopulate from 'xlsx-populate';
import { supabase } from './supabase.js';
import ExcelJS from 'exceljs';

// Function to check if a row is valid based on certain criteria.
function isValidRow(row) {
	// Add your criteria here to determine if the row is valid.
	// For example, you might check if all required columns have values.

	// For demonstration purposes, let's assume the first column (index 0) is required.
	return row[0] !== undefined && row[0] !== null && row[0] !== '';
}
export function readRowsAndConvertToJson(worksheet, headers, subheaders) {
	const data = [];
	const rowRange = xlsx.utils.decode_range(worksheet['!ref']);

	let currentMainHeader = null;
	let previousMainHeader = null;
	let rowObj = {};

	for (let rowIdx = rowRange.s.r + 2; rowIdx <= rowRange.e.r; rowIdx++) {
		const row = [];
		for (let colIdx = rowRange.s.c; colIdx <= rowRange.e.c; colIdx++) {
			const cellAddress = xlsx.utils.encode_cell({ r: rowIdx, c: colIdx });
			const cell = worksheet[cellAddress];
			const cellValue = cell ? cell.v : null;
			row.push(cellValue);
		}

		if (!isValidRow(row)) {
			// Skip this row if it's not valid based on the criteria.
			continue;
		}

		for (let colIdx = 0; colIdx < row.length; colIdx++) {
			const cellValue = row[colIdx];
			currentMainHeader = headers[colIdx];
			if (currentMainHeader == null || currentMainHeader == undefined || currentMainHeader == '') {
				currentMainHeader = previousMainHeader;
			}
			if (!(currentMainHeader in rowObj)) {
				rowObj[currentMainHeader] = {}; // Assign an empty object
			}
			if (isHasSubHeader(currentMainHeader, headers)) {
				const subHeader = subheaders[colIdx];
				if (cellValue !== undefined && cellValue !== null) {
					rowObj[currentMainHeader][subHeader] = cellValue;
				} else {
					rowObj[currentMainHeader][subHeader] = '';
				}
			} else {
				if (currentMainHeader && cellValue !== undefined && cellValue !== null) {
					rowObj[currentMainHeader] = cellValue;
				} else {
					rowObj[currentMainHeader] = '';
				}
			}

			previousMainHeader = currentMainHeader;
		}

		if (Object.keys(rowObj).length > 0) {
			data.push(rowObj);
		}

		// Reset the row object for the next row
		rowObj = {};
	}

	const result = data;

	// Output the JSON to a file
	const outputFilename = 'output.json';
	fs.writeFileSync(outputFilename, JSON.stringify(data, null, 4));

	console.log(`JSON data written to file: ${outputFilename}`);
	return result;
}
export function readDataAndAttendance(worksheet, valid_attendance_dates) {
	const data = [];
	const rowRange = xlsx.utils.decode_range(worksheet['!ref']);

	for (let rowIdx = rowRange.s.r + 2; rowIdx <= rowRange.e.r; rowIdx++) {
		const row = [];

		// Read columns 1 to 27
		for (let colIdx = rowRange.s.c; colIdx <= 26; colIdx++) {
			const cellAddress = xlsx.utils.encode_cell({ r: rowIdx, c: colIdx });
			const cell = worksheet[cellAddress];
			const cellValue = cell ? cell.v : null;
			row.push(cellValue);
		}

		// Read columns 27 to (27 + valid_attendance_dates.length)
		for (
			let colIdx = 27;
			colIdx < 27 + valid_attendance_dates['attendance_dates'].length;
			colIdx++
		) {
			const cellAddress = xlsx.utils.encode_cell({ r: rowIdx, c: colIdx });
			const cell = worksheet[cellAddress];
			const cellValue = cell ? cell.v : null;
			row.push(cellValue);
		}

		if (!isValidRow(row)) {
			// Skip this row if it's not valid based on the criteria.
			continue;
		}

		data.push(row);
	}

	return data;
}

export function readExcelFile(filePath, sheetName) {
	const workbook = xlsx.readFile(filePath);
	const worksheet = workbook.Sheets[sheetName];

	return worksheet;
}
// // Function to read the uploaded Excel file
// export function readExcelFileFromBuffer(file, sheetName) {
// 	const workbook = xlsx.read(file.data, { type: 'buffer' });
// 	// Change 'Sheet1' to your desired sheet name
// 	const worksheet = workbook.Sheets['Sheet1'];
// 	return xlsx.utils.sheet_to_json(worksheet);
// }
export function getRowValues(worksheet, range) {
	const cells = xlsx.utils.sheet_to_json(worksheet, { range, header: 1, defval: null });
	return Object.values(cells[0]);
}

function sanitizeHeader(header) {
	if (typeof header === 'string') {
		// Remove specific symbols, while preserving diacritics
		const sanitizedHeader = header.replace(/[`~!@.,<>?;'ơ\]=]/g, '').toLowerCase();
		return sanitizedHeader;
	}
	return '';
}

export function checkHeader(header, expectedMainHeader, expectedSubHeader, acceptableHeaders) {
	const errors = [];

	// Check if the acceptableHeaders object is empty
	if (Object.keys(acceptableHeaders).length === 0) {
		errors.push('The acceptableHeaders object is empty.');
	}

	for (let i = 0; i < header.length; i++) {
		const cellValue = sanitizeHeader(header[i]);
		const expectedMain = expectedMainHeader && expectedMainHeader[i];
		const expectedSub = expectedSubHeader && expectedSubHeader[i];

		if (cellValue !== sanitizeHeader(expectedMain) && expectedMain !== null) {
			if (acceptableHeaders[expectedMain]) {
				const acceptableValues = acceptableHeaders[expectedMain].map(sanitizeHeader);
				if (!acceptableValues.includes(cellValue)) {
					errors.push(
						`Expected main header "${expectedMain}" at index ${i}, found "${cellValue}" but acceptable values are ${acceptableValues}`
					);
				}
			} else {
				errors.push(`Expected main header "${expectedMain}" at index ${i}, found "${cellValue}"`);
			}
		}

		if (cellValue !== sanitizeHeader(expectedSub) && expectedSub !== null) {
			if (acceptableHeaders[expectedSub]) {
				const acceptableValues = acceptableHeaders[expectedSub].map(sanitizeHeader);
				if (!acceptableValues.includes(cellValue)) {
					errors.push(
						`Expected sub header "${expectedSub}" at index ${i}, found "${cellValue}" but acceptable values are ${acceptableValues}`
					);
				}
			} else {
				errors.push(`Expected sub header "${expectedSub}" at index ${i}, found "${cellValue}"`);
			}
		}
	}

	return errors;
}

export function checkHeaderFile(
	worksheet,
	mainHeaderRanges,
	subHeaderRange,
	expectedMainHeader,
	expectedSubHeader,
	acceptableHeaders
) {
	let mainHeader = [];
	for (const range of mainHeaderRanges) {
		const values = getRowValues(worksheet, range);
		mainHeader = mainHeader.concat(values);
	}

	// Check if the main header matches the expected format
	const mainHeaderErrors = checkHeader(mainHeader, expectedMainHeader, null, acceptableHeaders);
	if (mainHeaderErrors.length > 0) {
		console.error('Main header errors:', mainHeaderErrors);
		return false;
	}

	// Check if the sub header matches the expected format
	const subHeader = getRowValues(worksheet, subHeaderRange);
	const subHeaderErrors = checkHeader(subHeader, null, expectedSubHeader, acceptableHeaders);
	if (subHeaderErrors.length > 0) {
		console.error('Sub header errors:', subHeaderErrors);
		return false;
	}

	// console.log('Main header:', mainHeader);
	// console.log('Sub header:', subHeader);

	return true;
}
function isHasSubHeader(currentHeader, headers) {
	const currentIndex = headers.indexOf(currentHeader);
	const nextValue = headers[currentIndex + 1];

	return nextValue === null;
}
export const transformDataKeys = (data, mapping) => {
	const transformedData = {};
	for (const [key, value] of Object.entries(mapping)) {
		if (typeof value === 'object') {
			transformedData[key] = transformDataKeys(data, value);
		} else {
			transformedData[key] = data[value] || null;
		}
	}
	return transformedData;
};

export function extractData(dictData, mapping) {
	const result = [];

	for (const dict of dictData) {
		const entityData = {};

		for (const entity in mapping) {
			entityData[entity] = {};

			for (const attribute in mapping[entity]) {
				const key = mapping[entity][attribute];
				const keys = key.split('.');

				let value = dict;
				for (const k of keys) {
					value = value[k];
					if (typeof value === 'undefined') break;
				}

				// Handle special case for "sex" attribute
				if (attribute === 'sex' && typeof value === 'object' && value !== null) {
					const nonEmptyKeys = Object.keys(value).filter((key) => value[key] !== '');
					if (nonEmptyKeys.length > 0) {
						value = nonEmptyKeys[0];
					} else {
						value = null; // Set to null if all keys have empty values
					}
				}

				entityData[entity][attribute] = value;
			}
		}

		result.push(entityData);
	}
	return result;
}

export async function insertStudentAndParent(data, class_room_id, branch_id) {
	let transformedData = data.map((item) => {
		// Step 1: Transform student data
		const student = {
			class_room_id: class_room_id,
			grade: item['LỚP'],
			first_name: item['HỌ VÀ TÊN LÓT'],
			last_name: item['TÊN HỌC SINH'],
			enroll_date: item['NGÀY NHẬP HỌC'],
			dob: item['SINH NGÀY'],
			birth_year: item['NĂM SINH'],
			sex: item['GIỚI TÍNH']['NỮ'] === 'x' ? 'Nữ' : 'Nam',
			ethnic: item['DÂN TỘC'],
			birth_place: item['NƠI SINH HS'],
			temp_res: item['TẠM TRÚ'],
			perm_res_province: item['THƯỜNG TRÚ']['TỈNH'],
			perm_res_district: item['THƯỜNG TRÚ']['HUYỆN'],
			perm_res_commune: item['THƯỜNG TRÚ']['XÃ']
		};

		// Step 2: Transform parent data
		const father = {
			student_id: null,
			phone_number: item['SỐ ĐIỆN THOẠI']['ĐT1'],
			name: item['THÔNG TIN CHA']['HỌ VÀ TÊN'],
			dob: item['THÔNG TIN CHA']['NĂM SINH'],
			sex: 'NAM',
			occupation: item['THÔNG TIN CHA']['NGHỀ NGHIỆP'],
			zalo: item['ZALO'],
			landlord: item['CHỦ NHÀ TRỌ'],
			roi: item['BC PHỔ CẬP'],
			birthplace: item['KHAI SINH'],
			res_registration: item['HỘ KHẨU']
		};

		const mother = {
			student_id: null,
			phone_number: item['SỐ ĐIỆN THOẠI']['ĐT2'],
			name: item['THÔNG TIN MẸ']['HỌ VÀ TÊN'],
			dob: item['THÔNG TIN MẸ']['NĂM SINH'],
			sex: 'NỮ',
			occupation: item['THÔNG TIN MẸ']['NGHỀ NGHIỆP'],
			zalo: item['ZALO'],
			landlord: item['CHỦ NHÀ TRỌ'],
			roi: item['BC PHỔ CẬP'],
			birthplace: item['KHAI SINH'],
			res_registration: item['HỘ KHẨU']
		};

		return { student, parents: [father, mother] };
	});
	transformedData = preprocessDate('dob', transformedData);
	transformedData = preprocessDate('enroll_date', transformedData);
	let dbStudentParentData = await getStudentAndParentData(class_room_id);

	// Assuming transformedData and dbStudentParentData are already defined
	// let finalData = mergeData(
	// 	removeDuplicateStudents([...transformedData, ...dbStudentParentData]),
	// 	dbStudentParentData
	// );
	// Modify finalData to set id to null for students with no id
	// finalData.forEach((item) => {
	// 	if (!item.student.id) {
	// 	item.student.id = null;
	// 	}
	// });
	trimNamesInDataList(transformedData);
	trimNamesInDataList(dbStudentParentData);
	let { syncData, notSyncData } = await processSyncAndNotSyncData(
		transformedData,
		dbStudentParentData
	);
	let { toBeInserted, toBeDeletedIds } = await processInsertedAndNotInsertedData(notSyncData);
	await deleteChainAttendanceEventAttendanceStudentParent(toBeDeletedIds);
	trimNamesInDataList(syncData);
	trimNamesInDataList(toBeInserted);

	toBeInserted = preprocessDate('dob', toBeInserted);
	toBeInserted = preprocessDate('enroll_date', toBeInserted);
	syncData = preprocessDate('dob', syncData);
	syncData = preprocessDate('enroll_date', syncData);
	await pushToDatabase(toBeInserted);
	await pushToDatabase(syncData);
}
function trimNamesInDataList(dataList) {
	for (const data of dataList) {
		// Trim student name
		data.student.first_name = data.student.first_name.trim();
		data.student.last_name = data.student.last_name.trim();

		// Trim parent names
		for (const parent of data.parents) {
			parent.name = parent.name.trim();
		}
	}
}

export async function processInsertedAndNotInsertedData(notSyncData) {
	let toBeInserted = [];
	let toBeDeletedIds = [];

	for (const item of notSyncData) {
		if (!item.student.id) {
			toBeInserted.push(item);
		} else {
			toBeDeletedIds.push(item.student.id);
		}
	}

	return { toBeInserted, toBeDeletedIds };
}
// Step 3: Get mapped parent data based on student_id
function getMappedParentData(studentId, data) {
	const parentData = data.find((item) => item[0].id === studentId);
	return parentData ? parentData[1] : [];
}

// Step 3: Get mapped parent data based on student_id
async function getMappedDbStudentData(branch_id, class_room_id, data) {
	const studentIds = data.map((item) => item.student.id);
	const dbStudentData = await getStudentsFromClassroom(branch_id, class_room_id);
	return dbStudentData.map((student) => ({
		student: student,
		parent: getMappedParentData(student.id, studentIds, data)
	}));
}

function processParentData(data, studentIds) {
	const parentsArray = [];

	data.forEach((element, index) => {
		const father = {
			student_id: studentIds[index], // Add studentId to the father object
			phone_number: element.SỐ_ĐIỆN_THOẠI.ĐT1 || null, // Set to null if empty or undefined
			name: element.THÔNG_TIN_CHA['HỌ VÀ TÊN'] || null, // Set to null if empty or undefined
			dob: element.THÔNG_TIN_CHA['NĂM SINH'] || null, // Set to null if empty or undefined
			occupation: element.THÔNG_TIN_CHA['NGHỀ NGHIỆP'] || null, // Set to null if empty or undefined
			zalo: element.ZALO || null, // Set to null if empty or undefined
			landlord: element['THƯỜNG TRÚ']?.TỈNH || null, // Set to null if empty or undefined
			roi: element['THƯỜNG TRÚ']?.HUYỆN || null, // Set to null if empty or undefined
			birthplace: element.NƠI_SINH_HS || null, // Set to null if empty or undefined
			res_registration: element['THƯỜNG TRÚ']?.XÃ || null // Set to null if empty or undefined
		};

		const mother = {
			student_id: studentIds[index], // Add studentId to the mother object
			phone_number: element.SỐ_ĐIỆN_THOẠI.ĐT2 || null, // Set to null if empty or undefined
			name: element.THÔNG_TIN_MẸ['HỌ VÀ TÊN'] || null, // Set to null if empty or undefined
			dob: element.THÔNG_TIN_MẸ['NĂM SINH'] || null, // Set to null if empty or undefined
			occupation: element.THÔNG_TIN_MẸ['NGHỀ NGHIỆP'] || null, // Set to null if empty or undefined
			zalo: element.ZALO || null, // Set to null if empty or undefined
			landlord: element['THƯỜNG TRÚ']?.TỈNH || null, // Set to null if empty or undefined
			roi: element['THƯỜNG TRÚ']?.HUYỆN || null, // Set to null if empty or undefined
			birthplace: element.NƠI_SINH_HS || null, // Set to null if empty or undefined
			res_registration: element['THƯỜNG TRÚ']?.XÃ || null // Set to null if empty or undefined
		};

		// Check if both father and mother objects have at least one non-empty attribute
		if (
			Object.values(father).some((value) => value !== '' && value !== ' ' && value !== null) ||
			Object.values(mother).some((value) => value !== '' && value !== ' ' && value !== null)
		) {
			const parents = [father, mother];
			parentsArray.push(parents);
		}
	});

	return parentsArray;
}

function preprocessDate(dateKey, data) {
	data.forEach((item) => {
		const row = item.student; // Extracting the student object from the item
		if (row[dateKey]) {
			const dateFormats = [
				'dd/MM/yyyy',
				'dd/M/yyyy',
				'd/MM/yyyy',
				'd/M/yyyy',
				'dd/MM/yy',
				'dd/M/yy',
				'd/MM/yy',
				'd/M/yy',
				'yyyy/M/dd',
				'yyyy/MM/d',
				'yyyy/M/d',
				'yy/MM/dd',
				'yy/M/dd',
				'yy/MM/d',
				'yy/M/d',

				'dd-MM-yyyy',
				'dd-M-yyyy',
				'd-MM-yyyy',
				'd-M-yyyy',
				'dd-MM-yy',
				'dd-M-yy',
				'd-MM-yy',
				'd-M-yy',
				'yyyy-M-dd',
				'yyyy-MM-d',
				'yyyy-M-d',
				'yy-MM-dd',
				'yy-M-dd',
				'yy-MM-d',
				'yy-M-d'
			];

			let formattedDate = null;
			let validDateFound = false;

			for (const format of dateFormats) {
				try {
					const parsedDate = DateTime.fromFormat(row[dateKey], format);
					if (parsedDate.isValid) {
						formattedDate = parsedDate.toFormat('yyyy/MM/dd');
						validDateFound = true;
						break; // Exit the loop once a valid date is found
					}
				} catch (error) {
					// Ignore the error and continue to the next format
				}
			}

			if (!validDateFound) {
				formattedDate = null;
			}
			row[dateKey] = formattedDate;
		}
	});

	// Additional loop to handle empty date strings
	for (let i = 0; i < data.length; i++) {
		const row = data[i].student;
		if (row[dateKey] === '' || row[dateKey] === ' ' || row[dateKey] === null) {
			row[dateKey] = null;
		}
	}

	return data;
}

// Function to flatten the merged data into one flat data row.
function flattenData(studentData, fatherData, motherData) {
	const flatData = {};

	// Student data
	Object.keys(studentData).forEach((key) => {
		flatData[`student.${key}`] = studentData[key];
	});

	// Father data
	Object.keys(fatherData).forEach((key) => {
		flatData[`father.${key}`] = fatherData[key];
	});

	// Mother data
	Object.keys(motherData).forEach((key) => {
		flatData[`mother.${key}`] = motherData[key];
	});
	return attributeOrder.map((attribute) => flatData[attribute]);
}
function flattenDataWithAttendance(studentData, fatherData, motherData, attendanceData) {
	const flatData = {};

	// Student data
	Object.keys(studentData).forEach((key) => {
		flatData[`student.${key}`] = studentData[key];
	});

	// Father data
	Object.keys(fatherData).forEach((key) => {
		flatData[`father.${key}`] = fatherData[key];
	});

	// Mother data
	Object.keys(motherData).forEach((key) => {
		flatData[`mother.${key}`] = motherData[key];
	});

	// Attendance data
	attendanceData.forEach((status, index) => {
		flatData[`attendance.${index}`] = status; // Assuming the index starts from 1
	});

	const finalResult = attributeOrder.map((attribute) => flatData[attribute]);

	return finalResult;
}

// Function to merge student and parent data based on student_id
export function mergeStudentParentData(studentData, parentData) {
	const mergedData = [];

	studentData.forEach((student) => {
		const parents = parentData.filter((parent) => parent.student_id === student.id);

		if (parents.length > 0) {
			mergedData.push([student, parents]);
		}
	});

	return mergedData;
}

function modifyFlattenedData(flatData) {
	// Modify the flatData as needed
	const modifiedData = [...flatData]; // Create a copy of the original flatData

	let sex = flatData[9]; // Assuming 'student.sex' is at index 9
	sex = sex.toUpperCase();

	if (sex === 'NAM') {
		// Replace with 'x', ''
		modifiedData[9] = 'x';
		modifiedData.splice(10, 0, ''); // Insert an empty string at index 10
	} else if (sex === 'NỮ') {
		// Replace with '', 'x'
		modifiedData[9] = '';
		modifiedData.splice(10, 0, 'x'); // Insert 'x' at index 10
	}

	// Add other modifications if needed

	return modifiedData; // Return the modified data
}

export function getAttendanceDatesAndWeekdays(attendance_date) {
	const startDate = new Date(attendance_date);
	const currentDate = startDate.getDate();

	if (currentDate <= 9) {
		startDate.setMonth(startDate.getMonth() - 1); // Set the month to the previous month
		startDate.setDate(9); // Set the date to the 9th of the previous month
	} else {
		startDate.setDate(10); // Set the date to the 10th of the current month
	}

	const endDate = new Date(startDate);
	endDate.setMonth(endDate.getMonth() + 1); // Set the date to the 10th of the next month

	const attendance_dates = [];
	const weekdays = [];

	// Loop through the dates from start date to end date
	while (startDate <= endDate) {
		const year = startDate.getFullYear();
		const month = String(startDate.getMonth() + 1).padStart(2, '0'); // Get the month with leading zeros
		const date = String(startDate.getDate()).padStart(2, '0'); // Get the date with leading zeros
		const formattedDate = `${year}/${month}/${date}`;
		const dayOfWeek = startDate.getDay();

		// Skip Sunday (dayOfWeek === 0)
		if (dayOfWeek !== 0) {
			attendance_dates.push(formattedDate);
			weekdays.push(dayOfWeek);
		}

		startDate.setDate(startDate.getDate() + 1);
	}

	// Convert day of week to Vietnamese format
	const weekdaysVietnamese = weekdays.map((dayOfWeek) => {
		switch (dayOfWeek) {
			case 1:
				return 'T2';
			case 2:
				return 'T3';
			case 3:
				return 'T4';
			case 4:
				return 'T5';
			case 5:
				return 'T6';
			case 6:
				return 'T7';
			default:
				return '';
		}
	});

	return { attendance_dates, weekdaysVietnamese };
}

export function getDatesMonthAndWeekdays() {
	const startDate = new Date();
	const currentDate = startDate.getDate();

	if (currentDate <= 9) {
		startDate.setMonth(startDate.getMonth() - 1); // Set the month to the previous month
		startDate.setDate(9); // Set the date to the 9th of the previous month
	} else {
		startDate.setDate(10); // Set the date to the 10th of the current month
	}

	const endDate = new Date(startDate);
	endDate.setMonth(endDate.getMonth() + 1); // Set the date to the 10th of the next month

	const dates = [];
	const weekdays = [];

	// Loop through the dates from start date to end date
	while (startDate <= endDate) {
		const month = String(startDate.getMonth() + 1).padStart(2, '0'); // Get the month with leading zeros
		const date = String(startDate.getDate()).padStart(2, '0'); // Get the date with leading zeros
		const formattedDate = `${month}/${date}`;
		const dayOfWeek = startDate.getDay();

		// Skip Sunday (dayOfWeek === 0)
		if (dayOfWeek !== 0) {
			dates.push(formattedDate);
			weekdays.push(dayOfWeek);
		}

		startDate.setDate(startDate.getDate() + 1);
	}

	// Convert day of week to Vietnamese format
	const weekdaysVietnamese = weekdays.map((dayOfWeek) => {
		switch (dayOfWeek) {
			case 1:
				return 'T2';
			case 2:
				return 'T3';
			case 3:
				return 'T4';
			case 4:
				return 'T5';
			case 5:
				return 'T6';
			case 6:
				return 'T7';
			default:
				return '';
		}
	});

	return { dates, weekdaysVietnamese };
}
import b64 from './template.xlsx';
export async function writeDataToTemplate(dataItem, attendance_date) {
	try {
		// // Step 1: Load the template file using exceljs.
		// const sheetjs_workbook = read(b64);
		// const buffer = xlsx.write(sheetjs_workbook, { type: 'buffer', bookType: 'xlsx' });
		// // read from a stream

		const buffer = Buffer.from(b64, 'base64');

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(buffer);
		// await workbook.xlsx.readFile(templatePath);
		// Step 2: Get the target worksheet.
		// const worksheet = workbook.Sheets['Sheet1'];
		const worksheet = workbook.getWorksheet('Sheet1');
		let rowIndex = 3; // Start from the 3rd row

		// Step 3: Get the dates and weekdays
		const { attendance_dates, weekdaysVietnamese } = getAttendanceDatesAndWeekdays(attendance_date);

		let rowCount = dataItem.length; // Count the number of students
		// Output the row count in the second row, second column
		const studentCountCell = worksheet.getCell(2, 2);
		studentCountCell.value = `Sĩ số: ${rowCount}`;
		// Step 4: Loop through each element in dataItem and write the student, father, and mother data to the new worksheet.
		for (const data of dataItem) {
			const studentData = data.student;
			const fatherData = data.father[0];
			const motherData = data.mother[0];

			const flatData = flattenData(studentData, fatherData, motherData);
			// Call the function to modify the flattenedData array
			const modifiedData = modifyFlattenedData(flatData);

			// Write the modifiedData to the worksheet starting from the 3rd column (C)
			let colIndex = 0; // Start from column C
			for (const value of modifiedData) {
				const cell = worksheet.getCell(rowIndex, colIndex + 1);
				cell.value = value;
				colIndex++;
			}

			rowIndex++; // Move to the next row for the next data set
		}

		// Step 5: Write the dates and weekdays to the first and second rows of the worksheet starting from column AB (column index 27)
		let colIndex = 28; // Start from column AB (column index 27)
		for (const date of attendance_dates) {
			const day = new Date(date).getDate().toString().padStart(2, '0');
			const cell = worksheet.getCell(1, colIndex + 1);
			cell.value = day;
			cell.alignment = { horizontal: 'center' };
			worksheet.getColumn(colIndex + 1).width = 5; // Set column width to 5 (adjust as needed)
			colIndex++;
		}

		colIndex = 28; // Start from column AB (column index 27)
		for (const weekday of weekdaysVietnamese) {
			const cell = worksheet.getCell(2, colIndex + 1);
			cell.value = weekday;
			cell.alignment = { horizontal: 'center' };
			if (weekday === 'T7') {
				cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Set red fill for T7
			}
			colIndex++;
		}
		const attendancesStartColumnIndex = 28;
		const attendancesEndColumnIndex = attendancesStartColumnIndex + attendance_dates.length - 1;

		// Step 6: Write "Tăng ca /n T(Current month)/(Current year)" to the first row, centered, spanning two columns
		const mergedCell = worksheet.getCell(1, colIndex + 1, 1, colIndex + 2);
		mergedCell.value = `Tăng ca T${(attendance_date.getMonth() + 1)
			.toString()
			.padStart(2, '0')}/${new Date().getFullYear()}`;
		mergedCell.alignment = { horizontal: 'center', vertical: 'middle' };
		mergedCell.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: 'ebb134' } // Use your desired color code
		};
		const gioColIndex = attendancesEndColumnIndex + 1;

		// Step 7: Merge cells for "giờ" and "ăn tối" sub-columns, and set their values
		const hourCell = worksheet.getCell(2, colIndex + 1);
		hourCell.value = 'Giờ';
		hourCell.alignment = { horizontal: 'center' };

		const anToiCollIndex = gioColIndex + 1;
		const dinnerCell = worksheet.getCell(2, colIndex + 2);
		dinnerCell.value = 'Ăn tối';
		dinnerCell.alignment = { horizontal: 'center' };

		worksheet.mergeCells(1, colIndex + 1, 1, colIndex + 2);
		colIndex += 3;
		const phepColIndex = colIndex;

		const phepCell = worksheet.getCell(2, colIndex);
		phepCell.value = 'Phép';
		phepCell.alignment = { horizontal: 'center' };
		phepCell.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: 'FF0000' } // Use your desired color code
		};
		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
			const formula = `=COUNTIF(${excelColumnName(
				attendancesStartColumnIndex
			)}${rowIndex}:${excelColumnName(attendancesEndColumnIndex)}${rowIndex},"P")`;
			const formulaCell = worksheet.getCell(rowIndex, colIndex);
			formulaCell.value = { formula };
		}
		colIndex++;
		// Step 7: Write "CÁC KHOẢN PHẢI THU T05/2023"
		const headerCell = worksheet.getCell(1, colIndex);
		headerCell.value = `CÁC KHOẢN PHẢI THU T${(attendance_date.getMonth() + 1)
			.toString()
			.padStart(2, '0')}/${new Date().getFullYear()}`;
		headerCell.alignment = { horizontal: 'center' };
		worksheet.mergeCells(1, colIndex, 1, colIndex + 9);
		headerCell.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: 'FFA500' } // Use your desired color code
		};

		// Step 8: Write the subheader
		const subheader = [
			'NỢ T03',
			'TC T03',
			'CSVC',
			'ĐP',
			'Học toán',
			'Năng khiếu',
			'A.V',
			'Aerobic',
			'Tiền ăn T04',
			'HPT04'
		];
		const subHeaderStartIndex = colIndex;
		const subHeaderEndIndex = subHeaderStartIndex + subheader.length - 1;
		for (let i = 0; i < subheader.length; i++) {
			const cell = worksheet.getCell(2, colIndex + i);
			cell.value = subheader[i];
			cell.alignment = { horizontal: 'center' };
			cell.fill = {
				type: 'pattern',
				pattern: 'solid',
				fgColor: { argb: '46a642' } // Use your desired color code
			};
			worksheet.getColumn(colIndex + 1 + i).width = 12;
		}
		// TC T03 formula
		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
			const formula = `=${excelColumnName(gioColIndex)}${rowIndex}*10+${excelColumnName(
				anToiCollIndex
			)}${rowIndex}*10`;
			const cell = worksheet.getCell(rowIndex, colIndex + 1);
			cell.value = { formula };
		}
		colIndex += 10;
		const truTienAnIndex = colIndex;
		// Step 9: Create a merged cell spanning two rows and one column for "Trừ tiền ăn"
		const mergedCellTruTienAn = worksheet.getCell(1, colIndex, 2, colIndex + 1);
		mergedCellTruTienAn.value = 'Trừ tiền ăn';
		mergedCellTruTienAn.alignment = { horizontal: 'center', vertical: 'middle' };
		mergedCellTruTienAn.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
		mergedCellTruTienAn.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: 'FFA500' } // Use your desired color code
		};
		worksheet.mergeCells(1, colIndex, 2, colIndex);

		// Adjust the column width
		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed
		// Step 10: Add the formula "=+BH4*30" in the "Trừ tiền ăn" column for each row
		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
			const formula = `=+${excelColumnName(phepColIndex)}${rowIndex}*30`;
			const formulaCell = worksheet.getCell(rowIndex, colIndex);
			formulaCell.value = { formula };
		}
		colIndex++;
		const tongThuThangColIndex = colIndex;
		// Step 12: Write "TỔNG" at the first row
		const tongCell = worksheet.getCell(1, colIndex);
		tongCell.value = 'TỔNG';
		tongCell.alignment = { horizontal: 'center', vertical: 'middle' };
		tongCell.font = { bold: true };
		tongCell.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: '8842a6' } // Use your desired color code
		};
		// Step 13: Write "Thu T(current Month)" at the second row
		const thuCell = worksheet.getCell(2, colIndex);
		thuCell.value = `THU T${(attendance_date.getMonth() + 1).toString().padStart(2, '0')}`;
		thuCell.alignment = { horizontal: 'center', vertical: 'middle' };
		thuCell.font = { bold: true };
		thuCell.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: '8842a6' } // Use your desired color code
		};
		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed

		// Add the formula =SUM(BI4:BR4)-BS4 in the "Tổng / Thu Tháng" column for each row
		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
			const totalFormula = `=SUM(${excelColumnName(
				subHeaderStartIndex
			)}${rowIndex}:${excelColumnName(subHeaderEndIndex)}${rowIndex})-${excelColumnName(
				truTienAnIndex
			)}${rowIndex}`;
			const totalCell = worksheet.getCell(rowIndex, colIndex);
			totalCell.value = { formula: totalFormula };
		}
		// Step 14 Add more
		colIndex++;
		const cellSo1ColIndex = colIndex;
		const mergedCellSo1 = worksheet.getCell(1, colIndex, 2, colIndex + 1);
		mergedCellSo1.value = 'SỐ 2/03';
		mergedCellSo1.alignment = { horizontal: 'center', vertical: 'middle' };
		mergedCellSo1.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
		mergedCellSo1.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: '425da6' } // Use your desired color code
		};
		worksheet.mergeCells(1, colIndex, 2, colIndex);
		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed
		colIndex++;
		const mergedCellSo2 = worksheet.getCell(1, colIndex, 2, colIndex + 1);
		mergedCellSo2.value = 'SỐ 1/04';
		mergedCellSo2.alignment = { horizontal: 'center', vertical: 'middle' };
		mergedCellSo2.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: '425da6' } // Use your desired color code
		};
		worksheet.mergeCells(1, colIndex, 2, colIndex);
		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed
		colIndex++;
		const mergedCellSo3 = worksheet.getCell(1, colIndex, 2, colIndex + 1);
		mergedCellSo3.value = 'SỐ 2/04';
		mergedCellSo3.alignment = { horizontal: 'center', vertical: 'middle' };
		mergedCellSo3.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: '425da6' } // Use your desired color code
		};
		worksheet.mergeCells(1, colIndex, 2, colIndex);
		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed
		const cellSo3ColIndex = colIndex;
		// Step 15: Add "Đã Thu"
		colIndex++;
		const daThuColIndex = colIndex;
		const mergedCellDaThu = worksheet.getCell(1, colIndex, 2, colIndex + 1);
		mergedCellDaThu.value = 'ĐÃ THU';
		mergedCellDaThu.alignment = { horizontal: 'center', vertical: 'middle' };
		mergedCellDaThu.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: 'a64258' } // Use your desired color code
		};
		worksheet.mergeCells(1, colIndex, 2, colIndex);
		worksheet.getColumn(colIndex).width = 20; // Adjust the width as needed

		// Formula for đã thu

		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
			// Generate the formula to calculate the total
			const totalFormula = `=SUM(${excelColumnName(cellSo1ColIndex)}${rowIndex}:${excelColumnName(
				cellSo3ColIndex
			)}${rowIndex})`;

			// Get the cell in which you want to place the formula
			const totalCell = worksheet.getCell(rowIndex, colIndex);

			// Set the formula to the cell
			totalCell.value = { formula: totalFormula };
		}

		// Step 16 Còn Nợ cell
		colIndex++;
		const mergedCellConNo = worksheet.getCell(1, colIndex, 2, colIndex + 1);
		mergedCellConNo.value = 'CÒN NỢ';
		mergedCellConNo.alignment = { horizontal: 'center', vertical: 'middle' };
		mergedCellConNo.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: 'FFA500' } // Use your desired color code
		};
		worksheet.mergeCells(1, colIndex, 2, colIndex);
		worksheet.getColumn(colIndex).width = 20; // Adjust the width as needed

		// Formula for CÒN NỢ

		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
			// Generate the formula to calculate the total
			const formula = `=${excelColumnName(tongThuThangColIndex)}${rowIndex}-${excelColumnName(
				daThuColIndex
			)}${rowIndex}`;
			const cell = worksheet.getCell(rowIndex, colIndex);
			// Set the formula to the cell
			cell.value = { formula: formula };
		}

		// Step 17 Ghi Chú
		colIndex++;
		const mergedCellGhiChu = worksheet.getCell(1, colIndex, 2, colIndex + 1);
		mergedCellGhiChu.value = 'GHI CHÚ';
		mergedCellGhiChu.alignment = { horizontal: 'center', vertical: 'middle' };
		// Set text color
		mergedCellGhiChu.font = { color: { argb: '0000FF' } };
		// Set cell color
		mergedCellGhiChu.fill = {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: '46a642' } // Use your desired color code
		};
		worksheet.mergeCells(1, colIndex, 2, colIndex);
		worksheet.getColumn(colIndex).width = 20; // Adjust the width as needed
		// Step 14: Save the new workbook to a new file.
		// await workbook.xlsx.writeFile('test.csv');
		return workbook;
	} catch (error) {
		console.error('Error writing data to the new file:', error);
		return { error: error, message: 'Lỗi file' };
	}
}
// export async function writeDataToTemplate(dataItem, attendance_date) {
// 	const templatePath = './src/lib/db/Attendance.xlsx';
// 	const newFilePath = './src/lib/Attendance_tenplate.xlsx';

// 	try {
// 		// Step 1: Load the template file using exceljs.
// 		const workbook = new ExcelJS.Workbook();
// 		await workbook.xlsx.readFile(templatePath);

// 		// Step 2: Get the target worksheet.
// 		const worksheet = workbook.getWorksheet('Sheet1'); // Replace 'Sheet1' with the actual sheet name.
// 		let rowIndex = 3; // Start from the 3rd row

// 		// Step 3: Get the dates and weekdays
// 		const { attendance_dates, weekdaysVietnamese } = getAttendanceDatesAndWeekdays(attendance_date);

// 		let rowCount = dataItem.length; // Count the number of students
// 		// Output the row count in the second row, second column
// 		const studentCountCell = worksheet.getCell(2, 2);
// 		studentCountCell.value = `Sĩ số: ${rowCount}`;
// 		// Step 4: Loop through each element in dataItem and write the student, father, and mother data to the new worksheet.
// 		for (const data of dataItem) {
// 			const studentData = data.student;
// 			const fatherData = data.father[0];
// 			const motherData = data.mother[0];

// 			const flatData = flattenData(studentData, fatherData, motherData);
// 			// Call the function to modify the flattenedData array
// 			const modifiedData = modifyFlattenedData(flatData);

// 			// Write the modifiedData to the worksheet starting from the 3rd column (C)
// 			let colIndex = 0; // Start from column C
// 			for (const value of modifiedData) {
// 				const cell = worksheet.getCell(rowIndex, colIndex + 1);
// 				cell.value = value;
// 				colIndex++;
// 			}

// 			rowIndex++; // Move to the next row for the next data set
// 		}

// 		// Step 5: Write the dates and weekdays to the first and second rows of the worksheet starting from column AB (column index 27)
// 		let colIndex = 27; // Start from column AB (column index 27)
// 		for (const date of attendance_dates) {
// 			const day = new Date(date).getDate().toString().padStart(2, '0');
// 			const cell = worksheet.getCell(1, colIndex + 1);
// 			cell.value = day;
// 			cell.alignment = { horizontal: 'center' };
// 			worksheet.getColumn(colIndex + 1).width = 5; // Set column width to 5 (adjust as needed)
// 			colIndex++;
// 		}

// 		colIndex = 27; // Start from column AB (column index 27)
// 		for (const weekday of weekdaysVietnamese) {
// 			const cell = worksheet.getCell(2, colIndex + 1);
// 			cell.value = weekday;
// 			cell.alignment = { horizontal: 'center' };
// 			if (weekday === 'T7') {
// 				cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Set red fill for T7
// 			}
// 			colIndex++;
// 		}
// 		const attendancesStartColumnIndex = 28;
// 		const attendancesEndColumnIndex = attendancesStartColumnIndex + attendance_dates.length - 1;

// 		// Step 6: Write "Tăng ca /n T(Current month)/(Current year)" to the first row, centered, spanning two columns
// 		const mergedCell = worksheet.getCell(1, colIndex + 1, 1, colIndex + 2);
// 		mergedCell.value = `Tăng ca T${(new Date().getMonth() + 1)
// 			.toString()
// 			.padStart(2, '0')}/${new Date().getFullYear()}`;
// 		mergedCell.alignment = { horizontal: 'center', vertical: 'middle' };
// 		mergedCell.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: 'ebb134' } // Use your desired color code
// 		};
// 		const gioColIndex = attendancesEndColumnIndex + 1;

// 		// Step 7: Merge cells for "giờ" and "ăn tối" sub-columns, and set their values
// 		const hourCell = worksheet.getCell(2, colIndex + 1);
// 		hourCell.value = 'Giờ';
// 		hourCell.alignment = { horizontal: 'center' };

// 		const anToiCollIndex = gioColIndex + 1;
// 		const dinnerCell = worksheet.getCell(2, colIndex + 2);
// 		dinnerCell.value = 'Ăn tối';
// 		dinnerCell.alignment = { horizontal: 'center' };

// 		worksheet.mergeCells(1, colIndex + 1, 1, colIndex + 2);
// 		colIndex += 3;
// 		const phepColIndex = colIndex;

// 		const phepCell = worksheet.getCell(2, colIndex);
// 		phepCell.value = 'Phép';
// 		phepCell.alignment = { horizontal: 'center' };
// 		phepCell.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: 'FF0000' } // Use your desired color code
// 		};
// 		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
// 			const formula = `=COUNTIF(${excelColumnName(
// 				attendancesStartColumnIndex
// 			)}${rowIndex}:${excelColumnName(attendancesEndColumnIndex)}${rowIndex},"P")`;
// 			const formulaCell = worksheet.getCell(rowIndex, colIndex);
// 			formulaCell.value = { formula };
// 		}
// 		colIndex++;
// 		// Step 7: Write "CÁC KHOẢN PHẢI THU T05/2023"
// 		const headerCell = worksheet.getCell(1, colIndex);
// 		headerCell.value = `CÁC KHOẢN PHẢI THU T${(new Date().getMonth() + 1)
// 			.toString()
// 			.padStart(2, '0')}/${new Date().getFullYear()}`;
// 		headerCell.alignment = { horizontal: 'center' };
// 		worksheet.mergeCells(1, colIndex, 1, colIndex + 9);
// 		headerCell.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: 'FFA500' } // Use your desired color code
// 		};

// 		// Step 8: Write the subheader
// 		const subheader = [
// 			'NỢ T03',
// 			'TC T03',
// 			'CSVC',
// 			'ĐP',
// 			'Học toán',
// 			'Năng khiếu',
// 			'A.V',
// 			'Aerobic',
// 			'Tiền ăn T04',
// 			'HPT04'
// 		];
// 		const subHeaderStartIndex = colIndex;
// 		const subHeaderEndIndex = subHeaderStartIndex + subheader.length - 1;
// 		for (let i = 0; i < subheader.length; i++) {
// 			const cell = worksheet.getCell(2, colIndex + i);
// 			cell.value = subheader[i];
// 			cell.alignment = { horizontal: 'center' };
// 			cell.fill = {
// 				type: 'pattern',
// 				pattern: 'solid',
// 				fgColor: { argb: '46a642' } // Use your desired color code
// 			};
// 			worksheet.getColumn(colIndex + 1 + i).width = 12;
// 		}
// 		// TC T03 formula
// 		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
// 			const formula = `=${excelColumnName(gioColIndex)}${rowIndex}*10+${excelColumnName(
// 				anToiCollIndex
// 			)}${rowIndex}*10`;
// 			const cell = worksheet.getCell(rowIndex, colIndex + 1);
// 			cell.value = { formula };
// 		}
// 		colIndex += 10;
// 		const truTienAnIndex = colIndex;
// 		// Step 9: Create a merged cell spanning two rows and one column for "Trừ tiền ăn"
// 		const mergedCellTruTienAn = worksheet.getCell(1, colIndex, 2, colIndex + 1);
// 		mergedCellTruTienAn.value = 'Trừ tiền ăn';
// 		mergedCellTruTienAn.alignment = { horizontal: 'center', vertical: 'middle' };
// 		mergedCellTruTienAn.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
// 		mergedCellTruTienAn.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: 'FFA500' } // Use your desired color code
// 		};
// 		worksheet.mergeCells(1, colIndex, 2, colIndex);

// 		// Adjust the column width
// 		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed
// 		// Step 10: Add the formula "=+BH4*30" in the "Trừ tiền ăn" column for each row
// 		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
// 			const formula = `=+${excelColumnName(phepColIndex)}${rowIndex}*30`;
// 			const formulaCell = worksheet.getCell(rowIndex, colIndex);
// 			formulaCell.value = { formula };
// 		}
// 		colIndex++;
// 		const tongThuThangColIndex = colIndex;
// 		// Step 12: Write "TỔNG" at the first row
// 		const tongCell = worksheet.getCell(1, colIndex);
// 		tongCell.value = 'TỔNG';
// 		tongCell.alignment = { horizontal: 'center', vertical: 'middle' };
// 		tongCell.font = { bold: true };
// 		tongCell.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: '8842a6' } // Use your desired color code
// 		};
// 		// Step 13: Write "Thu T(current Month)" at the second row
// 		const thuCell = worksheet.getCell(2, colIndex);
// 		thuCell.value = `THU T${(new Date().getMonth() + 1).toString().padStart(2, '0')}`;
// 		thuCell.alignment = { horizontal: 'center', vertical: 'middle' };
// 		thuCell.font = { bold: true };
// 		thuCell.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: '8842a6' } // Use your desired color code
// 		};
// 		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed

// 		// Add the formula =SUM(BI4:BR4)-BS4 in the "Tổng / Thu Tháng" column for each row
// 		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
// 			const totalFormula = `=SUM(${excelColumnName(
// 				subHeaderStartIndex
// 			)}${rowIndex}:${excelColumnName(subHeaderEndIndex)}${rowIndex})-${excelColumnName(
// 				truTienAnIndex
// 			)}${rowIndex}`;
// 			const totalCell = worksheet.getCell(rowIndex, colIndex);
// 			totalCell.value = { formula: totalFormula };
// 		}
// 		// Step 14 Add more
// 		colIndex++;
// 		const cellSo1ColIndex = colIndex;
// 		const mergedCellSo1 = worksheet.getCell(1, colIndex, 2, colIndex + 1);
// 		mergedCellSo1.value = 'SỐ 2/03';
// 		mergedCellSo1.alignment = { horizontal: 'center', vertical: 'middle' };
// 		mergedCellSo1.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
// 		mergedCellSo1.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: '425da6' } // Use your desired color code
// 		};
// 		worksheet.mergeCells(1, colIndex, 2, colIndex);
// 		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed
// 		colIndex++;
// 		const mergedCellSo2 = worksheet.getCell(1, colIndex, 2, colIndex + 1);
// 		mergedCellSo2.value = 'SỐ 1/04';
// 		mergedCellSo2.alignment = { horizontal: 'center', vertical: 'middle' };
// 		mergedCellSo2.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: '425da6' } // Use your desired color code
// 		};
// 		worksheet.mergeCells(1, colIndex, 2, colIndex);
// 		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed
// 		colIndex++;
// 		const mergedCellSo3 = worksheet.getCell(1, colIndex, 2, colIndex + 1);
// 		mergedCellSo3.value = 'SỐ 2/04';
// 		mergedCellSo3.alignment = { horizontal: 'center', vertical: 'middle' };
// 		mergedCellSo3.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: '425da6' } // Use your desired color code
// 		};
// 		worksheet.mergeCells(1, colIndex, 2, colIndex);
// 		worksheet.getColumn(colIndex).width = 15; // Adjust the width as needed
// 		const cellSo3ColIndex = colIndex;
// 		// Step 15: Add "Đã Thu"
// 		colIndex++;
// 		const daThuColIndex = colIndex;
// 		const mergedCellDaThu = worksheet.getCell(1, colIndex, 2, colIndex + 1);
// 		mergedCellDaThu.value = 'ĐÃ THU';
// 		mergedCellDaThu.alignment = { horizontal: 'center', vertical: 'middle' };
// 		mergedCellDaThu.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: 'a64258' } // Use your desired color code
// 		};
// 		worksheet.mergeCells(1, colIndex, 2, colIndex);
// 		worksheet.getColumn(colIndex).width = 20; // Adjust the width as needed

// 		// Formula for đã thu

// 		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
// 			// Generate the formula to calculate the total
// 			const totalFormula = `=SUM(${excelColumnName(cellSo1ColIndex)}${rowIndex}:${excelColumnName(
// 				cellSo3ColIndex
// 			)}${rowIndex})`;

// 			// Get the cell in which you want to place the formula
// 			const totalCell = worksheet.getCell(rowIndex, colIndex);

// 			// Set the formula to the cell
// 			totalCell.value = { formula: totalFormula };
// 		}

// 		// Step 16 Còn Nợ cell
// 		colIndex++;
// 		const mergedCellConNo = worksheet.getCell(1, colIndex, 2, colIndex + 1);
// 		mergedCellConNo.value = 'CÒN NỢ';
// 		mergedCellConNo.alignment = { horizontal: 'center', vertical: 'middle' };
// 		mergedCellConNo.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: 'FFA500' } // Use your desired color code
// 		};
// 		worksheet.mergeCells(1, colIndex, 2, colIndex);
// 		worksheet.getColumn(colIndex).width = 20; // Adjust the width as needed

// 		// Formula for CÒN NỢ

// 		for (let rowIndex = 3; rowIndex <= rowCount; rowIndex++) {
// 			// Generate the formula to calculate the total
// 			const formula = `=${excelColumnName(tongThuThangColIndex)}${rowIndex}-${excelColumnName(
// 				daThuColIndex
// 			)}${rowIndex}`;
// 			const cell = worksheet.getCell(rowIndex, colIndex);
// 			// Set the formula to the cell
// 			cell.value = { formula: formula };
// 		}

// 		// Step 17 Ghi Chú
// 		colIndex++;
// 		const mergedCellGhiChu = worksheet.getCell(1, colIndex, 2, colIndex + 1);
// 		mergedCellGhiChu.value = 'GHI CHÚ';
// 		mergedCellGhiChu.alignment = { horizontal: 'center', vertical: 'middle' };
// 		// Set text color
// 		mergedCellGhiChu.font = { color: { argb: '0000FF' } };
// 		// Set cell color
// 		mergedCellGhiChu.fill = {
// 			type: 'pattern',
// 			pattern: 'solid',
// 			fgColor: { argb: '46a642' } // Use your desired color code
// 		};
// 		worksheet.mergeCells(1, colIndex, 2, colIndex);
// 		worksheet.getColumn(colIndex).width = 20; // Adjust the width as needed
// 		// Step 14: Save the new workbook to a new file.
// 		// await workbook.xlsx.writeFile(newFilePath);
// 		// console.log('Data written to the new file:', newFilePath);
// 		return workbook;
// 	} catch (error) {
// 		console.error('Error writing data to the new file:', error);
// 	}
// }
function excelColumnName(index) {
	let columnName = '';
	while (index > 0) {
		let remainder = (index - 1) % 26;
		columnName = String.fromCharCode(65 + remainder) + columnName;
		index = Math.floor((index - 1) / 26);
	}
	return columnName;
}
// export async function writeAttendance(dataItem, wanted_date, attendance_dates, weekdaysVietnamese) {
// 	const templatePath = './src/lib/Attendance.xlsx';
// 	const newFilePath = './src/lib/NewAttendance.xlsx';
// 	try {
// 		let workbook = await writeDataToTemplate(dataItem, wanted_date);

// 		// Step 2: Get the target worksheet.
// 		const worksheet = workbook.getWorksheet('Sheet1'); // Replace 'Sheet1' with the actual sheet name.
// 		let rowIndex = 3; // Start from the 3rd row

// 		// Step 4: Get the dates and weekdays
// 		const number_of_date = attendance_dates.length;
// 		// Step 3: Loop through each element in dataItem and write the student, father, and mother data to the new worksheet.
// 		for (const data of dataItem) {
// 			// const studentData = data.student;
// 			// const fatherData = data.father[0];
// 			// const motherData = data.mother[0];

// 			// const flatData = flattenData(studentData, fatherData, motherData);
// 			// Call the function to modify the flattenedData array
// 			// let modifiedData = modifyFlattenedData(flatData);
// 			// modifiedData = modifiedData.concat(attendanceData);
// 			const attendanceData = data.attendance_status;
// 			// Write the modifiedData to the worksheet starting from the 3rd column (C)
// 			let colIndex = 28; // Start from column AB (column index 27)
// 			for (const value of attendanceData) {
// 				const cell = worksheet.getCell(rowIndex, colIndex);
// 				cell.value = value;
// 				colIndex++;
// 			}
// 			rowIndex++; // Move to the next row for the next data set
// 		}

// 		// Step 8: Save the new workbook to a new file.
// 		await workbook.xlsx.writeFile(newFilePath);
// 		console.log('Data written to the new file:', newFilePath);
// 		return workbook;
// 	} catch (error) {
// 		console.error('Error writing data to the new file:', error);
// 	}
// }

export async function writeAttendance(dataItem, wanted_date) {
	try {
		let workbook = await writeDataToTemplate(dataItem, wanted_date);

		// Step 2: Get the target worksheet.
		const worksheet = workbook.getWorksheet('Sheet1'); // Replace 'Sheet1' with the actual sheet name.
		let rowIndex = 3; // Start from the 3rd row

		// Step 3: Loop through each element in dataItem and write the student, father, and mother data to the new worksheet.
		for (const data of dataItem) {
			const attendanceData = data.attendance_status;
			// Write the modifiedData to the worksheet starting from the 3rd column (C)
			let colIndex = 28; // Start from column AB (column index 27)
			for (const value of attendanceData) {
				const cell = worksheet.getCell(rowIndex, colIndex);
				cell.value = value;
				colIndex++;
			}
			rowIndex++; // Move to the next row for the next data set
		}
		return workbook;
	} catch (error) {
		console.error('Error writing data to the new file:', error);
	}
}
export async function getStudentAndParentData(class_room_id) {
	try {
		// Step 1: Fetch all students from the specified class room of the branch
		const { data: students, error: studentError } = await supabase
			.from('student')
			.select('*')
			.eq('class_room_id', class_room_id);

		if (studentError) {
			throw new Error(`Error fetching students: ${studentError.message}`);
		}

		// Step 2: Get parent data for each student using their ID
		const studentAndParentData = await Promise.all(
			students.map(async (student) => {
				try {
					// Fetch all parents associated with the student ID
					const { data: parents, error: parentError } = await supabase
						.from('parent')
						.select('*')
						.eq('student_id', student.id);

					if (parentError) {
						console.error(
							`Error fetching parents for student ID ${student.id}: ${parentError.message}`
						);
						return {
							student: student,
							parents: [] // Empty array for parents if an error occurs
						};
					}

					// Map the student and their parents
					return {
						student: student,
						parents: parents
					};
				} catch (error) {
					console.error(`Error fetching parents for student ID ${student.id}: ${error.message}`);
					return {
						student: student,
						parents: [] // Empty array for parents if an error occurs
					};
				}
			})
		);

		return studentAndParentData;
	} catch (error) {
		console.error('Error getting student and parent data:', error);
		return [];
	}
}
// Function to remove duplicates based on the specified rules
function removeDuplicateStudents(data) {
	const students = data.map((item) => item.student); // Extracting students from the data
	const duplicates = findDuplicates(students);
	const result = [];

	data.forEach((item, index) => {
		if (!duplicates.some((indexes) => indexes.includes(index))) {
			result.push(item);
		}
	});

	duplicates.forEach((duplicateIndexes) => {
		const duplicateStudents = duplicateIndexes.map((index) => students[index]);
		const hasNullDob = duplicateStudents.some((student) => !student.dob);

		if (hasNullDob) {
			// Insert all duplicate records as separate students
			duplicateIndexes.forEach((index) => {
				const { student, parents } = data[index];
				result.push({ student, parents });
			});
		} else {
			// Insert the first record and ignore the subsequent duplicates
			const [firstDuplicateIndex, ...otherDuplicateIndexes] = duplicateIndexes;
			result.push(data[firstDuplicateIndex]);

			otherDuplicateIndexes.forEach((index) => {
				const { parents } = data[index];
				result.push({ student: students[index], parents });
			});
		}
	});

	return result;
}

async function separateSyncAndNotSyncData(excelData, databaseData) {
	const syncData = [];
	const notSyncData = [];

	// Find synchronized data and populate syncData array
	excelData.forEach((excelEntry) => {
		const matchingDatabaseStudent = databaseData.find(
			(dbEntry) =>
				dbEntry.student.first_name === excelEntry.student.first_name &&
				dbEntry.student.last_name === excelEntry.student.last_name &&
				dbEntry.student.dob === excelEntry.student.dob
		);

		if (matchingDatabaseStudent) {
			const student_id = matchingDatabaseStudent.student.id;

			const mergedStudent = {
				student: {
					...excelEntry.student,
					id: student_id
				},
				parents: excelEntry.parents.map((parent) => {
					const matchingParent = matchingDatabaseStudent.parents.find(
						(dbParent) => dbParent.sex === parent.sex
					);

					if (matchingParent) {
						return {
							...parent,
							id: matchingParent.id,
							student_id: student_id
						};
					} else {
						// Handle case where matching parent is not found
						return parent;
					}
				})
			};

			syncData.push(mergedStudent);
		} else {
			notSyncData.push(excelEntry);
		}
	});

	// Find unsynchronized data from the database and add to notSyncData array
	databaseData.forEach((dbEntry) => {
		const matchingExcelStudent = excelData.find(
			(excelEntry) =>
				excelEntry.student.first_name === dbEntry.student.first_name &&
				excelEntry.student.last_name === dbEntry.student.last_name &&
				excelEntry.student.dob === dbEntry.student.dob
		);

		if (!matchingExcelStudent) {
			const student_id = dbEntry.student.id;

			const mergedStudent = {
				student: {
					...dbEntry.student,
					id: student_id
				},
				parents: dbEntry.parents.map((parent) => {
					return {
						...parent,
						student_id: student_id
					};
				})
			};

			notSyncData.push(mergedStudent);
		}
	});

	return { syncData, notSyncData };
}
async function processSyncAndNotSyncData(excelData, databaseData) {
	const { syncData, notSyncData } = await separateSyncAndNotSyncData(excelData, databaseData);

	// Filter out entries from notSyncData that have matching student data in syncData
	const finalNotSyncData = notSyncData.filter((notSyncEntry) => {
		const isDuplicate = syncData.some(
			(syncEntry) =>
				syncEntry.student.first_name === notSyncEntry.student.first_name &&
				syncEntry.student.last_name === notSyncEntry.student.last_name &&
				syncEntry.student.dob === notSyncEntry.student.dob
		);

		if (isDuplicate) {
			// Apply the duplicate handling rules
			if (notSyncEntry.student.dob !== null && notSyncEntry.student.dob !== undefined) {
				// If the notSyncEntry has a valid dob, insert it as a separate student
				return true;
			} else {
				// If the notSyncEntry has a null dob, ignore it
				return false;
			}
		}

		return true; // Keep entries that are not duplicates
	});

	return { syncData, notSyncData: finalNotSyncData };
}

function assignIdsToParents(data) {
	const newData = { ...data };

	const studentId = newData.student.id; // Get the student's ID

	newData.parents.forEach((parent, index) => {
		parent.id = index + 1; // Assign IDs based on the index (+1 to avoid 0)
		parent.student_id = studentId; // Link parent to student
	});

	return newData;
}

function removeDuplicatesByKey(arr, keyFn) {
	const seen = new Set();
	return arr.filter((item) => {
		const key = keyFn(item);
		return seen.has(key) ? false : seen.add(key);
	});
}
// Helper function to merge student data and parent data
function mergeData(transformedData, dbStudentParentData) {
	const finalData = [];
	const studentIdMap = new Map();

	// Remove duplicates from transformedData based on first_name, last_name, and dob
	transformedData = removeDuplicatesByKey(transformedData, (item) => {
		const { first_name, last_name, dob } = item.student;
		return `${first_name}-${last_name}-${dob}`;
	});

	// Process dbStudentParentData
	dbStudentParentData.forEach((item) => {
		const { student, parents } = item;
		const key = `${student.first_name}-${student.last_name}-${student.dob}`;

		if (!studentIdMap.has(key)) {
			// If the student doesn't exist in the map, add it to the map with its ID
			const insertedIndex = finalData.push(item) - 1;
			studentIdMap.set(key, insertedIndex);
		} else {
			// If the student already exists in the map, update the parent data (if available)
			const index = studentIdMap.get(key);
			const existingParents = finalData[index].parents;

			if (parents) {
				// Update the parent data with student ID
				const updatedParents = parents.map((parent) => {
					return { ...parent, student_id: finalData[index].student.id };
				});

				// Ensure both transformedData and dbStudentParentData have parent data
				const mergedParents = [
					...(updatedParents || existingParents || [null]), // First parent from dbStudentParentData (or null if not available)
					...(existingParents || updatedParents || [null]) // Second parent from transformedData (or null if not available)
				];

				// Update the parents data in the finalData
				finalData[index].parents = mergedParents;
			}
		}
	});

	// Process transformedData
	transformedData.forEach((item) => {
		const { student, parents } = item;
		const key = `${student.first_name}-${student.last_name}-${student.dob}`;

		if (!studentIdMap.has(key)) {
			// If the student doesn't exist in the map, add it to the finalData with its ID
			finalData.push(item);
			studentIdMap.set(key, finalData.length - 1);
		} else {
			// If the student already exists in the map, update the student_id and parent data (if available)
			const index = studentIdMap.get(key);
			const existingItem = finalData[index];

			// If the existing item doesn't have a student_id but the new item does, update the student_id
			if (!existingItem.student.id && item.student.id) {
				existingItem.student.id = item.student.id;
			}

			if (parents) {
				// Update the parent data with student ID
				const updatedParents = parents.map((parent) => {
					return { ...parent, student_id: existingItem.student.id };
				});

				// Ensure both transformedData and dbStudentParentData have parent data
				const mergedParents = [
					...(updatedParents || existingItem.parents || [null]), // First parent from transformedData (or null if not available)
					...(existingItem.parents || updatedParents || [null]) // Second parent from dbStudentParentData (or null if not available)
				];

				// Update the parents data in the finalData
				existingItem.parents = mergedParents;
			}
		}
	});

	return finalData;
}

function areDatesEqual(date1, date2) {
	if (!date1 || !date2) {
		return false;
	}
	const date1Obj = new Date(date1);
	const date2Obj = new Date(date2);
	return date1Obj.getTime() === date2Obj.getTime();
}

// Helper function to find duplicates in an array of student objects
function findDuplicates(students) {
	const duplicates = new Map();
	students.forEach((student, index) => {
		const key = `${student.first_name}-${student.last_name}`;
		if (duplicates.has(key)) {
			duplicates.get(key).push(index);
		} else {
			duplicates.set(key, [index]);
		}
	});
	return Array.from(duplicates.values()).filter((indexes) => indexes.length > 1);
}

// Insertion or Update Logic:

// For each record in the provided data array:
// If student_id is null/undefined, insert new student and retrieve the generated student_id.
// If insert fails, query and get student_id based on name and dob.
// If student_id is available, update existing student's data.
// Iterate through parents, updating or inserting based on conditions.
// Exception Handling:

// If any insert/update operation fails, catch the error.
// If inserting a new student fails due to uniqueness, query and recover student_id.
// If errors occur during parent processing, log and continue.
async function pushToDatabase(data) {
	try {
		for (const item of data) {
			const studentData = item.student;
			let student_id = studentData.id;

			// Step 1: Insert or update the student
			if (student_id === null || student_id === undefined) {
				// If student_id is not present, it means the student is new and needs to be inserted.
				// Insert the student and get the newly generated student_id.
				try {
					const { data: newStudentData, error: newStudentError } = await supabase
						.from('student')
						.insert({
							...studentData
						})
						.single();

					if (newStudentError) {
						throw newStudentError;
					}
					if (studentData.dob != null || studentData.dob != undefined) {
						// Query for the student to get the student_id
						const { data: newStudentQueryData, error: newStudentQueryError } = await supabase
							.from('student')
							.select('id')
							.eq('first_name', studentData.first_name) // Adjust the query criteria as needed
							.eq('last_name', studentData.last_name)
							.eq('dob', studentData.dob)
							.single();

						if (newStudentQueryError) {
							throw newStudentQueryError;
						}
						student_id = newStudentQueryData.id;
					} else {
						// Query for the student to get the student_id
						const { data: newStudentQueryData, error: newStudentQueryError } = await supabase
							.from('student')
							.select('id')
							.eq('first_name', studentData.first_name) // Adjust the query criteria as needed
							.eq('last_name', studentData.last_name)
							.single();

						if (newStudentQueryError) {
							throw newStudentQueryError;
						}
						student_id = newStudentQueryData.id;
					}
				} catch (error) {
					console.error('Error inserting student:', error);
					// Query for the student to get the student_id
					const { data: newStudentQueryData, error: newStudentQueryError } = await supabase
						.from('student')
						.select('id')
						.eq('first_name', studentData.first_name) // Adjust the query criteria as needed
						.eq('last_name', studentData.last_name)
						.single();

					if (newStudentQueryError) {
						throw newStudentQueryError;
					}

					student_id = newStudentQueryData.id;
					continue;
				}
			} else {
				// If student_id is present, it means the student already exists and needs to be updated.
				// Update the student with the provided student_id.
				try {
					const { error: updateStudentError } = await supabase
						.from('student')
						.update({
							grade: studentData.grade,
							enroll_date: studentData.enroll_date,
							dob: studentData.dob,
							birth_year: studentData.birth_year,
							ethnic: studentData.ethnic,
							birth_place: studentData.birth_place,
							temp_res: studentData.temp_res,
							perm_res_province: studentData.perm_res_province,
							perm_res_district: studentData.perm_res_district,
							perm_res_commune: studentData.perm_res_commune
						})
						.eq('id', student_id)
						.single();

					if (updateStudentError) {
						throw updateStudentError;
					}
				} catch (error) {
					console.error('Error updating student:', error);
					continue;
				}
			}

			// Step 2: Insert or update parents with the obtained student_id
			for (const parent of item.parents) {
				// Insert or update parent with the obtained student_id
				let parentStudentId = parent.student_id;

				if (parentStudentId === null || parentStudentId === undefined) {
					parentStudentId = student_id;
					parent.student_id = parentStudentId;
				}

				if (parent.sex != null || parent.sex != undefined) {
					// If parent's name is not present, check if the parent exists based on name and student_id
					const { data: existingParentData } = await supabase
						.from('parent')
						.select('id')
						.eq('student_id', parentStudentId)
						.eq('sex', parent.sex)
						.single();

					if (existingParentData != null || existingParentData != undefined) {
						// If parent with the student_id and sex exists, update the parent with the existingParentData.id
						try {
							const { error: updateParentError } = await supabase
								.from('parent')
								.update(parent)
								.eq('id', existingParentData.id)
								.single();

							if (updateParentError) {
								throw updateParentError;
							}
						} catch (error) {
							console.error('Error updating parent:', error);
						}
					} else {
						// If parent with the student_id and sex does not exist, insert the parent with the student_id
						try {
							const { error: parentInsertError } = await supabase.from('parent').insert({
								student_id: parentStudentId,
								name: parent.name,
								dob: parent.dob,
								sex: parent.sex,
								phone_number: parent.phone_number,
								zalo: parent.zalo,
								occupation: parent.occupation,
								landlord: parent.landlord,
								roi: parent.roi,
								birthplace: parent.birthplace,
								res_registration: parent.res_registration
							});

							if (parentInsertError) {
								throw parentInsertError;
							}
						} catch (error) {
							console.error('Error inserting parent:', error);
						}
					}
				} else {
					// If parent's name is present, it means the parent already exists and needs to be updated.
					// Update the parent with the provided parent.id
					try {
						const { error: updateParentError } = await supabase
							.from('parent')
							.update({ ...parent })
							.eq('id', parent.id)
							.single();

						if (updateParentError) {
							throw updateParentError;
						}
					} catch (error) {
						console.error('Error updating parent:', error);
					}
				}
			}
		}
		console.log('Data pushed to the database successfully.');
	} catch (error) {
		console.error('Error pushing data to the database:', error);
	}
}

// Function to get the current date in the format yyyy-mm-dd
export const getCurrentDate = () => {
	const currentDate = new Date();
	const year = currentDate.getFullYear();
	const month = String(currentDate.getMonth() + 1).padStart(2, '0');
	const day = String(currentDate.getDate()).padStart(2, '0');
	return `${year}-${month}-${day}`;
};

// Function to calculate start_date and end_date based on current date
export const calculateDates = (currentDate) => {
	const ninthDayOfCurrentMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 9);

	let start_date, end_date;

	if (currentDate <= ninthDayOfCurrentMonth) {
		// If current date is before or on the 9th of the current month
		// start_date = 10th of the previous month, end_date = 9th of the current month
		const startDateOfPreviousMonth = new Date(
			currentDate.getFullYear(),
			currentDate.getMonth() - 1,
			10
		);
		const endDateOfCurrentMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 9);
		start_date = formatDate(startDateOfPreviousMonth);
		end_date = formatDate(endDateOfCurrentMonth);
	} else {
		// If current date is after the 9th of the current month
		// start_date = 10th of the current month, end_date = 9th of the next month
		const startDateOfCurrentMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 10);
		const endDateOfNextMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 9);
		start_date = formatDate(startDateOfCurrentMonth);
		end_date = formatDate(endDateOfNextMonth);
	}

	return { start_date, end_date };
};

// Function to format a date as yyyy-mm-dd
const formatDate = (date) => {
	const year = date.getFullYear();
	const month = String(date.getMonth() + 1).padStart(2, '0');
	const day = String(date.getDate()).padStart(2, '0');
	return `${year}-${month}-${day}`;
};
