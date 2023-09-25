import fs from 'fs';
import {
	readExcelFile,
	getRowValues,
	checkHeaderFile,
	readRowsAndConvertToJson,
	getCurrentDate,
	calculateDates,
	insertStudentAndParent,
	mergeStudentParentData,
	writeDataToTemplate,
	writeAttendance,
	getAttendanceDatesAndWeekdays,
	getDatesMonthAndWeekdays,
	readDataAndAttendance
	// readExcelFileFromBuffer
} from './util.js';
import { insertClassRoom } from './class_room.js';
import { getParentsByStudentId } from './parent.js';
import { insertAttendance, getCurrentAttendance } from './attendance.js';
import {
	insertAttendanceEvent,
	getAllAttendanceEventsByAttendanceId,
	updateOrInsertAttendanceEvent,
	deleteAttendanceEvent
} from './attendance_event.js';
import { getStudentsFromClassroom, selectIdFromStudentData } from './student.js';

// Export Template file
export function exportTemplateStudentList(class_name) {
	let originalFileName = 'src/lib/Danh_sach_hoc_sinh.xlsx';
	let newFileName = `src/lib/Danh_sach_lop_${class_name}.xlsx`;

	let readStream = fs.createReadStream(originalFileName);
	let writeStream = fs.createWriteStream(newFileName);

	readStream.pipe(writeStream);

	readStream.on('error', (err) => {
		console.error('Error reading the original file:', err);
	});

	writeStream.on('error', (err) => {
		console.error('Error creating the new file:', err);
	});

	writeStream.on('finish', () => {
		console.log('File renamed successfully.');
	});
}

// export async function importStudentListFile(filePath, sheetName, class_room_id, branch_id) {
// 	let worksheet = readExcelFile(filePath, sheetName);

// 	// Define the expected main and sub header rows
// 	let expectedMainHeader = [
// 		'LỚP',
// 		'HỌ VÀ TÊN LÓT',
// 		'TÊN HỌC SINH',
// 		'NGÀY NHẬP HỌC',
// 		'SINH NGÀY',
// 		'SỐ ĐIỆN THOẠI',
// 		null,
// 		'ZALO',
// 		'NĂM SINH',
// 		'GIỚI TÍNH',
// 		null,
// 		'DÂN TỘC',
// 		'NƠI SINH HS',
// 		'THÔNG TIN CHA',
// 		null,
// 		null,
// 		'THÔNG TIN MẸ',
// 		null,
// 		null,
// 		'THƯỜNG TRÚ',
// 		null,
// 		null,
// 		'TẠM TRÚ',
// 		'CHỦ NHÀ TRỌ',
// 		'BC PHỔ CẬP',
// 		'KHAI SINH',
// 		'HỘ KHẨU'
// 	];

// 	let expectedSubHeader = [
// 		null,
// 		null,
// 		null,
// 		null,
// 		null,
// 		'ĐT1',
// 		'ĐT2',
// 		null,
// 		null,
// 		'NAM',
// 		'NỮ',
// 		null,
// 		null,
// 		'HỌ VÀ TÊN',
// 		'NĂM SINH',
// 		'NGHỀ NGHIỆP',
// 		'HỌ VÀ TÊN',
// 		'NĂM SINH',
// 		'NGHỀ NGHIỆP',
// 		'TỈNH',
// 		'HUYỆN',
// 		'XÃ',
// 		null,
// 		null,
// 		null,
// 		null,
// 		null
// 	];

// 	let acceptableHeaders = {
// 		'HỌ VÀ TÊN LÓT': ['HỌ VÀ TÊN', 'HỌ'],
// 		'TÊN HỌC SINH': ['TÊN', 'TÊN RIÊNG', 'TÊN HS'],
// 		'NGÀY NHẬP HỌC': ['NHẬP HỌC'],
// 		'SINH NGÀY': ['SINH NHẬT', 'NGÀY SINH'],
// 		'SỐ ĐIỆN THOẠI': ['SDT', 'SĐT', 'SỐ ĐT', 'SỐ DT'],
// 		DT1: ['ĐT1', 'SĐT1', 'SDT1'],
// 		DT2: ['ĐT2', 'SĐT2', 'SDT2']
// 	};

// 	let mainHeaderRanges = [
// 		'A1:E1', // Range before "Số điện thoại"
// 		'F1:G1', // Range of "Số điện thoại" merged cells
// 		'H1:I1', // Range after "Số điện thoại" merged cells and before "GIỚI TÍNH"
// 		'J1:K1', // Range of "GIỚI TÍNH" merged cells
// 		'L1:M1', // Range after "GIỚI TÍNH" merged cells and before "THÔNG TIN CHA"
// 		'N1:P1', // Range of "THÔNG TIN CHA" merged cells
// 		'Q1:S1', // Range of "THÔNG TIN MẸ" merged cells
// 		'T1:V1' // Range of "THƯỜNG TRÚ" merged cells
// 	];
// 	let mainHeaderRange = 'A1:AA1';
// 	let subHeaderRange = 'A2:AA2';

// 	let headers = getRowValues(worksheet, mainHeaderRange);
// 	let subheaders = getRowValues(worksheet, subHeaderRange);

// 	let checkHeadersResult = checkHeaderFile(
// 		worksheet,
// 		mainHeaderRanges,
// 		subHeaderRange,
// 		expectedMainHeader,
// 		expectedSubHeader,
// 		acceptableHeaders
// 	);
// 	if (!checkHeadersResult) {
// 		console.error('Header check failed.');
// 		return;
// 	} else {
// 		headers = expectedMainHeader;
// 		subHeaderRange = expectedSubHeader;
// 	}
// 	let data = readRowsAndConvertToJson(worksheet, headers, subheaders);
// 	// const studentData = transformDataKeys(data, attributeMappings.student);

// 	await insertStudentAndParent(data, class_room_id, branch_id);
// }
export async function importStudentListFile(worksheet, class_room_id, branch_id) {
	// Define the expected main and sub header rows
	let expectedMainHeader = [
		'LỚP',
		'HỌ VÀ TÊN LÓT',
		'TÊN HỌC SINH',
		'NGÀY NHẬP HỌC',
		'SINH NGÀY',
		'SỐ ĐIỆN THOẠI',
		null,
		'ZALO',
		'NĂM SINH',
		'GIỚI TÍNH',
		null,
		'DÂN TỘC',
		'NƠI SINH HS',
		'THÔNG TIN CHA',
		null,
		null,
		'THÔNG TIN MẸ',
		null,
		null,
		'THƯỜNG TRÚ',
		null,
		null,
		'TẠM TRÚ',
		'CHỦ NHÀ TRỌ',
		'BC PHỔ CẬP',
		'KHAI SINH',
		'HỘ KHẨU'
	];

	let expectedSubHeader = [
		null,
		null,
		null,
		null,
		null,
		'ĐT1',
		'ĐT2',
		null,
		null,
		'NAM',
		'NỮ',
		null,
		null,
		'HỌ VÀ TÊN',
		'NĂM SINH',
		'NGHỀ NGHIỆP',
		'HỌ VÀ TÊN',
		'NĂM SINH',
		'NGHỀ NGHIỆP',
		'TỈNH',
		'HUYỆN',
		'XÃ',
		null,
		null,
		null,
		null,
		null
	];

	let acceptableHeaders = {
		'HỌ VÀ TÊN LÓT': ['HỌ VÀ TÊN', 'HỌ'],
		'TÊN HỌC SINH': ['TÊN', 'TÊN RIÊNG', 'TÊN HS'],
		'NGÀY NHẬP HỌC': ['NHẬP HỌC'],
		'SINH NGÀY': ['SINH NHẬT', 'NGÀY SINH'],
		'SỐ ĐIỆN THOẠI': ['SDT', 'SĐT', 'SỐ ĐT', 'SỐ DT'],
		DT1: ['ĐT1', 'SĐT1', 'SDT1'],
		DT2: ['ĐT2', 'SĐT2', 'SDT2']
	};

	let mainHeaderRanges = [
		'A1:E1', // Range before "Số điện thoại"
		'F1:G1', // Range of "Số điện thoại" merged cells
		'H1:I1', // Range after "Số điện thoại" merged cells and before "GIỚI TÍNH"
		'J1:K1', // Range of "GIỚI TÍNH" merged cells
		'L1:M1', // Range after "GIỚI TÍNH" merged cells and before "THÔNG TIN CHA"
		'N1:P1', // Range of "THÔNG TIN CHA" merged cells
		'Q1:S1', // Range of "THÔNG TIN MẸ" merged cells
		'T1:V1' // Range of "THƯỜNG TRÚ" merged cells
	];
	let mainHeaderRange = 'A1:AA1';
	let subHeaderRange = 'A2:AA2';

	let headers = getRowValues(worksheet, mainHeaderRange);
	let subheaders = getRowValues(worksheet, subHeaderRange);

	let checkHeadersResult = checkHeaderFile(
		worksheet,
		mainHeaderRanges,
		subHeaderRange,
		expectedMainHeader,
		expectedSubHeader,
		acceptableHeaders
	);
	if (!checkHeadersResult) {
		console.error('Header check failed.');
		return;
	} else {
		headers = expectedMainHeader;
		subHeaderRange = expectedSubHeader;
	}
	let data = readRowsAndConvertToJson(worksheet, headers, subheaders);
	// const studentData = transformDataKeys(data, attributeMappings.student);

	await insertStudentAndParent(data, class_room_id, branch_id);
}
// export async function importAttendanceFile(filePath, sheetName, class_room_id, branch_id, date) {
// 	let worksheet = readExcelFile(filePath, sheetName);

// 	// Define the expected main and sub header rows
// 	let expectedMainHeader = [
// 		'LỚP',
// 		'HỌ VÀ TÊN LÓT',
// 		'TÊN HỌC SINH',
// 		'NGÀY NHẬP HỌC',
// 		'SINH NGÀY',
// 		'SỐ ĐIỆN THOẠI',
// 		null,
// 		'ZALO',
// 		'NĂM SINH',
// 		'GIỚI TÍNH',
// 		null,
// 		'DÂN TỘC',
// 		'NƠI SINH HS',
// 		'THÔNG TIN CHA',
// 		null,
// 		null,
// 		'THÔNG TIN MẸ',
// 		null,
// 		null,
// 		'THƯỜNG TRÚ',
// 		null,
// 		null,
// 		'TẠM TRÚ',
// 		'CHỦ NHÀ TRỌ',
// 		'BC PHỔ CẬP',
// 		'KHAI SINH',
// 		'HỘ KHẨU'
// 	];

// 	let expectedSubHeader = [
// 		null,
// 		null,
// 		null,
// 		null,
// 		null,
// 		'ĐT1',
// 		'ĐT2',
// 		null,
// 		null,
// 		'NAM',
// 		'NỮ',
// 		null,
// 		null,
// 		'HỌ VÀ TÊN',
// 		'NĂM SINH',
// 		'NGHỀ NGHIỆP',
// 		'HỌ VÀ TÊN',
// 		'NĂM SINH',
// 		'NGHỀ NGHIỆP',
// 		'TỈNH',
// 		'HUYỆN',
// 		'XÃ',
// 		null,
// 		null,
// 		null,
// 		null,
// 		null
// 	];

// 	let acceptableHeaders = {
// 		'HỌ VÀ TÊN LÓT': ['HỌ VÀ TÊN', 'HỌ'],
// 		'TÊN HỌC SINH': ['TÊN', 'TÊN RIÊNG', 'TÊN HS'],
// 		'NGÀY NHẬP HỌC': ['NHẬP HỌC'],
// 		'SINH NGÀY': ['SINH NHẬT', 'NGÀY SINH'],
// 		'SỐ ĐIỆN THOẠI': ['SDT', 'SĐT', 'SỐ ĐT', 'SỐ DT'],
// 		DT1: ['ĐT1', 'SĐT1', 'SDT1'],
// 		DT2: ['ĐT2', 'SĐT2', 'SDT2']
// 	};

// 	let mainHeaderRanges = [
// 		'A1:E1', // Range before "Số điện thoại"
// 		'F1:G1', // Range of "Số điện thoại" merged cells
// 		'H1:I1', // Range after "Số điện thoại" merged cells and before "GIỚI TÍNH"
// 		'J1:K1', // Range of "GIỚI TÍNH" merged cells
// 		'L1:M1', // Range after "GIỚI TÍNH" merged cells and before "THÔNG TIN CHA"
// 		'N1:P1', // Range of "THÔNG TIN CHA" merged cells
// 		'Q1:S1', // Range of "THÔNG TIN MẸ" merged cells
// 		'T1:V1' // Range of "THƯỜNG TRÚ" merged cells
// 	];
// 	let mainHeaderRange = 'A1:AA1';
// 	let subHeaderRange = 'A2:AA2';

// 	let headers = getRowValues(worksheet, mainHeaderRange);
// 	let subheaders = getRowValues(worksheet, subHeaderRange);

// 	let checkHeadersResult = checkHeaderFile(
// 		worksheet,
// 		mainHeaderRanges,
// 		subHeaderRange,
// 		expectedMainHeader,
// 		expectedSubHeader,
// 		acceptableHeaders
// 	);
// 	if (!checkHeadersResult) {
// 		console.error('Header check failed.');
// 		return;
// 	} else {
// 		headers = expectedMainHeader;
// 		subHeaderRange = expectedSubHeader;
// 	}

// 	let valid_attendance_dates = getAttendanceDatesAndWeekdays(date);
// 	let data = readDataAndAttendance(worksheet, valid_attendance_dates);
// 	// Extract student data and attendance status for each date
// 	let studentAttendanceData = [];

// 	for (const row of data) {
// 		const studentData = {
// 			id: null,
// 			first_name: row[1], // Assuming first name is at index 1
// 			last_name: row[2], // Assuming last name is at index 2
// 			dob: row[4] // Assuming DOB is at index 3
// 		};

// 		const attendanceStatus = row.slice(27); // Start from index 27 for attendance status

// 		const studentAttendanceRecord = {
// 			studentData: studentData,
// 			attendanceStatus: attendanceStatus
// 		};

// 		studentAttendanceData.push(studentAttendanceRecord);
// 	}
// 	// Step 4: Get the current date and dates of the month
// 	const { start_date, end_date } = calculateDates(date);
// 	console.log('Start Date:', start_date);
// 	console.log('End Date:', end_date);
// 	let currentAttendance = await getCurrentAttendance(class_room_id, start_date, end_date);
// 	let databaseStudentData = await getStudentsFromClassroom(branch_id, class_room_id);
// 	let attendance_events = await getAllAttendanceEventsByAttendanceId(currentAttendance.id);
// 	// Compare studentData with databaseStudentData
// 	const missingStudents = [];
// 	const validStudents = []
// 	for (const studentAttendanceRecord of studentAttendanceData) {
// 		const { first_name, last_name, dob } = studentAttendanceRecord.studentData;
// 		const matchingStudent = databaseStudentData.find((student) => {
// 			if (dob !== null && dob !== undefined) {
// 				return (
// 					student.first_name.trim() === first_name.trim() &&
// 					student.last_name.trim() === last_name.trim() &&
// 					student.dob !== null && student.dob.trim() === dob.trim()
// 				);
// 			} else {
// 				return (
// 					student.first_name.trim() === first_name.trim() &&
// 					student.last_name.trim() === last_name.trim()
// 				);
// 			}

// 		});

// 		if (matchingStudent == undefined || matchingStudent == null) {
// 			missingStudents.push(studentAttendanceRecord);
// 		}
// 		else{
// 			studentAttendanceRecord.studentData.id = matchingStudent.id;
// 			validStudents.push(studentAttendanceRecord);
// 		}
// 	}

// 	if (missingStudents.length > 0) {
// 		console.error('Missing students from the database:');
// 		for (const missingStudent of missingStudents) {
// 			const { firstName, lastName, dob } = missingStudent.studentData;
// 			console.error(`First Name: ${firstName}, Last Name: ${lastName}, DOB: ${dob}`);
// 		}
// 		return; // Return early if missing students are found
// 	}
// 	const absentDatesByStudent = studentAttendanceData.map((studentAttendanceData) => {
// 		const absentDatesWithStatus = studentAttendanceData.attendanceStatus.reduce((absentDates, status, index) => {
// 		  if (String(status).toLowerCase() === "p" || String(status).toLowerCase() === "k") {
// 			absentDates.push({ date: valid_attendance_dates.attendance_dates[index], status: status.toUpperCase() });
// 		  }
// 		  return absentDates;
// 		}, []);

// 		return {
// 		  student: studentAttendanceData.studentData,
// 		  absentDates: absentDatesWithStatus,
// 		};
// 	  });

// 	if (!currentAttendance) {
// 		// If current attendance does not exist, insert a new one
// 		console.log('Current attendance does not exist. Inserting a new attendance record...');
// 		await insertAttendance(class_room_id, start_date, end_date);
// 		currentAttendance = await getCurrentAttendance(class_room_id, start_date, end_date);
// 	}
// 	let unmatchedRecords = []
// 	for (const studentAbsences of absentDatesByStudent) {
// 		const { student, absentDates } = studentAbsences;

// 		const studentId = student.id;
// 		if (absentDates.length > 0){
// 			for (const absentDate of absentDates) {
// 				await updateOrInsertAttendanceEvent(currentAttendance.id, studentId, absentDate.date, absentDate.status);
// 			}
// 		}
// 	  }
// 	for (const attendanceEvent of attendance_events) {
// 		const { student_id, date } = attendanceEvent;

// 		const matchingLocalRecord = absentDatesByStudent.find(studentAbsences =>
// 			studentAbsences.student.id === student_id &&
// 			studentAbsences.absentDates.some(absentDate =>
// 			  absentDate.date === date && absentDate.date.length > 0
// 			)
// 		  );

// 		if (matchingLocalRecord) {
// 			continue
// 		} else {
// 		  unmatchedRecords.push(attendanceEvent);
// 		}
// 	  }

// 	  console.log('Unmatched records:', unmatchedRecords);

// 	  deleteAttendanceEvent (unmatchedRecords)
// }
export async function importAttendanceFile(worksheet, class_room_id, branch_id, date) {
	// Define the expected main and sub header rows
	let expectedMainHeader = [
		'LỚP',
		'HỌ VÀ TÊN LÓT',
		'TÊN HỌC SINH',
		'NGÀY NHẬP HỌC',
		'SINH NGÀY',
		'SỐ ĐIỆN THOẠI',
		null,
		'ZALO',
		'NĂM SINH',
		'GIỚI TÍNH',
		null,
		'DÂN TỘC',
		'NƠI SINH HS',
		'THÔNG TIN CHA',
		null,
		null,
		'THÔNG TIN MẸ',
		null,
		null,
		'THƯỜNG TRÚ',
		null,
		null,
		'TẠM TRÚ',
		'CHỦ NHÀ TRỌ',
		'BC PHỔ CẬP',
		'KHAI SINH',
		'HỘ KHẨU'
	];

	let expectedSubHeader = [
		null,
		null,
		null,
		null,
		null,
		'ĐT1',
		'ĐT2',
		null,
		null,
		'NAM',
		'NỮ',
		null,
		null,
		'HỌ VÀ TÊN',
		'NĂM SINH',
		'NGHỀ NGHIỆP',
		'HỌ VÀ TÊN',
		'NĂM SINH',
		'NGHỀ NGHIỆP',
		'TỈNH',
		'HUYỆN',
		'XÃ',
		null,
		null,
		null,
		null,
		null
	];

	let acceptableHeaders = {
		'HỌ VÀ TÊN LÓT': ['HỌ VÀ TÊN', 'HỌ'],
		'TÊN HỌC SINH': ['TÊN', 'TÊN RIÊNG', 'TÊN HS'],
		'NGÀY NHẬP HỌC': ['NHẬP HỌC'],
		'SINH NGÀY': ['SINH NHẬT', 'NGÀY SINH'],
		'SỐ ĐIỆN THOẠI': ['SDT', 'SĐT', 'SỐ ĐT', 'SỐ DT'],
		DT1: ['ĐT1', 'SĐT1', 'SDT1'],
		DT2: ['ĐT2', 'SĐT2', 'SDT2']
	};

	let mainHeaderRanges = [
		'A1:E1', // Range before "Số điện thoại"
		'F1:G1', // Range of "Số điện thoại" merged cells
		'H1:I1', // Range after "Số điện thoại" merged cells and before "GIỚI TÍNH"
		'J1:K1', // Range of "GIỚI TÍNH" merged cells
		'L1:M1', // Range after "GIỚI TÍNH" merged cells and before "THÔNG TIN CHA"
		'N1:P1', // Range of "THÔNG TIN CHA" merged cells
		'Q1:S1', // Range of "THÔNG TIN MẸ" merged cells
		'T1:V1' // Range of "THƯỜNG TRÚ" merged cells
	];
	let mainHeaderRange = 'A1:AA1';
	let subHeaderRange = 'A2:AA2';

	let headers = getRowValues(worksheet, mainHeaderRange);
	let subheaders = getRowValues(worksheet, subHeaderRange);

	let checkHeadersResult = checkHeaderFile(
		worksheet,
		mainHeaderRanges,
		subHeaderRange,
		expectedMainHeader,
		expectedSubHeader,
		acceptableHeaders
	);
	if (!checkHeadersResult) {
		console.error('Header check failed.');
		return;
	} else {
		headers = expectedMainHeader;
		subHeaderRange = expectedSubHeader;
	}

	let valid_attendance_dates = getAttendanceDatesAndWeekdays(date);
	let data = readDataAndAttendance(worksheet, valid_attendance_dates);
	// Extract student data and attendance status for each date
	let studentAttendanceData = [];

	for (const row of data) {
		const studentData = {
			id: null,
			first_name: row[1], // Assuming first name is at index 1
			last_name: row[2], // Assuming last name is at index 2
			dob: row[4] // Assuming DOB is at index 3
		};

		const attendanceStatus = row.slice(27); // Start from index 27 for attendance status

		const studentAttendanceRecord = {
			studentData: studentData,
			attendanceStatus: attendanceStatus
		};

		studentAttendanceData.push(studentAttendanceRecord);
	}
	// Step 4: Get the current date and dates of the month
	const { start_date, end_date } = calculateDates(date);
	console.log('Start Date:', start_date);
	console.log('End Date:', end_date);
	let currentAttendance = await getCurrentAttendance(class_room_id, start_date, end_date);
	let databaseStudentData = await getStudentsFromClassroom(branch_id, class_room_id);
	let attendance_events = await getAllAttendanceEventsByAttendanceId(currentAttendance.id);
	// Compare studentData with databaseStudentData
	const missingStudents = [];
	const validStudents = [];
	for (const studentAttendanceRecord of studentAttendanceData) {
		const { first_name, last_name, dob } = studentAttendanceRecord.studentData;
		const matchingStudent = databaseStudentData.find((student) => {
			if (dob !== null && dob !== undefined) {
				return (
					student.first_name.trim() === first_name.trim() &&
					student.last_name.trim() === last_name.trim() &&
					student.dob !== null &&
					student.dob.trim() === dob.trim()
				);
			} else {
				return (
					student.first_name.trim() === first_name.trim() &&
					student.last_name.trim() === last_name.trim()
				);
			}
		});

		if (matchingStudent == undefined || matchingStudent == null) {
			missingStudents.push(studentAttendanceRecord);
		} else {
			studentAttendanceRecord.studentData.id = matchingStudent.id;
			validStudents.push(studentAttendanceRecord);
		}
	}

	if (missingStudents.length > 0) {
		console.error('Missing students from the database:');
		for (const missingStudent of missingStudents) {
			const { firstName, lastName, dob } = missingStudent.studentData;
			console.error(`First Name: ${firstName}, Last Name: ${lastName}, DOB: ${dob}`);
		}
		return; // Return early if missing students are found
	}
	const absentDatesByStudent = studentAttendanceData.map((studentAttendanceData) => {
		const absentDatesWithStatus = studentAttendanceData.attendanceStatus.reduce(
			(absentDates, status, index) => {
				if (String(status).toLowerCase() === 'p' || String(status).toLowerCase() === 'k') {
					absentDates.push({
						date: valid_attendance_dates.attendance_dates[index],
						status: status.toUpperCase()
					});
				}
				return absentDates;
			},
			[]
		);

		return {
			student: studentAttendanceData.studentData,
			absentDates: absentDatesWithStatus
		};
	});

	if (!currentAttendance) {
		// If current attendance does not exist, insert a new one
		console.log('Current attendance does not exist. Inserting a new attendance record...');
		await insertAttendance(class_room_id, start_date, end_date);
		currentAttendance = await getCurrentAttendance(class_room_id, start_date, end_date);
	}
	let unmatchedRecords = [];
	for (const studentAbsences of absentDatesByStudent) {
		const { student, absentDates } = studentAbsences;

		const studentId = student.id;
		if (absentDates.length > 0) {
			for (const absentDate of absentDates) {
				await updateOrInsertAttendanceEvent(
					currentAttendance.id,
					studentId,
					absentDate.date,
					absentDate.status
				);
			}
		}
	}
	for (const attendanceEvent of attendance_events) {
		const { student_id, date } = attendanceEvent;

		const matchingLocalRecord = absentDatesByStudent.find(
			(studentAbsences) =>
				studentAbsences.student.id === student_id &&
				studentAbsences.absentDates.some(
					(absentDate) => absentDate.date === date && absentDate.date.length > 0
				)
		);

		if (matchingLocalRecord) {
			continue;
		} else {
			unmatchedRecords.push(attendanceEvent);
		}
	}

	console.log('Unmatched records:', unmatchedRecords);

	deleteAttendanceEvent(unmatchedRecords);
}

function convertDataToDictionary(dataArr) {
	const resultArr = [];

	for (const data of dataArr) {
		const result = {};

		// Extract student data
		const studentData = data[0];
		result['student'] = studentData;

		// Extract parents data
		const parentsData = data[1];
		const fathers = [];
		const mothers = [];

		for (const parent of parentsData) {
			if (parent.sex === 'NAM') {
				// Map parent with sex "NAM" to father key
				fathers.push({ ...parent, sex: undefined }); // Remove the "sex" attribute
			} else if (parent.sex === 'NỮ') {
				// Map parent with sex "NỮ" to mother key
				mothers.push({ ...parent, sex: undefined }); // Remove the "sex" attribute
			}
		}

		result['father'] = fathers;
		result['mother'] = mothers;

		resultArr.push(result);
	}

	return resultArr;
}

export async function exportAttendanceTemplate(branch_id, class_room_id, attendance_date) {
	let studentData = await getStudentsFromClassroom(branch_id, class_room_id);
	const studentIds = studentData.map((student) => student.id);
	let parentData = await getParentsByStudentId(studentIds);
	// Merge student and parent data based on student_id
	let mergedData = mergeStudentParentData(studentData, parentData);
	mergedData = convertDataToDictionary(mergedData);
	// Calculate start_date and end_date based on the current date
	const { start_date, end_date } = calculateDates(attendance_date);
	// console.log('Start Date:', start_date);
	// console.log('End Date:', end_date);
	insertAttendance(class_room_id, start_date, end_date);
	return writeDataToTemplate(mergedData, attendance_date);
}
// export async function exportAttendance(branch_id, class_room_id, date) {
// 	try {
// 		// Step 1: Fetch student data for the provided branch_id and class_room_id
// 		let studentData = await getStudentsFromClassroom(branch_id, class_room_id);
// 		const studentIds = studentData.map((student) => student.id);

// 		// Step 2: Fetch parent data for the retrieved studentIds
// 		let parentData = await getParentsByStudentId(studentIds);

// 		// Step 3: Merge student and parent data based on student_id
// 		let mergedData = mergeStudentParentData(studentData, parentData);
// 		mergedData = convertDataToDictionary(mergedData);

// 		// Step 4: Get the current date and dates of the month
// 		const { attendance_dates, weekdaysVietnamese } = getAttendanceDatesAndWeekdays(date);
// 		// Calculate start_date and end_date based on the current date
// 		const { start_date, end_date } = calculateDates(date);
// 		console.log('Start Date:', start_date);
// 		console.log('End Date:', end_date);

// 		// Step 5: Get the current attendance for the classroom based on the start_date and end_date
// 		let currentAttendance = await getCurrentAttendance(class_room_id, start_date, end_date);

// 		if (!currentAttendance) {
// 			// If current attendance does not exist, insert a new one
// 			console.log('Current attendance does not exist. Inserting a new attendance record...');
// 			await insertAttendance(class_room_id, start_date, end_date);
// 			currentAttendance = await getCurrentAttendance(class_room_id, start_date, end_date);
// 		}

// 		// Step 6: Get the attendance_id and attendance_event based on the currentAttendance
// 		let attendance_id = currentAttendance['id'];
// 		let attendance_event = await getAllAttendanceEventsByAttendanceId(attendance_id);

// 		// Step 7: Create the attendanceStatusMapping and update it based on attendance_event
// 		const attendanceStatusMapping = {};

// 		const days = Object.values(attendance_dates); // Map student_id to their attendance status array
// 		for (const event of attendance_event) {
// 			const eventDate = event.date; // Convert the date format to match 'YYYY-MM-DD'
// 			const index = attendance_dates.indexOf(eventDate); // Find the index of the day in the days array
// 			if (index !== -1) {
// 				const studentId = event.student_id;
// 				if (!attendanceStatusMapping[studentId]) {
// 					attendanceStatusMapping[studentId] = Array(days.length).fill('1');
// 				}
// 				attendanceStatusMapping[studentId][index] = event.status;
// 			}
// 		}

// 		// Step 8: Update the attendance_status for each student in mergedData
// 		mergedData.forEach((student) => {
// 			const studentId = student.student.id;
// 			if (attendanceStatusMapping[studentId]) {
// 				student.attendance_status = attendanceStatusMapping[studentId];
// 			} else {
// 				// If attendance status array doesn't exist for the student, add an empty array
// 				student.attendance_status = Array(attendance_dates.length).fill('1');
// 			}
// 		});

// 		// Step 9: Write data and attendance to a file or perform any other required actions
// 		writeAttendance(mergedData, date, attendance_dates, weekdaysVietnamese);
// 	} catch (error) {
// 		console.error('Error exporting attendance:', error);
// 	}
// }

export async function exportAttendance(branch_id, class_room_id, date) {
	try {
		// Step 1: Fetch student data for the provided branch_id and class_room_id
		let studentData = await getStudentsFromClassroom(branch_id, class_room_id);
		const studentIds = studentData.map((student) => student.id);

		// Step 2: Fetch parent data for the retrieved studentIds
		let parentData = await getParentsByStudentId(studentIds);

		// Step 3: Merge student and parent data based on student_id
		let mergedData = mergeStudentParentData(studentData, parentData);
		mergedData = convertDataToDictionary(mergedData);

		// Step 4: Get the current date and dates of the month
		const { attendance_dates, weekdaysVietnamese } = getAttendanceDatesAndWeekdays(date);
		// Calculate start_date and end_date based on the current date
		const { start_date, end_date } = calculateDates(date);
		console.log('Start Date:', start_date);
		console.log('End Date:', end_date);

		// Step 5: Get the current attendance for the classroom based on the start_date and end_date
		let currentAttendance = await getCurrentAttendance(class_room_id, start_date, end_date);

		if (!currentAttendance) {
			// If current attendance does not exist, insert a new one
			console.log('Current attendance does not exist. Inserting a new attendance record...');
			await insertAttendance(class_room_id, start_date, end_date);
			currentAttendance = await getCurrentAttendance(class_room_id, start_date, end_date);
		}

		// Step 6: Get the attendance_id and attendance_event based on the currentAttendance
		let attendance_id = currentAttendance['id'];
		let attendance_event = await getAllAttendanceEventsByAttendanceId(attendance_id);

		// Step 7: Create the attendanceStatusMapping and update it based on attendance_event
		const attendanceStatusMapping = {};

		const days = Object.values(attendance_dates); // Map student_id to their attendance status array
		for (const event of attendance_event) {
			const eventDate = event.date; // Convert the date format to match 'YYYY-MM-DD'
			const index = attendance_dates.indexOf(eventDate); // Find the index of the day in the days array
			if (index !== -1) {
				const studentId = event.student_id;
				if (!attendanceStatusMapping[studentId]) {
					attendanceStatusMapping[studentId] = Array(days.length).fill('1');
				}
				attendanceStatusMapping[studentId][index] = event.status;
			}
		}

		// Step 8: Update the attendance_status for each student in mergedData
		mergedData.forEach((student) => {
			const studentId = student.student.id;
			if (attendanceStatusMapping[studentId]) {
				student.attendance_status = attendanceStatusMapping[studentId];
			} else {
				// If attendance status array doesn't exist for the student, add an empty array
				student.attendance_status = Array(attendance_dates.length).fill('1');
			}
		});

		// Step 9: Write data and attendance to a file or perform any other required actions
		return writeAttendance(mergedData, date);
	} catch (error) {
		console.error('Error exporting attendance:', error);
	}
}
