import {
	exportTemplateStudentList,
	importStudentListFile,
	exportAttendanceTemplate,
	exportAttendance,
	importAttendanceFile
} from './excel.js';

const class_name = 'Lop01';

const filePath = `src/lib/Danh_sach_lop_${class_name}.xlsx`;
const sheetName = 'Sheet1';
let branch_id = 1;
let class_room_id = 1;

const currentDate = new Date();
const customDate = new Date('2023-01-01');
await importStudentListFile(filePath, sheetName, class_room_id, branch_id);
await exportAttendanceTemplate(branch_id, class_room_id, currentDate);
await exportAttendance(branch_id, class_room_id, currentDate);
await importAttendanceFile(sheetName, class_room_id, branch_id, currentDate);
