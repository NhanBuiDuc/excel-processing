import { fail } from '@sveltejs/kit';
import * as classroom from '$lib/db/class_room';
import { exportAttendanceTemplate } from '../../lib/db/excel';
import { read, utils, writeFileXLSX } from 'xlsx';

/** @type {import('./$types').PageServerLoad} */
export const load = async () => {
	const loadClassRoomData = async () => {
		let branch_id = 1; //code cứng
		const data = await classroom.getAllClassRoomByBranchId(branch_id);
		const parsedData = data.map((item) => ({
			text: item.name,
			value: item.id
		}));
		return parsedData;
	};
	return {
		class_room_data: loadClassRoomData()
	};
};
/** @type {import('@sveltejs/kit').Handle} */
export const actions = {
	download: async ({ request }) => {
		let branch_id = 1; //code cứng
		const data = await request.formData();
		const classRoomId = parseInt(data.get('class_room_id'), 10);
		let classRoomData = await classroom.getClassRoomById(classRoomId);
		const class_room_name = classRoomData[0].name;
		const fromDateString = data.get('fromDate');
		const dateObject = parseDateStringToDate(fromDateString);
		const filename = 'DiemDanhLop' + class_room_name + 'Ngay' + fromDateString;
		if (
			classRoomId === null ||
			classRoomId === undefined ||
			Number.isNaN(classRoomId) ||
			!classRoomId ||
			fromDateString === '' ||
			dateObject.toString() === 'Invalid Date'
		)
			return fail(500, {
				error: true,
				message:
					'Xin hãy nhập định dạng .xlsx file excel, chọn lớp học, ngày tháng và nhập đúng tên của Excel Sheet'
			});
		try {
			let workbook = await exportAttendanceTemplate(branch_id, classRoomId, dateObject);
			let buffer = await workbook.xlsx.writeBuffer();
			let base64String = Buffer.from(buffer).toString('base64');
			// const sheetjs_workbook = read(buffer);
			return { success: true, workbook: base64String, filename: filename };
		} catch (error) {
			console.error('Error processing the file:', error);
			return fail(500, { error: true });
		}
	}
};
function parseDateStringToDate(dateString) {
	const parts = dateString.split('/');
	const day = parseInt(parts[0], 10); // Parse the day as an integer
	const month = parseInt(parts[1], 10); // Parse the month as an integer
	const year = parseInt(parts[2], 10); // Parse the year as an integer

	// Create a Date object (months are zero-based, so subtract 1 from the month)
	const date = new Date(year, month - 1, day);

	return date;
}
