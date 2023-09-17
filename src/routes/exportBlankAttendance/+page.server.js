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
		let currentMonth = new Date();
		try {
			let workbook = await exportAttendanceTemplate(branch_id, classRoomId, currentMonth);
			let buffer = await workbook.xlsx.writeBuffer();
			let base64String = Buffer.from(buffer).toString('base64');
			// const sheetjs_workbook = read(buffer);
			return { success: true, workbook: base64String };
		} catch (error) {
			console.error('Error processing the file:', error);
			return fail(500, { error: true });
		}
	}
};
