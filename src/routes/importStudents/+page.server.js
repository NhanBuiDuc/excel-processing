import { fail } from '@sveltejs/kit';
import * as classroom from '$lib/db/class_room';
import { importStudentListFile } from '../../lib/db/excel';
import { read } from 'xlsx';
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

export const actions = {
	upload: async ({ request }) => {
		let branch_id = 1; //code cứng
		const data = await request.formData();
		const classRoomId = parseInt(data.get('class_room_id'), 10);
		console.log('classs id: ', classRoomId);
		const worksheet_name = data.get('worksheet_name');
		console.log('worksheet: ', worksheet_name);
		const file = data.get('file-upload');
		const fileTypes = ['.xlsx', '.xls', '.xlsm', '.csv'];
		let includesFileType = false;
		for (let i = 0; i < fileTypes.length; i++) {
			if (file.name.endsWith(fileTypes[i])) {
				includesFileType = true;
				break;
			}
		}

		if (
			!includesFileType ||
			classRoomId === null ||
			classRoomId === undefined ||
			Number.isNaN(classRoomId) ||
			!classRoomId ||
			worksheet_name === null ||
			worksheet_name === undefined ||
			Number.isNaN(worksheet_name) ||
			!worksheet_name ||
			
		)
			return fail(500, {
				error: true,
				message:
					'Xin hãy nhập định dạng .xlsx file excel, chọn lớp học, và nhập đúng tên của Excel Sheet'
			});
		try {
			// Read the uploaded file as an array buffer
			const arrayBuffer = await file.arrayBuffer();
			const workbook = read(arrayBuffer);
			// Get the worksheet by sheet name
			const worksheet = workbook.Sheets[worksheet_name];
			if (worksheet === undefined) {
				return fail(500, { error: true, message: 'Tên Sheet của file excel không tồn tại' });
			} else {
				// Call your 'importAttendanceFile' function with the 'worksheet' data
				return importStudentListFile(worksheet, classRoomId, branch_id);
			}
		} catch (error) {
			return fail(500, { error: true, message: error });
		}
	}
};
