import { supabase } from './supabase.js';

// Function to insert an attendance record
export const insertAttendance = async (class_room_id, start_date, end_date) => {
	const attendanceData = {
		class_room_id,
		start_date,
		end_date
	};

	try {
		const { data, error } = await supabase.from('attendance').insert([attendanceData]).single();

		if (error) {
			console.error('Error inserting attendance:', error);
		} else {
			// console.log('Attendance inserted successfully:', data);
		}
	} catch (err) {
		console.error('Error inserting attendance:', err.message);
	}
};

export async function getCurrentAttendance(class_room_id, start_date, end_date) {
	try {
		// Step 1: Fetch the attendance based on the provided class_room_id, start_date, and end_date
		const { data: currentAttendance, error: attendanceError } = await supabase
			.from('attendance')
			.select('*')
			.eq('class_room_id', class_room_id)
			.eq('start_date', start_date)
			.eq('end_date', end_date);

		if (attendanceError) {
			console.error('Error fetching attendance:', attendanceError);
			return null;
		}

		if (!currentAttendance || currentAttendance.length === 0) {
			console.log(
				'No attendance data found for the given class_room_id, start_date, and end_date.'
			);
			return null;
		}

		return currentAttendance[0]; // Return the first (and only) attendance record found
	} catch (error) {
		console.error('Error getting current attendance:', error);
		return null;
	}
}
