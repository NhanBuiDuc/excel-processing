import { attributeMappings } from './keyword_mapping.js';
import { supabase } from './supabase.js';

// Function to insert a class room record
export const insertClassRoom = async (classRoomData) => {
	const { data, error } = await supabase.from('class_room').insert([classRoomData]).single();

	if (error) {
		console.error('Error inserting class room:', error);
	} else {
		console.log('Class room inserted successfully:', data);
	}
};
// Function to get all attendance events by attendance_id
export async function getAllClassRoomByBranchId(branch_id) {
	try {
		// Replace 'attendance_events' with the name of your table that holds attendance events
		const { data, error } = await supabase
			.from('class_room')
			.select('*')
			.eq('branch_id', branch_id);

		if (error) {
			throw new Error('Error while fetching class room data:', error.message);
		}
		console.log(data);
		return data;
	} catch (error) {
		// Handle any errors that might occur during the data retrieval process
		console.error(error.message);
		return null;
	}
}
