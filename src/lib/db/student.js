process.argv.push('--experimental-modules');
import { attributeMappings } from './keyword_mapping.js';
import { supabase } from './supabase.js';

export const insertStudent = async (studentData, class_room_id) => {
	try {
		// Prepare data for insertion/upsertion
		const dataToInsert = [];
		const dataToUpdate = [];

		for (const student of studentData) {
			const { first_name, last_name, dob } = student;

			// Check if the student already exists in the database based on the unique identifier
			const { data: existingStudent, error } = await supabase
				.from('student')
				.select('*')
				.eq('first_name', first_name)
				.eq('last_name', last_name)
				.eq('dob', dob);

			if (error) {
				throw new Error(`Error fetching existing student data: ${error.message}`);
			}

			if (!existingStudent || existingStudent.length === 0) {
				// Student does not exist, add to dataToInsert
				dataToInsert.push({ ...student, class_room_id });
			} else {
				// Student already exists, add to dataToUpdate
				dataToUpdate.push({ ...student, id: existingStudent[0].id });
			}
		}

		// Perform the insertion/upsertion operations
		const { data: insertedData, error: upsertError } = await supabase
			.from('student')
			.upsert([...dataToInsert, ...dataToUpdate], {
				returning: 'minimal',
				onConflict: ['first_name', 'last_name', 'dob']
			});

		if (upsertError) {
			throw new Error(`Error inserting/updating student data: ${upsertError.message}`);
		}

		console.log('Student inserted/updated successfully:', insertedData);
		return insertedData;
	} catch (error) {
		console.error('Error inserting/updating student data:', error);
		return [];
	}
};

export async function selectIdFromStudentData(studentData) {
	try {
		const studentIds = [];
		for (const student of studentData) {
			// Initialize the query without the 'dob' filter
			let query = supabase
				.from('student')
				.select('id')
				.eq('first_name', student.first_name)
				.eq('last_name', student.last_name);

			// If 'dob' is not null, include the 'dob' filter in the query
			if (student.dob !== null) {
				query = query.eq('dob', student.dob);
			} else {
				// If 'dob' is null, also include null dob records in the query
				query = query.is('dob', null);
			}

			const { data, error } = await query;

			if (error) {
				console.error('Error fetching student:', error);
				studentIds.push(null); // Push null to indicate that an error occurred for this student
			} else if (data && data.length > 0) {
				// Push the first matching student ID to the studentIds array
				studentIds.push(data[0].id);
			} else {
				console.error(
					'Student not found with first name:',
					student.first_name,
					'and last name:',
					student.last_name
				);
				studentIds.push(null); // Push null to indicate that the student was not found in the database
			}
		}

		return studentIds;
	} catch (error) {
		console.error('Error selecting student IDs:', error);
		return []; // Return an empty array in case of an error
	}
}

// Assuming you have the correct import for the Supabase client
// Function to get all students from a class room of a branch
export async function getStudentsFromClassroom(branch_id, class_room_id) {
	try {
		// Fetch the class_room_id based on the provided branch_id
		const { data: classRoomData, error: classRoomError } = await supabase
			.from('class_room')
			.select('id')
			.eq('branch_id', branch_id)
			.eq('id', class_room_id);

		if (classRoomError) {
			console.error('Error fetching class room:', classRoomError);
			return [];
		}

		if (!classRoomData || classRoomData.length === 0) {
			console.log('Class room not found with the provided branch_id and class_room_id.');
			return [];
		}

		// Fetch all students from the specified class room of the branch
		const { data: studentData, error: studentError } = await supabase
			.from('student')
			.select('*')
			.eq('class_room_id', classRoomData[0].id);

		if (studentError) {
			console.error('Error fetching students:', studentError);
			return [];
		}

		if (studentData && studentData.length > 0) {
			// console.log('Students found in the class room:', studentData);
			return studentData;
		} else {
			console.log('No students found in the class room.');
			return [];
		}
	} catch (error) {
		console.error('Error getting students from the class room:', error);
		return [];
	}
}
export async function deleteChainAttendanceEventAttendanceStudentParent(idsToDelete) {
	try {
		for (const id of idsToDelete) {
			// Fetch student and parent data using the provided ID
			const { data: studentData, error: studentError } = await supabase
				.from('student')
				.select('*')
				.eq('id', id);

			if (studentError) {
				console.error('Error fetching student data:', studentError);
				continue;
			}

			const { data: parentData, error: parentError } = await supabase
				.from('parent')
				.select('*')
				.eq('student_id', id);

			if (parentError) {
				console.error('Error fetching parent data:', parentError);
				continue;
			}

			// Delete parent's related attendance_event records
			const { error: parentAttendanceError } = await supabase
				.from('attendance_event')
				.delete()
				.eq('student_id', id); // Assuming 'student_id' is the foreign key in attendance_event

			if (parentAttendanceError) {
				console.error('Error deleting parent attendance_event:', parentAttendanceError);
			}

			// Delete parent records
			for (const parent of parentData) {
				const { error: deleteParentError } = await supabase
					.from('parent')
					.delete()
					.eq('id', parent.id); // Assuming 'id' is the primary key for parents table

				if (deleteParentError) {
					console.error('Error deleting parent:', deleteParentError);
				}
			}

			// Delete student's related attendance_event records
			const { error: studentAttendanceError } = await supabase
				.from('attendance_event')
				.delete()
				.eq('student_id', id); // Assuming 'student_id' is the foreign key in attendance_event

			if (studentAttendanceError) {
				console.error('Error deleting student attendance_event:', studentAttendanceError);
			}

			// Delete student record
			const { error: deleteStudentError } = await supabase.from('student').delete().eq('id', id); // Assuming 'id' is the primary key for students table

			if (deleteStudentError) {
				console.error('Error deleting student:', deleteStudentError);
			}
		}
		console.log('Deletion process completed.');
	} catch (error) {
		console.error('Error during deletion process:', error);
	}
}
