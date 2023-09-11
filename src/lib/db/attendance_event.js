import { supabase } from './supabase.js';

// Function to insert an attendance event record
export const insertAttendanceEvent = async (attendance_id, student_id, date, status) => {
  const attendanceEventData = {
    attendance_id,
    student_id,
    date,
    status,
  };

  try {
    const { data, error } = await supabase.from('attendance_event').insert([attendanceEventData]).single();

    if (error) {
      console.error('Error inserting attendance event:', error);
    } else {
      console.log('Attendance event inserted successfully:', data);
    }
  } catch (err) {
    console.error('Error inserting attendance event:', err.message);
  }
};

// Function to get all attendance events by attendance_id
export async function getAllAttendanceEventsByAttendanceId(attendance_id) {
    try {
      // Replace 'attendance_events' with the name of your table that holds attendance events
      const { data, error } = await supabase
        .from('attendance_event')
        .select('*')
        .eq('attendance_id', attendance_id);
  
      if (error) {
        throw new Error('Error while fetching attendance events:', error.message);
      }
  
      return data;
    } catch (error) {
      // Handle any errors that might occur during the data retrieval process
      console.error(error.message);
      return null;
    }
  }

  // Function to update or insert an attendance event record
export const updateOrInsertAttendanceEvent = async (attendance_id, student_id, date, status) => {
  try {
    // Check if the attendance event record already exists
    const { data: existingRecord, error } = await supabase
      .from('attendance_event')
      .select('*')
      .eq('attendance_id', attendance_id)
      .eq('student_id', student_id)
      .eq('date', date)
      .single();

    if (error) {
      console.error('Error querying attendance event:', error);
      if (error.code === 'PGRST116'){
          // Insert a new attendance event record
          const attendanceEventData = {
            attendance_id,
            student_id,
            date,
            status,
          };
    
          const { data: insertedData, error: insertError } = await supabase
            .from('attendance_event')
            .insert([attendanceEventData])
            .single();
    
          if (insertError) {
            console.error('Error inserting attendance event:', insertError);
          } else {
            console.log('Attendance event inserted successfully:', insertedData);
          }
      }
      else{
        return;
      }

    }

    if (existingRecord) {
      // Update the status of the existing attendance event record
      const { data: updatedData, error: updateError } = await supabase
        .from('attendance_event')
        .update({ status })
        .eq('id', existingRecord.id)
        .single();

      if (updateError) {
        console.error('Error updating attendance event:', updateError);
      } else {
        console.log('Attendance event updated successfully:', updatedData);
      }
    } else {
      // Insert a new attendance event record
      const attendanceEventData = {
        attendance_id,
        student_id,
        date,
        status,
      };

      const { data: insertedData, error: insertError } = await supabase
        .from('attendance_event')
        .insert([attendanceEventData])
        .single();

      if (insertError) {
        console.error('Error inserting attendance event:', insertError);
      } else {
        console.log('Attendance event inserted successfully:', insertedData);
      }
    }
  } catch (err) {
    console.error('Error updating or inserting attendance event:', err.message);
  }
};

export async function deleteAttendanceEvent(data) {
  for (const record of data) {
    const { data, error } = await supabase
      .from('attendance_event')
      .delete()
      .eq('id', record.id);

    if (error) {
      console.error(`Error deleting record with ID ${record.id}:`, error);
    } else {
      console.log(`Record with ID ${record.id} deleted successfully.`);
    }
  }
}
