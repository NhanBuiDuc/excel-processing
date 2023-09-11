import {attributeMappings} from './keyword_mapping.js'
import { supabase } from './supabase.js';

// Function to insert parent records with foreign key relationships to the student
export const insertParent = async (parentData) => {
  try {
    const { data, error } = await supabase
      .from('parent')
      .insert(parentData)

    if (error) {
      console.error('Error inserting parents:', error);
    } else {
      console.log('Parents inserted successfully:', data);
    }
  } catch (error) {
    console.error('Error inserting parents:', error);
  }
};
export async function getParentsByStudentId(student_id) {
  try {
    // Fetch all parents associated with the provided student_id
    const { data, error } = await supabase
      .from('parent')
      .select('*')
      .in('student_id', student_id);

    if (error) {
      console.error('Error fetching parents:', error);
      return [];
    }

    if (data && data.length > 0) {
      console.log('Parents found for the student:', data);
      return data;
    } else {
      console.log('No parents found for the student.');
      return [];
    }
  } catch (error) {
    console.error('Error getting parents by student_id:', error);
    return [];
  }
}
