import { supabase } from './supabase.js'; // Import your Supabase client

// Function to insert a branch record
export const insertBranch = async (branchData) => {
  try {
    const { data, error } = await supabase
      .from('branch')
      .insert([branchData])
      .single();

    if (error) {
      console.error('Error inserting branch:', error);
    } else {
      console.log('Branch inserted successfully:', data);
    }
  } catch (error) {
    console.error('Error:', error);
  }
};