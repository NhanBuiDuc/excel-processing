import { supabase } from './supabase.js';
import {attributeMappings} from './keyword_mapping.js'
export const fetchUserById = async (userId) => {
  const { data, error } = await supabase
    .from('users')
    .select('*')
    .eq('id', userId)
    .single();

  if (error) {
    console.error(error);
  } else {
    console.log(data);
  }
};
export const insertUser = async (email, password, role) => {
    const { data, error } = await supabase
      .from('users')
      .insert([{ email, password, role }])
      .single();
  
    if (error) {
      console.error(error);
    } else {
      console.log('User inserted successfully:', data);
    }
};
export const getAllUsers = async () => {
    const { data, error } = await supabase.from('users').select('*');

    if (error) {
        console.error(error);
    } else {
        console.log('All users:', data);
    }
};

