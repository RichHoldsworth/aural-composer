// Supabase configuration for Aural Composer.
// Both values are PUBLIC keys — designed to be exposed in client-side code.
// The actual security is enforced by Row Level Security policies on the database.
import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://jwbguqllixwnppeieunw.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_r62ONskQKJ40m1s3ThxTTw_gOqqX3QQ';

export const supabase = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
});

// Helper: get the current user's profile row
export async function getMyProfile() {
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return null;
  const { data, error } = await supabase
    .from('profiles')
    .select('*')
    .eq('id', user.id)
    .single();
  if (error) {
    console.warn('Could not fetch profile:', error.message);
    return null;
  }
  return data;
}

// Helper: list exams visible to the current user (their own + shared-with-all if approved)
export async function listExams() {
  const { data, error } = await supabase
    .from('exams')
    .select('*, owner:profiles!owner_id(email, display_name)')
    .order('updated_at', { ascending: false });
  if (error) throw new Error(error.message);
  return data;
}

// Helper: save (or update) an exam
export async function saveExam({ id, name, config, shared_with_all }) {
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) throw new Error('Not signed in');
  if (id) {
    const { data, error } = await supabase
      .from('exams')
      .update({ name, config, shared_with_all })
      .eq('id', id)
      .select()
      .single();
    if (error) throw new Error(error.message);
    return data;
  } else {
    const { data, error } = await supabase
      .from('exams')
      .insert({ owner_id: user.id, name, config, shared_with_all: !!shared_with_all })
      .select()
      .single();
    if (error) throw new Error(error.message);
    return data;
  }
}

export async function deleteExam(id) {
  const { error } = await supabase.from('exams').delete().eq('id', id);
  if (error) throw new Error(error.message);
}

// Admin: list all users
export async function listAllProfiles() {
  const { data, error } = await supabase
    .from('profiles')
    .select('*')
    .order('created_at', { ascending: false });
  if (error) throw new Error(error.message);
  return data;
}

// Admin: approve / unapprove a user
export async function setUserApproved(userId, approved) {
  const { error } = await supabase
    .from('profiles')
    .update({ approved })
    .eq('id', userId);
  if (error) throw new Error(error.message);
}
