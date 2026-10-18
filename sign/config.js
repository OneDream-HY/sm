// ODC client-sign configuration.
// After deploying the Supabase Edge Function, put its URL here.
//   supabase functions deploy sign-session --no-verify-jwt
//   → https://<project-ref>.supabase.co/functions/v1/sign-session
window.ODC_SIGN_CONFIG = {
  functionUrl: 'https://YOUR-PROJECT-REF.supabase.co/functions/v1/sign-session',
  // Only needed if the function is deployed WITH JWT verification (default is --no-verify-jwt).
  anonKey: '',
  // Shown to the participant on the confirmation screen and in error messages.
  officePhone: '02 9161 7567',
  officeEmail: 'agedcare@onedreamcommunity.com.au',
  // How long a link stays valid (informational — the real limit is set in the Edge Function).
  linkValidDays: 30
};
