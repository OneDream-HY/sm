-- One Dream Community · client e-sign sessions
-- One row per link sent to a participant. The documents themselves (PDFs) are
-- NOT stored here — they go to SharePoint via Microsoft Graph. This table
-- holds the pre-filled form payload (so the participant sees exactly what the
-- coordinator prepared) and the audit trail of the signing.

create table if not exists public.sign_sessions (
  id                 uuid primary key default gen_random_uuid(),
  created_at         timestamptz not null default now(),
  expires_at         timestamptz not null,
  page               text not null check (page in ('contract', 'onboarding')),
  coordinator_email  text not null,
  participant_name   text not null,
  participant_email  text,
  payload            jsonb not null,           -- the coordinator's snapshot (form values, staff signatures)
  status             text not null default 'sent' check (status in ('sent', 'opened', 'submitting', 'signed', 'failed')),
  opened_at          timestamptz,
  opened_ip          text,
  signed_at          timestamptz,
  signer_ip          text,
  signer_ua          text,
  signer_tz          text,
  signed_payload     jsonb,                    -- snapshot at submission (participant signatures included)
  files              jsonb,                    -- [{name, size, sha256, webUrl}]
  email_result       jsonb,                    -- {to, cc, messageId}
  last_error         text,
  created_ip         text
);

create index if not exists sign_sessions_created_ip_idx on public.sign_sessions (created_ip, created_at desc);
create index if not exists sign_sessions_status_idx     on public.sign_sessions (status, created_at desc);

-- Only the Edge Function (service role) touches this table.
alter table public.sign_sessions enable row level security;
revoke all on public.sign_sessions from anon, authenticated;
