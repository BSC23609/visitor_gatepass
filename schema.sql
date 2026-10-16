-- Visitor Gate Pass register
-- Run this once in the Neon SQL Editor (Neon console > your project > SQL Editor > paste > Run)

CREATE TABLE IF NOT EXISTS visitor_passes (
  id                 BIGSERIAL PRIMARY KEY,
  pass_no            TEXT        NOT NULL UNIQUE,        -- e.g. VP-261016-130812
  company            TEXT,                               -- name on letterhead
  visitor_name       TEXT        NOT NULL,
  person_to_meet     TEXT        NOT NULL,
  department         TEXT        NOT NULL,
  visit_date         DATE        NOT NULL,
  entry_time         TIME        NOT NULL,
  purpose            TEXT        NOT NULL,
  photo              TEXT,                               -- visitor photo (JPEG, base64 data URL)
  security_signature TEXT,                               -- security officer signature (PNG, base64 data URL)
  created_at         TIMESTAMPTZ NOT NULL DEFAULT now()  -- when the pass was saved
);

CREATE INDEX IF NOT EXISTS visitor_passes_visit_date_idx ON visitor_passes (visit_date);

-- Handy queries
-- Today's visitors:
--   SELECT pass_no, visitor_name, person_to_meet, department, entry_time, purpose
--   FROM visitor_passes WHERE visit_date = CURRENT_DATE ORDER BY entry_time;
-- Last 50 passes:
--   SELECT pass_no, visit_date, entry_time, visitor_name, person_to_meet, department, purpose
--   FROM visitor_passes ORDER BY created_at DESC LIMIT 50;
