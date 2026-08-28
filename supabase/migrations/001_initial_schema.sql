-- ============================================================
-- P&L Audit Desk — Full Database Migration
-- Run this ONCE in the Supabase SQL Editor:
-- https://supabase.com/dashboard/project/xsgcpsoxdrshlejrcbgv/sql/new
-- ============================================================

-- Enable UUID generation
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- ============================================================
-- TABLE: clients
-- One row per accounting client (restaurant, hotel, etc.)
-- ============================================================
CREATE TABLE IF NOT EXISTS clients (
  id                    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  display_name          TEXT NOT NULL,
  knowledge_initialized BOOLEAN NOT NULL DEFAULT FALSE,
  initialized_at        TIMESTAMPTZ,
  init_file_name        TEXT,
  init_row_count        INTEGER,
  init_item_count       INTEGER,
  authoritative_column  TEXT CHECK (authoritative_column IN ('account', 'purchase_account')),
  created_at            TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ============================================================
-- TABLE: account_heads
-- Chart of Accounts per client — populated from historical
-- sheet or entered manually.
-- ============================================================
CREATE TABLE IF NOT EXISTS account_heads (
  id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id   UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
  name        TEXT NOT NULL,
  category    TEXT,         -- e.g. 'food', 'beverages', 'cleaning', etc.
  is_active   BOOLEAN NOT NULL DEFAULT TRUE,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (client_id, name)
);

-- ============================================================
-- TABLE: client_item_knowledge
-- The persistent, client-specific knowledge base.
-- This is the PRIMARY source of truth for classification.
-- ============================================================
CREATE TABLE IF NOT EXISTS client_item_knowledge (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id         UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
  item_name_raw     TEXT NOT NULL,
  item_name_norm    TEXT NOT NULL,          -- normalized for matching (lowercase, no special chars)
  account_head      TEXT NOT NULL,           -- the verified/authoritative account
  purchase_account  TEXT,                    -- raw purchase_account column value if present
  conflict_flag     BOOLEAN NOT NULL DEFAULT FALSE,  -- true if account != purchase_account
  source            TEXT NOT NULL DEFAULT 'historical_import'
                    CHECK (source IN ('historical_import', 'human_approved', 'rule_engine')),
  verified          BOOLEAN NOT NULL DEFAULT TRUE,
  confidence        INTEGER NOT NULL DEFAULT 100 CHECK (confidence BETWEEN 0 AND 100),
  notes             TEXT,
  created_at        TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at        TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (client_id, item_name_norm)        -- one canonical mapping per client per item
);

-- Index for fast lookup during audit matching
CREATE INDEX IF NOT EXISTS idx_knowledge_client_norm
  ON client_item_knowledge (client_id, item_name_norm);

-- ============================================================
-- TABLE: client_periods
-- One row per monthly audit session uploaded for a client
-- ============================================================
CREATE TABLE IF NOT EXISTS client_periods (
  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id    UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
  period_label TEXT NOT NULL,     -- e.g. "August 2026"
  file_name    TEXT NOT NULL,
  row_count    INTEGER NOT NULL DEFAULT 0,
  status       TEXT NOT NULL DEFAULT 'active'
               CHECK (status IN ('active', 'completed')),
  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ============================================================
-- TABLE: imports
-- Raw ledger rows imported per period — the source transactions
-- ============================================================
CREATE TABLE IF NOT EXISTS imports (
  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id    UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
  period_id    UUID NOT NULL REFERENCES client_periods(id) ON DELETE CASCADE,
  import_date  TEXT,
  vendor       TEXT,
  bill_no      TEXT,
  account      TEXT,
  item_name    TEXT,
  qty          NUMERIC,
  rate         NUMERIC,
  total        NUMERIC,
  branch       TEXT,
  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_imports_client_period
  ON imports (client_id, period_id);

-- ============================================================
-- TABLE: classification_history
-- Every classification decision made during an audit.
-- Tracks: what was known, what rule fired, what AI said, 
-- and what the accountant finally approved.
-- ============================================================
CREATE TABLE IF NOT EXISTS classification_history (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id           UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
  period_id           UUID NOT NULL REFERENCES client_periods(id) ON DELETE CASCADE,
  item_name           TEXT NOT NULL,
  vendor              TEXT,
  detected_account    TEXT,        -- what account was used in the ledger
  knowledge_match     BOOLEAN NOT NULL DEFAULT FALSE,
  knowledge_account   TEXT,        -- what the KB says it should be
  rule_match          BOOLEAN NOT NULL DEFAULT FALSE,
  rule_account        TEXT,        -- what the rule engine says
  ai_called           BOOLEAN NOT NULL DEFAULT FALSE,
  ai_account          TEXT,        -- what Groq AI says
  ai_confidence       INTEGER,
  final_account       TEXT,        -- the final decided account head
  is_misclassified    BOOLEAN NOT NULL DEFAULT FALSE,
  human_reviewed      BOOLEAN NOT NULL DEFAULT FALSE,
  human_approved      BOOLEAN,     -- null = not reviewed yet
  saved_to_knowledge  BOOLEAN NOT NULL DEFAULT FALSE,
  total_amount        NUMERIC,
  created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_classification_client_period
  ON classification_history (client_id, period_id);

CREATE INDEX IF NOT EXISTS idx_classification_pending_review
  ON classification_history (client_id, human_reviewed, human_approved)
  WHERE human_reviewed = FALSE AND is_misclassified = TRUE;

-- ============================================================
-- Row Level Security (RLS)
-- All tables: service_role bypasses RLS (for API routes)
-- anon key: read-only access for now (no auth in Phase 1)
-- ============================================================
ALTER TABLE clients                 ENABLE ROW LEVEL SECURITY;
ALTER TABLE account_heads           ENABLE ROW LEVEL SECURITY;
ALTER TABLE client_item_knowledge   ENABLE ROW LEVEL SECURITY;
ALTER TABLE client_periods          ENABLE ROW LEVEL SECURITY;
ALTER TABLE imports                 ENABLE ROW LEVEL SECURITY;
ALTER TABLE classification_history  ENABLE ROW LEVEL SECURITY;

-- Permissive policies for anon (Phase 1 — no auth yet)
CREATE POLICY "anon_all_clients"               ON clients               FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all_account_heads"         ON account_heads         FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all_knowledge"             ON client_item_knowledge FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all_periods"               ON client_periods        FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all_imports"               ON imports               FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all_classification"        ON classification_history FOR ALL TO anon USING (true) WITH CHECK (true);

-- ============================================================
-- Auto-update updated_at on clients + knowledge rows
-- ============================================================
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_clients_updated_at
  BEFORE UPDATE ON clients
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_knowledge_updated_at
  BEFORE UPDATE ON client_item_knowledge
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- ============================================================
-- DONE. All tables created.
-- ============================================================
