/*
  # Add depreciation and accounting balance columns

  1. New Columns
    - `depr_acum` (numeric) - Accumulated depreciation
    - `saldo_contabil` (numeric) - Accounting balance (residual value)
  
  2. Changes
    - Add depr_acum column to ativos table
    - Add saldo_contabil column to ativos table
    - Both default to 0
*/

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'ativos' AND column_name = 'depr_acum'
  ) THEN
    ALTER TABLE ativos ADD COLUMN depr_acum numeric DEFAULT 0;
  END IF;
  
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'ativos' AND column_name = 'saldo_contabil'
  ) THEN
    ALTER TABLE ativos ADD COLUMN saldo_contabil numeric DEFAULT 0;
  END IF;
END $$;
