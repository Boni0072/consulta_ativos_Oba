/*
  # Create ativos (assets) table

  1. New Tables
    - `ativos`
      - `id` (uuid, primary key) - Unique identifier for each asset
      - `placa` (text) - License plate number (vehicle/equipment plate)
      - `numero_loja` (text) - Store number
      - `descricao` (text) - Description of the asset/item
      - `status` (text, default 'ativo') - Current status of the asset
      - `categoria` (text) - Category of the asset
      - `localizacao` (text) - Physical location
      - `data_aquisicao` (date) - Acquisition date
      - `valor` (numeric) - Value of the asset
      - `observacao` (text) - Additional notes
      - `created_at` (timestamptz) - Record creation timestamp
      - `updated_at` (timestamptz) - Record update timestamp

  2. Security
    - Enable RLS on `ativos` table
    - Add policy for authenticated users to read all assets
    - Add policy for authenticated users to insert assets
    - Add policy for authenticated users to update assets
    - Add policy for authenticated users to delete assets

  3. Indexes
    - Index on `placa` for fast filtering by plate number
    - Index on `numero_loja` for fast filtering by store number
*/

CREATE TABLE IF NOT EXISTS ativos (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  placa text NOT NULL DEFAULT '',
  numero_loja text NOT NULL DEFAULT '',
  descricao text NOT NULL DEFAULT '',
  status text NOT NULL DEFAULT 'ativo',
  categoria text DEFAULT '',
  localizacao text DEFAULT '',
  data_aquisicao date,
  valor numeric DEFAULT 0,
  depr_acum numeric DEFAULT 0,
  saldo_contabil numeric DEFAULT 0,
  observacao text DEFAULT '',
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE ativos ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view ativos"
  ON ativos FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Authenticated users can insert ativos"
  ON ativos FOR INSERT
  TO authenticated
  WITH CHECK (true);

CREATE POLICY "Authenticated users can update ativos"
  ON ativos FOR UPDATE
  TO authenticated
  USING (true)
  WITH CHECK (true);

CREATE POLICY "Authenticated users can delete ativos"
  ON ativos FOR DELETE
  TO authenticated
  USING (true);

-- Allow anonymous reads for this public-facing app
CREATE POLICY "Anyone can view ativos"
  ON ativos FOR SELECT
  TO anon
  USING (true);

-- Allow anonymous inserts
CREATE POLICY "Anyone can insert ativos"
  ON ativos FOR INSERT
  TO anon
  WITH CHECK (true);

-- Allow anonymous updates
CREATE POLICY "Anyone can update ativos"
  ON ativos FOR UPDATE
  TO anon
  USING (true)
  WITH CHECK (true);

-- Allow anonymous deletes
CREATE POLICY "Anyone can delete ativos"
  ON ativos FOR DELETE
  TO anon
  USING (true);

CREATE INDEX IF NOT EXISTS idx_ativos_placa ON ativos (placa);
CREATE INDEX IF NOT EXISTS idx_ativos_numero_loja ON ativos (numero_loja);

-- Insert sample data
INSERT INTO ativos (placa, numero_loja, descricao, status, categoria, localizacao, data_aquisicao, valor, observacao) VALUES
  ('ABC-1234', 'LOJA-01', 'Computador Desktop Dell', 'ativo', 'Informática', 'Sala 101', '2024-01-15', 3500.00, 'Equipamento em bom estado'),
  ('DEF-5678', 'LOJA-02', 'Impressora HP LaserJet', 'ativo', 'Impressão', 'Sala 202', '2024-03-20', 1200.00, 'Necessita manutenção preventiva'),
  ('GHI-9012', 'LOJA-01', 'Notebook Lenovo ThinkPad', 'ativo', 'Informática', 'Sala 103', '2024-02-10', 4800.00, 'Com garantia até 2026'),
  ('JKL-3456', 'LOJA-03', 'Monitor Samsung 27"', 'inativo', 'Informática', 'Depósito', '2023-06-05', 1800.00, 'Tela com defeito'),
  ('MNO-7890', 'LOJA-02', 'Cadeira Ergonômica', 'ativo', 'Mobiliário', 'Sala 202', '2024-04-01', 900.00, ''),
  ('PQR-2345', 'LOJA-01', 'Projetor Epson', 'ativo', 'Apresentação', 'Auditório', '2024-01-25', 2800.00, 'Lâmpada substituída recentemente'),
  ('STU-6789', 'LOJA-04', 'Servidor Dell PowerEdge', 'ativo', 'Infraestrutura', 'Data Center', '2023-11-15', 15000.00, 'Crítico para operações'),
  ('VWX-0123', 'LOJA-03', 'Ar Condicionado Split 18000 BTU', 'ativo', 'Climatização', 'Sala 301', '2024-05-10', 2200.00, 'Manutenção anual agendada'),
  ('YZA-4567', 'LOJA-02', 'Scanner Fujitsu', 'ativo', 'Digitalização', 'Sala 205', '2024-02-28', 1600.00, ''),
  ('BCD-8901', 'LOJA-01', 'UPS Nobreak 1500VA', 'ativo', 'Energia', 'Data Center', '2023-09-20', 750.00, 'Bateria substituída em 2024');
