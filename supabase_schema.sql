-- ==========================================
-- Script de criação da tabela no Supabase
-- ==========================================
-- Execute este SQL no SQL Editor do painel do Supabase
-- (https://supabase.com/dashboard → seu projeto → SQL Editor)
-- ==========================================

-- 1. Criar a tabela de validações
CREATE TABLE IF NOT EXISTS validacoes (
    id BIGSERIAL PRIMARY KEY,
    timestamp TIMESTAMPTZ DEFAULT NOW(),
    data_validacao TEXT,
    hora_validacao TEXT,
    operador TEXT,
    serial TEXT NOT NULL,
    tecnico TEXT,
    tecnico_vicky TEXT,
    codigo_produto TEXT,
    sku TEXT,
    descricao TEXT,
    classificacao_sgm TEXT,
    estado_vicky TEXT,
    tipo_material TEXT,
    status TEXT NOT NULL,
    categoria TEXT NOT NULL,
    motivo TEXT,
    alerta_classificacao TEXT,
    origem_planilha TEXT
);

-- 2. Criar índices para pesquisas rápidas
CREATE INDEX IF NOT EXISTS idx_validacoes_serial ON validacoes(serial);
CREATE INDEX IF NOT EXISTS idx_validacoes_status ON validacoes(status);
CREATE INDEX IF NOT EXISTS idx_validacoes_timestamp ON validacoes(timestamp);
CREATE INDEX IF NOT EXISTS idx_validacoes_operador ON validacoes(operador);

-- 3. Habilitar Row Level Security (RLS) — IMPORTANTE para segurança
ALTER TABLE validacoes ENABLE ROW LEVEL SECURITY;

-- 4. Política: Permitir INSERT para usuários autenticados ou com chave anon
CREATE POLICY "Permitir inserção de validações"
ON validacoes FOR INSERT
TO anon, authenticated
WITH CHECK (true);

-- 5. Política: Permitir leitura de todas as validações
CREATE POLICY "Permitir leitura de validações"
ON validacoes FOR SELECT
TO anon, authenticated
USING (true);

-- 6. (Opcional) Visualização para métricas agregadas — evita carregar todas as linhas
CREATE OR REPLACE VIEW validacoes_metricas AS
SELECT
    COUNT(*) AS total,
    COUNT(*) FILTER (WHERE status = 'LIBERADO') AS aprovados,
    COUNT(*) FILTER (WHERE status = 'BLOQUEADO') AS bloqueados,
    COUNT(DISTINCT serial) FILTER (WHERE status = 'LIBERADO') AS seriais_unicos_liberados
FROM validacoes;
