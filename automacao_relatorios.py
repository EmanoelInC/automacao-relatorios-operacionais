"""
Automação de Relatórios Operacionais
=====================================
Autor: Emanoel Cavalcante

Script que automatiza a geração de relatório executivo
a partir de dados brutos — eliminando trabalho manual
repetitivo, exatamente como feito na operação CDDNPA
via Power Query e Excel.

O que o script faz automaticamente:
 1. Lê dados brutos (CSV simulado em memória)
 2. Limpa e valida os dados
 3. Calcula KPIs e indicadores
 4. Identifica anomalias e alertas
 5. Gera relatório executivo em Excel (.xlsx)
 6. Gera gráfico de acompanhamento (.png)
 7. Exporta resumo em texto (.txt)

Tempo manual equivalente: ~2 horas/semana
Tempo automatizado: ~8 segundos
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

np.random.seed(7)

print("=" * 60)
print("  AUTOMAÇÃO DE RELATÓRIOS OPERACIONAIS")
print("  Emanoel Cavalcante · emanoelinc.github.io")
print("=" * 60)

# ──────────────────────────────────────────────
# 1. GERAÇÃO DE DADOS BRUTOS (simula input real)
# ──────────────────────────────────────────────

print("\n[1/6] Lendo dados brutos...")

datas = pd.date_range('2025-01-01', '2025-05-31', freq='D')
n = len(datas)

df_raw = pd.DataFrame({
    'data':            datas,
    'tons_expedidas':  np.random.normal(60000, 8000, n).clip(20000, 95000),
    'tons_devolvidas': np.random.normal(6000, 2000, n).clip(0, 25000),
    'veiculos_exp':    np.random.randint(40, 70, n),
    'veiculos_rec':    np.random.randint(3, 10, n),
    'avaria_valor':    np.random.choice([0]*18 + [0.02, 0.05], n),
    'sla_exp_pct':     np.random.normal(99.2, 1.2, n).clip(90, 100),
    'sla_rec_pct':     np.random.normal(99.5, 0.8, n).clip(92, 100),
    'temperatura_ok':  np.random.choice([True]*19 + [False], n),
    'cliente':         np.random.choice(['Ferrero', 'Danone', 'Nutricia', 'Colog'], n),
})

# Injeta picos reais conhecidos
df_raw.loc[df_raw['data'] == '2025-05-02', 'tons_devolvidas'] *= 3.2
df_raw.loc[df_raw['data'] == '2025-05-21', 'tons_devolvidas'] *= 2.1

print(f"   → {len(df_raw)} registros carregados ({df_raw['data'].min().date()} a {df_raw['data'].max().date()})")

# ──────────────────────────────────────────────
# 2. LIMPEZA E VALIDAÇÃO
# ──────────────────────────────────────────────

print("\n[2/6] Limpando e validando dados...")

antes = len(df_raw)
df = df_raw.copy()
df['data'] = pd.to_datetime(df['data'])
df['mes']  = df['data'].dt.to_period('M')
df['pct_devolucao'] = (df['tons_devolvidas'] / df['tons_expedidas'] * 100).round(2)

# Flag registros com % devolução acima do SLA
df['alerta_devolucao'] = df['pct_devolucao'] > 10

nulos = df.isnull().sum().sum()
print(f"   → Nulos encontrados: {nulos}")
print(f"   → Registros válidos: {len(df)}/{antes}")
print(f"   → Alertas de devolução: {df['alerta_devolucao'].sum()} dias acima de 10%")

# ──────────────────────────────────────────────
# 3. CÁLCULO DE KPIs
# ──────────────────────────────────────────────

print("\n[3/6] Calculando KPIs...")

mensal = df.groupby('mes').agg(
    tons_expedidas   = ('tons_expedidas',  'sum'),
    tons_devolvidas  = ('tons_devolvidas', 'sum'),
    veiculos_exp     = ('veiculos_exp',    'sum'),
    avaria_total     = ('avaria_valor',    'sum'),
    sla_exp_media    = ('sla_exp_pct',     'mean'),
    sla_rec_media    = ('sla_rec_pct',     'mean'),
    dias_alerta      = ('alerta_devolucao','sum'),
).reset_index()

mensal['pct_devolucao_mes'] = (mensal['tons_devolvidas'] / mensal['tons_expedidas'] * 100).round(2)
mensal['sla_exp_media']     = mensal['sla_exp_media'].round(2)
mensal['sla_rec_media']     = mensal['sla_rec_media'].round(2)
mensal['tons_expedidas']    = mensal['tons_expedidas'].round(0).astype(int)
mensal['tons_devolvidas']   = mensal['tons_devolvidas'].round(0).astype(int)
mensal['status_avaria']     = mensal['avaria_total'].apply(
    lambda x: '✓ R$0,00' if x == 0 else f'⚠ R${x:.2f}'
)
mensal['status_sla_exp']    = mensal['sla_exp_media'].apply(
    lambda x: '✓ OK' if x >= 95 else '✗ Abaixo'
)

print(f"   → {len(mensal)} meses processados")
for _, row in mensal.iterrows():
    print(f"   {row['mes']}: {row['tons_expedidas']:,} tons | SLA {row['sla_exp_media']}% | Dev {row['pct_devolucao_mes']}%")

# ──────────────────────────────────────────────
# 4. IDENTIFICAÇÃO DE ANOMALIAS
# ──────────────────────────────────────────────

print("\n[4/6] Identificando anomalias...")

picos = df[df['pct_devolucao'] > 15].sort_values('pct_devolucao', ascending=False)
print(f"   → {len(picos)} dias com devolução > 15%:")
for _, row in picos.head(5).iterrows():
    print(f"      {row['data'].date()} → {row['pct_devolucao']:.1f}% ({row['cliente']})")

temp_nok = df[~df['temperatura_ok']]
print(f"   → {len(temp_nok)} registros com temperatura fora do range")

# ──────────────────────────────────────────────
# 5. EXPORTAR EXCEL
# ──────────────────────────────────────────────

print("\n[5/6] Gerando relatório Excel...")

try:
    with pd.ExcelWriter('relatorio_operacional.xlsx', engine='openpyxl') as writer:
        mensal.to_excel(writer, sheet_name='Resumo Mensal', index=False)
        df[['data', 'cliente', 'tons_expedidas', 'tons_devolvidas',
            'pct_devolucao', 'alerta_devolucao', 'sla_exp_pct',
            'avaria_valor', 'temperatura_ok']].to_excel(
            writer, sheet_name='Dados Diários', index=False)
        picos.head(20).to_excel(writer, sheet_name='Alertas', index=False)
    print("   → relatorio_operacional.xlsx gerado ✓")
except Exception as e:
    print(f"   → Excel não gerado (instale openpyxl): {e}")

# ──────────────────────────────────────────────
# 6. GRÁFICO DE ACOMPANHAMENTO
# ──────────────────────────────────────────────

print("\n[6/6] Gerando gráfico...")

import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec

BG    = '#0a0d14'; SURF = '#111520'; VERDE = '#00d4a4'
AZUL  = '#4f7fff'; AMAR = '#fbbf24'; VERM  = '#ff6b6b'
TEXTO = '#e8ecf5'; MUT  = '#8290aa'

fig, axes = plt.subplots(2, 2, figsize=(16, 10), facecolor=BG)
fig.suptitle('RELATÓRIO OPERACIONAL — CDDNPA (Jan–Mai 2025)\nGerado automaticamente por automacao_relatorios.py',
             color=TEXTO, fontsize=13, fontweight='bold', y=0.98)

meses_str = [str(m) for m in mensal['mes']]

# Tons expedidas
ax = axes[0, 0]; ax.set_facecolor(SURF)
ax.bar(meses_str, mensal['tons_expedidas']/1000, color=AZUL, alpha=0.8)
ax.set_title('Tons Expedidas (mil)', color=TEXTO, fontsize=10)
ax.tick_params(colors=MUT, labelsize=8)
[s.set_color('#1e2740') for s in ax.spines.values()]

# % Devolução
ax = axes[0, 1]; ax.set_facecolor(SURF)
cores = [VERM if v > 8 else AMAR if v > 5 else VERDE for v in mensal['pct_devolucao_mes']]
ax.bar(meses_str, mensal['pct_devolucao_mes'], color=cores)
ax.axhline(5, color=VERDE, linestyle='--', linewidth=1.5, label='SLA 5%')
ax.set_title('% Devolução Mensal', color=TEXTO, fontsize=10)
ax.tick_params(colors=MUT, labelsize=8)
ax.legend(fontsize=8, facecolor=SURF, labelcolor=MUT, edgecolor='#1e2740')
[s.set_color('#1e2740') for s in ax.spines.values()]

# SLA Expedição
ax = axes[1, 0]; ax.set_facecolor(SURF)
ax.plot(meses_str, mensal['sla_exp_media'], color=VERDE, marker='o', linewidth=2.5, markersize=7)
ax.axhline(95, color=AMAR, linestyle='--', linewidth=1.2, label='Meta 95%')
ax.set_ylim(90, 101)
ax.set_title('SLA Expedição % (Meta ≥ 95%)', color=TEXTO, fontsize=10)
ax.tick_params(colors=MUT, labelsize=8)
ax.legend(fontsize=8, facecolor=SURF, labelcolor=MUT, edgecolor='#1e2740')
[s.set_color('#1e2740') for s in ax.spines.values()]

# Dias de alerta
ax = axes[1, 1]; ax.set_facecolor(SURF)
ax.bar(meses_str, mensal['dias_alerta'], color=VERM, alpha=0.8)
ax.set_title('Dias com Devolução > 10% (Alertas)', color=TEXTO, fontsize=10)
ax.tick_params(colors=MUT, labelsize=8)
[s.set_color('#1e2740') for s in ax.spines.values()]

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig('relatorio_operacional.png', dpi=150, bbox_inches='tight',
            facecolor=BG, edgecolor='none')
print("   → relatorio_operacional.png gerado ✓")

# ──────────────────────────────────────────────
# 7. RESUMO EM TEXTO
# ──────────────────────────────────────────────

resumo = f"""
╔══════════════════════════════════════════════════════════╗
║         RESUMO EXECUTIVO — OPERAÇÃO CDDNPA               ║
║         Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}                       ║
╠══════════════════════════════════════════════════════════╣
║  Período analisado : Jan/2025 – Mai/2025                 ║
║  Total de registros: {len(df)} dias                           ║
╠══════════════════════════════════════════════════════════╣
║  EXPEDIÇÃO                                               ║
║  Total tons expedidas : {df['tons_expedidas'].sum():>15,.0f}             ║
║  Média SLA expedição  : {df['sla_exp_pct'].mean():>14.2f}%             ║
║  Total veículos exp.  : {df['veiculos_exp'].sum():>15,}             ║
╠══════════════════════════════════════════════════════════╣
║  DEVOLUÇÃO                                               ║
║  Total tons devolvidas: {df['tons_devolvidas'].sum():>15,.0f}             ║
║  Média % devolução    : {df['pct_devolucao'].mean():>14.2f}%             ║
║  Dias acima de 10%    : {df['alerta_devolucao'].sum():>15}             ║
╠══════════════════════════════════════════════════════════╣
║  AVARIA                                                  ║
║  Total avaria (R$)    : R${df['avaria_valor'].sum():>13.2f}             ║
║  Meta contratual      :       ≤ R$0,10/ton               ║
╚══════════════════════════════════════════════════════════╝
"""

with open('resumo_executivo.txt', 'w', encoding='utf-8') as f:
    f.write(resumo)

print(resumo)
print("   → resumo_executivo.txt salvo ✓")
print("\n✅ Automação concluída com sucesso!")
print(f"   Arquivos gerados: relatorio_operacional.xlsx | .png | resumo_executivo.txt")
