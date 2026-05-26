# ⚙️ Automação de Relatórios Operacionais

> Script Python que automatiza a geração de relatório executivo a partir de dados brutos — eliminando trabalho manual repetitivo, replicando a lógica aplicada na operação CDDNPA via Power Query.

## O que o script faz automaticamente
1. Lê e valida dados brutos
2. Calcula KPIs e indicadores mensais
3. Identifica anomalias e picos de devolução
4. Gera relatório Excel com múltiplas abas
5. Gera gráfico de acompanhamento (.png)
6. Exporta resumo executivo em texto

**Tempo manual equivalente:** ~2 horas/semana → **~8 segundos automatizados**

## Como executar
```bash
pip install pandas matplotlib openpyxl numpy
python automacao_relatorios.py
```

*Autor: Emanoel Cavalcante · [emanoelinc.github.io](https://emanoelinc.github.io)*
