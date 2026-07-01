# Dashboard de Risco de Enchente - Industrial 4.0

Dashboard em Python/Streamlit para monitorar colaboradores impactados por nível de enchente.

## Recursos

- KPIs de colaboradores, impacto, acesso e setores críticos
- Simulação do nível atual da enchente
- Curva de pessoas impactadas por nível do rio
- Gráficos por setor, cota, transporte e acesso
- Matriz de Prioridade Operacional
- Ranking de acionamento por setor
- Tabela pesquisável e exportação da base filtrada em Excel

## Matriz de Prioridade Operacional

- **P1 - Ação imediata:** impacto imediato na cota atual ou N/P com acesso NÃO
- **P2 - Alta atenção:** cota até 1 metro acima do nível atual e acesso NÃO
- **P3 - Monitorar:** cota até 2 metros acima do nível atual e acesso NÃO
- **P4 - Rotina:** sem ação imediata

## Como executar

```bash
pip install -r requirements.txt
streamlit run app.py
```

O arquivo Excel base já está incluído na pasta `base`.
