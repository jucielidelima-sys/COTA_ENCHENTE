# Dashboard de Risco de Enchente

Dashboard em Python/Streamlit para monitorar colaboradores impactados por nível do rio, setor, acesso à empresa e meio de transporte.

## Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estrutura

- `app.py`: dashboard principal.
- `base/COTA_ENCHENTE.xlsx`: base de dados padrão.
- `requirements.txt`: dependências.
- `.streamlit/config.toml`: configuração visual.

## Correções desta versão

- Leitura automática da base Excel mesmo no Streamlit Cloud.
- Campo lateral para atualizar/substituir a base Excel.
- Cálculo acumulado de pessoas impactadas por nível do rio.
- Matriz de Prioridade Operacional P1/P2/P3/P4.
- Sem dependência do `xlsxwriter`.
