# Apontamento de Produção

Sistema mobile de apontamento diário de produção industrial.

## Como usar

Abra o arquivo `index.html` no navegador do celular ou computador.

### Funcionalidades
- 📋 **Pedidos** — lista todas as ordens de produção com status e progresso
- 📝 **Registros** — histórico de apontamentos por turno e operador
- 📊 **Dashboard** — resumo executivo com métricas
- ⬇️ **Exportar** — gera arquivo CSV compatível com Excel

### Status das OPs
- 🟢 No prazo — dentro do prazo de entrega
- 🟡 Em risco — vence em até 2 dias
- 🔴 Atrasado — prazo vencido
- ✅ Concluído — 100% produzido

## Exportação Excel (.xlsx)

Para gerar o relatório Excel formatado, instale o Python e rode:

```bash
pip install openpyxl
echo '{"pedidos":[],"apontamentos":[]}' | python exportar.py
```

## Tecnologias
- HTML + CSS + JavaScript puro
- Dados salvos no localStorage do navegador
- Sem dependências externas
