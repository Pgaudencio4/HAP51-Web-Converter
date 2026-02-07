---
description: "Extrair E3A para Excel. Usa quando o utilizador quer exportar um ficheiro HAP para Excel, extrair dados de e3a, ver o conteúdo de um e3a, analisar e3a."
---

# Extrair E3A para Excel

Exporta todos os dados de um ficheiro HAP 5.1 (.E3A) para um Excel organizado.

## Workflow

### 1. Pedir ao utilizador os ficheiros

Pergunta ao utilizador:
- **Ficheiro E3A de input**
- **Nome do Excel de output** (opcional, por defeito usa o mesmo nome com extensão .xlsx)

### 2. Executar a extracção

```bash
python extractor/hap_extractor.py "<input.E3A>" "<output.xlsx>"
```

Se o utilizador não especificar o output, o segundo argumento é opcional e o script gera automaticamente.

### 3. Reportar resultado

O Excel gerado tem 4 folhas:
- **Espacos**: Todos os 147 campos de cada espaço (área, pé-direito, paredes, coberturas, etc.)
- **Windows**: Nome, U-Value, SHGC, Altura, Largura
- **Walls**: Nome, U-Value, Espessura, Massa
- **Roofs**: Nome, U-Value, Espessura, Massa

Mostra ao utilizador:
- Quantos espaços/janelas/paredes/coberturas foram extraídos
- Confirma que o Excel foi criado com sucesso

## Notas importantes

- Todos os valores são convertidos de Imperial (interno) para SI (métrico) automaticamente
- O formato do Excel é compatível com o template `HAP_Template_RSECE.xlsx`
- O extractor lê schedules, assemblies e todos os tipos de registos
