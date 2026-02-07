---
description: "Adaptar Excel HAP 5.2 para formato standard. Usa quando o utilizador tem um excel no formato hap 5.2, converter formato 5.2, adaptar folha hap 5.2."
---

# Adaptar Excel HAP 5.2 para Formato Standard

Converte um Excel no formato HAP 5.2 para o formato standard do template, compatível com o conversor.

## Workflow

### 1. Pedir ao utilizador os ficheiros

Pergunta ao utilizador:
- **Ficheiro Excel no formato HAP 5.2** (com abas INPUT SPACES HAP, INPUT WALLS HAP, etc.)
- **Nome do Excel de output** (opcional)

### 2. Executar a adaptação

```bash
python adaptador/adapter_hap52.py "<input_hap52.xlsx>" "<output.xlsx>"
```

Se o utilizador não especificar output, o script gera automaticamente.

### 3. Reportar resultado e próximo passo

O Excel gerado tem o formato standard (Espacos, Walls, Roofs, Windows) e pode ser usado directamente com o conversor:

```bash
python conversor/excel_to_hap.py "<output.xlsx>" conversor/templates/Modelo_RSECE.E3A "<projecto.E3A>"
```

Sugere ao utilizador que use `/hap-converter` para o passo seguinte.

## Formato de entrada esperado (HAP 5.2)

O Excel de input deve ter estas abas:
- **INPUT SPACES HAP** (147 colunas, mesma estrutura)
- **INPUT WALLS HAP** (Nome, U-Value, Peso, Espessura)
- **INPUT ROOFS HAP** (Nome, U-Value, Peso, Espessura)
- **INPUT VIDROS HAP** (Nome, U-Value, SHGC, Altura, Largura)
