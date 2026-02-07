---
description: "Validar ficheiros E3A ou Excel. Usa quando o utilizador quer validar e3a, verificar excel, check ficheiro hap, diagnosticar problemas e3a."
---

# Validar Ficheiros HAP

Valida ficheiros E3A e/ou Excel para detectar e corrigir problemas.

## Workflow

### 1. Pedir ao utilizador o que validar

Pergunta ao utilizador:
- **Ficheiro a validar** (E3A ou Excel)
- **Corrigir automaticamente?** (só para E3A)

### 2. Detectar tipo de ficheiro e validar

#### Se for um ficheiro .E3A:

```bash
python validar_e3a.py "<ficheiro.E3A>" --fix
```

Sem `--fix` para apenas reportar (sem corrigir):
```bash
python validar_e3a.py "<ficheiro.E3A>"
```

Verifica:
- Calendário dos schedules (valores inválidos >8)
- Default Space com schedule IDs != 0
- Schedule IDs inválidos nos spaces
- Estrutura do ZIP e ficheiros internos

#### Se for um ficheiro .xlsx (Excel):

```bash
python conversor/validar_excel_hap.py "<ficheiro.xlsx>"
```

Verifica campo a campo:
- Sheets obrigatórias (Espacos, Windows, Walls, Roofs, Tipos, Legenda, Schedules_RSECE)
- Valores válidos para campos enumerados (OA Unit, Activity Level, Fixture Type, etc.)
- Ranges de valores aceitáveis
- Referências a schedules, windows, walls e roofs existentes
- Consistência entre sheets

### 3. Reportar resultado

Mostra ao utilizador:
- Erros críticos (impedem conversão)
- Warnings (podem causar problemas)
- Correcções aplicadas (se usou --fix)
- Recomendações

## Notas importantes

- O validador de E3A com `--fix` modifica o ficheiro in-place (sobrescreve)
- Se o utilizador quiser manter o original, sugerir fazer cópia primeiro
- O validador de Excel não modifica o ficheiro, apenas reporta
