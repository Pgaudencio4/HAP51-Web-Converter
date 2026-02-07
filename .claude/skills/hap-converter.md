---
description: "Converter Excel para E3A (HAP 5.1). Usa quando o utilizador quer criar um ficheiro HAP a partir de Excel, converter excel para e3a, gerar e3a, criar projecto hap."
---

# Converter Excel para E3A (HAP 5.1)

Converte um ficheiro Excel preenchido com o template RSECE para um ficheiro HAP 5.1 (.E3A).

## Workflow

### 1. Pedir ao utilizador os ficheiros

Pergunta ao utilizador:
- **Ficheiro Excel de input** (o Excel preenchido com os dados do projecto)
- **Nome do ficheiro E3A de output** (opcional, por defeito usa o mesmo nome com extensão .E3A)

### 2. Validar o Excel antes de converter

```bash
python conversor/validar_excel_hap.py "<input.xlsx>"
```

Analisa o resultado. Se houver erros críticos, mostra ao utilizador e pergunta se quer continuar.
Se houver apenas warnings, informa e continua.

### 3. Converter Excel para E3A

```bash
python conversor/excel_to_hap.py "<input.xlsx>" conversor/templates/Modelo_RSECE.E3A "<output.E3A>"
```

O template base é sempre `conversor/templates/Modelo_RSECE.E3A` (contém os 82 schedules RSECE).

### 4. Validar o E3A gerado

```bash
python validar_e3a.py "<output.E3A>" --fix
```

O `--fix` corrige automaticamente problemas conhecidos (calendários inválidos, default space, etc).

### 5. Reportar resultado

Mostra ao utilizador:
- Quantos espaços foram criados
- Quantas janelas/paredes/coberturas
- Se houve erros ou warnings
- Confirma que o ficheiro está pronto para abrir no HAP 5.1

## Notas importantes

- Todos os caminhos são relativos à raiz do projecto HAPPXXXX
- O Excel deve seguir o formato do template `conversor/templates/HAP_Template_RSECE.xlsx` (147 colunas)
- O conversor faz todas as conversões de unidades SI → Imperial automaticamente
- Os schedules RSECE (82) vêm do modelo base, não do Excel
