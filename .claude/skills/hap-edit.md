---
description: "Editar E3A existente sem perder sistemas AVAC. Usa quando o utilizador quer modificar um e3a, editar espaços, alterar valores num hap existente, actualizar e3a."
---

# Editar E3A Existente

Modifica campos de um E3A existente **sem perder sistemas AVAC, schedules ou resultados de simulação**.

## Workflow

O editor tem duas fases: EXTRAIR e APLICAR.

### Fase 1: Extrair E3A para Excel de edição

Pergunta ao utilizador:
- **Ficheiro E3A original**
- **Nome do Excel de edição** (opcional)

```bash
python editor/editor_e3a.py extrair "<ficheiro.E3A>" "<editor.xlsx>"
```

Explica ao utilizador:
- O Excel gerado tem colunas **PREV** (vazio) e **REF** (valores actuais)
- Preencher a coluna **PREV** apenas com os valores que quer alterar
- Deixar **PREV vazio** = campo não é alterado
- Suporta espaços, windows, walls e roofs

Aguarda que o utilizador edite o Excel e avise que está pronto.

### Fase 2: Aplicar alterações

Pergunta ao utilizador:
- **Ficheiro E3A original** (o mesmo da fase 1)
- **Excel de edição** (o que editou na fase 1)
- **Nome do E3A de output** (o ficheiro modificado)

```bash
python editor/editor_e3a.py aplicar "<original.E3A>" "<editor.xlsx>" "<output.E3A>"
```

### 3. Reportar resultado

Mostra ao utilizador:
- Quantos campos foram alterados
- Que o ficheiro mantém todos os sistemas AVAC intactos
- Confirma que o novo E3A está pronto

## Notas importantes

- NUNCA sobrescrever o E3A original — usar sempre um nome diferente para output
- O editor usa `data_only=True` ao ler o Excel, por isso fórmulas Excel são avaliadas correctamente
- Suporta edição de: Espaços (147 campos), Windows (U-Value, SHGC, Altura, Largura), Walls (U-Value, Espessura, Massa), Roofs (U-Value, Espessura, Massa)
