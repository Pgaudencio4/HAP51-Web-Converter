---
description: "Comparar dois E3A lado a lado. Usa quando o utilizador quer comparar e3a, ver diferenças entre previsto e referência, comparar projectos hap, diff e3a."
---

# Comparar Dois E3A

Compara dois ficheiros E3A (ex: Previsto vs Referência) e gera um Excel com as diferenças.

## Workflow

### 1. Pedir ao utilizador os ficheiros

Pergunta ao utilizador:
- **Ficheiro E3A 1** (ex: Previsto)
- **Ficheiro E3A 2** (ex: Referência)
- **Nome do Excel de comparação** (opcional)

### 2. Extrair ambos os E3A para Excel

Executa as duas extracções (podem correr em paralelo):

```bash
python extractor/hap_extractor.py "<e3a_1>" "<temp_extract_1.xlsx>"
```

```bash
python extractor/hap_extractor.py "<e3a_2>" "<temp_extract_2.xlsx>"
```

### 3. Executar a comparação

```bash
python comparador/comparar_com_template.py comparador/Template_Comparacao_v7.xlsx "<temp_extract_1.xlsx>" "<temp_extract_2.xlsx>" "<output_comparacao.xlsx>"
```

O template de comparação é sempre `comparador/Template_Comparacao_v7.xlsx`.

### 4. Limpar ficheiros temporários

Remove os Excels temporários de extracção (temp_extract_1.xlsx e temp_extract_2.xlsx).

### 5. Reportar resultado

O Excel de comparação tem cores:
- **Verde (OK)**: Valores iguais entre os dois ficheiros
- **Vermelho (DIFF)**: Valores diferentes
- **F1/F2**: Valor só existe num dos ficheiros

Compara:
- Espaços (147 campos × 3 colunas: F1, F2, Check)
- Windows (Nome, U-Value, SHGC, Dimensões)
- Walls (Nome, U-Value, Espessura, Massa)
- Roofs (Nome, U-Value, Espessura, Massa)

Mostra um resumo das diferenças encontradas.

## Notas importantes

- O template de comparação formatado deve existir em `comparador/Template_Comparacao_v7.xlsx`
- Se não existir, pode usar `comparador/criar_template_v7.py` para o gerar
- Os ficheiros temporários devem ser limpos após comparação
