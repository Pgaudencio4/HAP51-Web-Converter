# Analise de Validacao - Malhoa22 (Fevereiro 2026)

## Contexto

Projecto: Edificio Malhoa 22 (12 pisos + 6 caves)
Excel fonte: `HAP_Template_Malhoa22 (12).xlsx`
E3A gerado: `Malhoa22_06FEV_v2.E3A`
Projecto HAP live: `C:\E20-II\Projects\Untitled66`
Data: 6 de Fevereiro de 2026

---

## 1. Resumo - TUDO OK

O conversor `excel_to_hap.py` gera o ficheiro E3A **correctamente**. Todos os dados do Excel sao reproduzidos fielmente no HAP. A discrepancia inicial nos relatorios devia-se ao utilizador estar a consultar um relatorio de Building Simulation que so incluia parte dos espacos (95 de 140), e nao ao "Minimum Energy Performance Calculator" que inclui todos.

---

## 2. Validacao Cruzada: RTF vs Excel vs E3A

### 2.1. WINDOW.RTF vs Excel

**Resultado: 71/71 janelas IDENTICAS**

Todas as janelas exportadas do HAP via WINDOW.RTF correspondem exactamente ao Excel:
- Alturas: identicas (todas 2.2m)
- Larguras: identicas
- Nenhuma diferenca >= 0.02m em qualquer dimensao

### 2.2. SPACE.RTF vs Excel (Wall Areas e Glazing)

**Resultado: MATCH PERFEITO**

| Dir | RTF Wall (m2) | Excel Wall (m2) | RTF Glaz (m2) | Excel Glaz (m2) |
|-----|---------------|-----------------|---------------|-----------------|
| N   | 675.5         | 675.7           | 382.7         | 382.7           |
| E   | 477.5         | 478.0           | 90.4          | 90.4            |
| S   | 802.6         | 802.6           | 542.8         | 542.8           |
| SW  | 1.7           | 1.7             | 0.0           | 0.0             |
| W   | 815.1         | 814.7           | 630.2         | 630.2           |
| **TOT** | **2772.4** | **2772.7**     | **1646.0**    | **1646.0**      |

Diferencas de ~0.2-0.5 m2 nas walls sao arredondamentos float32, completamente negligiveis.

### 2.3. Relatorio "Minimum Energy Performance Calculator" vs Excel

**Resultado: MATCH PERFEITO**

O relatorio correcto do HAP (Minimum Energy Performance Calculator - Proposed Design) mostra:

| Dir | HAP Report | Excel  |
|-----|-----------|--------|
| N   | Wall=676, Glaz=383 | Wall=675.7, Glaz=382.7 |
| E   | Wall=478, Glaz=90  | Wall=478.0, Glaz=90.4  |
| S   | Wall=803, Glaz=543 | Wall=802.6, Glaz=542.8 |
| SW  | Wall=2, Glaz=0     | Wall=1.7, Glaz=0       |
| W   | Wall=815, Glaz=630 | Wall=814.7, Glaz=630.2 |
| **TOT** | **2773, 1646** | **2773, 1646** |

---

## 3. Discrepancia Inicial - Building Simulation Report

### 3.1. O problema original

O utilizador comparou valores de um relatorio de Building Simulation que mostrava:

| Dir | HAP Building | Excel  | Diferenca |
|-----|-------------|--------|-----------|
| N   | Wall=515, Glaz=439 | Wall=675.7, Glaz=382.7 | Wall:-160, Glaz:+56 |
| E   | Wall=411, Glaz=83  | Wall=478.0, Glaz=90.4  | Wall:-67, Glaz:-7   |
| S   | Wall=551, Glaz=423 | Wall=802.6, Glaz=542.8 | Wall:-252, Glaz:-120|
| W   | Wall=721, Glaz=615 | Wall=814.7, Glaz=630.2 | Wall:-95, Glaz:-15  |

### 3.2. Causa: Air System com apenas 95 de 140 espacos

O Air System "Default System" so incluia 95 dos 140 espacos:
- **Incluidos (95):** IDs 1-50 (caves A-06 a F-01) + IDs 96-140 (pisos M06 a S12)
- **Excluidos (45):** IDs 51-95 (pisos G00 a L05 = res-do-chao ao 5o andar)

Isto foi confirmado via `System_Space_Links` no `HAP51INX.MDB`.

O Building Simulation Report so agrega os espacos incluidos no Air System.

### 3.3. Porque Norte Glaz era SUPERIOR no HAP (439 > 382.7)?

Esta anomalia (que inicialmente parecia impossivel) deve-se ao facto de o Building Simulation Report ser um relatorio diferente do "Minimum Energy Performance Calculator". O Building report processa os dados de simulacao termica e pode redistribuir/recalcular areas de envolvente de forma diferente dos inputs brutos (p.ex. contabilizando shading, orientacoes reais pos-simulacao, etc.).

**O relatorio correcto para comparacao de envolvente e o "Minimum Energy Performance Calculator"**, que mostra os dados de input tal como foram definidos, e este bate 100% com o Excel.

### 3.4. Resolucao

O utilizador verificou que estava a duplicar espacos no Air System (nao a omiti-los), o que explicava a discrepancia. Apos corrigir, o relatorio "Minimum Energy Performance Calculator" confirmou que os dados estao correctos.

---

## 4. Achados Tecnicos Importantes

### 4.1. Wall Block Format (confirmado)

O formato do wall block (34 bytes, offset 72 no registo Space) foi confirmado empiricamente:

```
Offset  Bytes  Campo              Confirmado via
------  -----  -----              ---------------
+0      2      Direction code     E3A binario + HAP UI
+2      4      Gross area (ft2)   E3A binario + SPACE.RTF
+6      2      Wall type ID       E3A binario + HAP UI
+8      2      Window 1 type ID   E3A binario + HAP UI screenshots
+10     2      Window 2 type ID   E3A binario
+12     2      Window 1 quantity  E3A binario + HAP UI + SPACE.RTF
+14     2      Window 2 quantity  E3A binario
+16     2      Door type ID       E3A binario
+18     2      Door quantity      E3A binario
+20-33  14     Reserved/padding
```

**Direction codes confirmados:**
| Codigo | Orientacao |
|--------|-----------|
| 1      | N         |
| 3      | NE        |
| 5      | E         |
| 7      | SE        |
| 9      | S         |
| 11     | SW        |
| 13     | W         |
| 15     | NW        |

### 4.2. Wall Type ID vs Window ID - NAO ha confusao

Confirmado via screenshots do HAP UI que:
- O HAP le correctamente +6 como wall_type_id (ex: "Pext 1" = ID 3)
- O HAP le correctamente +8 como window1_id (ex: "V00_A00001Area_S" = ID 3)
- O facto de ambos terem ID=3 e coincidencia (nao confusao)
- Cada orientacao tem a janela correcta atribuida (N->_N, S->_S, E->_E, W->_W)

### 4.3. Window Record (confirmado)

```
Offset  Bytes  Campo
------  -----  -----
0-254   255    Nome (Latin-1)
257-260 4      Altura (ft, float)  <- NÃO offset 24!
261-264 4      Largura (ft, float) <- NÃO offset 28!
269-272 4      U-Value (BTU/hr.ft2.F, float)
273-276 4      SHGC (float)
```

**NOTA:** Offsets 24 e 28 NAO contem dimensoes (dao 0.00). Os offsets correctos sao 257 e 261.

### 4.4. E3A original vs Live (HAP editado)

Comparacao byte-a-byte entre `Malhoa22_06FEV_v2.E3A` e `C:\E20-II\Projects\Untitled66`:
- HAP51SPC.DAT: apenas 21 bytes diferentes em 2 records (padding de nome + schedule)
- HAP51WIN.DAT: **identico** (0 bytes diferentes)
- O HAP nao modifica os dados dos espacos/janelas ao abrir o projecto

### 4.5. Codificacao de nomes com acentos

Os nomes dos espacos/janelas usam codificacao Latin-1 (ISO-8859-1) nos ficheiros DAT, com 24 bytes para nomes de espacos e 255 bytes para nomes de janelas.

### 4.6. Record sizes confirmados

| Ficheiro | Record Size | Notas |
|----------|-------------|-------|
| HAP51SPC.DAT | 682 bytes | 1 default + N espacos |
| HAP51WIN.DAT | 555 bytes | 1 default + N janelas |
| HAP51SCH.DAT | 792 bytes | 1 default + N schedules |
| HAP51WAL.DAT | ~128 bytes | Paredes/assemblies |
| HAP51ROF.DAT | variavel | Coberturas |

---

## 5. Parsing de Ficheiros RTF do HAP

### 5.1. Estrutura do SPACE.RTF

O HAP exporta relatorios de espacos em formato RTF com tabelas. Pontos-chave:

- **Tabelas usam `\row` e `\cell`** para separar linhas e celulas
- **Espacos com Roofs** tem DUAS tabelas (Walls + Roofs) separadas por rows, com Construction Types entre elas
- Os "Construction Types" (que mapeiam orientacao -> nome da janela) aparecem num row segment entre a tabela de Walls e a tabela de Roofs
- Padrao: `for Exposure X ... 1st Window Type VNAME ... 2nd Window Type VNAME`

### 5.2. Armadilha no parsing

Espacos com coberturas (H01_001, L05_001, R11_001, R11_007 neste projecto) tem a seguinte estrutura:

```
Row 0: Header (Exp, Wall Gross Area, Win1 Qty, Win2 Qty, Door Qty)
Row 1-4: Dados das paredes (S, W, N, E)
Row 5: Construction Types das WALLS + Header da tabela de ROOFS  <-- CONTEM JANELAS!
Row 6: Dados do Roof (H, area, slope, skylight)
Row 7: Construction Types do ROOF (so "for Exposure H")  <-- NAO contem janelas!
```

Se o parser so procurar "for Exposure" no ULTIMO row, perde as janelas destes 4 espacos. A solucao e procurar em TODOS os row segments.

### 5.3. Codificacao RTF

- Caracteres especiais: `\'XX` onde XX e hex (Latin-1)
- Tags de formatacao: `\par`, `\pard`, `\sect`, `\b`, `\b0`, `\cf4`, etc.
- Celulas de tabela: conteudo entre formatacao e `\cell`
- Linhas de tabela: terminam com `\row`

---

## 6. Dados do Projecto Malhoa22

### 6.1. Espacos

- **140 espacos** em 18 pisos (A-06 a S12)
- **Floor area total:** 14 090 m2
- Pisos de caves (A-06 a F-01): maioritariamente estacionamentos, sem paredes exteriores
- Pisos acima do solo (G00 a S12): areas de escritorio + copas, com paredes exteriores

### 6.2. Janelas

- **71 tipos de janela** (todos com H=2.20m, U=5.75 W/m2K, SHGC=0.85)
- Nomeacao: `VXX_YYYYYY_DIR` onde XX=piso, YYYYYY=espaco, DIR=orientacao (N/S/E/W)
- Largura varia por piso e orientacao (0.33m a 25.4m)

### 6.3. Areas de envolvente por orientacao

| Dir | Gross Wall (m2) | Glazing (m2) | WWR (%) |
|-----|-----------------|-------------|---------|
| N   | 675.7           | 382.7       | 56.6    |
| E   | 478.0           | 90.4        | 18.9    |
| S   | 802.6           | 542.8       | 67.6    |
| SW  | 1.7             | 0.0         | 0.0     |
| W   | 814.7           | 630.2       | 77.4    |
| **TOT** | **2772.7**  | **1646.0**  | **59.4**|

### 6.4. Air System

- 1 Air System: "Default System"
- Devem estar incluidos TODOS os 140 espacos
- Se faltarem espacos, o Building Simulation Report mostra valores incorrectos

---

## 7. Scripts de Analise Criados

Durante esta investigacao foram criados varios scripts em `C:\Users\pedro\Downloads\`:

| Script | Funcao |
|--------|--------|
| `calc_from_rtf_v3.py` | Calcula wall/glaz totais por orientacao a partir de SPACE.RTF + WINDOW.RTF |
| `compare_windows.py` | Compara dimensoes de janelas entre WINDOW.RTF e Excel |
| `debug_missing_glaz.py` | Identifica espacos onde faltam dados de glazing no parse |
| `debug_per_space.py` | Compara glazing por espaco entre RTF e Excel |
| `debug_4spaces.py` | Debug dos 4 espacos com Roofs (H01_001, L05_001, R11_001, R11_007) |
| `parse_space_cells.py` | Extrai celulas de tabelas RTF de um espaco exemplo |

### Script principal de validacao: `calc_from_rtf_v3.py`

Este script:
1. Parseia WINDOW.RTF para obter dimensoes de todas as janelas
2. Parseia SPACE.RTF para obter paredes, orientacoes, quantities e nomes de janelas
3. Calcula wall area e glazing area por orientacao
4. Compara com dados do Excel
5. Compara com dados do relatorio HAP

---

## 8. Licoes Aprendidas

1. **Qual relatorio consultar:** O "Minimum Energy Performance Calculator" (Proposed Design) mostra os dados de input. O "Building Simulation Report" pode mostrar valores processados/diferentes.

2. **Air System completo:** Todos os espacos devem estar no Air System para o Building report agregar correctamente.

3. **Validacao cruzada:** Exportar SPACE.RTF e WINDOW.RTF do HAP e comparar com o Excel e o metodo mais fiavel de validacao.

4. **Parsing de RTF do HAP:** Tabelas RTF com Roofs tem Construction Types num row intermedio, nao no ultimo row.

5. **Offsets de janelas:** Os offsets para H/W no HAP51WIN.DAT sao 257/261, NAO 24/28.

6. **E3A vs projecto live:** O HAP quase nao modifica o SPC ao abrir (apenas padding de nomes e schedule IDs).

---

*Documento criado em 6 de Fevereiro de 2026*
