# Historico de Desenvolvimento - HAP 5.1 Tools

Este documento resume todo o trabalho de desenvolvimento realizado neste projecto.

---

## Resumo do Projecto

**Objectivo:** Criar ferramentas Python para automatizar a criacao de ficheiros HAP 5.1 (.E3A) a partir de templates Excel, com perfis RSECE pre-configurados.

**Data:** Janeiro 2026

---

## Descobertas Tecnicas Principais

### Formato do Ficheiro .E3A

O ficheiro `.E3A` do HAP 5.1 e um arquivo **ZIP** contendo:

| Ficheiro | Descricao | Tamanho Registo |
|----------|-----------|-----------------|
| HAP51SPC.DAT | Espacos | 682 bytes |
| HAP51SCH.DAT | Schedules | 792 bytes |
| HAP51WAL.DAT | Paredes | ~128 bytes |
| HAP51WIN.DAT | Janelas | 555 bytes |
| HAP51DOR.DAT | Portas | variavel |
| HAP51ROF.DAT | Coberturas | variavel |
| HAP51INX.MDB | Base de dados Access | - |

### Estrutura Binaria do Registo Space (682 bytes)

```
Offset  Bytes  Campo
------  -----  -----
0-23    24     Nome (Latin-1)
24-27   4      Area (ft2, float)
28-31   4      Altura (ft, float)
32-35   4      Peso (lb/ft2, float)
46-49   4      OA (encoded, float)
50-51   2      OA Unit (uint16)
72-343  272    8 Walls (34 bytes cada)
344-439 96     4 Roofs (24 bytes cada)
440-465 26     Ceiling Partition
466-491 26     Wall Partition
492-541 50     Floor
554-571 18     Infiltration
580-595 16     People
600-617 18     Lighting
632-645 14     Misc
656-661 6      Equipment
```

### Estrutura Binaria do Registo Window (555 bytes)

```
Offset  Bytes  Campo
------  -----  -----
0-254   255    Nome (Latin-1)
257-260 4      Altura (ft, float)
261-264 4      Largura (ft, float)
269-272 4      U-Value (BTU/hr.ft2.F, float)
273-276 4      SHGC (float)
```

### Codificacao do Outdoor Air (OA)

O valor de OA e codificado usando uma funcao exponencial:

```python
# Encoding (valor visivel -> interno)
OA_A = 0.00470356
OA_B = 2.71147770
interno = log(valor / OA_A) / OA_B

# Decoding (interno -> valor visivel)
valor = OA_A * exp(OA_B * interno)
```

### Conversoes de Unidades

| De (SI) | Para (IP) | Formula |
|---------|-----------|---------|
| m2 | ft2 | x 10.7639 |
| m | ft | x 3.28084 |
| kg/m2 | lb/ft2 | / 4.8824 |
| W/m2K | BTU/hr.ft2.F | / 5.678 |
| W | BTU/hr | x 3.412 |
| W/m2 | W/ft2 | / 10.764 |
| C | F | x 1.8 + 32 |

---

## Problemas Resolvidos

### 1. Perfil "Hospital" nao existe no RSECE

**Problema:** O nome "Hospital" nao corresponde a nomenclatura oficial do RSECE.

**Solucao:** Renomeado para "Saude Com Intern" (Estabelecimentos de saude com internamento), seguindo o Anexo XV do RSECE.

### 2. Schedules nao vinham do modelo base

**Problema:** O script `excel_to_hap.py` so lia 3 schedules da sheet "Tipos" do Excel, ignorando os 82 schedules do modelo.

**Solucao:** Adicionado codigo para ler schedules automaticamente do `HAP51SCH.DAT` do modelo base:

```python
sch_path = os.path.join(temp_dir, 'HAP51SCH.DAT')
if os.path.exists(sch_path):
    with open(sch_path, 'rb') as f:
        sch_data = f.read()
    SCHEDULE_RECORD_SIZE = 792
    num_schedules = len(sch_data) // SCHEDULE_RECORD_SIZE
    for i in range(num_schedules):
        offset = i * SCHEDULE_RECORD_SIZE
        name_bytes = sch_data[offset:offset+24]
        name = name_bytes.rstrip(b'\x00').decode('latin-1').strip()
        if name and name not in types['schedules']:
            types['schedules'][name] = i
```

### 3. MDB nIndex=0 nao permitido

**Problema:** A tabela ScheduleIndex do MDB tem validacao que nao permite nIndex=0.

**Solucao:** Iniciar indices a partir de 1:
```python
cursor.execute("INSERT INTO ScheduleIndex (nIndex, szName) VALUES (1, 'Sample Schedule')")
for i, schedule in enumerate(schedules):
    sch_id = i + 2  # Comeca em 2
```

### 4. Valores de Windows nao correspondiam ao Excel

**Problema:** Ao criar janelas, apenas o nome era alterado - os valores binarios permaneciam do template.

**Solucao:** Descobertos os offsets correctos e escrita dos valores:
```python
OFFSET_WIN_HEIGHT = 257
OFFSET_WIN_WIDTH = 261
OFFSET_WIN_UVALUE = 269
OFFSET_WIN_SHGC = 273

struct.pack_into('<f', record, OFFSET_WIN_HEIGHT, m_to_ft(altura))
struct.pack_into('<f', record, OFFSET_WIN_WIDTH, m_to_ft(largura))
struct.pack_into('<f', record, OFFSET_WIN_UVALUE, u_si_to_ip(u_value))
struct.pack_into('<f', record, OFFSET_WIN_SHGC, shgc)
```

---

## Ficheiros Criados

### Scripts Principais

| Script | Funcao |
|--------|--------|
| excel_to_hap.py | Converter Excel -> HAP (principal) |
| hap_to_excel.py | Exportar HAP -> Excel |
| criar_perfis_rsece.py | Criar modelo com 82 perfis RSECE |
| criar_schedule.py | Listar/ver/criar schedules |
| adicionar_dropdowns_rsece.py | Adicionar dropdowns a Excel |

### Bibliotecas

| Biblioteca | Funcao |
|------------|--------|
| hap_library.py | Estruturas binarias dos espacos |
| hap_schedule_library.py | Estruturas binarias dos schedules |

### Templates

| Ficheiro | Descricao |
|----------|-----------|
| Modelo_RSECE.E3A | Modelo com 82 schedules RSECE |
| HAP_Template_RSECE.xlsx | Template Excel com dropdowns |

---

## Perfis RSECE Implementados (27 tipologias)

### Comercio (8)
1. Hipermercado
2. Venda Grosso
3. Supermercado
4. Centro Comercial
5. Pequena Loja
6. Restaurante
7. Pastelaria
8. Pronto-a-Comer

### Hotelaria/Lazer (7)
9. Hotel 4-5 Estrelas
10. Hotel 1-3 Estrelas
11. Cinema Teatro
12. Discoteca
13. Bingo Clube Social
14. Clube Desp Piscina
15. Clube Desportivo

### Servicos (6)
16. Escritorio
17. Banco Sede
18. Banco Filial
19. Comunicacoes
20. Biblioteca
21. Museu Galeria

### Institucional/Outros (6)
22. Tribunal Camara
23. Prisao
24. Escola
25. Universidade
26. Saude Sem Intern
27. Saude Com Intern

**Total:** 27 tipologias x 3 tipos (Ocup, Ilum, Equip) = 81 schedules + 1 Sample = **82 schedules**

---

## Template Excel - Estrutura

### Sheet "Espacos" (147 colunas)

| Colunas | Categoria | Campos |
|---------|-----------|--------|
| 1-6 | GENERAL | Nome, Area, Altura, Peso, OA valor, OA unidade |
| 7-11 | PEOPLE | Ocupacao, Actividade, Sens, Lat, Schedule |
| 12-16 | LIGHTING | Task, General, Fixture, Ballast, Schedule |
| 17-18 | EQUIPMENT | W/m2, Schedule |
| 19-22 | MISC | Sens, Lat, Schedules |
| 23-26 | INFILTRATION | Metodo, ACH (Clg, Htg, Energy) |
| 27-39 | FLOORS | Tipo, Area, U-value, etc. |
| 40-51 | PARTITIONS | Ceiling, Wall |
| 52-123 | WALLS | 8 paredes x 9 campos |
| 124-147 | ROOFS | 4 coberturas x 6 campos |

### Outras Sheets

- **Tipos**: Mapeamento nome -> ID
- **Windows**: Definicao de vaos (Nome, U-Value, SHGC, Altura, Largura)
- **Walls**: Definicao de paredes (Nome, U-Value, Peso, Espessura)
- **Roofs**: Definicao de coberturas (Nome, U-Value, Peso, Espessura)
- **Schedules_RSECE**: Lista de todos os schedules disponiveis
- **Legenda**: Explicacao de cada campo

---

## Workflow de Utilizacao

```
1. Copiar HAP_Template_RSECE.xlsx para novo ficheiro

2. Preencher dados:
   - Sheet "Espacos": dados dos espacos
   - Seleccionar schedules nos dropdowns (colunas K, P, R)
   - Sheet "Windows/Walls/Roofs": tipos customizados (opcional)

3. Executar conversor:
   python excel_to_hap.py MeusDados.xlsx Modelo_RSECE.E3A Output.E3A

4. Abrir Output.E3A no HAP 5.1
```

---

## Notas para Desenvolvimento Futuro

### Estruturas Parcialmente Mapeadas

- **HAP51WAL.DAT**: Estrutura de ~128 bytes, parcialmente mapeada
- **HAP51ROF.DAT**: Estrutura similar a WAL, parcialmente mapeada
- **HAP51DOR.DAT**: Estrutura de portas, nao mapeada

### Tabelas MDB Importantes

- **SpaceIndex**: nIndex, szName, fFloorArea, fNumPeople, fLightingDensity
- **ScheduleIndex**: nIndex, szName
- **WallIndex**: nIndex, szName, fOverallUValue, fOverallWeight, fThickness
- **WindowIndex**: nIndex, szName, fOverallUValue, fOverallShadeCo, fHeight, fWidth
- **RoofIndex**: nIndex, szName, fOverallUValue, fOverallWeight, fThickness
- **DoorIndex**: nIndex, szName, fSolidUValue, fGlassUValue, fArea

### Links (tabelas de relacao)

- Space_Schedule_Links
- Space_Wall_Links
- Space_Window_Links
- Space_Door_Links
- Space_Roof_Links

---

## Ficheiros de Referencia

A conversa completa de desenvolvimento esta guardada em:
`_documentacao/CONVERSA_DESENVOLVIMENTO.jsonl`

Este ficheiro contem todos os detalhes tecnicos, tentativas, erros e solucoes encontradas durante o desenvolvimento.

---

## Requisitos

- Python 3.8+
- openpyxl (obrigatorio)
- pyodbc (opcional, para actualizar MDB)

```bash
pip install openpyxl pyodbc
```

---

*Documento gerado em Janeiro 2026*
*Actualizado em Fevereiro 2026*

---

## Validacao Malhoa22 - Fevereiro 2026

### Resultado: CONVERSOR VALIDADO - Tudo Correcto

Analise exaustiva do projecto Malhoa22 (140 espacos, 71 janelas, 12 pisos + 6 caves) confirmou que:

1. **E3A = Excel:** Os dados binarios no E3A correspondem exactamente ao Excel fonte
2. **HAP = E3A:** O HAP le correctamente todos os dados (confirmado via SPACE.RTF, WINDOW.RTF e screenshots do HAP UI)
3. **Relatorio = Excel:** O relatorio "Minimum Energy Performance Calculator" do HAP mostra Wall=2773 m2 e Glaz=1646 m2, identico ao Excel

### Discrepancia inicial (resolvida)

O Building Simulation Report mostrava valores diferentes (Wall=2197, Glaz=1559) porque o utilizador tinha espacos duplicados no Air System. Apos corrigir, o relatorio bateu certo.

### Achados tecnicos documentados

Ver `docs/ANALISE_MALHOA22_FEV2026.md` para detalhes completos sobre:
- Confirmacao empirica do wall block format (34 bytes)
- Confirmacao dos direction codes (N=1, E=5, S=9, SW=11, W=13)
- Confirmacao dos offsets de janela (H=257, W=261)
- Armadilhas no parsing de SPACE.RTF (espacos com Roofs)
- Diferenca entre relatorios HAP (Building Sim vs Min Energy Performance)

---

## Analise Coluna-a-Coluna - Fevereiro 2026

### Resultado: 2 bugs encontrados e corrigidos

Analise exaustiva de todas as 147 colunas do Excel vs codigo do conversor.

### BUG 1 (CRITICO): Wall block offsets desalinhados - `excel_to_hap.py`

O wall block (34 bytes) tem o layout: +0=dir, +2=area, +6=wall_id, +8=win1_id, **+10=win2_id**, +12=win1_qty, +14=win2_qty, +16=door_id, +18=door_qty.

O codigo tinha um comentario errado ("+10 reservado") e escrevia win2_id no offset +14, win2_qty no +16, door_id no +18 e door_qty no +20 - tudo deslocado 2-4 bytes. Nao afectou Malhoa22 porque nenhum espaco usa Window 2 ou Doors (valores todos 0).

**Corrigido:** Offsets alinhados com o layout correcto confirmado em `hap_library.py:parse_wall_block()`.

### BUG 2 (MEDIO): Lighting schedule offset - `hap_library.py`

`excel_to_hap.py` escreve lighting_schedule_id no offset **616** (validado com HAP). `hap_library.py` lia/escrevia no offset 614 (errado).

**Corrigido:** `hap_library.py` parse_space e encode_space alterados de 614 para 616.

### Campos todos verificados OK

Todas as 147 colunas do Excel foram verificadas: GENERAL (6), PEOPLE (5), LIGHTING (5), EQUIPMENT (2), MISC (4), INFILTRATION (4), FLOORS (13), PARTITIONS (12), WALLS (72), ROOFS (24) - tudo lido e escrito correctamente (excepto os 2 bugs acima, ja corrigidos).

---

## Auditoria Completa Documentacao vs Codigo - Fevereiro 2026

### Resultado: 5 bugs adicionais encontrados e corrigidos

Auditoria cruzando toda a documentacao existente com todos os ficheiros de codigo (conversor, extractor, editor, library).

### BUG 3 (CRITICO): `hap_extractor.py` wall block offsets errados

O extractor lia win2_type_id de +14 (correcto: +10), win2_qty de +18 (correcto: +14), door_type_id de +20 (correcto: +16), door_qty de +22 (correcto: +18). Mesma confusao que o BUG 1 do conversor.

**Corrigido:** Offsets alinhados com `hap_library.py:parse_wall_block()`.

### BUG 4 (MEDIO): `hap_extractor.py` Window SHGC offset 433 -> 273

O extractor lia SHGC do offset 433 em vez de 273 (confirmado em ANALISE_MALHOA22 e excel_to_hap.py).

**Corrigido:** Offset alterado para 273.

### BUG 5 (MEDIO): `editor_e3a.py` Window SHGC offset 433 -> 273

O editor escrevia SHGC no offset 433 em vez de 273 (mesma inconsistencia do extractor).

**Corrigido:** Offset alterado para 273.

### BUG 6 (MEDIO): `hap_library.py` infiltration offsets 492 -> 554

`parse_space()` e `encode_space()` usavam offsets 492-526 para infiltracao, que sao na verdade a zona de FLOOR data. Os offsets correctos sao 554-572 (confirmado em excel_to_hap.py e na documentacao do proprio ficheiro).

**Corrigido:** parse_space e encode_space alterados para usar offsets 556/562.

### BUG 7 (MEDIO): `hap_library.py` thermostat vs partition overlap

`parse_space()` lia "thermostat" dos offsets 440-492 (thermostat_schedule_id, cooling/heating setpoints), mas estes offsets sao na verdade Partition 1 (440-465) e Partition 2 (466-491), conforme confirmado em excel_to_hap.py e hap_extractor.py.

**Corrigido:** parse_space e encode_space agora leem/escrevem correctamente Partition 1 e Partition 2 com todos os campos (type, area, u-value, 4 temperaturas).

### Documentacao header `hap_library.py` limpa

Removidas entradas duplicadas/conflituantes na tabela de offsets do header (thermostat vs partition2, floor temps vs partition temps).

---

## Auditoria Editor + Formula OA - Fevereiro 2026

### Resultado: 2 bugs adicionais encontrados e corrigidos

### BUG 8 (MEDIO): `editor_e3a.py` missing `data_only=True`

`openpyxl.load_workbook(editor_xlsx)` nao avaliava formulas Excel. Quando o utilizador preenchia OA com formulas como `=M4/0.8`, o valor era lido como string "=M4/0.8" em vez do valor numerico. O `float()` falhava silenciosamente no try/except, ignorando a alteracao.

**Corrigido:** Linha 299 alterada para `openpyxl.load_workbook(editor_xlsx, data_only=True)`.

### BUG 9 (CRITICO): Formula OA errada em editor, library e extractor

O conversor (`excel_to_hap.py`) usava a formula piecewise correcta `fast_exp2/fast_log2` (descoberta 2026-02-05), mas os outros 3 ficheiros ainda usavam a formula antiga:
- `editor_e3a.py`: `log(val / 0.00470356) / 2.71147770`
- `hap_library.py`: `decode_oa_value` e `encode_oa_value` com mesma formula antiga
- `hap_extractor.py`: `decode_oa` com mesma formula antiga

**Impacto:** Para OA=1323 L/s, a formula antiga dava 606 L/s no HAP (erro de 54%!). Para OA=500 L/s, erro de 26%.

**Corrigido:** Todos os 3 ficheiros actualizados com formula piecewise identica ao conversor:
- `y = Y0 * fast_exp2(k * (x - 4))` com k=4 para x<4, k=2 para x>=4
- Y0 = 512 CFM em L/s = 241.637...
- Header de `hap_library.py` actualizado com formula correcta

### Auditoria exaustiva final (9 bugs totais)

Auditoria completa dos 4 ficheiros (conversor, library, extractor, editor) confirmou que apos os 9 bugs corrigidos, **todos os offsets, formulas, conversoes e estruturas binarias estao 100% consistentes** entre todos os ficheiros.
