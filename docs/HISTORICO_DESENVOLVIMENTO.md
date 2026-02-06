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
