---
description: "Referência técnica completa para desenvolvimento e alterações ao código HAP. Usa SEMPRE que vais modificar código Python do projecto, corrigir bugs, adicionar funcionalidades, alterar offsets, mexer em binários, alterar conversor/extractor/editor/library. Também quando o utilizador diz alterar código, modificar script, corrigir, rectificar, mexer no conversor, mexer no extractor, mexer no editor, mexer na library, desenvolver, programar."
---

# Referência Técnica HAP 5.1 - Para Desenvolvimento

Este documento contém TODO o conhecimento técnico necessário para alterar o código deste projecto. Lê-o antes de fazer qualquer modificação.

## Arquitectura do código

### Ficheiros principais e responsabilidades

```
conversor/excel_to_hap.py      ESCREVE E3A (Excel → binário)
conversor/hap_library.py       BIBLIOTECA PARTILHADA (parse/encode de todas as estruturas)
conversor/hap_schedule_library.py  Gestão de schedules
conversor/validar_excel_hap.py VALIDA Excel antes de converter
conversor/validar_e3a.py       Wrapper → importa de ../validar_e3a.py
extractor/hap_extractor.py     LÊ E3A (binário → Excel)
editor/editor_e3a.py           LÊ + ESCREVE E3A (modifica campos selectivos)
validar_e3a.py                 VALIDA E3A (schedules, default space, IDs)
adaptador/adapter_hap52.py     CONVERTE formato HAP 5.2 → standard
iee/iee_completo_v3.py         CALCULA IEE (CSV → Excel com fórmulas)
```

### Regra de ouro: consistência entre ficheiros

Quando alteras um offset ou fórmula, TENS de alterar em TODOS os ficheiros que o usam:
- `excel_to_hap.py` (escrita)
- `hap_library.py` (parse + encode)
- `hap_extractor.py` (leitura)
- `editor_e3a.py` (leitura + escrita)

Já aconteceram 9 bugs por falta de consistência. Verifica SEMPRE os 4 ficheiros.

---

## Formato E3A

Ficheiros .E3A são ZIPs contendo ficheiros binários. Valores numéricos em **little-endian**. Strings em **Latin-1** (ISO-8859-1).

### Ficheiros dentro do ZIP

| Ficheiro | Conteúdo | Bytes/registo |
|----------|----------|---------------|
| HAP51SPC.DAT | Espaços | **682** |
| HAP51SCH.DAT | Schedules | **792** |
| HAP51WAL.DAT | Wall assemblies | **3187** |
| HAP51ROF.DAT | Roof assemblies | **3187** |
| HAP51WIN.DAT | Janelas | **555** |
| HAP51DOR.DAT | Portas | variável |
| HAP51INX.MDB | Base dados Access (índices + links) | - |
| PROJECT.E3P | Configuração (INI) | - |

---

## Space Record (682 bytes) — MAPA COMPLETO

O registo 0 (offset 0-681) é o Default Space (template, NÃO MODIFICAR).
Espaços reais começam no offset 682.

### GENERAL (bytes 0-71)

| Offset | Size | Type | Campo | Unidade interna | Conversão SI→IP |
|--------|------|------|-------|-----------------|-----------------|
| 0-23 | 24 | char[] | Nome | Latin-1, null-padded | - |
| 24-27 | 4 | float | Floor Area | ft² | × 10.7639 |
| 28-31 | 4 | float | Ceiling Height | ft | × 3.28084 |
| 32-35 | 4 | float | Building Weight | lb/ft² | ÷ 4.8824 |
| 36-39 | 4 | uint32 | Type Flag | valor típico: 4 | - |
| 44-45 | 2 | bytes | OA Auxiliary | non-zero para L/s e L/s/person | - |
| 46-49 | 4 | float | OA Internal Value | **encoded** (ver fórmula OA) | - |
| 50-51 | 2 | uint16 | OA Unit Code | 1=L/s, 2=L/s/m², 3=L/s/person, 4=% | - |

### WALLS (bytes 72-343) — 8 blocos × 34 bytes

Início: offset 72. Cada bloco = 34 bytes.

| Offset relativo | Size | Type | Campo |
|-----------------|------|------|-------|
| +0 | 2 | uint16 | Direction code (N=1, NNE=2, NE=3, ENE=4, E=5, ESE=6, SE=7, SSE=8, S=9, SSW=10, SW=11, WSW=12, W=13, WNW=14, NW=15, NNW=16, H=17) |
| +2 | 4 | float | Gross wall area (ft²) |
| +6 | 2 | uint16 | Wall type ID |
| +8 | 2 | uint16 | Window 1 type ID |
| **+10** | **2** | **uint16** | **Window 2 type ID** (NÃO é reservado!) |
| **+12** | **2** | **uint16** | **Window 1 quantity** (NÃO é +10!) |
| +14 | 2 | uint16 | Window 2 quantity |
| +16 | 2 | uint16 | Door type ID |
| +18 | 2 | uint16 | Door quantity |
| +20-33 | 14 | bytes | Shading/overhangs/reserved |

**ARMADILHA HISTÓRICA (BUG 1 e 3):** Os offsets +10/+12 foram trocados no passado. Window 2 type ID está em +10, Window 1 quantity em +12. Já corrigido.

### ROOFS (bytes 344-439) — 4 blocos × 24 bytes

Início: offset 344. Cada bloco = 24 bytes.

| Offset relativo | Size | Type | Campo |
|-----------------|------|------|-------|
| +0 | 2 | uint16 | Direction code (mesmo que walls, H=17 para horizontal) |
| +2 | 2 | uint16 | Slope (graus) |
| +4 | 4 | float | Gross area (ft²) |
| +8 | 2 | uint16 | Roof type ID |
| +10 | 2 | uint16 | Skylight type ID (usa WindowIndex!) |
| +12 | 2 | uint16 | Skylight quantity |
| +14-23 | 10 | bytes | Reserved |

### PARTITIONS (bytes 440-491) — 2 partições × 26 bytes

**Partition 1 (Ceiling) — bytes 440-465:**

| Offset | Size | Type | Campo |
|--------|------|------|-------|
| 440-441 | 2 | uint16 | Type (1=Ceiling, 2=Wall) |
| 442-445 | 4 | float | Area (ft²) |
| 446-449 | 4 | float | U-value (BTU/hr·ft²·°F) — SI÷5.678 |
| 450-453 | 4 | float | Uncond Space Max Temp (°F) |
| 454-457 | 4 | float | Ambient at Max Temp (°F) |
| 458-461 | 4 | float | Uncond Space Min Temp (°F) |
| 462-465 | 4 | float | Ambient at Min Temp (°F) |

**Partition 2 (Wall) — bytes 466-491:** Mesma estrutura.

**ARMADILHA HISTÓRICA (BUG 7):** Estes offsets foram confundidos com "thermostat" no passado. São PARTITIONS, não thermostat. Já corrigido.

### FLOOR (bytes 492-541)

| Offset | Size | Type | Campo |
|--------|------|------|-------|
| 492-493 | 2 | uint16 | Floor Type (1=Above Cond, 2=Above Uncond, 3=Slab On Grade, 4=Slab Below) |
| 494-497 | 4 | float | Floor Area (ft²) |
| 498-501 | 4 | float | Floor U-value (IP) — SI÷5.678 |
| 502-505 | 4 | float | Slab Exposed Perimeter (ft) — para tipos 3,4 |
| 506-509 | 4 | float | Slab Edge Insulation R (IP) — para tipo 3 |
| 510-513 | 4 | float | Slab Floor Depth (ft) — para tipo 4 |
| 514-517 | 4 | float | Basement Wall U (IP) — para tipo 4 |
| 518-521 | 4 | float | Wall Insulation R (IP) — para tipo 4 |
| 522-525 | 4 | float | Depth of Wall Insulation (ft) — para tipo 4 |

### INFILTRATION (bytes 554-571)

| Offset | Size | Type | Campo |
|--------|------|------|-------|
| 554-555 | 2 | uint16 | Design Cooling flag (2=ACH mode) |
| 556-559 | 4 | float | Design Cooling ACH |
| 560-561 | 2 | uint16 | Design Heating flag (2=ACH mode) |
| 562-565 | 4 | float | Design Heating ACH |
| 566-567 | 2 | uint16 | Energy Analysis flag (2=ACH mode) |
| 568-571 | 4 | float | Energy Analysis ACH |

**ARMADILHA HISTÓRICA (BUG 6):** Estes offsets foram confundidos com 492-526 (zona de Floor). Offsets correctos são 554-571. Já corrigido.

### PEOPLE (bytes 580-595)

| Offset | Size | Type | Campo |
|--------|------|------|-------|
| 580-583 | 4 | float | Occupancy (pessoas) |
| 584-585 | 2 | uint16 | Activity Level ID |
| 586-589 | 4 | float | Sensible heat (BTU/hr) — W×3.412 |
| 590-593 | 4 | float | Latent heat (BTU/hr) — W×3.412 |
| **594-595** | **2** | **uint16** | **People Schedule ID** |

### LIGHTING (bytes 600-617)

| Offset | Size | Type | Campo |
|--------|------|------|-------|
| 600-603 | 4 | float | Task Lighting (W total) |
| 604-605 | 2 | uint16 | Fixture Type (0=Recessed Unvented, 1=Vented Return, 2=Vented Supply+Return, 3=Surface/Pendant) |
| 606-609 | 4 | float | Overhead Lighting (W total) |
| 610-613 | 4 | float | Ballast Multiplier |
| **616-617** | **2** | **uint16** | **Lighting Schedule ID** |

**ARMADILHA HISTÓRICA (BUG 2):** O schedule ID está no offset **616**, NÃO 614! O offset 614 é outro campo. Já corrigido.

### MISC LOADS (bytes 632-645)

| Offset | Size | Type | Campo |
|--------|------|------|-------|
| 632-635 | 4 | float | Misc Sensible (BTU/hr) — W×3.412 |
| 636-639 | 4 | float | Misc Latent (BTU/hr) — W×3.412 |
| 640-641 | 2 | uint16 | Misc Sensible Schedule ID |
| 644-645 | 2 | uint16 | Misc Latent Schedule ID |

### EQUIPMENT (bytes 656-661)

| Offset | Size | Type | Campo |
|--------|------|------|-------|
| 656-659 | 4 | float | Equipment density (W/ft²) — W/m²÷10.764 |
| **660-661** | **2** | **uint16** | **Equipment Schedule ID** |

---

## Window Record (555 bytes)

| Offset | Size | Type | Campo | Conversão |
|--------|------|------|-------|-----------|
| 0-254 | 255 | char[] | Nome (Latin-1) | - |
| **257-260** | 4 | float | Altura | ft — m×3.28084 |
| **261-264** | 4 | float | Largura | ft — m×3.28084 |
| **269-272** | 4 | float | U-Value | BTU/hr·ft²·°F — SI÷5.678 |
| **273-276** | 4 | float | SHGC | adimensional |

**ARMADILHA HISTÓRICA (BUG 4, 5):** SHGC está no offset **273**, NÃO 433! Altura/Largura em **257/261**, NÃO 24/28! Já corrigido.

---

## Wall/Roof Assembly (3187 bytes)

| Offset | Size | Campo |
|--------|------|-------|
| 0-254 | 255 | Nome (Latin-1) |
| 377+ | - | Layers (281 bytes cada) |
| ~269 | 4 | U-value overall (float, IP) |

Layer structure: 281 bytes com material name, thickness, conductivity, etc.
R-values superficiais fixos: R_inside = 0.68, R_outside = 0.33 (IP units).

---

## Schedule Record (792 bytes)

| Offset | Size | Campo |
|--------|------|-------|
| 0-79 | 80 | Nome |
| 80-191 | 112 | Nomes dos 8 profiles |
| 192-575 | 384 | Valores horários (8 profiles × 24h × 2 bytes) |
| **576-791** | **216** | **Calendário** (12 meses × 9 day-types × 2 bytes) |

**VALORES VÁLIDOS no calendário: 1 a 8.** Valor 100 causa Error 9 "Subscript out of range".

---

## Fórmula OA (piecewise fast_exp2) — CRÍTICA

NÃO usar fórmula exponencial simples. A fórmula correcta é piecewise:

```python
import math

Y0 = 512.0 * (28.316846592 / 60.0)  # = 241.637... (512 CFM em L/s)

def fast_exp2(t):
    """Aproximação HAP de 2^t"""
    fi = math.floor(t)
    ff = t - fi
    return (2.0 ** fi) * (1.0 + ff)

def fast_log2(v):
    """Inverso de fast_exp2"""
    if v <= 0:
        return -16.0
    fi = math.floor(math.log2(v))
    ff = v / (2.0 ** fi) - 1.0
    return fi + ff

def encode_oa(user_value_ls):
    """User L/s → internal float"""
    v = user_value_ls / Y0
    t = fast_log2(v)
    if t < 0:
        return t / 4.0 + 4.0   # k=4 para valores < Y0
    else:
        return t / 2.0 + 4.0   # k=2 para valores >= Y0

def decode_oa(internal):
    """Internal float → user L/s"""
    if internal < 4.0:
        t = (internal - 4.0) * 4.0  # k=4
    else:
        t = (internal - 4.0) * 2.0  # k=2
    return Y0 * fast_exp2(t)
```

**ARMADILHA HISTÓRICA (BUG 9):** A fórmula antiga `OA_A * exp(OA_B * x)` dava erros de 26-54% para valores altos. A fórmula piecewise é a correcta. Já corrigido nos 4 ficheiros.

---

## Conversões de unidades

```python
# SI → Imperial (para ESCRITA no E3A)
def m2_to_ft2(v):  return v * 10.7639
def m_to_ft(v):    return v * 3.28084
def kg_m2_to_lb_ft2(v): return v / 4.8824
def u_si_to_ip(v): return v / 5.678      # W/m²K → BTU/hr·ft²·°F
def w_to_btu(v):   return v * 3.412      # W → BTU/hr
def wm2_to_wft2(v): return v / 10.764    # W/m² → W/ft²
def c_to_f(v):     return v * 1.8 + 32   # °C → °F

# Imperial → SI (para LEITURA do E3A)
def ft2_to_m2(v):  return v / 10.7639
def ft_to_m(v):    return v / 3.28084
def lb_ft2_to_kg_m2(v): return v * 4.8824
def u_ip_to_si(v): return v * 5.678
def btu_to_w(v):   return v / 3.412
def wft2_to_wm2(v): return v * 10.764
def f_to_c(v):     return (v - 32) / 1.8
```

---

## MDB (HAP51INX.MDB)

Tabelas de índice:
- **SpaceIndex**: nIndex, szName, fFloorArea, fNumPeople, fLightingDensity
- **ScheduleIndex**: nIndex, szName
- **WallIndex**: nIndex, szName, fOverallUValue, fOverallWeight, fThickness
- **WindowIndex**: nIndex, szName, fOverallUValue, fOverallShadeCo, fHeight, fWidth
- **RoofIndex**: nIndex, szName, fOverallUValue, fOverallWeight, fThickness
- **DoorIndex**: nIndex, szName, fSolidUValue, fGlassUValue, fArea

Tabelas de links (OBRIGATÓRIAS):
- **Space_Schedule_Links**: Space_ID, Schedule_ID
- **Space_Wall_Links**: Space_ID, Wall_ID
- **Space_Window_Links**: Space_ID, Window_ID (inclui skylights!)
- **Space_Door_Links**: Space_ID, Door_ID
- **Space_Roof_Links**: Space_ID, Roof_ID

**IMPORTANTE:** IDs no MDB começam em 1 (não 0). nIndex=0 não é permitido em ScheduleIndex.

---

## Default Space — NÃO MODIFICAR

O registo 0 do HAP51SPC.DAT é o template interno. Os seguintes offsets DEVEM ser 0:
- 554 (Infiltration flag 1)
- 560 (Infiltration flag 2)
- 566 (Infiltration flag 3)
- 594 (People Schedule ID)
- 616 (Lighting Schedule ID)
- 660 (Equipment Schedule ID)

---

## 9 Bugs Históricos (todos corrigidos — para referência)

| # | Severidade | Ficheiro | Bug | Fix |
|---|-----------|----------|-----|-----|
| 1 | CRÍTICO | excel_to_hap.py | Wall block offsets desalinhados (+10 reservado vs +10=win2_id) | Corrigido offsets |
| 2 | MÉDIO | hap_library.py | Lighting schedule offset 614 | Alterado para 616 |
| 3 | CRÍTICO | hap_extractor.py | Wall block offsets errados (mesma confusão que bug 1) | Corrigido offsets |
| 4 | MÉDIO | hap_extractor.py | Window SHGC offset 433 | Alterado para 273 |
| 5 | MÉDIO | editor_e3a.py | Window SHGC offset 433 | Alterado para 273 |
| 6 | MÉDIO | hap_library.py | Infiltration offsets 492-526 (zona Floor!) | Alterado para 554-572 |
| 7 | MÉDIO | hap_library.py | Thermostat nos offsets 440-492 (são Partitions!) | Corrigido para Partitions |
| 8 | MÉDIO | editor_e3a.py | Missing data_only=True no openpyxl | Adicionado |
| 9 | CRÍTICO | editor/library/extractor | Fórmula OA exponencial simples | Corrigido para piecewise |

---

## Checklist para qualquer alteração

1. [ ] Identificar TODOS os ficheiros afectados (conversor, library, extractor, editor)
2. [ ] Verificar offset no header do hap_library.py (fonte de verdade)
3. [ ] Alterar de forma consistente em todos os ficheiros
4. [ ] Testar com ficheiro de exemplo (exemplos/Malhoa22.E3A)
5. [ ] Verificar que o E3A abre no HAP sem Error 9
6. [ ] Verificar que os valores no HAP correspondem ao Excel
