# Formato HAP 5.1 (.E3A)

## Estrutura do Ficheiro

O ficheiro .E3A é um arquivo ZIP contendo:

| Ficheiro | Descrição | Tamanho Registo |
|----------|-----------|-----------------|
| HAP51SPC.DAT | Spaces | 682 bytes |
| HAP51SCH.DAT | Schedules | 792 bytes |
| HAP51WAL.DAT | Wall Assemblies | 3187 bytes |
| HAP51ROF.DAT | Roof Assemblies | 3187 bytes |
| HAP51WIN.DAT | Windows | 555 bytes |
| HAP51INX.MDB | Índice Access (UI) | - |
| PROJECT.E3P | Info do projecto | - |

---

## Estrutura de um Space (682 bytes)

```
GENERAL:
  0-23:     Nome do space (24 bytes, null-terminated)
  24-27:    Floor Area (float, ft²)
  28-31:    Ceiling Height (float, ft)
  32-35:    Building Weight (float, lb/ft²)
  46-49:    Outside Air (float)
  50-51:    OA Unit Code (short)

WALLS (8 walls x 34 bytes = 272 bytes):
  72-343:   Wall blocks

ROOFS (4 roofs x 24 bytes = 96 bytes):
  344-439:  Roof blocks

PARTITIONS:
  440-465:  Ceiling partition
  466-491:  Wall partition
  492-541:  Floor

INFILTRATION (CRÍTICO!):
  554-555:  Infiltration Schedule ID 1 (short) - NÃO É FLAG!
  556-559:  ACH Design Clg (float)
  560-561:  Infiltration Schedule ID 2 (short) - NÃO É FLAG!
  562-565:  ACH Design Htg (float)
  566-567:  Infiltration Schedule ID 3 (short) - NÃO É FLAG!
  568-571:  ACH Energy (float)

PEOPLE:
  580-583:  Occupancy (float, pessoas)
  584-585:  Activity code (short)
  586-589:  Sensible (float, BTU/hr)
  590-593:  Latent (float, BTU/hr)
  594-595:  People Schedule ID (short)

LIGHTING:
  600-603:  Task Light (float)
  604-605:  Fixture code (short)
  606-609:  General Light (float)
  610-613:  Ballast factor (float)
  616-617:  Light Schedule ID (short)

EQUIPMENT:
  656-659:  Equipment (float, W/ft²)
  660-661:  Equip Schedule ID (short)
```

### Default Space (Space 0)

O Default Space é um template interno do HAP que **NÃO DEVE SER MODIFICADO**.

Valores obrigatórios:
```
Offset 554: 0 (Schedule ID)
Offset 560: 0 (Schedule ID)
Offset 566: 0 (Schedule ID)
Offset 594: 0 (People Schedule)
Offset 616: 0 (Light Schedule)
Offset 660: 0 (Equip Schedule)
```

---

## Estrutura de um Schedule (792 bytes)

```
HEADER:
  0-79:     Nome (80 bytes)
  80-191:   Profile names (8 profiles x ~14 bytes)

HOURLY DATA:
  192-575:  Valores horários por profile
            8 profiles x 24 horas x 2 bytes = 384 bytes

CALENDAR (CRÍTICO!):
  576-791:  Day-type to Profile mapping
            108 valores de 2 bytes = 216 bytes
            (12 meses x 9 day-types)

            VALORES VÁLIDOS: 1, 2, 3, 4, 5, 6, 7, 8
            VALOR INVÁLIDO: 100 (causa Error 9!)
```

---

## Assemblies (Walls/Roofs)

- Tamanho fixo: **3187 bytes** cada
- Nome: bytes 0-254
- U-value: offset ~269 (float)

---

## MDB (HAP51INX.MDB)

Base de dados Access usada pela UI do HAP para mostrar listas.

Tabelas principais:
- SpaceIndex
- ScheduleIndex
- WindowIndex
- WallIndex
- RoofIndex

**IMPORTANTE:** Deve ser actualizado em sincronia com os ficheiros .DAT

---

## Conversões de Unidades

| SI | Imperial | Fórmula |
|----|----------|---------|
| m² | ft² | × 10.7639 |
| m | ft | × 3.28084 |
| W/m²·K | BTU/hr·ft²·°F | × 0.176110 |
| kg/m² | lb/ft² | × 0.204816 |

---

Última actualização: 2026-02-04
