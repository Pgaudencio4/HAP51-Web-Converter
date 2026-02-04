# Descobertas sobre o formato HAP 5.1 (.E3A)

## Ficheiros no arquivo .E3A (ZIP):

| Ficheiro | Descrição | Tamanho registo |
|----------|-----------|-----------------|
| HAP51SPC.DAT | Spaces | 682 bytes |
| HAP51SCH.DAT | Schedules | 792 bytes |
| HAP51ROF.DAT | Roof assemblies | 3187 bytes |
| HAP51WAL.DAT | Wall assemblies | 3187 bytes |
| HAP51WIN.DAT | Windows | 555 bytes |
| HAP51INX.MDB | Índice Access | - |

---

## Estrutura de um Space (682 bytes)

### Offsets principais:

```
GENERAL:
  0-23:     Nome do space (24 bytes, null-terminated)
  24-27:    Floor Area (float, ft²)
  28-31:    Ceiling Height (float, ft)
  32-35:    Building Weight (float, lb/ft²)
  46-49:    Outside Air (float, encoded)
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

MISC:
  632-635:  Misc Sensible (float, BTU/hr)
  636-639:  Misc Latent (float, BTU/hr)
  640-641:  Misc Sens Schedule ID (short)
  644-645:  Misc Lat Schedule ID (short)

EQUIPMENT:
  656-659:  Equipment (float, W/ft²)
  660-661:  Equip Schedule ID (short)
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

## Default Space (Space 0)

O Default Space é um template interno do HAP que **NÃO DEVE SER MODIFICADO**.

### Valores correctos (do Template):
```
Offset 554: 0 (Schedule ID - deve ser 0)
Offset 560: 0 (Schedule ID - deve ser 0)
Offset 566: 0 (Schedule ID - deve ser 0)
Offset 594: 0 (People Schedule - deve ser 0)
Offset 616: 0 (Light Schedule - deve ser 0)
Offset 660: 0 (Equip Schedule - deve ser 0)
```

---

## ERRO 9 "Subscript out of range"

### Causa raiz descoberta (2026-02-04):

O erro ocorre quando o HAP tenta ler `HourlyValue` de um Schedule com um **Profile ID inválido** no calendário.

### Problema específico encontrado:

Todos os 83 schedules tinham no calendário (bytes 576-791) o valor **100** em vez de valores válidos (1-8).

```
ERRADO:  Calendário com valor 100 (Profile 100 não existe!)
CORRECTO: Calendário com valores 1, 2, 3, 4 (Profiles existentes)
```

### Manifestação do erro:
- O HAP carrega o ficheiro normalmente
- Ao simular, tenta ler `Schedule.HourlyValue(profile_id)`
- Se `profile_id = 100`, dá "Subscript out of range"
- O erro pode aparecer em diferentes funções:
  - `CalcInfiltrationObj()` - se Schedule de infiltração corrompido
  - `CalcPeopleObj()` - se People Schedule corrompido
  - etc.

### Solução:
1. Corrigir o calendário de TODOS os schedules
2. Substituir valor 100 por valor 1 (ou copiar do template)

---

## Correcção no excel_to_hap.py

### ERRO no código original (linhas 559-565):

```python
# ERRADO - está a escrever Schedule IDs, não flags!
struct.pack_into('<H', data, 554, 2)  # Flag ACH mode <- ERRADO!
struct.pack_into('<f', data, 556, safe_float(space.get('ach_clg')))
struct.pack_into('<H', data, 560, 2)  # Flag ACH mode <- ERRADO!
```

### Código correcto:

```python
# Usar Schedule ID do modelo ou 0 para usar Sample Schedule
# Os offsets 554, 560, 566 são SCHEDULE IDs, não flags!
infil_sch_id = get_type_id(space.get('infiltration_sch'), types['schedules'], 0)
struct.pack_into('<H', data, 554, infil_sch_id)
struct.pack_into('<f', data, 556, safe_float(space.get('ach_clg')))
struct.pack_into('<H', data, 560, infil_sch_id)
struct.pack_into('<f', data, 562, safe_float(space.get('ach_htg')))
struct.pack_into('<H', data, 566, infil_sch_id)
struct.pack_into('<f', data, 568, safe_float(space.get('ach_energy')))
```

---

## CHECKLIST de Validação

Antes de gerar um ficheiro E3A, verificar:

- [ ] **Schedules**: Todos os calendários (bytes 576-791) têm valores 1-8
- [ ] **Default Space**: Offsets 554, 560, 566, 594, 616, 660 = 0
- [ ] **Spaces reais**: Schedule IDs < número total de schedules
- [ ] **Assemblies**: Cada wall/roof tem 3187 bytes

### Script de validação:

```python
import zipfile
import struct

def validate_e3a(path):
    errors = []

    with zipfile.ZipFile(path, 'r') as z:
        # Verificar schedules
        with z.open('HAP51SCH.DAT') as f:
            sch_data = f.read()

        num_schedules = len(sch_data) // 792

        for i in range(num_schedules):
            sch = sch_data[i*792:(i+1)*792]
            for j in range(576, 792, 2):
                val = struct.unpack('<H', sch[j:j+2])[0]
                if val > 8 and val != 0:
                    errors.append(f"Schedule {i}: calendário tem valor {val} (inválido)")
                    break

        # Verificar Default Space
        with z.open('HAP51SPC.DAT') as f:
            spc_data = f.read()

        ds = spc_data[0:682]
        for offset in [554, 560, 566, 594, 616, 660]:
            val = struct.unpack('<H', ds[offset:offset+2])[0]
            if val != 0:
                errors.append(f"Default Space: offset {offset} = {val} (deve ser 0)")

        # Verificar schedule IDs nos spaces
        num_spaces = len(spc_data) // 682
        for i in range(1, num_spaces):
            space = spc_data[i*682:(i+1)*682]
            for offset in [554, 560, 566, 594, 616, 660]:
                val = struct.unpack('<H', space[offset:offset+2])[0]
                if val >= num_schedules and val != 0 and val != 65535:
                    name = space[0:24].decode('latin1', errors='ignore')
                    errors.append(f"Space {i} '{name}': offset {offset} = {val} >= {num_schedules}")

    return errors

# Uso:
errors = validate_e3a("meu_ficheiro.E3A")
if errors:
    print("ERROS ENCONTRADOS:")
    for e in errors:
        print(f"  - {e}")
else:
    print("Ficheiro válido!")
```

---

## Histórico de Erros Resolvidos

| Data | Erro | Causa | Solução |
|------|------|-------|---------|
| 2026-02-04 | Error 341 | Assembly < 3187 bytes | Usar WALL_ASSEMBLY_SIZE = 3187 |
| 2026-02-04 | Error 9 CalcInfiltrationObj | Schedule calendário = 100 | Corrigir calendário para 1-4 |
| 2026-02-04 | Error 9 CalcPeopleObj | Schedule calendário = 100 | Corrigir calendário para 1-4 |

---

## Referências

- Ficheiro modelo: `HAP_Template_Clean.E3A`
- Ficheiro bom de referência: `CasaAlecrim2025_Prev (1).E3A`
- Conversor: `excel_to_hap.py`

---
Última actualização: 2026-02-04
