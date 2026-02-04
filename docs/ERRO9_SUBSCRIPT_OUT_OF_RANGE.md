# ERRO 9 "Subscript out of range" no HAP 5.1

## Resumo

O erro "Subscript out of range" (Error 9) ocorre durante a simulação quando o HAP tenta aceder a um índice de array inválido.

### Stack Trace típico:
```
Error Number: 9
Description: Subscript out of range
Source:
   frmMain.mnuReportSimulation_Click
   HAP.frmMain.DoAppReport
   HAPCalc.SC_Controller.Run()
   HAPCalc.ASD_Calculations.Calculate()
   HAPCalc.ASD_Calculations.SpaceLoadController()
   HAPCalc.ASD_Calculations.CalcInfiltrationObj()   <- ou CalcPeopleObj()
   HAPDataHandler20.SCH_Schedule.HourlyValue [PropertyGet]
```

---

## Causa Raiz

O problema está nos **Schedules** - especificamente no **calendário** interno de cada schedule.

### Estrutura do Schedule (792 bytes):

| Offset | Tamanho | Descrição |
|--------|---------|-----------|
| 0-79 | 80 bytes | Nome do schedule |
| 80-191 | 112 bytes | Nomes dos 8 profiles |
| 192-575 | 384 bytes | Valores horários (8 profiles x 24 horas x 2 bytes) |
| **576-791** | **216 bytes** | **CALENDÁRIO (108 valores x 2 bytes)** |

### O que é o calendário:

O calendário define qual **Profile** (1-8) usar para cada combinação de mês/dia-type.
- 12 meses x 9 day-types = 108 valores
- Cada valor é um short (2 bytes)
- **Valores válidos: 1, 2, 3, 4, 5, 6, 7, 8**

### O problema:

Os schedules tinham no calendário o valor **100** em vez de valores válidos (1-8).

Quando o HAP executa `Schedule.HourlyValue(profile_id)` com `profile_id = 100`, tenta aceder ao Profile 100 que não existe (só há 8 profiles), resultando em "Subscript out of range".

---

## Solução

### Usar o validador/corrector automático:

```bash
# Só validar
python validar_e3a.py MeuFicheiro.E3A

# Validar e corrigir
python validar_e3a.py MeuFicheiro.E3A --fix
```

### Correcção manual (se necessário):

```python
import zipfile
import struct

with zipfile.ZipFile('ficheiro.E3A', 'r') as z:
    files = {name: z.read(name) for name in z.namelist()}

sch_data = bytearray(files['HAP51SCH.DAT'])
num_schedules = len(sch_data) // 792

# Corrigir calendário de todos os schedules
for i in range(num_schedules):
    for j in range(576, 792, 2):
        offset = i * 792 + j
        val = struct.unpack('<H', sch_data[offset:offset+2])[0]
        if val > 8:
            sch_data[offset:offset+2] = struct.pack('<H', 1)

files['HAP51SCH.DAT'] = bytes(sch_data)

# Gravar
with zipfile.ZipFile('ficheiro_corrigido.E3A', 'w', zipfile.ZIP_DEFLATED) as z:
    for name, content in files.items():
        z.writestr(name, content)
```

---

## Checklist de Validação

Antes de usar um ficheiro E3A no HAP:

- [ ] Executar `validar_e3a.py` sem erros
- [ ] Schedules: calendário com valores 1-8 (não 100)
- [ ] Default Space: offsets 554, 560, 566, 594, 616, 660 = 0
- [ ] Schedule IDs válidos (< número de schedules)
- [ ] Assemblies com 3187 bytes cada

---

**Data:** 2026-02-04
