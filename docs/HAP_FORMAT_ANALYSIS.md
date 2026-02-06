# Analise do Formato de Ficheiro HAP 5.1 (.E3A)

## Estrutura do Arquivo

O ficheiro `.E3A` e um arquivo **ZIP** contendo:

| Ficheiro | Descricao |
|----------|-----------|
| HAP51SPC.DAT | Dados dos espacos (spaces) |
| HAP51A00.DAT | Sistemas de ar (air systems) |
| HAP51SCH.DAT | Schedules (horarios) |
| HAP51WAL.DAT | Paredes (walls) |
| HAP51WIN.DAT | Janelas (windows) |
| HAP51DOR.DAT | Portas (doors) |
| HAP51ROF.DAT | Telhados (roofs) |
| HAP51CHL.DAT | Chillers |
| HAP51BLR.DAT | Boilers |
| HAP51INX.MDB | Indices e referencias |
| HAP51WIZ.MDB | Wizard data (inclui tabela ASHRAE 62.1) |
| Project.mdb | Base de dados do projeto |
| PROJECT.E3P | Configuracao do projeto (INI file) |

---

## HAP51SPC.DAT - Estrutura dos Espacos

### Layout do Ficheiro

- **Bytes 0-3409**: Default Space (registo especial com configuracoes padrao)
- **Bytes 3410+**: Registos de espacos, **682 bytes cada**

### Campos Confirmados (por registo de espaco)

| Offset | Bytes | Tipo | Campo | Unidade Interna | Conversao |
|--------|-------|------|-------|-----------------|-----------|
| 0-23 | 24 | char[] | Nome do espaco | texto | - |
| 24-27 | 4 | float | Floor Area | ft^2 | x 0.0929 = m^2 |
| 28-31 | 4 | float | Ceiling Height | ft | x 0.3048 = m |
| 32-35 | 4 | float | Building Weight | lb/ft^2 | x 4.8824 = kg/m^2 |
| 472-475 | 4 | float | Sensible Fraction | decimal (0-1) | - |
| 476-479 | 4 | float | Cooling Setpoint | F | (x-32)/1.8 = C |
| 480-483 | 4 | float | Cooling HR | % | - |
| 484-487 | 4 | float | Heating Setpoint | F | (x-32)/1.8 = C |
| 488-491 | 4 | float | Heating HR | % | - |
| 580-583 | 4 | float | Occupancy | pessoas | - |
| 596-599 | 4 | float | (desconhecido) | ~0.48 | possivelmente fator |
| 606-609 | 4 | float | Lighting Wattage | W (total) | - |
| 656-659 | 4 | float | Equipment Density | W/ft^2 | x 10.764 = W/m^2 |

### OA Requirement (Outdoor Air) - CONFIRMADO

| Offset | Bytes | Tipo | Campo | Notas |
|--------|-------|------|-------|-------|
| 44-45 | 2 | bytes | Dados auxiliares OA | Nao-zero para L/s e L/s/pessoa |
| 46-49 | 4 | float | OA Value (interno) | Valor codificado (ver formula) |
| 50-51 | 2 | uint16 | OA Unit Code | 1=L/s, 2=L/s/m2, 3=L/s/pessoa, 4=% |

**Formulas de Conversao (Exactas - descobertas 2026-02-05):**

**AVISO:** As formulas lineares originais estavam ERRADAS. A codificacao OA
usa uma funcao `fast_exp2` (nao-linear). Ver `docs/OA_FORMULA.md` para detalhes completos.

Para **LER** o valor de OA do ficheiro:
```python
import math

Y0 = 512.0 * (28.316846592 / 60.0)  # 241.637 L/s

float_interno = struct.unpack('<f', data[offset+46:offset+50])[0]
unit_code = struct.unpack('<H', data[offset+50:offset+52])[0]

if unit_code in (1, 2, 3):  # L/s, L/s/m2, L/s/pessoa
    k = 4.0 if float_interno < 4.0 else 2.0
    t = k * (float_interno - 4.0)
    n = math.floor(t)
    OA = Y0 * (2.0 ** n) * (1.0 + (t - n))
elif unit_code == 4:  # %
    OA = float_interno * 28.5714
```

Para **ESCREVER** o valor de OA no ficheiro:
```python
if unit_code in (1, 2, 3):  # L/s, L/s/m2, L/s/pessoa
    v = OA / Y0
    n = math.floor(math.log2(v))
    f = v / (2.0 ** n) - 1.0
    t = n + f  # fast_log2
    float_interno = t / 4.0 + 4.0 if t < 0 else t / 2.0 + 4.0
elif unit_code == 4:  # %
    float_interno = OA / 28.5714
```

### Campos Parcialmente Identificados

| Offset | Bytes | Possivel Campo | Notas |
|--------|-------|----------------|-------|
| 36-39 | 4 | Flags/Type | Valor tipico: 4 |
| 72-78 | 6 | Wall Data 1 | Dados de paredes (direcao S) |
| 108-114 | 6 | Wall Data 2 | Dados de paredes (direcao SW) |
| 142-148 | 6 | Wall Data 3 | Dados de paredes (direcao W) |
| 176-182 | 6 | Wall Data 4 | Dados de paredes |
| 210-216 | 6 | Wall Data 5 | Dados de paredes |
| 244-250 | 6 | Wall Data 6 | Dados de paredes |

---

## Campos NAO Encontrados Diretamente

Os seguintes campos ainda NAO foram localizados nos registos de espaco:

1. **Sensible W/person** - Vem do Activity Level
2. **Latent W/person** - Vem do Activity Level
3. **Task Lighting** (W/m^2)
4. **Misc Loads** (W)
5. **Infiltration rates**

---

## OA Requirement - Validacao com Ficheiros de Teste

### Ficheiros de Teste Criados

| Ficheiro | Espacos | Configuracao OA |
|----------|---------|-----------------|
| TESTE2_3.E3A | 4 | 1: 50 L/s, 2: 50 L/s/m2, 3: 5 L/s/pessoa, 4: 50% |

### Dados Extraidos (HAP51SPC.DAT)

| Espaco | OA Config | Bytes 44-51 (hex) | Float@46 | Unit Code |
|--------|-----------|-------------------|----------|-----------|
| 1 | 50 L/s | 00a06a7c5a400100 | 3.4138 | 1 |
| 2 | 50 L/s/m2 | 00005faf23400200 | 2.5576 | 2 |
| 3 | 5 L/s/pessoa | 0060553025400300 | 2.5811 | 3 |
| 4 | 50 % | 00000000e03f0400 | 1.7500 | 4 |

### Verificacao das Formulas

| Espaco | Valor Real | Valor Calculado | Erro |
|--------|------------|-----------------|------|
| 1 (50 L/s) | 50.00 | 50.00 | 0.00% |
| 2 (50 L/s/m2) | 50.00 | 50.00 | 0.00% |
| 3 (5 L/s/p) | 5.00 | 5.00 | 0.00% |
| 4 (50 %) | 50.00 | 50.00 | 0.00% |

---

## Para Descobrir os Campos Restantes

### Metodo Recomendado

1. **Comparacao antes/depois**: Modificar um valor no HAP e comparar os ficheiros
2. **Debugging do executavel**: Usar debugger para ver como HAP le/escreve os campos
3. **Mais amostras**: Obter projetos com diferentes configuracoes

### Offsets a Investigar

- Bytes 36-72: Podem conter referencias adicionais
- Bytes 584-605: Zona entre ocupacao e iluminacao
- Bytes 610-655: Zona antes do equipment density

---

## Codigo Python para Leitura e Escrita

```python
import struct
import zipfile
import math

# OA encoding uses fast_exp2 (NOT linear). See docs/OA_FORMULA.md
_OA_Y0 = 512.0 * (28.316846592 / 60.0)  # 241.637 L/s

def _fast_exp2(t):
    n = math.floor(t)
    return (2.0 ** n) * (1.0 + (t - n))

def _fast_log2(v):
    n = math.floor(math.log2(v))
    f = v / (2.0 ** n) - 1.0
    if f < 0: n -= 1; f = v / (2.0 ** n) - 1.0
    if f >= 1.0: n += 1; f = v / (2.0 ** n) - 1.0
    return n + f

def decode_oa_value(x, unit_code):
    """Descodifica valor interno OA. Ver docs/OA_FORMULA.md"""
    if unit_code in (1, 2, 3):
        k = 4.0 if x < 4.0 else 2.0
        return _OA_Y0 * _fast_exp2(k * (x - 4.0))
    elif unit_code == 4:
        return x * 28.5714
    return 0

def encode_oa_value(oa_value, unit_code):
    """Codifica valor OA do utilizador. Ver docs/OA_FORMULA.md"""
    if unit_code in (1, 2, 3):
        t = _fast_log2(oa_value / _OA_Y0)
        return t / 4.0 + 4.0 if t < 0 else t / 2.0 + 4.0
    elif unit_code == 4:
        return oa_value / 28.5714
    return 0

OA_UNITS = {1: 'L/s', 2: 'L/s/m2', 3: 'L/s/pessoa', 4: '%'}

def read_hap_space(data, offset):
    """Le um registo de espaco do HAP51SPC.DAT"""
    space = {}
    space['name'] = data[offset:offset+24].decode('latin-1').rstrip('\x00').strip()
    space['area_ft2'] = struct.unpack('<f', data[offset+24:offset+28])[0]
    space['area_m2'] = space['area_ft2'] * 0.0929
    space['height_ft'] = struct.unpack('<f', data[offset+28:offset+32])[0]
    space['height_m'] = space['height_ft'] * 0.3048
    space['weight_lb_ft2'] = struct.unpack('<f', data[offset+32:offset+36])[0]

    # OA Requirement
    oa_float = struct.unpack('<f', data[offset+46:offset+50])[0]
    oa_unit_code = struct.unpack('<H', data[offset+50:offset+52])[0]
    space['oa_unit_code'] = oa_unit_code
    space['oa_unit'] = OA_UNITS.get(oa_unit_code, 'unknown')
    space['oa_value'] = decode_oa_value(oa_float, oa_unit_code)

    space['sensible_fraction'] = struct.unpack('<f', data[offset+472:offset+476])[0]
    space['cooling_setpoint_f'] = struct.unpack('<f', data[offset+476:offset+480])[0]
    space['cooling_setpoint_c'] = (space['cooling_setpoint_f'] - 32) / 1.8
    space['cooling_hr'] = struct.unpack('<f', data[offset+480:offset+484])[0]
    space['heating_setpoint_f'] = struct.unpack('<f', data[offset+484:offset+488])[0]
    space['heating_setpoint_c'] = (space['heating_setpoint_f'] - 32) / 1.8
    space['heating_hr'] = struct.unpack('<f', data[offset+488:offset+492])[0]
    space['occupancy'] = struct.unpack('<f', data[offset+580:offset+584])[0]
    space['lighting_w'] = struct.unpack('<f', data[offset+606:offset+610])[0]
    space['equip_density_w_ft2'] = struct.unpack('<f', data[offset+656:offset+660])[0]
    space['equip_density_w_m2'] = space['equip_density_w_ft2'] * 10.764
    return space

def read_e3a_spaces(filepath):
    """Le todos os espacos de um ficheiro .E3A"""
    with zipfile.ZipFile(filepath, 'r') as z:
        data = z.read('HAP51SPC.DAT')

    spaces = []
    # Primeiro registo (682 bytes) e o Default Space
    # Espacos reais comecam em offset 682
    offset = 682
    while offset + 682 <= len(data):
        space = read_hap_space(data, offset)
        if space['name']:  # Ignorar registos vazios
            spaces.append(space)
        offset += 682
    return spaces

# Exemplo de uso:
# spaces = read_e3a_spaces('projeto.E3A')
# for s in spaces:
#     print(f"{s['name']}: OA={s['oa_value']:.1f} {s['oa_unit']}")
```

---

## Referencias

- [HAP e-Help 006 - Ventilation In HAP](https://www.shareddocs.com/hvac/docs/1004/public/07/hap_ehelp_006.pdf)
- [HAP eHelp 011 - ASHRAE 62.1 Ventilation](https://www.shareddocs.com/hvac/docs/1004/public/0c/hap_ehelp_011.pdf)
- [How to Calculate Ventilation Air - MEP Academy](https://mepacademy.com/how-to-calculate-ventilation-air/)

---

*Analise realizada em Janeiro 2026*
