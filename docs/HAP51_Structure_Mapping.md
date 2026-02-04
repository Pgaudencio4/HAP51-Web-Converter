# Mapeamento da Estrutura de Ficheiros HAP 5.1

## Resumo da Análise

Análise de engenharia reversa dos ficheiros do Carrier HAP (Hourly Analysis Program) versão 5.1.
Projeto analisado: **Vale Formoso - Porto**

---

## HAP51SPC.DAT - Estrutura de Espaços

### Informação Geral
- **Tamanho por registo:** 682 bytes
- **Total de espaços:** 34 (1 Default + 33 espaços do projeto)
- **Unidades internas:** Sistema Imperial (ft, ft², °F, lb/ft²)

### Campos Confirmados ✓

| Offset | Bytes | Tipo | Campo | Descrição | Unidade HAP | Conversão |
|--------|-------|------|-------|-----------|-------------|-----------|
| 0 | 24 | char[24] | szName | Nome do espaço | texto | - |
| 24 | 4 | float | fFloorArea | Área do piso | ft² | ÷ 10.7639 → m² |
| 28 | 4 | float | fCeilingHeight | Altura do teto | ft | ÷ 3.28084 → m |
| 32 | 4 | float | fBuildingWeight | Peso do edifício | lb/ft² | ÷ 0.2048 → kg/m² |
| 472 | 4 | float | fSensibleFraction | Fração calor sensível | decimal | direto |
| 476 | 4 | float | fCoolingSetpoint | Setpoint arrefecimento | °F | (°F-32)×5/9 → °C |
| 480 | 4 | float | fCoolingHR | HR arrefecimento | % | direto |
| 484 | 4 | float | fHeatingSetpoint | Setpoint aquecimento | °F | (°F-32)×5/9 → °C |
| 488 | 4 | float | fHeatingHR | HR aquecimento | % | direto |
| 580 | 4 | float | fNumPeople | Ocupação | pessoas | direto |
| 606 | 4 | float | fLightingWattage | Potência iluminação | W | direto |

### Campos Parcialmente Identificados ⚠️

| Offset | Bytes | Tipo | Campo Provável | Notas |
|--------|-------|------|----------------|-------|
| 108 | 4 | float | fWallArea1 | Área parede exposição 1 (ft²) |
| 176 | 4 | float | fWallArea2 | Área parede exposição 2 (ft²) |
| 244 | 4 | float | fWallArea3 | Área parede exposição 3 (ft²) |
| 596 | 4 | float | ? | Valor ~0.48 constante |
| 656 | 4 | float | fElecEquipDensity? | Valor ~0.186 (pode ser W/ft²) |

### Verificação Cruzada (5 espaços)

| Campo | Offset | -1.04_IS | 0.01_Circ | 0.02_Sala1 | 1.03_Coz | 2.08_Sala4 |
|-------|--------|----------|-----------|------------|----------|------------|
| Área (ft²) | 24 | 34.4 | 6451.9 | 207.7 | 221.7 | 341.2 |
| Altura (ft) | 28 | 7.55 | 13.12 | 9.84 | 9.84 | 8.86 |
| Peso (lb/ft²) | 32 | 56.32 | 56.32 | 56.32 | 56.32 | 56.32 |
| Ocupação | 580 | 2 | 300 | 10 | 0 | 16 |
| Iluminação (W) | 606 | 23.4 | 1939.6 | 712.4 | 257.4 | 70.2 |
| Cool Setpoint | 476 | 75°F | 75°F | 75°F | 75°F | 75°F |
| Heat Setpoint | 484 | 75°F | 75°F | 75°F | 75°F | 75°F |

---

## HAP51INX.MDB - Base de Dados de Índices

### Tabela: SpaceIndex
| Campo | Tipo | Descrição |
|-------|------|-----------|
| nIndex | int | Índice do espaço (1-based) |
| szName | text | Nome do espaço |
| fFloorArea | float | Área (ft²) - **duplicado do DAT** |
| fNumPeople | float | Ocupação - **duplicado do DAT** |
| fLightingDensity | float | Iluminação (W) - **duplicado do DAT** |

### Tabelas de Relacionamento
- Space_Wall_Links
- Space_Window_Links
- Space_Roof_Links
- Space_Door_Links
- Space_Schedule_Links
- Building_System_Links
- System_Space_Links
- Plant_Equipment_Links

---

## Campos NÃO Encontrados no HAP51SPC.DAT

Os seguintes campos da interface HAP **não** foram localizados no ficheiro de espaços:

- **OA Requirement 1** (L/s) - Provavelmente no sistema de ar (HAP51A00.DAT)
- **OA Requirement 2** (L/s·m²)
- **Task Lighting** (W/m²)
- **Electrical Equipment** (W/m²)
- **Misc. Loads** (Sensible/Latent W)
- **Partition Areas** e U-Values detalhados
- **Infiltration rates**
- **Sensible/Latent W/person** (71.8/60.1)

Estes campos podem estar:
1. Noutros ficheiros DAT (A00, P00, etc.)
2. Calculados a partir de outros valores
3. Armazenados como referências (IDs de schedules, types)

---

## Fórmulas de Conversão

```
Área:       m² = ft² / 10.7639
Altura:     m = ft / 3.28084  
Peso:       kg/m² = lb/ft² / 0.2048
Temperatura: °C = (°F - 32) × 5/9
Ventilação: L/s = CFM / 2.11888
```

---

## Recomendações para Edição Externa

### ✅ Seguro Editar (com backup)
- Nomes de espaços (offset 0-23 no DAT + szName no MDB)
- Áreas (offset 24 no DAT + fFloorArea no MDB) - **editar ambos!**
- Ocupação (offset 580 no DAT + fNumPeople no MDB) - **editar ambos!**
- Iluminação (offset 606 no DAT + fLightingDensity no MDB) - **editar ambos!**

### ⚠️ Editar com Cuidado
- Altura do teto (apenas DAT)
- Setpoints temperatura (apenas DAT)
- Building Weight (apenas DAT)

### ❌ Não Recomendado
- Campos não identificados
- Estrutura de links MDB
- Ficheiros de sistema (A00, P00, B00)

---

## Estrutura de Ficheiros do Projeto

```
Vale_Formoso_-_Porto/
├── PROJECT.E3P          # Metadados (INI text)
├── HAP51INX.MDB         # Índices e links (Access)
├── HAP51SPC.DAT         # Dados dos espaços (binário)
├── HAP51WAL.DAT         # Tipos de parede
├── HAP51WIN.DAT         # Tipos de janela
├── HAP51ROF.DAT         # Tipos de cobertura
├── HAP51SCH.DAT         # Schedules
├── HAP51A00.DAT         # Sistemas de ar
├── HAP51B00/B01.DAT     # Edifícios
├── HAP51P00/P01/P02.DAT # Plantas
├── HAP51CHL.DAT         # Chillers
├── HAP51BLR.DAT         # Caldeiras
├── HAP51TWR.DAT         # Torres de arrefecimento
└── HAP51WTA/WTD.DAT     # Dados meteorológicos
```

---

## Código Python para Leitura

```python
import struct

SPACE_SIZE = 682

def read_space(data, index):
    """Lê dados de um espaço do HAP51SPC.DAT"""
    offset = index * SPACE_SIZE
    
    name = data[offset:offset+24].decode('latin-1').strip('\x00')
    area_ft2 = struct.unpack('<f', data[offset+24:offset+28])[0]
    height_ft = struct.unpack('<f', data[offset+28:offset+32])[0]
    weight_lbft2 = struct.unpack('<f', data[offset+32:offset+36])[0]
    occupancy = struct.unpack('<f', data[offset+580:offset+584])[0]
    lighting_w = struct.unpack('<f', data[offset+606:offset+610])[0]
    cool_sp_f = struct.unpack('<f', data[offset+476:offset+480])[0]
    heat_sp_f = struct.unpack('<f', data[offset+484:offset+488])[0]
    
    return {
        'name': name,
        'area_m2': area_ft2 / 10.7639,
        'height_m': height_ft / 3.28084,
        'weight_kgm2': weight_lbft2 / 0.2048,
        'occupancy': occupancy,
        'lighting_w': lighting_w,
        'cool_setpoint_c': (cool_sp_f - 32) * 5/9,
        'heat_setpoint_c': (heat_sp_f - 32) * 5/9,
    }

# Exemplo de uso
with open('HAP51SPC.DAT', 'rb') as f:
    data = f.read()

for i in range(1, 34):  # Skip Default (index 0)
    space = read_space(data, i)
    print(f"{space['name']}: {space['area_m2']:.1f} m², {space['occupancy']:.0f} pessoas")
```

---

*Análise realizada em 26/01/2026*
*Projeto: Vale Formoso - Porto*
*Software: Carrier HAP 5.1*
