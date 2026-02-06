# HAP 5.1 File Tools

Ferramentas para ler, modificar e criar ficheiros Carrier HAP 5.1 (.E3A) programaticamente.

## Ficheiros Incluídos

### Documentação
| Ficheiro | Descrição |
|----------|-----------|
| `HAP_COMPLETE_FIELD_MAP.md` | Mapa completo de campos do registo de espaço (682 bytes) |
| `HAP_FILE_SPECIFICATION.md` | Especificação técnica detalhada do formato |
| `HAP_FILE_SPEC.json` | Especificação em formato JSON (machine-readable) |
| `HAP_FORMAT_ANALYSIS.md` | Notas de análise do formato binário |

### Biblioteca Python
| Ficheiro | Descrição |
|----------|-----------|
| `hap_library.py` | Biblioteca principal para ler/escrever ficheiros HAP |
| `hap_to_excel.py` | Exporta espaços HAP para Excel |
| `excel_to_hap.py` | Importa dados do Excel para ficheiro HAP |

### Templates Excel
| Ficheiro | Descrição |
|----------|-----------|
| `HAP_Input_Template_Complete.xlsx` | Template completo para entrada de dados |
| `create_excel_template_final.py` | Script para gerar o template |

## Uso Rápido

### Ler um ficheiro HAP existente
```python
from hap_library import HAPProject

project = HAPProject.open("projeto.E3A")

for space in project.spaces:
    print(f"{space.name}: {space.floor_area_m2:.1f} m²")
```

### Modificar um espaço
```python
space = project.get_space_by_name("Escritório 1")
space.occupancy = 10
space.overhead_lighting_w = 500
space.equipment_w_m2 = 15.0

project.save("projeto_modificado.E3A")
```

### Exportar para Excel
```bash
python hap_to_excel.py projeto.E3A
# Gera: projeto_export.xlsx
```

### Importar do Excel
```bash
python excel_to_hap.py dados.xlsx template.E3A projeto_novo.E3A
```

## Estrutura do Registo de Espaço (682 bytes)

| Bytes | Secção | Status |
|-------|--------|--------|
| 0-35 | Identificação e Dimensões | ✅ |
| 36-71 | Flags e Ar Exterior (OA) | ✅ |
| 72-359 | Paredes/Janelas/Portas (8 direcções) | ✅ |
| 360-439 | Pavimentos e Coberturas | ⚠️ |
| 440-491 | Termostato e Schedules | ✅ |
| 492-527 | Infiltração | ✅ |
| 528-579 | Partições | ✅ |
| 580-599 | Ocupação e Actividade | ✅ |
| 600-623 | Iluminação | ✅ |
| 624-655 | Cargas Miscelâneas | ✅ |
| 656-681 | Equipamento Eléctrico | ✅ |

## Conversões de Unidades

O HAP armazena valores em unidades imperiais. A biblioteca converte automaticamente:

| Campo | HAP (Imperial) | Métrico |
|-------|----------------|---------|
| Área | ft² | m² (× 0.0929) |
| Altura | ft | m (× 0.3048) |
| Temperatura | °F | °C ((x-32)/1.8) |
| Peso construção | lb/ft² | kg/m² (× 4.8824) |
| Equipamento | W/ft² | W/m² (× 10.764) |
| Calor | BTU/hr | W (÷ 3.412) |

## Fórmulas OA (Ar Exterior)

**NOTA:** As fórmulas lineares abaixo estavam INCORRECTAS.
A fórmula correcta é não-linear (fast_exp2). Ver `docs/OA_FORMULA.md`.

```python
# Formula correcta (descoberta 2026-02-05):
# Y0 = 512 CFM em L/s = 241.637
# Decode: y = Y0 * fast_exp2(k * (x - 4))   k=4 se x<4, k=2 se x>=4
# fast_exp2(t) = 2^floor(t) * (1 + frac(t))
# %: interno = valor / 28.5714 (esta sim é linear)
```

## Ficheiros HAP Relacionados

| Ficheiro | Conteúdo | Tamanho Registo |
|----------|----------|-----------------|
| HAP51SPC.DAT | Espaços | 682 bytes |
| HAP51WAL.DAT | Paredes | ~300 bytes |
| HAP51WIN.DAT | Janelas | 555 bytes |
| HAP51DOR.DAT | Portas | ~200 bytes |
| HAP51ROF.DAT | Coberturas | ~300 bytes |
| HAP51SCH.DAT | Schedules | variável |
| HAP51A00.DAT | Sistemas AVAC | variável |

## Limitações

- Estrutura de Pavimentos (bytes 360-399) parcialmente mapeada
- Estrutura de Coberturas (bytes 400-439) parcialmente mapeada
- Alguns campos de schedules ainda não identificados
- Referências a sistemas AVAC não implementadas

## Requisitos

- Python 3.7+
- openpyxl (para funcionalidades Excel)

```bash
pip install openpyxl
```

## Notas

- Ficheiros .E3A são arquivos ZIP contendo ficheiros .DAT binários e .MDB (Access)
- O primeiro registo de HAP51SPC.DAT (offset 0-681) é o template default
- Espaços do utilizador começam no offset 682

---
Última actualização: 2026-01-26
