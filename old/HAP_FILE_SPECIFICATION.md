# Carrier HAP 5.1 File Format Specification

## Overview

This document describes the binary file format used by Carrier HAP (Hourly Analysis Program) version 5.1 for HVAC load calculation projects. The file extension is `.E3A`.

## File Structure

An `.E3A` file is a **ZIP archive** containing multiple data files:

| File | Description | Format |
|------|-------------|--------|
| `HAP51SPC.DAT` | Space (room) definitions | Binary records |
| `HAP51A00.DAT` | Air handling systems | Binary records |
| `HAP51SCH.DAT` | Schedules | Binary records |
| `HAP51WAL.DAT` | Wall constructions | Binary records |
| `HAP51WIN.DAT` | Window constructions | Binary records |
| `HAP51DOR.DAT` | Door constructions | Binary records |
| `HAP51ROF.DAT` | Roof constructions | Binary records |
| `HAP51CHL.DAT` | Chillers | Binary records |
| `HAP51BLR.DAT` | Boilers | Binary records |
| `HAP51TWR.DAT` | Cooling towers | Binary records |
| `HAP51P00.DAT` | Plant equipment | Binary records |
| `HAP51WTD.DAT` | Design weather data | Binary records |
| `HAP51WTA.DAT` | Simulation weather data | Binary records |
| `HAP51INX.MDB` | Project index database | MS Access |
| `HAP51WIZ.MDB` | Wizard/ASHRAE tables | MS Access |
| `Project.mdb` | Project database | MS Access |
| `PROJECT.E3P` | Project configuration | INI text file |

---

## HAP51SPC.DAT - Space Records

### Record Layout

- **Record size**: 682 bytes per space
- **First record (offset 0-681)**: Default Space template
- **Subsequent records**: User-defined spaces starting at offset 682

### Number of Spaces Calculation

```python
num_spaces = (file_size - 682) // 682
```

### Field Definitions

All numeric values are **little-endian**. Floats are IEEE 754 single-precision (4 bytes).

#### Basic Properties

| Offset | Size | Type | Field | Internal Unit | To Metric |
|--------|------|------|-------|---------------|-----------|
| 0-23 | 24 | char[] | Space Name | ASCII/Latin-1 | - |
| 24-27 | 4 | float | Floor Area | ft² | × 0.0929 = m² |
| 28-31 | 4 | float | Ceiling Height | ft | × 0.3048 = m |
| 32-35 | 4 | float | Building Weight | lb/ft² | × 4.8824 = kg/m² |

#### Wall/Window Data by Direction

Each space has 8 direction blocks (36 bytes each) for exterior walls and windows:

| Direction | Offset | Description |
|-----------|--------|-------------|
| S (South) | 72-107 | South-facing wall |
| SW (Southwest) | 108-143 | Southwest-facing wall |
| W (West) | 144-179 | West-facing wall |
| NW (Northwest) | 180-215 | Northwest-facing wall |
| N (North) | 216-251 | North-facing wall |
| NE (Northeast) | 252-287 | Northeast-facing wall |
| E (East) | 288-323 | East-facing wall |
| SE (Southeast) | 324-359 | Southeast-facing wall |

**Structure of each 36-byte direction block:**

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 0-1 | 2 | uint16 | Wall Type ID | Reference to wall type in HAP51WAL.DAT |
| 2-5 | 4 | float | Gross Wall Area | Wall area in ft² (× 0.0929 = m²) |
| 6-7 | 2 | uint16 | Unknown | Always 1 when wall exists |
| 8-9 | 2 | uint16 | Window Type ID | Reference to window type in HAP51WIN.DAT |
| 10-11 | 2 | uint16 | Unknown | Flag (0 or 1) |
| 12-35 | 24 | bytes | Additional data | May contain window area, door data |

**Example - South wall with 6.0 m² area:**
```
Hex: 0900 BC2A8142 0100 0100 0100 000000...
     ---- -------- ---- ---- ----
     |    |        |    |    |
     |    |        |    |    +-- Unknown (1)
     |    |        |    +------- Window Type ID (1)
     |    |        +------------ Unknown (1)
     |    +--------------------- Wall Area: 64.58 ft² = 6.0 m²
     +-------------------------- Wall Type ID: 9
```

#### Outdoor Air (OA) Requirement

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 44-45 | 2 | bytes | OA Auxiliary | Additional data (non-zero for L/s, L/s/person) |
| 46-49 | 4 | float | OA Internal Value | Encoded OA value (see conversion formulas) |
| 50-51 | 2 | uint16 | OA Unit Code | 1=L/s, 2=L/s/m², 3=L/s/person, 4=% |

**OA Unit Codes:**
- `1` = L/s (total airflow)
- `2` = L/s/m² (airflow per floor area)
- `3` = L/s/person (airflow per occupant)
- `4` = % (percentage of supply air)

**Conversion Formulas - Reading:**
```python
def decode_oa(float_value, unit_code):
    if unit_code == 1:  # L/s
        return float_value * 31.0336 / 2.11888
    elif unit_code == 2:  # L/s/m²
        return float_value * 3.8484 / 0.19685
    elif unit_code == 3:  # L/s/person
        return float_value * 4.1046 / 2.11888
    elif unit_code == 4:  # %
        return float_value * 28.5714
    return 0
```

**Conversion Formulas - Writing:**
```python
def encode_oa(oa_value, unit_code):
    if unit_code == 1:  # L/s
        return oa_value * 2.11888 / 31.0336
    elif unit_code == 2:  # L/s/m²
        return oa_value * 0.19685 / 3.8484
    elif unit_code == 3:  # L/s/person
        return oa_value * 2.11888 / 4.1046
    elif unit_code == 4:  # %
        return oa_value / 28.5714
    return 0
```

#### Thermostat Settings

| Offset | Size | Type | Field | Internal Unit | To Metric |
|--------|------|------|-------|---------------|-----------|
| 472-475 | 4 | float | Sensible Heat Ratio | decimal (0-1) | - |
| 476-479 | 4 | float | Cooling Setpoint | °F | (x-32)/1.8 = °C |
| 480-483 | 4 | float | Cooling Relative Humidity | % | - |
| 484-487 | 4 | float | Heating Setpoint | °F | (x-32)/1.8 = °C |
| 488-491 | 4 | float | Heating Relative Humidity | % | - |

#### Internal Loads

| Offset | Size | Type | Field | Internal Unit | To Metric |
|--------|------|------|-------|---------------|-----------|
| 580-583 | 4 | float | Occupancy | people | - |
| 596-599 | 4 | float | Unknown Factor | ~0.48 | Purpose unknown |
| 606-609 | 4 | float | Lighting Power | W (total) | - |
| 656-659 | 4 | float | Equipment Density | W/ft² | × 10.764 = W/m² |

#### Wall/Window References (Partially Decoded)

| Offset | Size | Description |
|--------|------|-------------|
| 72-78 | 6 | Wall data direction 1 (S) |
| 108-114 | 6 | Wall data direction 2 (SW) |
| 142-148 | 6 | Wall data direction 3 (W) |
| 176-182 | 6 | Wall data direction 4 |
| 210-216 | 6 | Wall data direction 5 |
| 244-250 | 6 | Wall data direction 6 |

---

## Unit Conversion Constants

### Length
- 1 ft = 0.3048 m
- 1 m = 3.28084 ft

### Area
- 1 ft² = 0.0929 m²
- 1 m² = 10.7639 ft²

### Temperature
- °C = (°F - 32) / 1.8
- °F = °C × 1.8 + 32

### Airflow
- 1 CFM = 0.4719 L/s
- 1 L/s = 2.11888 CFM
- 1 CFM/ft² = 5.0799 L/s/m²
- 1 L/s/m² = 0.19685 CFM/ft²

### Power/Heat
- 1 BTU/hr = 0.2931 W
- 1 W = 3.412 BTU/hr
- 1 W/ft² = 10.764 W/m²

### Mass
- 1 lb/ft² = 4.8824 kg/m²

---

## Complete Python Implementation

```python
import struct
import zipfile
from dataclasses import dataclass
from typing import List, Optional

# Constants
OA_UNITS = {1: 'L/s', 2: 'L/s/m²', 3: 'L/s/person', 4: '%'}
RECORD_SIZE = 682

@dataclass
class HAPSpace:
    name: str
    area_m2: float
    area_ft2: float
    height_m: float
    height_ft: float
    building_weight_kg_m2: float
    oa_value: float
    oa_unit: str
    oa_unit_code: int
    sensible_ratio: float
    cooling_setpoint_c: float
    cooling_rh: float
    heating_setpoint_c: float
    heating_rh: float
    occupancy: float
    lighting_w: float
    equipment_w_m2: float


def decode_oa_value(float_interno: float, unit_code: int) -> float:
    """Convert internal OA value to user units."""
    if unit_code == 1:  # L/s
        return float_interno * 31.0336 / 2.11888
    elif unit_code == 2:  # L/s/m²
        return float_interno * 3.8484 / 0.19685
    elif unit_code == 3:  # L/s/person
        return float_interno * 4.1046 / 2.11888
    elif unit_code == 4:  # %
        return float_interno * 28.5714
    return 0.0


def encode_oa_value(oa_value: float, unit_code: int) -> float:
    """Convert user OA value to internal format."""
    if unit_code == 1:  # L/s
        return oa_value * 2.11888 / 31.0336
    elif unit_code == 2:  # L/s/m²
        return oa_value * 0.19685 / 3.8484
    elif unit_code == 3:  # L/s/person
        return oa_value * 2.11888 / 4.1046
    elif unit_code == 4:  # %
        return oa_value / 28.5714
    return 0.0


def read_space_record(data: bytes, offset: int) -> Optional[HAPSpace]:
    """Read a single space record from HAP51SPC.DAT data."""
    if offset + RECORD_SIZE > len(data):
        return None

    # Name (24 bytes, null-terminated string)
    name = data[offset:offset+24].decode('latin-1').rstrip('\x00').strip()
    if not name:
        return None

    # Basic properties
    area_ft2 = struct.unpack('<f', data[offset+24:offset+28])[0]
    height_ft = struct.unpack('<f', data[offset+28:offset+32])[0]
    weight_lb_ft2 = struct.unpack('<f', data[offset+32:offset+36])[0]

    # OA Requirement
    oa_float = struct.unpack('<f', data[offset+46:offset+50])[0]
    oa_unit_code = struct.unpack('<H', data[offset+50:offset+52])[0]
    oa_value = decode_oa_value(oa_float, oa_unit_code)
    oa_unit = OA_UNITS.get(oa_unit_code, 'unknown')

    # Thermostat
    sensible_ratio = struct.unpack('<f', data[offset+472:offset+476])[0]
    cooling_f = struct.unpack('<f', data[offset+476:offset+480])[0]
    cooling_rh = struct.unpack('<f', data[offset+480:offset+484])[0]
    heating_f = struct.unpack('<f', data[offset+484:offset+488])[0]
    heating_rh = struct.unpack('<f', data[offset+488:offset+492])[0]

    # Internal loads
    occupancy = struct.unpack('<f', data[offset+580:offset+584])[0]
    lighting_w = struct.unpack('<f', data[offset+606:offset+610])[0]
    equip_w_ft2 = struct.unpack('<f', data[offset+656:offset+660])[0]

    return HAPSpace(
        name=name,
        area_m2=area_ft2 * 0.0929,
        area_ft2=area_ft2,
        height_m=height_ft * 0.3048,
        height_ft=height_ft,
        building_weight_kg_m2=weight_lb_ft2 * 4.8824,
        oa_value=oa_value,
        oa_unit=oa_unit,
        oa_unit_code=oa_unit_code,
        sensible_ratio=sensible_ratio,
        cooling_setpoint_c=(cooling_f - 32) / 1.8,
        cooling_rh=cooling_rh,
        heating_setpoint_c=(heating_f - 32) / 1.8,
        heating_rh=heating_rh,
        occupancy=occupancy,
        lighting_w=lighting_w,
        equipment_w_m2=equip_w_ft2 * 10.764
    )


def read_e3a_file(filepath: str) -> List[HAPSpace]:
    """Read all spaces from a HAP .E3A file."""
    spaces = []

    with zipfile.ZipFile(filepath, 'r') as z:
        data = z.read('HAP51SPC.DAT')

    # Skip default space (first 682 bytes)
    offset = RECORD_SIZE

    while offset + RECORD_SIZE <= len(data):
        space = read_space_record(data, offset)
        if space:
            spaces.append(space)
        offset += RECORD_SIZE

    return spaces


def modify_space_oa(data: bytearray, space_index: int, new_oa_value: float,
                    new_unit_code: Optional[int] = None) -> bytearray:
    """Modify OA value for a specific space in HAP51SPC.DAT data."""
    offset = RECORD_SIZE * (space_index + 1)  # +1 to skip default space

    if new_unit_code is None:
        # Keep existing unit code
        new_unit_code = struct.unpack('<H', data[offset+50:offset+52])[0]

    # Encode and write new OA value
    internal_value = encode_oa_value(new_oa_value, new_unit_code)
    data[offset+46:offset+50] = struct.pack('<f', internal_value)
    data[offset+50:offset+52] = struct.pack('<H', new_unit_code)

    return data


# Example usage
if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("Usage: python hap_reader.py <file.E3A>")
        sys.exit(1)

    spaces = read_e3a_file(sys.argv[1])

    print(f"Found {len(spaces)} spaces:\n")
    for i, s in enumerate(spaces):
        print(f"{i+1}. {s.name}")
        print(f"   Area: {s.area_m2:.1f} m²")
        print(f"   Height: {s.height_m:.2f} m")
        print(f"   OA: {s.oa_value:.1f} {s.oa_unit}")
        print(f"   Cooling: {s.cooling_setpoint_c:.1f}°C / {s.cooling_rh:.0f}% RH")
        print(f"   Heating: {s.heating_setpoint_c:.1f}°C / {s.heating_rh:.0f}% RH")
        print(f"   Occupancy: {s.occupancy:.0f} people")
        print(f"   Lighting: {s.lighting_w:.0f} W")
        print(f"   Equipment: {s.equipment_w_m2:.1f} W/m²")
        print()
```

---

## Fields Not Yet Decoded

The following fields have not been located in the binary structure:

| Field | Expected Location | Notes |
|-------|-------------------|-------|
| Sensible W/person | Unknown | May be in Activity Level reference |
| Latent W/person | Unknown | May be in Activity Level reference |
| Task Lighting | Unknown | Separate from general lighting |
| Misc Sensible Load | Unknown | Additional sensible heat gains |
| Misc Latent Load | Unknown | Additional latent heat gains |
| Infiltration Rate | Unknown | Air infiltration CFM or ACH |
| Space Usage Type | Bytes 36-43? | Reference to ASHRAE 62.1 table |

---

## Binary Data Examples

### Example: 50 L/s OA Requirement

```
Offset 44-51: 00 A0 6A 7C 5A 40 01 00

Breakdown:
- Bytes 44-45: 00 A0 (auxiliary data)
- Bytes 46-49: 6A 7C 5A 40 = float 3.4138
- Bytes 50-51: 01 00 = unit code 1 (L/s)

Calculation: 3.4138 * 31.0336 / 2.11888 = 50.00 L/s
```

### Example: 50% OA Requirement

```
Offset 44-51: 00 00 00 00 E0 3F 04 00

Breakdown:
- Bytes 44-45: 00 00 (zero)
- Bytes 46-49: 00 00 E0 3F = float 1.75
- Bytes 50-51: 04 00 = unit code 4 (%)

Calculation: 1.75 * 28.5714 = 50.00%
```

---

## Validation Test Cases

| Input | Unit | Expected Float | Actual Float | Match |
|-------|------|----------------|--------------|-------|
| 50 | L/s | 3.4138 | 3.4138 | Yes |
| 50 | L/s/m² | 2.5576 | 2.5576 | Yes |
| 5 | L/s/person | 2.5811 | 2.5811 | Yes |
| 50 | % | 1.7500 | 1.7500 | Yes |

---

## Version History

- **2026-01-26**: Initial specification created
- **2026-01-26**: OA Requirement encoding discovered and validated

---

## References

- Carrier HAP 5.1 Software
- ASHRAE Standard 62.1 (Ventilation for Acceptable Indoor Air Quality)
- IEEE 754 Floating-Point Standard
