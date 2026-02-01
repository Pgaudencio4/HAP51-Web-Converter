# HAP 5.1 Space Record - Complete Field Map

## Record Size: 682 bytes

Each space in `HAP51SPC.DAT` occupies exactly 682 bytes. The first record (offset 0-681) is the Default Space template.

---

## IDENTIFICATION & DIMENSIONS (Bytes 0-35)

| Offset | Size | Type | Field | Unit | Conversion |
|--------|------|------|-------|------|------------|
| 0-23 | 24 | char[] | Space Name | text | Latin-1 encoded, null-terminated |
| 24-27 | 4 | float | Floor Area | ft² | × 0.0929 = m² |
| 28-31 | 4 | float | Ceiling Height | ft | × 0.3048 = m |
| 32-35 | 4 | float | Building Weight | lb/ft² | × 4.8824 = kg/m² |

---

## FLAGS & OUTDOOR AIR (Bytes 36-71)

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 36-39 | 4 | uint32 | Type Flag | Typical value: 4 |
| 40-43 | 4 | - | Unknown | |
| 44-45 | 2 | bytes | OA Auxiliary | Non-zero for L/s and L/s/person units |
| 46-49 | 4 | float | OA Internal Value | Encoded value (see formulas below) |
| 50-51 | 2 | uint16 | OA Unit Code | 1=L/s, 2=L/s/m², 3=L/s/person, 4=% |
| 52-59 | 8 | - | OA Requirement 2 | Second OA requirement (same structure) |
| 60-63 | 4 | uint32 | Unknown | Typical value: 2 |
| 64-71 | 8 | - | Unknown | |

**OA Conversion Formulas (Verified):**
```python
# Reading (internal float to user value):
def decode_oa(internal_float, unit_code):
    if unit_code == 1:  # L/s
        return internal_float * 31.0336 / 2.11888
    elif unit_code == 2:  # L/s/m²
        return internal_float * 3.8484 / 0.19685
    elif unit_code == 3:  # L/s/person
        return internal_float * 4.1046 / 2.11888
    elif unit_code == 4:  # %
        return internal_float * 28.5714

# Writing (user value to internal float):
def encode_oa(user_value, unit_code):
    if unit_code == 1:  # L/s
        return user_value * 2.11888 / 31.0336
    elif unit_code == 2:  # L/s/m²
        return user_value * 0.19685 / 3.8484
    elif unit_code == 3:  # L/s/person
        return user_value * 2.11888 / 4.1046
    elif unit_code == 4:  # %
        return user_value / 28.5714
```

---

## WALLS, WINDOWS, DOORS (Bytes 72-359)

8 direction blocks, 36 bytes each:

| Offset | Direction |
|--------|-----------|
| 72-107 | South (S) |
| 108-143 | Southwest (SW) |
| 144-179 | West (W) |
| 180-215 | Northwest (NW) |
| 216-251 | North (N) |
| 252-287 | Northeast (NE) |
| 288-323 | East (E) |
| 324-359 | Southeast (SE) |

**Structure of each 36-byte direction block:**

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| +0 | 2 | uint16 | Wall Type ID | Reference to HAP51WAL.DAT |
| +2 | 4 | float | Gross Wall Area | ft² (× 0.0929 = m²) |
| +6 | 2 | uint16 | Flag | 1 if wall exists, 0 otherwise |
| +8 | 2 | uint16 | Window 1 Type ID | Reference to HAP51WIN.DAT |
| +10 | 2 | uint16 | Window 1 Quantity | Number of windows |
| +12 | 2 | uint16 | Window 2 Type ID | Second window type |
| +14 | 2 | uint16 | Window 2 Quantity | Number of windows |
| +16 | 2 | uint16 | Door Type ID | Reference to HAP51DOR.DAT |
| +18 | 2 | uint16 | Door Quantity | Number of doors |
| +20-35 | 16 | bytes | Additional data | Shades, overhangs, etc. |

**Note on Window References:**
- Window Type IDs reference records in HAP51WIN.DAT
- HAP51WIN.DAT has 555-byte records
- Each window is defined with glass layers, frame, area, etc.
- Window names often include space name + direction (e.g., "V1.05_Circulação2N")

---

## FLOORS & ROOFS (Bytes 360-439)

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 360-399 | 40 | struct | Floor Data | Floor type, area, U-value |
| 400-439 | 40 | struct | Roof Data | Roof type, area, slope, skylights |

**Floor Type Codes:**
- 1 = Floor Above Conditioned Space
- 2 = Floor Above Unconditioned Space
- 3 = Slab Floor On Grade
- 4 = Slab Floor Below Grade

**Roof Exposure Codes:**
- 0 = Not used
- 1-16 = Cardinal directions (N, NNE, NE, ENE, E, ESE, SE, SSE, S, SSW, SW, WSW, W, WNW, NW, NNW)
- 17 = Horizontal (H) - para coberturas planas

**Roof Block Structure (bytes 344-440, 4 roofs × 24 bytes):**
| Offset | Size | Type | Field |
|--------|------|------|-------|
| +0 | 2 | uint16 | Exposure Code |
| +2 | 2 | uint16 | Slope (degrees) |
| +4 | 4 | float | Gross Area (ft²) |
| +8 | 2 | uint16 | Roof Type ID |
| +10 | 2 | uint16 | Skylight Type ID |
| +12 | 2 | uint16 | Skylight Quantity |

**Note:** Structure partially mapped. Interior spaces typically have zeros here.

---

## THERMOSTAT & SCHEDULES (Bytes 440-491)

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 440-441 | 2 | uint16 | Thermostat Schedule ID | Reference to HAP51SCH.DAT |
| 442-445 | 4 | - | Unknown | |
| 446-471 | 26 | - | Duplicate/Cache | Contains duplicate of some thermostat values |
| 472-475 | 4 | float | Sensible Heat Ratio | 0.0-1.0 (typical: 0.5-0.7) |
| 476-479 | 4 | float | Cooling Setpoint | °F → (x-32)/1.8 = °C |
| 480-483 | 4 | float | Cooling RH | % (0-100) |
| 484-487 | 4 | float | Heating Setpoint | °F → (x-32)/1.8 = °C |
| 488-491 | 4 | float | Heating RH | % (0-100) |

---

## INFILTRATION (Bytes 492-527)

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 492-493 | 2 | uint16 | Infiltration Mode | 1=When Fan Off, 2=All Hours |
| 494-497 | 4 | float | Design Cooling CFM | Convert: CFM × 0.4719 = L/s |
| 498-501 | 4 | float | Design Cooling CFM/ft² | Convert: × 5.08 = L/s/m² |
| 502-505 | 4 | float | Design Cooling ACH | Air changes per hour |
| 506-509 | 4 | float | Design Heating CFM | |
| 510-513 | 4 | float | Design Heating CFM/ft² | |
| 514-517 | 4 | float | Design Heating ACH | |
| 518-521 | 4 | float | Energy Analysis CFM | |
| 522-525 | 4 | float | Energy Analysis CFM/ft² | |
| 526-527 | 2 | - | Energy Analysis ACH (partial) | |

---

## PARTITIONS (Bytes 528-579)

Two partition structures, 26 bytes each:

**Partition 1 (528-553):**

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 528-531 | 4 | float | Area | ft² |
| 532-535 | 4 | float | U-Value | BTU/(hr·ft²·°F) |
| 536-539 | 4 | float | Uncond. Space Max Temp | °F |
| 540-543 | 4 | float | Ambient at Max Temp | °F |
| 544-547 | 4 | float | Uncond. Space Min Temp | °F |
| 548-551 | 4 | float | Ambient at Min Temp | °F |
| 552-553 | 2 | uint16 | Partition Type | 1=Ceiling, 2=Wall |

**Partition 2 (554-579):** Same structure

---

## PEOPLE & ACTIVITY (Bytes 580-599)

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 580-583 | 4 | float | Occupancy | Number of people |
| 584-585 | 2 | uint16 | Activity Level ID | Reference to internal activity table |
| 586-589 | 4 | float | Sensible Heat | BTU/hr per person (÷3.412=W) |
| 590-593 | 4 | float | Latent Heat | BTU/hr per person (÷3.412=W) |
| 594-595 | 2 | uint16 | People Schedule ID | Reference to HAP51SCH.DAT |
| 596-599 | 4 | float | Unknown | Typical value: ~0.48 |

**Typical Values:**
- Office work: Sensible ~250 BTU/hr (73W), Latent ~200 BTU/hr (59W)
- Light activity: Sensible ~245 BTU/hr (72W), Latent ~205 BTU/hr (60W)

---

## LIGHTING (Bytes 600-623)

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 600-603 | 4 | float | Task Lighting | W total |
| 604-605 | 2 | uint16 | Fixture Type ID | See codes below |
| 606-609 | 4 | float | Overhead Lighting | W total |
| 610-613 | 4 | float | Ballast Multiplier | Typical: 1.0 |
| 614-615 | 2 | uint16 | Lighting Schedule ID | Reference to HAP51SCH.DAT |
| 616-619 | 4 | - | Unknown | |
| 620-623 | 4 | float | Task Lighting W/ft² | Alternative input method |

**Fixture Type Codes:**
- 0 = Recessed, unvented
- 1 = Recessed, vented to return air
- 2 = Recessed, vented to supply and return
- 3 = Surface mount / pendant

---

## MISCELLANEOUS LOADS (Bytes 624-655)

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 624-627 | 4 | - | Unknown | |
| 628-631 | 4 | float | Unknown | |
| 632-635 | 4 | float | Misc Sensible | BTU/hr (÷3.412=W) |
| 636-639 | 4 | float | Misc Latent | BTU/hr (÷3.412=W) |
| 640-641 | 2 | uint16 | Misc Sensible Schedule | Reference |
| 642-643 | 2 | - | Unknown | |
| 644-645 | 2 | uint16 | Misc Latent Schedule | Reference |
| 646-655 | 10 | - | Unknown | |

---

## ELECTRICAL EQUIPMENT (Bytes 656-681)

| Offset | Size | Type | Field | Description |
|--------|------|------|-------|-------------|
| 656-659 | 4 | float | Equipment Density | W/ft² (×10.764=W/m²) |
| 660-661 | 2 | uint16 | Equipment Schedule ID | Reference to HAP51SCH.DAT |
| 662-681 | 20 | - | Unknown / Padding | |

---

## RELATED FILES

| File | Record Size | Description |
|------|-------------|-------------|
| HAP51SPC.DAT | 682 bytes | Space definitions (this file) |
| HAP51WAL.DAT | ~300 bytes | Wall construction types |
| HAP51WIN.DAT | 555 bytes | Window assemblies |
| HAP51DOR.DAT | ~200 bytes | Door types |
| HAP51ROF.DAT | ~300 bytes | Roof constructions |
| HAP51SCH.DAT | variable | Schedules |
| HAP51A00.DAT | variable | Air handling systems |

---

## Summary by HAP Interface Tab

### General Tab ✅
- Name (0-23)
- Floor Area (24-27)
- Ceiling Height (28-31)
- Building Weight (32-35)
- Space Usage (36-39) - partial
- OA Requirement 1 (46-51)
- OA Requirement 2 (52-59) - same structure

### Internals Tab ✅
- Overhead Lighting Fixture Type (604-605)
- Overhead Lighting Wattage (606-609)
- Ballast Multiplier (610-613)
- Task Lighting (600-603)
- Electrical Equipment (656-659)
- Occupancy (580-583)
- Activity Level ID (584-585)
- Sensible/Latent Heat per person (586-593)
- Miscellaneous Loads (632-639)
- All Schedule IDs (594, 614, 640, 644, 660)

### Walls, Windows, Doors Tab ✅
- 8 Directions (72-359)
- Wall Type ID, Gross Area
- Window 1 & 2 Type ID, Quantity
- Door Type ID, Quantity

### Roofs, Skylights Tab ⚠️
- Roof data (400-439) - structure partially mapped

### Infiltration Tab ✅
- Mode (492-493)
- Design Cooling (494-505)
- Design Heating (506-517)
- Energy Analysis (518-527)

### Floors Tab ⚠️
- Floor Type (360-399) - structure partially mapped

### Partitions Tab ✅
- Partition 1 (528-553)
- Partition 2 (554-579)

---

## Python Library

A complete Python library is available in `hap_library.py`:

```python
from hap_library import HAPProject, HAPSpace

# Open existing project
project = HAPProject.open("project.E3A")

# List spaces
for space in project.spaces:
    print(f"{space.name}: {space.floor_area_m2:.1f} m²")

# Modify a space
space = project.get_space_by_name("Office 1")
space.occupancy = 10
space.overhead_lighting_w = 500
space.equipment_w_m2 = 15.0

# Save
project.save("modified.E3A")
```

---

## Version

- Last updated: 2026-01-26
- Status: ~90% complete
- Verified: OA encoding, wall blocks, thermostat, people, lighting, equipment, infiltration
- Partially verified: Floor/Roof structure, some schedule references
