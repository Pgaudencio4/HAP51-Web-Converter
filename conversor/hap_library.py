"""
HAP 5.1 File Library
====================
Read and write Carrier HAP 5.1 .E3A files programmatically.

Author: Generated from reverse engineering
Version: 1.2
Date: 2026-01-26

FILE STRUCTURE SUMMARY
======================

.E3A files are ZIP archives containing:
  - HAP51SPC.DAT: Space records (682 bytes each)
  - HAP51SCH.DAT: Schedule records (792 bytes each)
  - HAP51INX.MDB: MS Access database with indexes and links
  - HAP51WAL.DAT: Wall type definitions
  - HAP51WIN.DAT: Window type definitions
  - HAP51ROF.DAT: Roof type definitions
  - HAP51DOR.DAT: Door type definitions
  - HAP51SHD.DAT: Shade type definitions (592 bytes)
  - HAP51WTA/WTD.DAT: Weather data
  - And others...

SPACE RECORD STRUCTURE (682 bytes)
----------------------------------
Offset    Size  Field
------    ----  -----
[0-24]      24  Name (null-padded string, latin-1)
[24-28]      4  Floor area (ft², float)
[28-32]      4  Ceiling height (ft, float)
[32-36]      4  Building weight (lb/ft², float)
[36-40]      4  Type flag (uint32)
[44-46]      2  OA auxiliary
[46-50]      4  OA internal value (float, exponential encoding)
[50-52]      2  OA unit code (1=L/s, 2=L/s/m², 3=L/s/person, 4=%)
[72-344]   272  Wall blocks (8 x 34 bytes each)
[344-440]   96  Roof blocks (4 x 24 bytes each)
[440-466]   26  Partition 1: Type(H) Area(f) U(f) UncMax(f) OutMax(f) UncMin(f) OutMin(f)
[466-492]   26  Partition 2: Type(H) Area(f) U(f) UncMax(f) OutMax(f) UncMin(f) OutMin(f)
[492-494]    2  Floor Type (1=Above Cond, 2=Above Uncond, 3=Slab On Grade, 4=Slab Below)
[494-498]    4  Floor Area (ft², float)
[498-502]    4  Floor U-value (IP units, float) - convert: SI/5.678
[502-506]    4  Slab Exposed Perimeter (ft, float) - for types 3,4
[506-510]    4  Slab Edge Insulation R-value (IP, float) - for type 3
[510-514]    4  Slab Floor Depth (ft, float) - for type 4 (Below Grade)
[514-518]    4  Basement Wall U-value (IP, float) - for type 4
[518-522]    4  Wall Insulation R-value (IP, float) - for type 4
[522-526]    4  Depth of Wall Insulation (ft, float) - for type 4
[440-442]    2  Partition 1 Type (1=Ceiling, 2=Wall)
[442-446]    4  Partition 1 Area (ft², float)
[446-450]    4  Partition 1 U-value (IP, float) - convert: SI/5.678
[450-454]    4  Partition 1 Uncond Max Temp (°F, float)
[454-458]    4  Partition 1 Ambient Max Temp (°F, float)
[458-462]    4  Partition 1 Uncond Min Temp (°F, float)
[462-466]    4  Partition 1 Ambient Min Temp (°F, float)
[466-468]    2  Partition 2 Type (1=Ceiling, 2=Wall)
[468-472]    4  Partition 2 Area (ft², float)
[472-476]    4  Partition 2 U-value (IP, float)
[476-480]    4  Partition 2 Uncond Max Temp (°F, float)
[480-484]    4  Partition 2 Ambient Max Temp (°F, float)
[484-488]    4  Partition 2 Uncond Min Temp (°F, float)
[488-492]    4  Partition 2 Ambient Min Temp (°F, float)
[492-494]    2  Floor Type (1=Above Cond, 2=Above Uncond, 3=Slab On Grade, 4=Slab Below)
[494-542]   48  Floor data (Area, U-value, Perimeter, Depths, Temps)
[554-556]    2  Design Cooling flag (2=ACH mode)
[556-560]    4  Design Cooling ACH (float)
[560-562]    2  Design Heating flag (2=ACH mode)
[562-566]    4  Design Heating ACH (float)
[566-568]    2  Energy Analysis flag (2=ACH mode)
[568-572]    4  Energy Analysis ACH (float)
[580-584]    4  Occupancy (float)
[584-586]    2  Activity level ID
[586-590]    4  Sensible heat (BTU/hr, float)
[590-594]    4  Latent heat (BTU/hr, float)
[594-596]    2  People schedule ID
[600-604]    4  Task lighting (W, float)
[604-606]    2  Fixture type ID
[606-610]    4  Overhead lighting (W, float)
[610-614]    4  Ballast multiplier (float)
[616-618]    2  Lighting schedule ID (NOT [614]!)
[656-660]    4  Equipment (W/ft², float)
[660-662]    2  Equipment schedule ID

WALL BLOCK STRUCTURE (34 bytes each)
------------------------------------
Offset  Size  Field
------  ----  -----
[0-2]     2   Direction code (1=N, 2=NNE, 3=NE, ... 16=NNW)
[2-6]     4   Wall gross area (ft², float)
[6-8]     2   Wall type ID (WallIndex)
[8-10]    2   Window 1 type ID (WindowIndex)
[10-12]   2   Window 2 type ID (WindowIndex)
[12-14]   2   Window 1 quantity
[14-16]   2   Window 2 quantity
[16-18]   2   Door type ID (DoorIndex)
[18-20]   2   Door quantity
[20-34]  14   Reserved

ROOF BLOCK STRUCTURE (24 bytes each, 4 blocks starting at offset 344)
---------------------------------------------------------------------
Offset  Size  Field
------  ----  -----
[0-2]     2   Direction code (1=N, 2=NNE, 3=NE, ... 16=NNW)
[2-4]     2   Roof slope (degrees)
[4-8]     4   Roof gross area (ft², float)
[8-10]    2   Roof type ID (RoofIndex)
[10-12]   2   Skylight type ID (WindowIndex - reuses windows!)
[12-14]   2   Skylight quantity
[14-24]  10   Reserved

TESTED AND CONFIRMED (2026-01-26):
- Roof type ID must exist in RoofIndex table in MDB
- Skylight uses WindowIndex (same as windows)
- Must add link in Space_Roof_Links (Space_ID, Roof_ID)
- Must add link in Space_Window_Links for each skylight used

OA ENCODING (piecewise fast_exp2 formula - corrected 2026-02-05)
----------------------------------------------------------------
Y0 = 512 CFM in L/s = 512 * 28.316846592 / 60 = 241.637...
fast_exp2(t) = 2^floor(t) * (1 + frac(t))
Decode: if x<4: t=(x-4)*4, else: t=(x-4)*2; value = Y0 * fast_exp2(t)
Encode: t=fast_log2(value/Y0); if t<0: x=t/4+4, else: x=t/2+4

MDB TABLES REQUIRED FOR LINKS
-----------------------------
- Space_Schedule_Links: (Space_ID, Schedule_ID) - for schedules
- Space_Window_Links: (Space_ID, Window_ID) - for windows AND skylights
- Space_Wall_Links: (Space_ID, Wall_ID) - for walls
- Space_Door_Links: (Space_ID, Door_ID) - for doors
- Space_Roof_Links: (Space_ID, Roof_ID) - for roofs

IMPORTANT NOTES
---------------
1. Para mudar schedules: alterar binário + adicionar link no MDB + [80],[82]=0
2. Para janelas aparecerem: Window ID deve existir no WindowIndex + link no MDB
3. Áreas são sempre em ft² internamente (converter de m² * 10.7639)
4. First record (index 0) in HAP51SPC.DAT is the default template
5. Infiltration ACH: flag=2 antes de cada valor ACH (offsets 554, 560, 566)
6. Floor temps must be written to 3 locations: 450-462, 476-488, 526-538
7. Floor U-value/R-value conversions: SI to IP divide/multiply by 5.678
"""

import struct
import zipfile
import os
import tempfile
import shutil
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from pathlib import Path


# =============================================================================
# CONSTANTS
# =============================================================================

RECORD_SIZE = 682  # bytes per space record

# Direction indices for wall blocks (HAP uses 1-16)
# Order in HAP UI dropdown:
#   N, NNE, NE, ENE, E, ESE, SE, SSE, S, SSW, SW, WSW, W, WNW, NW, NNW
DIRECTION_CODES = {
    'N': 1, 'NNE': 2, 'NE': 3, 'ENE': 4,
    'E': 5, 'ESE': 6, 'SE': 7, 'SSE': 8,
    'S': 9, 'SSW': 10, 'SW': 11, 'WSW': 12,
    'W': 13, 'WNW': 14, 'NW': 15, 'NNW': 16
}
DIRECTION_NAMES = {v: k for k, v in DIRECTION_CODES.items()}

# =============================================================================
# WALL BLOCK STRUCTURE (discovered 2026-01-26)
# =============================================================================
# Each space can have up to 8 wall blocks, starting at offset 72.
# Each block is 34 bytes.
#
# STRUCTURE (34 bytes per wall block):
#   Offset  Size  Type    Field
#   ------  ----  ------  -----
#   [0-2]     2   uint16  Direction code (1-16, see DIRECTION_CODES)
#   [2-6]     4   float   Wall gross area (ft²)
#   [6-8]     2   uint16  Wall type ID (from WallIndex table)
#   [8-10]    2   uint16  Window 1 type ID (from WindowIndex table)
#   [10-12]   2   uint16  Window 2 type ID (from WindowIndex table)
#   [12-14]   2   uint16  Window 1 quantity
#   [14-16]   2   uint16  Window 2 quantity
#   [16-18]   2   uint16  Door type ID (from DoorIndex table)
#   [18-20]   2   uint16  Door quantity
#   [20-34]  14   bytes   Reserved/unknown
#
# IMPORTANTE: Para que as janelas apareçam no HAP:
#   1. O Window ID deve existir na tabela WindowIndex do MDB
#   2. Deve existir um link na tabela Space_Window_Links (Space_ID, Window_ID)
#
# TESTADO E CONFIRMADO (2026-01-26):
#   - Áreas em m² são convertidas para ft² (multiplicar por 10.7639)
#   - Múltiplas janelas por parede funcionam (Win1 e Win2)
#   - Quantidades funcionam correctamente
#   - Wall type ID refere-se a WallIndex no MDB
#
WALL_BLOCK_SIZE = 34
WALL_BLOCK_START = 72
WALL_BLOCK_COUNT = 8  # Maximum 8 walls per space

# Legacy compatibility
DIRECTIONS = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE']  # First 8 directions
DIRECTION_OFFSETS = {d: 72 + (i * 34) for i, d in enumerate(DIRECTIONS)}

# OA Unit codes
OA_UNITS = {1: 'L/s', 2: 'L/s/m²', 3: 'L/s/person', 4: '%'}
OA_UNIT_CODES = {'L/s': 1, 'L/s/m²': 2, 'L/s/m2': 2, 'L/s/person': 3, '%': 4}

# Fixture type codes
FIXTURE_TYPES = {
    0: 'Recessed, unvented',
    1: 'Recessed, vented to return air',
    2: 'Recessed, vented to supply and return',
    3: 'Surface mount / pendant'
}

# Floor type codes
FLOOR_TYPES = {
    1: 'Floor Above Conditioned Space',
    2: 'Floor Above Unconditioned Space',
    3: 'Slab Floor On Grade',
    4: 'Slab Floor Below Grade'
}

# Infiltration mode codes
INFILTRATION_MODES = {1: 'When Fan Off', 2: 'All Hours'}

# =============================================================================
# SCHEDULE FIELD OFFSETS (discovered 2026-01-26)
# =============================================================================
# IMPORTANTE: Para mudar schedules, é necessário:
#   1. Alterar o campo no binário (HAP51SPC.DAT)
#   2. Adicionar link na tabela Space_Schedule_Links do MDB
#   3. Colocar [80] e [82] a 0 (flags de validação?)
#
# Lighting Schedule:
#   - Offset [616-618] = Schedule ID (NÃO é [614]!)
#   - [614] parece ser Task Lighting Schedule ou outro campo
#
# Equipment Schedule:
#   - Offset [660-662] = Schedule ID
#
# People Schedule:
#   - Offset [594-596] = Schedule ID
#
# Activity Level IDs:
#   0: Seated at Rest
#   1: User-defined
#   2: Seated at Rest (duplicate?)
#   3: Office Work
#   4: Sedentary Work
#   5: Medium Work
#   6: Heavy Work
#   7: Dancing
#   8: Athletics


# =============================================================================
# UNIT CONVERSION FUNCTIONS
# =============================================================================

def ft2_to_m2(ft2: float) -> float:
    """Convert square feet to square meters."""
    return ft2 * 0.0929

def m2_to_ft2(m2: float) -> float:
    """Convert square meters to square feet."""
    return m2 / 0.0929

def ft_to_m(ft: float) -> float:
    """Convert feet to meters."""
    return ft * 0.3048

def m_to_ft(m: float) -> float:
    """Convert meters to feet."""
    return m / 0.3048

def f_to_c(f: float) -> float:
    """Convert Fahrenheit to Celsius."""
    return (f - 32) / 1.8

def c_to_f(c: float) -> float:
    """Convert Celsius to Fahrenheit."""
    return c * 1.8 + 32

def lb_ft2_to_kg_m2(lb_ft2: float) -> float:
    """Convert lb/ft² to kg/m²."""
    return lb_ft2 * 4.8824

def kg_m2_to_lb_ft2(kg_m2: float) -> float:
    """Convert kg/m² to lb/ft²."""
    return kg_m2 / 4.8824

def w_ft2_to_w_m2(w_ft2: float) -> float:
    """Convert W/ft² to W/m²."""
    return w_ft2 * 10.764

def w_m2_to_w_ft2(w_m2: float) -> float:
    """Convert W/m² to W/ft²."""
    return w_m2 / 10.764

def btu_hr_to_w(btu_hr: float) -> float:
    """Convert BTU/hr to Watts."""
    return btu_hr / 3.412

def w_to_btu_hr(w: float) -> float:
    """Convert Watts to BTU/hr."""
    return w * 3.412

def cfm_to_ls(cfm: float) -> float:
    """Convert CFM to L/s."""
    return cfm * 0.4719

def ls_to_cfm(ls: float) -> float:
    """Convert L/s to CFM."""
    return ls * 2.11888


# =============================================================================
# OA ENCODING/DECODING
# =============================================================================

import math

# OA encoding - exact closed-form formula (2026-02-05)
# HAP 5.1 uses a piecewise-linear "fast_exp2" approximation:
#   y = Y0 * fast_exp2(k * (x - 4))
#   k=4 for x<4 (base 16), k=2 for x>=4 (base 4)
#   Y0 = 512 CFM in L/s = 241.637...
_OA_Y0 = 512.0 * (28.316846592 / 60.0)

def _fast_exp2(t):
    n = math.floor(t)
    f = t - n
    return (2.0 ** n) * (1.0 + f)

def _fast_log2(v):
    n = math.floor(math.log2(v))
    f = v / (2.0 ** n) - 1.0
    if f < 0:
        n -= 1
        f = v / (2.0 ** n) - 1.0
    if f >= 1.0:
        n += 1
        f = v / (2.0 ** n) - 1.0
    return n + f

def decode_oa_value(internal_float: float, unit_code: int) -> float:
    """Decode internal OA value to user value.
    Uses piecewise fast_exp2 formula (same as excel_to_hap.py)."""
    if internal_float <= 0:
        return 0.0
    if unit_code == 4:  # %
        return internal_float * 28.5714
    try:
        x = internal_float
        if x < 4.0:
            t = (x - 4.0) * 4.0
        else:
            t = (x - 4.0) * 2.0
        return _OA_Y0 * _fast_exp2(t)
    except:
        return 0.0

def encode_oa_value(user_value: float, unit_code: int) -> float:
    """Encode user OA value to internal float.
    Uses piecewise fast_log2 formula (same as excel_to_hap.py)."""
    if user_value <= 0:
        return 0.0
    if unit_code == 4:  # %
        return user_value / 28.5714
    try:
        v = float(user_value) / _OA_Y0
        t = _fast_log2(v)
        if t < 0:
            return t / 4.0 + 4.0
        else:
            return t / 2.0 + 4.0
    except:
        return 0.0


# =============================================================================
# DATA CLASSES
# =============================================================================

@dataclass
class WallBlock:
    """Represents a wall/window/door block for one direction."""
    direction: str = ''
    wall_type_id: int = 0
    wall_area_ft2: float = 0.0
    has_wall: bool = False
    window1_type_id: int = 0
    window1_quantity: int = 0
    window2_type_id: int = 0
    window2_quantity: int = 0
    door_type_id: int = 0
    door_quantity: int = 0
    extra_data: bytes = field(default_factory=lambda: bytes(16))

    @property
    def wall_area_m2(self) -> float:
        return ft2_to_m2(self.wall_area_ft2)

    @wall_area_m2.setter
    def wall_area_m2(self, value: float):
        self.wall_area_ft2 = m2_to_ft2(value)


@dataclass
class Infiltration:
    """Infiltration settings."""
    mode: int = 1  # 1=When Fan Off, 2=All Hours
    design_cooling_cfm: float = 0.0
    design_cooling_cfm_ft2: float = 0.0
    design_cooling_ach: float = 0.0
    design_heating_cfm: float = 0.0
    design_heating_cfm_ft2: float = 0.0
    design_heating_ach: float = 0.0
    energy_cfm: float = 0.0
    energy_cfm_ft2: float = 0.0


@dataclass
class Partition:
    """Partition to adjacent space."""
    area_ft2: float = 0.0
    u_value: float = 0.0
    uncond_max_temp_f: float = 0.0
    ambient_at_max_f: float = 0.0
    uncond_min_temp_f: float = 0.0
    ambient_at_min_f: float = 0.0
    partition_type: int = 0  # 1=Ceiling, 2=Wall


@dataclass
class HAPSpace:
    """Represents a single space in HAP."""
    # Identification
    name: str = ''

    # Dimensions (stored in imperial, properties for metric)
    floor_area_ft2: float = 0.0
    ceiling_height_ft: float = 0.0
    building_weight_lb_ft2: float = 70.0

    # Flags
    type_flag: int = 4

    # Outdoor Air
    oa_value: float = 0.0
    oa_unit_code: int = 1  # 1=L/s, 2=L/s/m², 3=L/s/person, 4=%
    oa_auxiliary: bytes = field(default_factory=lambda: bytes(2))

    # Thermostat
    thermostat_schedule_id: int = 1
    sensible_heat_ratio: float = 0.5
    cooling_setpoint_f: float = 75.0
    cooling_rh: float = 50.0
    heating_setpoint_f: float = 68.0
    heating_rh: float = 50.0

    # Walls (8 directions)
    walls: Dict[str, WallBlock] = field(default_factory=dict)

    # Floor/Roof (raw bytes for now)
    floor_data: bytes = field(default_factory=lambda: bytes(40))
    roof_data: bytes = field(default_factory=lambda: bytes(40))

    # Infiltration
    infiltration: Infiltration = field(default_factory=Infiltration)

    # Partitions
    partition1: Partition = field(default_factory=Partition)
    partition2: Partition = field(default_factory=Partition)

    # People
    occupancy: float = 0.0
    activity_level_id: int = 0
    sensible_heat_btu_hr: float = 0.0
    latent_heat_btu_hr: float = 0.0
    people_schedule_id: int = 1

    # Lighting
    task_lighting_w: float = 0.0
    fixture_type_id: int = 0
    overhead_lighting_w: float = 0.0
    ballast_multiplier: float = 1.0
    lighting_schedule_id: int = 1

    # Miscellaneous
    misc_sensible_btu_hr: float = 0.0
    misc_latent_btu_hr: float = 0.0
    misc_sensible_schedule_id: int = 1
    misc_latent_schedule_id: int = 1

    # Equipment
    equipment_w_ft2: float = 0.0
    equipment_schedule_id: int = 1

    # Raw data for unknown bytes
    _raw_data: bytes = field(default_factory=lambda: bytes(RECORD_SIZE), repr=False)

    def __post_init__(self):
        # Initialize walls dict if empty
        if not self.walls:
            for direction in DIRECTIONS:
                self.walls[direction] = WallBlock(direction=direction)

    # Metric properties
    @property
    def floor_area_m2(self) -> float:
        return ft2_to_m2(self.floor_area_ft2)

    @floor_area_m2.setter
    def floor_area_m2(self, value: float):
        self.floor_area_ft2 = m2_to_ft2(value)

    @property
    def ceiling_height_m(self) -> float:
        return ft_to_m(self.ceiling_height_ft)

    @ceiling_height_m.setter
    def ceiling_height_m(self, value: float):
        self.ceiling_height_ft = m_to_ft(value)

    @property
    def building_weight_kg_m2(self) -> float:
        return lb_ft2_to_kg_m2(self.building_weight_lb_ft2)

    @building_weight_kg_m2.setter
    def building_weight_kg_m2(self, value: float):
        self.building_weight_lb_ft2 = kg_m2_to_lb_ft2(value)

    @property
    def cooling_setpoint_c(self) -> float:
        return f_to_c(self.cooling_setpoint_f)

    @cooling_setpoint_c.setter
    def cooling_setpoint_c(self, value: float):
        self.cooling_setpoint_f = c_to_f(value)

    @property
    def heating_setpoint_c(self) -> float:
        return f_to_c(self.heating_setpoint_f)

    @heating_setpoint_c.setter
    def heating_setpoint_c(self, value: float):
        self.heating_setpoint_f = c_to_f(value)

    @property
    def sensible_heat_w(self) -> float:
        return btu_hr_to_w(self.sensible_heat_btu_hr)

    @sensible_heat_w.setter
    def sensible_heat_w(self, value: float):
        self.sensible_heat_btu_hr = w_to_btu_hr(value)

    @property
    def latent_heat_w(self) -> float:
        return btu_hr_to_w(self.latent_heat_btu_hr)

    @latent_heat_w.setter
    def latent_heat_w(self, value: float):
        self.latent_heat_btu_hr = w_to_btu_hr(value)

    @property
    def equipment_w_m2(self) -> float:
        return w_ft2_to_w_m2(self.equipment_w_ft2)

    @equipment_w_m2.setter
    def equipment_w_m2(self, value: float):
        self.equipment_w_ft2 = w_m2_to_w_ft2(value)

    @property
    def oa_unit(self) -> str:
        return OA_UNITS.get(self.oa_unit_code, 'Unknown')

    @oa_unit.setter
    def oa_unit(self, value: str):
        self.oa_unit_code = OA_UNIT_CODES.get(value, 1)


# =============================================================================
# BINARY PARSING
# =============================================================================

def parse_wall_block(data: bytes, index: int = 0) -> WallBlock:
    """Parse a 34-byte wall block.

    Structure (34 bytes):
      [0-2]   direction code (1-16, see DIRECTION_CODES)
      [2-6]   wall area (ft², float)
      [6-8]   wall type ID
      [8-10]  window1 type ID
      [10-12] window2 type ID
      [12-14] window1 quantity
      [14-16] window2 quantity
      [16-18] door type ID
      [18-20] door quantity
      [20-34] reserved/extra
    """
    block = WallBlock()

    direction_code = struct.unpack('<H', data[0:2])[0]
    block.direction = DIRECTION_NAMES.get(direction_code, f'?{direction_code}')
    block.wall_area_ft2 = struct.unpack('<f', data[2:6])[0]
    block.wall_type_id = struct.unpack('<H', data[6:8])[0]
    block.window1_type_id = struct.unpack('<H', data[8:10])[0]
    block.window2_type_id = struct.unpack('<H', data[10:12])[0]
    block.window1_quantity = struct.unpack('<H', data[12:14])[0]
    block.window2_quantity = struct.unpack('<H', data[14:16])[0]
    block.door_type_id = struct.unpack('<H', data[16:18])[0]
    block.door_quantity = struct.unpack('<H', data[18:20])[0]
    block.extra_data = data[20:34]

    # has_wall flag based on area > 0
    block.has_wall = block.wall_area_ft2 > 0

    return block

def encode_wall_block(block: WallBlock) -> bytes:
    """Encode a WallBlock to 34 bytes."""
    data = bytearray(34)

    direction_code = DIRECTION_CODES.get(block.direction, 0)
    struct.pack_into('<H', data, 0, direction_code)
    struct.pack_into('<f', data, 2, block.wall_area_ft2)
    struct.pack_into('<H', data, 6, block.wall_type_id)
    struct.pack_into('<H', data, 8, block.window1_type_id)
    struct.pack_into('<H', data, 10, block.window2_type_id)
    struct.pack_into('<H', data, 12, block.window1_quantity)
    struct.pack_into('<H', data, 14, block.window2_quantity)
    struct.pack_into('<H', data, 16, block.door_type_id)
    struct.pack_into('<H', data, 18, block.door_quantity)

    # Extra data (14 bytes)
    if block.extra_data and len(block.extra_data) >= 14:
        data[20:34] = block.extra_data[:14]

    return bytes(data)

def parse_space(data: bytes) -> HAPSpace:
    """Parse a 682-byte space record."""
    space = HAPSpace()
    space._raw_data = data

    # Name (0-23)
    space.name = data[0:24].decode('latin-1').rstrip('\x00')

    # Dimensions (24-35)
    space.floor_area_ft2 = struct.unpack('<f', data[24:28])[0]
    space.ceiling_height_ft = struct.unpack('<f', data[28:32])[0]
    space.building_weight_lb_ft2 = struct.unpack('<f', data[32:36])[0]

    # Flags (36-39)
    space.type_flag = struct.unpack('<I', data[36:40])[0]

    # OA (44-51)
    space.oa_auxiliary = data[44:46]
    oa_internal = struct.unpack('<f', data[46:50])[0]
    space.oa_unit_code = struct.unpack('<H', data[50:52])[0]
    space.oa_value = decode_oa_value(oa_internal, space.oa_unit_code)

    # Walls (72-343, 8 blocks of 34 bytes each)
    for i in range(8):
        offset = WALL_BLOCK_START + (i * WALL_BLOCK_SIZE)
        wall = parse_wall_block(data[offset:offset+WALL_BLOCK_SIZE], i)
        if wall.direction:
            space.walls[wall.direction] = wall

    # Floor/Roof (344-439)
    # After walls: 72 + (8 * 34) = 344
    space.floor_data = data[344:392]  # 48 bytes for floor
    space.roof_data = data[392:440]   # 48 bytes for roof

    # Partitions (440-491) - two partitions of 26 bytes each
    # Partition 1 (440-465): Type(H), Area(f), U-Value(f), UncMax(f), OutMax(f), UncMin(f), OutMin(f)
    space.partition1.partition_type = struct.unpack('<H', data[440:442])[0]
    space.partition1.area_ft2 = struct.unpack('<f', data[442:446])[0]
    space.partition1.u_value = struct.unpack('<f', data[446:450])[0]
    space.partition1.uncond_max_temp_f = struct.unpack('<f', data[450:454])[0]
    space.partition1.ambient_at_max_f = struct.unpack('<f', data[454:458])[0]
    space.partition1.uncond_min_temp_f = struct.unpack('<f', data[458:462])[0]
    space.partition1.ambient_at_min_f = struct.unpack('<f', data[462:466])[0]
    # Partition 2 (466-491): Type(H), Area(f), U-Value(f), UncMax(f), OutMax(f), UncMin(f), OutMin(f)
    space.partition2.partition_type = struct.unpack('<H', data[466:468])[0]
    space.partition2.area_ft2 = struct.unpack('<f', data[468:472])[0]
    space.partition2.u_value = struct.unpack('<f', data[472:476])[0]
    space.partition2.uncond_max_temp_f = struct.unpack('<f', data[476:480])[0]
    space.partition2.ambient_at_max_f = struct.unpack('<f', data[480:484])[0]
    space.partition2.uncond_min_temp_f = struct.unpack('<f', data[484:488])[0]
    space.partition2.ambient_at_min_f = struct.unpack('<f', data[488:492])[0]

    # Infiltration (554-572) - ACH values
    # Offsets 554, 560, 566 are flags (2=ACH mode), followed by float ACH values
    space.infiltration.design_cooling_ach = struct.unpack('<f', data[556:560])[0]
    space.infiltration.design_heating_ach = struct.unpack('<f', data[562:566])[0]
    # Note: energy ACH is at 568-572 but not stored in Infiltration dataclass

    # People (580-599)
    space.occupancy = struct.unpack('<f', data[580:584])[0]
    space.activity_level_id = struct.unpack('<H', data[584:586])[0]
    space.sensible_heat_btu_hr = struct.unpack('<f', data[586:590])[0]
    space.latent_heat_btu_hr = struct.unpack('<f', data[590:594])[0]
    space.people_schedule_id = struct.unpack('<H', data[594:596])[0]

    # Lighting (600-623)
    space.task_lighting_w = struct.unpack('<f', data[600:604])[0]
    space.fixture_type_id = struct.unpack('<H', data[604:606])[0]
    space.overhead_lighting_w = struct.unpack('<f', data[606:610])[0]
    space.ballast_multiplier = struct.unpack('<f', data[610:614])[0]
    space.lighting_schedule_id = struct.unpack('<H', data[616:618])[0]

    # Misc (632-647)
    space.misc_sensible_btu_hr = struct.unpack('<f', data[632:636])[0]
    space.misc_latent_btu_hr = struct.unpack('<f', data[636:640])[0]
    space.misc_sensible_schedule_id = struct.unpack('<H', data[640:642])[0]
    space.misc_latent_schedule_id = struct.unpack('<H', data[644:646])[0]

    # Equipment (656-661)
    space.equipment_w_ft2 = struct.unpack('<f', data[656:660])[0]
    space.equipment_schedule_id = struct.unpack('<H', data[660:662])[0]

    return space

def encode_space(space: HAPSpace) -> bytes:
    """Encode a HAPSpace to 682 bytes."""
    # Start with raw data (preserves unknown bytes)
    data = bytearray(space._raw_data) if space._raw_data else bytearray(RECORD_SIZE)

    # Name (0-23)
    name_bytes = space.name.encode('latin-1')[:24].ljust(24, b'\x00')
    data[0:24] = name_bytes

    # Dimensions (24-35)
    struct.pack_into('<f', data, 24, space.floor_area_ft2)
    struct.pack_into('<f', data, 28, space.ceiling_height_ft)
    struct.pack_into('<f', data, 32, space.building_weight_lb_ft2)

    # Flags (36-39)
    struct.pack_into('<I', data, 36, space.type_flag)

    # OA (44-51)
    data[44:46] = space.oa_auxiliary
    oa_internal = encode_oa_value(space.oa_value, space.oa_unit_code)
    struct.pack_into('<f', data, 46, oa_internal)
    struct.pack_into('<H', data, 50, space.oa_unit_code)

    # Walls (72-343, 8 blocks of 34 bytes each)
    wall_index = 0
    for direction in DIRECTIONS:
        if direction in space.walls:
            offset = WALL_BLOCK_START + (wall_index * WALL_BLOCK_SIZE)
            data[offset:offset+WALL_BLOCK_SIZE] = encode_wall_block(space.walls[direction])
            wall_index += 1
            if wall_index >= 8:
                break

    # Floor/Roof (344-439)
    data[344:392] = space.floor_data[:48] if len(space.floor_data) >= 48 else space.floor_data.ljust(48, b'\x00')
    data[392:440] = space.roof_data[:48] if len(space.roof_data) >= 48 else space.roof_data.ljust(48, b'\x00')

    # Partitions (440-491) - two partitions of 26 bytes each
    # Partition 1 (440-465)
    struct.pack_into('<H', data, 440, space.partition1.partition_type)
    struct.pack_into('<f', data, 442, space.partition1.area_ft2)
    struct.pack_into('<f', data, 446, space.partition1.u_value)
    struct.pack_into('<f', data, 450, space.partition1.uncond_max_temp_f)
    struct.pack_into('<f', data, 454, space.partition1.ambient_at_max_f)
    struct.pack_into('<f', data, 458, space.partition1.uncond_min_temp_f)
    struct.pack_into('<f', data, 462, space.partition1.ambient_at_min_f)
    # Partition 2 (466-491)
    struct.pack_into('<H', data, 466, space.partition2.partition_type)
    struct.pack_into('<f', data, 468, space.partition2.area_ft2)
    struct.pack_into('<f', data, 472, space.partition2.u_value)
    struct.pack_into('<f', data, 476, space.partition2.uncond_max_temp_f)
    struct.pack_into('<f', data, 480, space.partition2.ambient_at_max_f)
    struct.pack_into('<f', data, 484, space.partition2.uncond_min_temp_f)
    struct.pack_into('<f', data, 488, space.partition2.ambient_at_min_f)

    # Infiltration (554-572) - ACH values
    # Offsets 554, 560, 566 are flags (2=ACH mode), followed by float ACH values
    ACH_MODE = 2
    struct.pack_into('<H', data, 554, ACH_MODE)
    struct.pack_into('<f', data, 556, space.infiltration.design_cooling_ach)
    struct.pack_into('<H', data, 560, ACH_MODE)
    struct.pack_into('<f', data, 562, space.infiltration.design_heating_ach)

    # People (580-599)
    struct.pack_into('<f', data, 580, space.occupancy)
    struct.pack_into('<H', data, 584, space.activity_level_id)
    struct.pack_into('<f', data, 586, space.sensible_heat_btu_hr)
    struct.pack_into('<f', data, 590, space.latent_heat_btu_hr)
    struct.pack_into('<H', data, 594, space.people_schedule_id)

    # Lighting (600-623)
    struct.pack_into('<f', data, 600, space.task_lighting_w)
    struct.pack_into('<H', data, 604, space.fixture_type_id)
    struct.pack_into('<f', data, 606, space.overhead_lighting_w)
    struct.pack_into('<f', data, 610, space.ballast_multiplier)
    struct.pack_into('<H', data, 616, space.lighting_schedule_id)

    # Misc (632-647)
    struct.pack_into('<f', data, 632, space.misc_sensible_btu_hr)
    struct.pack_into('<f', data, 636, space.misc_latent_btu_hr)
    struct.pack_into('<H', data, 640, space.misc_sensible_schedule_id)
    struct.pack_into('<H', data, 644, space.misc_latent_schedule_id)

    # Equipment (656-661)
    struct.pack_into('<f', data, 656, space.equipment_w_ft2)
    struct.pack_into('<H', data, 660, space.equipment_schedule_id)

    return bytes(data)


# =============================================================================
# HAP PROJECT CLASS
# =============================================================================

class HAPProject:
    """Represents a HAP 5.1 project (.E3A file)."""

    def __init__(self):
        self.filepath: Optional[Path] = None
        self.default_space: Optional[HAPSpace] = None
        self.spaces: List[HAPSpace] = []
        self._archive_files: Dict[str, bytes] = {}

    @classmethod
    def open(cls, filepath: str) -> 'HAPProject':
        """Open an existing .E3A file."""
        project = cls()
        project.filepath = Path(filepath)

        with zipfile.ZipFile(filepath, 'r') as zf:
            # Store all files
            for name in zf.namelist():
                project._archive_files[name] = zf.read(name)

            # Parse spaces
            if 'HAP51SPC.DAT' in project._archive_files:
                spc_data = project._archive_files['HAP51SPC.DAT']
                num_records = len(spc_data) // RECORD_SIZE

                # First record is default template
                if num_records > 0:
                    project.default_space = parse_space(spc_data[0:RECORD_SIZE])

                # Remaining records are user spaces
                for i in range(1, num_records):
                    offset = i * RECORD_SIZE
                    space = parse_space(spc_data[offset:offset+RECORD_SIZE])
                    project.spaces.append(space)

        return project

    def save(self, filepath: Optional[str] = None):
        """Save the project to an .E3A file."""
        if filepath:
            self.filepath = Path(filepath)

        if not self.filepath:
            raise ValueError("No filepath specified")

        # Rebuild HAP51SPC.DAT
        spc_data = bytearray()

        # Default space template
        if self.default_space:
            spc_data.extend(encode_space(self.default_space))
        else:
            spc_data.extend(bytes(RECORD_SIZE))

        # User spaces
        for space in self.spaces:
            spc_data.extend(encode_space(space))

        self._archive_files['HAP51SPC.DAT'] = bytes(spc_data)

        # Write to ZIP
        with zipfile.ZipFile(self.filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, data in self._archive_files.items():
                zf.writestr(name, data)

    def add_space(self, space: HAPSpace) -> None:
        """Add a new space to the project."""
        self.spaces.append(space)

    def remove_space(self, index: int) -> HAPSpace:
        """Remove and return a space by index."""
        return self.spaces.pop(index)

    def get_space_by_name(self, name: str) -> Optional[HAPSpace]:
        """Find a space by name."""
        for space in self.spaces:
            if space.name == name:
                return space
        return None

    def list_spaces(self) -> List[str]:
        """Return list of space names."""
        return [s.name for s in self.spaces]

    def print_summary(self):
        """Print project summary."""
        print(f"HAP Project: {self.filepath}")
        print(f"Number of spaces: {len(self.spaces)}")
        print("\nSpaces:")
        for i, space in enumerate(self.spaces):
            print(f"  {i+1}. {space.name}")
            print(f"      Area: {space.floor_area_m2:.1f} m² ({space.floor_area_ft2:.1f} ft²)")
            print(f"      Height: {space.ceiling_height_m:.2f} m ({space.ceiling_height_ft:.2f} ft)")
            print(f"      Occupancy: {space.occupancy:.0f} people")
            print(f"      Lighting: {space.overhead_lighting_w:.0f} W")
            print(f"      Equipment: {space.equipment_w_m2:.1f} W/m²")


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def create_space_from_dict(data: dict) -> HAPSpace:
    """Create a HAPSpace from a dictionary (e.g., from Excel/JSON)."""
    space = HAPSpace()

    # Required fields
    space.name = str(data.get('name', ''))[:24]

    # Dimensions - accept metric or imperial
    if 'area_m2' in data:
        space.floor_area_m2 = float(data['area_m2'])
    elif 'area_ft2' in data:
        space.floor_area_ft2 = float(data['area_ft2'])

    if 'height_m' in data:
        space.ceiling_height_m = float(data['height_m'])
    elif 'height_ft' in data:
        space.ceiling_height_ft = float(data['height_ft'])

    if 'weight_kg_m2' in data:
        space.building_weight_kg_m2 = float(data['weight_kg_m2'])
    elif 'weight_lb_ft2' in data:
        space.building_weight_lb_ft2 = float(data['weight_lb_ft2'])

    # OA
    if 'oa_value' in data:
        space.oa_value = float(data['oa_value'])
    if 'oa_unit' in data:
        space.oa_unit = str(data['oa_unit'])

    # Thermostat - accept C or F
    if 'cooling_setpoint_c' in data:
        space.cooling_setpoint_c = float(data['cooling_setpoint_c'])
    elif 'cooling_setpoint_f' in data:
        space.cooling_setpoint_f = float(data['cooling_setpoint_f'])

    if 'heating_setpoint_c' in data:
        space.heating_setpoint_c = float(data['heating_setpoint_c'])
    elif 'heating_setpoint_f' in data:
        space.heating_setpoint_f = float(data['heating_setpoint_f'])

    if 'cooling_rh' in data:
        space.cooling_rh = float(data['cooling_rh'])
    if 'heating_rh' in data:
        space.heating_rh = float(data['heating_rh'])

    # Occupancy
    if 'occupancy' in data:
        space.occupancy = float(data['occupancy'])
    if 'sensible_w' in data:
        space.sensible_heat_w = float(data['sensible_w'])
    if 'latent_w' in data:
        space.latent_heat_w = float(data['latent_w'])

    # Lighting
    if 'lighting_w' in data:
        space.overhead_lighting_w = float(data['lighting_w'])
    if 'ballast_multiplier' in data:
        space.ballast_multiplier = float(data['ballast_multiplier'])

    # Equipment - accept W/m² or W/ft²
    if 'equipment_w_m2' in data:
        space.equipment_w_m2 = float(data['equipment_w_m2'])
    elif 'equipment_w_ft2' in data:
        space.equipment_w_ft2 = float(data['equipment_w_ft2'])

    # Walls by direction
    for direction in DIRECTIONS:
        dir_lower = direction.lower()
        prefix = f'{dir_lower}_'

        if f'{prefix}wall_type' in data:
            space.walls[direction].wall_type_id = int(data[f'{prefix}wall_type'])
        if f'{prefix}wall_area_m2' in data:
            space.walls[direction].wall_area_m2 = float(data[f'{prefix}wall_area_m2'])
        elif f'{prefix}wall_area_ft2' in data:
            space.walls[direction].wall_area_ft2 = float(data[f'{prefix}wall_area_ft2'])
        if f'{prefix}window_type' in data:
            space.walls[direction].window1_type_id = int(data[f'{prefix}window_type'])
        if f'{prefix}window_qty' in data:
            space.walls[direction].window1_quantity = int(data[f'{prefix}window_qty'])

        # Set has_wall flag if wall type or area is set
        if space.walls[direction].wall_type_id > 0 or space.walls[direction].wall_area_ft2 > 0:
            space.walls[direction].has_wall = True

    return space


# =============================================================================
# MAIN (DEMO)
# =============================================================================

if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("HAP 5.1 File Library")
        print("====================")
        print("\nUsage:")
        print(f"  python {sys.argv[0]} <file.E3A>  - Read and display project info")
        print(f"  python {sys.argv[0]} --demo      - Create demo project")
        sys.exit(0)

    if sys.argv[1] == '--demo':
        # Create demo project
        print("Creating demo project...")

        # Open existing project as template
        # project = HAPProject()
        # ... would need template files

        print("Demo requires a template .E3A file to copy structure from.")
        print("Use: project = HAPProject.open('template.E3A')")

    else:
        # Read existing project
        filepath = sys.argv[1]
        print(f"Opening: {filepath}")

        project = HAPProject.open(filepath)
        project.print_summary()

        print("\n\nDetailed space info:")
        for space in project.spaces[:3]:  # First 3 spaces
            print(f"\n{'-'*60}")
            print(f"Space: {space.name}")
            print(f"  Area: {space.floor_area_m2:.1f} m²")
            print(f"  Height: {space.ceiling_height_m:.2f} m")
            print(f"  Weight: {space.building_weight_kg_m2:.0f} kg/m²")
            print(f"  OA: {space.oa_value:.1f} {space.oa_unit}")
            print(f"  Cooling: {space.cooling_setpoint_c:.1f}°C, {space.cooling_rh:.0f}% RH")
            print(f"  Heating: {space.heating_setpoint_c:.1f}°C, {space.heating_rh:.0f}% RH")
            print(f"  Occupancy: {space.occupancy:.0f} people")
            print(f"  Lighting: {space.overhead_lighting_w:.0f} W (fixture type {space.fixture_type_id})")
            print(f"  Equipment: {space.equipment_w_m2:.1f} W/m²")

            # Walls
            has_walls = any(w.has_wall for w in space.walls.values())
            if has_walls:
                print("  Walls:")
                for direction, wall in space.walls.items():
                    if wall.has_wall:
                        print(f"    {direction}: type={wall.wall_type_id}, area={wall.wall_area_m2:.1f}m²", end='')
                        if wall.window1_type_id > 0:
                            print(f", window type={wall.window1_type_id}", end='')
                        print()
