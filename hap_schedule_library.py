"""
HAP 5.1 Schedule Library
========================
Read and write Carrier HAP 5.1 schedule records from HAP51SCH.DAT.

Author: Generated from reverse engineering
Version: 2.0
Date: 2026-01-26

SCHEDULE RECORD STRUCTURE (792 bytes)
=====================================
Offset      Size    Field
------      ----    -----
[0-24]        24    Schedule name (null-padded string, latin-1)
[24-32]        8    Flags/padding (usually 0x20 0x20 0x20 0x20 0x20 0x20 0x00 0x00)
[32-192]     160    8 profile names (20 bytes each, null-padded)
[192-208]     16    8 uint16 values (min values for each profile)
[208-592]    384    Hourly values: 8 profiles × 24 hours × 2 bytes
[592-792]    200    Day mapping: 100 uint16 values

DAY MAPPING STRUCTURE (100 × uint16 = 200 bytes)
================================================
ESTRUTURA CONFIRMADA POR TESTES:
- 12 meses × 8 tipos de dia = 96 valores (índices 0-95)
- 4 valores Design extra (índices 96-99)

Fórmula: índice = (mês × 8) + dia

Onde:
  mês = 0 (Janeiro) a 11 (Dezembro)
  dia = 0 (Monday), 1 (Tuesday), 2 (Wednesday), 3 (Thursday),
        4 (Friday), 5 (Saturday), 6 (Sunday), 7 (Holiday)

Índices 96-99 = Design Day (usado para cálculos de pico)

Cada valor uint16 (1-8) indica qual Profile usar para esse dia/mês.
Profile numbering é 1-based (Profile 1 = valor 1, não 0).

EXEMPLO DE MAPEAMENTO PARA ESCRITÓRIO:
======================================
Para todos os 12 meses:
  Monday (dia 0)    → Profile 1 (Dias de Semana)
  Tuesday (dia 1)   → Profile 1 (Dias de Semana)
  Wednesday (dia 2) → Profile 1 (Dias de Semana)
  Thursday (dia 3)  → Profile 1 (Dias de Semana)
  Friday (dia 4)    → Profile 1 (Dias de Semana)
  Saturday (dia 5)  → Profile 2 (Fins de Semana)
  Sunday (dia 6)    → Profile 2 (Fins de Semana)
  Holiday (dia 7)   → Profile 3 (Feriados)
  Design (96-99)    → Profile 1 (Dias de Semana)

PROFILE HOURLY VALUES (384 bytes)
=================================
8 profiles, cada um com 24 valores horários (uint16).
Valores típicos: 0-100 (percentagem)
Valores especiais:
  - 65535 (0xFFFF) = ON para schedules de ventiladores (Fan/Thermostat)
  - 0 = OFF

SCHEDULE TYPES (MDB ScheduleIndex.nScheduleType)
================================================
  0 = Fractional (percentagens 0-100%)
  1 = Fan/Thermostat (On/Off: 0 ou 65535)

MDB DATABASE
============
A tabela ScheduleIndex no HAP51INX.MDB TAMBÉM precisa ser actualizada:
  - nIndex: índice do schedule (1-based)
  - szName: nome do schedule
  - nScheduleType: tipo (0=fractional, 1=fan)

IMPORTANTE: O HAP lê os nomes dos schedules do MDB, não do DAT!

UTILIZAÇÃO TÍPICA EM PORTUGAL
=============================
Normalmente definimos 3 profiles:
  - Profile 1: Dias de Semana (Segunda a Sexta)
  - Profile 2: Fins de Semana (Sábado e Domingo)
  - Profile 3: Feriados

E no Assignments atribuímos:
  - Mon-Fri → Profile 1
  - Sat-Sun → Profile 2
  - Holiday → Profile 3
"""

import struct
import zipfile
import os
from dataclasses import dataclass, field
from typing import List, Optional, Dict
from pathlib import Path


# =============================================================================
# CONSTANTS
# =============================================================================

SCHEDULE_RECORD_SIZE = 792

# Schedule types from ScheduleIndex.nScheduleType
SCHEDULE_TYPE_FRACTIONAL = 0  # Percentages (0-100)
SCHEDULE_TYPE_ONOFF = 1       # On/Off (0 or 65535)

# Day indices for day mapping
DAY_MONDAY = 0
DAY_TUESDAY = 1
DAY_WEDNESDAY = 2
DAY_THURSDAY = 3
DAY_FRIDAY = 4
DAY_SATURDAY = 5
DAY_SUNDAY = 6
DAY_HOLIDAY = 7

# Day names
DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday', 'Holiday']
DAY_NAMES_PT = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo', 'Feriado']


# =============================================================================
# DATA CLASSES
# =============================================================================

@dataclass
class ScheduleProfile:
    """Represents a single profile within a schedule."""
    name: str = ''
    hourly_values: List[int] = field(default_factory=lambda: [100] * 24)

    def get_value(self, hour: int) -> int:
        """Get percentage value for a specific hour (0-23)."""
        if 0 <= hour < 24:
            return self.hourly_values[hour]
        return 0

    def set_value(self, hour: int, value: int):
        """Set percentage value for a specific hour (0-23)."""
        if 0 <= hour < 24:
            self.hourly_values[hour] = max(0, min(65535, value))


@dataclass
class HAPSchedule:
    """Represents a schedule in HAP."""
    name: str = ''
    flags: bytes = field(default_factory=lambda: bytes(8))
    profiles: List[ScheduleProfile] = field(default_factory=list)
    unknown_values: List[int] = field(default_factory=lambda: [100] * 8)
    day_mapping: List[int] = field(default_factory=lambda: [1] * 100)
    schedule_type: int = SCHEDULE_TYPE_FRACTIONAL
    _raw_data: bytes = field(default_factory=lambda: bytes(SCHEDULE_RECORD_SIZE), repr=False)

    def __post_init__(self):
        # Initialize 8 profiles if empty
        if not self.profiles:
            for i in range(8):
                self.profiles.append(ScheduleProfile(
                    name=f'{i+1}:Profile {["One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight"][i]}'
                ))

    def get_profile(self, index: int) -> Optional[ScheduleProfile]:
        """Get profile by 0-based index."""
        if 0 <= index < len(self.profiles):
            return self.profiles[index]
        return None

    def _get_day_index(self, month: int, day: int) -> int:
        """Calculate day mapping index.

        Args:
            month: 0-11 (Jan-Dec)
            day: 0-7 (Mon, Tue, Wed, Thu, Fri, Sat, Sun, Holiday)

        Returns:
            Index into day_mapping array (0-95)
        """
        return month * 8 + day

    def set_day_profile(self, day: int, profile_num: int, month: int = -1):
        """Set profile for a specific day type.

        Args:
            day: 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun, 7=Holiday
            profile_num: Profile number (1-8)
            month: Month (0-11) or -1 for all months
        """
        if month == -1:
            # Set for all months
            for m in range(12):
                idx = self._get_day_index(m, day)
                if idx < 96:
                    self.day_mapping[idx] = profile_num
        else:
            idx = self._get_day_index(month, day)
            if idx < 96:
                self.day_mapping[idx] = profile_num

    def set_weekday_profile(self, profile_num: int):
        """Set profile for all weekdays (Mon-Fri) in all months."""
        for day in range(DAY_MONDAY, DAY_FRIDAY + 1):  # 0-4
            self.set_day_profile(day, profile_num)

    def set_weekend_profile(self, profile_num: int):
        """Set profile for all weekends (Sat-Sun) in all months."""
        self.set_day_profile(DAY_SATURDAY, profile_num)  # 5
        self.set_day_profile(DAY_SUNDAY, profile_num)    # 6

    def set_holiday_profile(self, profile_num: int):
        """Set profile for all holidays in all months."""
        self.set_day_profile(DAY_HOLIDAY, profile_num)  # 7
        # Also set the Design slots (96-99)
        for i in range(96, 100):
            self.day_mapping[i] = profile_num

    def set_design_profile(self, profile_num: int):
        """Set profile for Design days (indices 96-99)."""
        for i in range(96, 100):
            self.day_mapping[i] = profile_num

    def set_profile_hourly(self, profile_index: int, hourly_values: List[int]):
        """Set hourly values for a profile (0-based index)."""
        if 0 <= profile_index < 8 and len(hourly_values) == 24:
            self.profiles[profile_index].hourly_values = hourly_values.copy()

    def get_assignment(self, month: int, day: int) -> int:
        """Get profile assignment for a specific day/month.

        Args:
            month: 0-11 (Jan-Dec)
            day: 0-7 (Mon-Holiday)

        Returns:
            Profile number (1-8)
        """
        idx = self._get_day_index(month, day)
        if idx < len(self.day_mapping):
            return self.day_mapping[idx]
        return 1

    def print_assignments(self):
        """Print day mapping in a table format like HAP shows."""
        months = ['J', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D']

        print("       " + " ".join(months))
        for day_idx, day_name in enumerate(DAY_NAMES):
            values = [str(self.get_assignment(m, day_idx)) for m in range(12)]
            print(f"{day_name[:7]:7} " + " ".join(values))

        print(f"Design  {self.day_mapping[96]} {self.day_mapping[97]} {self.day_mapping[98]} {self.day_mapping[99]}")


# =============================================================================
# PARSING FUNCTIONS
# =============================================================================

def parse_schedule(data: bytes) -> HAPSchedule:
    """Parse a 792-byte schedule record."""
    schedule = HAPSchedule()
    schedule._raw_data = data

    # Name (0-24)
    schedule.name = data[0:24].decode('latin-1').rstrip('\x00')

    # Flags (24-32)
    schedule.flags = data[24:32]

    # Profile names (32-192)
    schedule.profiles = []
    for i in range(8):
        offset = 32 + (i * 20)
        name = data[offset:offset+20].decode('latin-1').rstrip('\x00')
        schedule.profiles.append(ScheduleProfile(name=name))

    # Unknown/min values (192-208)
    schedule.unknown_values = []
    for i in range(8):
        val = struct.unpack('<H', data[192+i*2:194+i*2])[0]
        schedule.unknown_values.append(val)

    # Hourly values for each profile (208-592)
    for p_idx in range(8):
        profile_offset = 208 + (p_idx * 48)  # 24 hours * 2 bytes
        hourly = []
        for h in range(24):
            val = struct.unpack('<H', data[profile_offset+h*2:profile_offset+h*2+2])[0]
            hourly.append(val)
        schedule.profiles[p_idx].hourly_values = hourly

    # Day mapping (592-792)
    schedule.day_mapping = []
    for i in range(100):
        val = struct.unpack('<H', data[592+i*2:594+i*2])[0]
        schedule.day_mapping.append(val)

    return schedule


def encode_schedule(schedule: HAPSchedule) -> bytes:
    """Encode a HAPSchedule to 792 bytes."""
    # Start with zeros
    data = bytearray(SCHEDULE_RECORD_SIZE)

    # Name (0-24)
    name_bytes = schedule.name.encode('latin-1')[:24].ljust(24, b'\x00')
    data[0:24] = name_bytes

    # Flags (24-32)
    if schedule.flags and len(schedule.flags) >= 8:
        data[24:32] = schedule.flags[:8]
    else:
        data[24:32] = b'\x20\x20\x20\x20\x20\x20\x00\x00'  # Default padding

    # Profile names (32-192)
    for i in range(8):
        offset = 32 + (i * 20)
        if i < len(schedule.profiles):
            name_bytes = schedule.profiles[i].name.encode('latin-1')[:20].ljust(20, b'\x00')
        else:
            name_bytes = b'\x00' * 20
        data[offset:offset+20] = name_bytes

    # Unknown/min values (192-208)
    for i in range(8):
        val = schedule.unknown_values[i] if i < len(schedule.unknown_values) else 100
        struct.pack_into('<H', data, 192 + i*2, val)

    # Hourly values (208-592)
    for p_idx in range(8):
        profile_offset = 208 + (p_idx * 48)
        if p_idx < len(schedule.profiles):
            for h in range(24):
                val = schedule.profiles[p_idx].hourly_values[h] if h < len(schedule.profiles[p_idx].hourly_values) else 100
                struct.pack_into('<H', data, profile_offset + h*2, val)
        else:
            for h in range(24):
                struct.pack_into('<H', data, profile_offset + h*2, 100)

    # Day mapping (592-792)
    for i in range(100):
        val = schedule.day_mapping[i] if i < len(schedule.day_mapping) else 1
        struct.pack_into('<H', data, 592 + i*2, val)

    return bytes(data)


# =============================================================================
# SCHEDULE MANAGER CLASS
# =============================================================================

class ScheduleManager:
    """Manages schedules in a HAP project."""

    def __init__(self):
        self.schedules: List[HAPSchedule] = []
        self._raw_data: bytes = b''

    @classmethod
    def from_dat_file(cls, data: bytes) -> 'ScheduleManager':
        """Load schedules from HAP51SCH.DAT content."""
        manager = cls()
        manager._raw_data = data

        num_records = len(data) // SCHEDULE_RECORD_SIZE
        for i in range(num_records):
            offset = i * SCHEDULE_RECORD_SIZE
            schedule = parse_schedule(data[offset:offset+SCHEDULE_RECORD_SIZE])
            manager.schedules.append(schedule)

        return manager

    @classmethod
    def from_e3a_file(cls, filepath: str) -> 'ScheduleManager':
        """Load schedules from a .E3A file."""
        with zipfile.ZipFile(filepath, 'r') as zf:
            sch_data = zf.read('HAP51SCH.DAT')
        return cls.from_dat_file(sch_data)

    def to_dat_file(self) -> bytes:
        """Export schedules to HAP51SCH.DAT format."""
        data = bytearray()
        for schedule in self.schedules:
            data.extend(encode_schedule(schedule))
        return bytes(data)

    def get_schedule_by_name(self, name: str) -> Optional[HAPSchedule]:
        """Find a schedule by name."""
        for schedule in self.schedules:
            if schedule.name.strip() == name.strip():
                return schedule
        return None

    def get_schedule_by_index(self, index: int) -> Optional[HAPSchedule]:
        """Get schedule by 0-based index."""
        if 0 <= index < len(self.schedules):
            return self.schedules[index]
        return None

    def add_schedule(self, schedule: HAPSchedule) -> int:
        """Add a new schedule and return its index."""
        self.schedules.append(schedule)
        return len(self.schedules) - 1

    def list_schedules(self) -> List[str]:
        """Return list of schedule names."""
        return [s.name.strip() for s in self.schedules]

    def print_summary(self):
        """Print summary of all schedules."""
        print(f"Total schedules: {len(self.schedules)}")
        print()
        for i, sch in enumerate(self.schedules):
            active_profiles = [p.name.strip() for p in sch.profiles if p.name.strip()]
            used_profiles = sorted(set(sch.day_mapping))
            print(f"{i}. {sch.name.strip()}")
            print(f"   Profiles: {', '.join(active_profiles[:3])}")
            print(f"   Profiles usados: {used_profiles}")

            # Show first profile hourly values summary
            if sch.profiles:
                p1 = sch.profiles[0]
                min_val = min(p1.hourly_values)
                max_val = max(p1.hourly_values)
                print(f"   Profile 1 range: {min_val}% - {max_val}%")
            print()


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def create_simple_schedule(
    name: str,
    weekday_values: List[int],
    weekend_values: Optional[List[int]] = None,
    holiday_values: Optional[List[int]] = None,
    schedule_type: int = SCHEDULE_TYPE_FRACTIONAL
) -> HAPSchedule:
    """Create a simple schedule with weekday/weekend/holiday profiles.

    Args:
        name: Schedule name (max 24 chars)
        weekday_values: List of 24 hourly percentage values for weekdays (Mon-Fri)
        weekend_values: List of 24 hourly percentage values for weekends (Sat-Sun)
        holiday_values: List of 24 hourly percentage values for holidays (optional)
        schedule_type: 0 for fractional, 1 for on/off

    Returns:
        HAPSchedule object with proper day mapping
    """
    schedule = HAPSchedule()
    schedule.name = name[:24]
    schedule.schedule_type = schedule_type
    schedule.flags = b'\x20\x20\x20\x20\x20\x20\x00\x00'

    # Profile 1: Dias de Semana (Mon-Fri)
    schedule.profiles[0].name = '1:Dias de Semana'
    wk_vals = weekday_values[:24] if len(weekday_values) >= 24 else weekday_values + [0] * (24 - len(weekday_values))
    schedule.profiles[0].hourly_values = wk_vals

    # Profile 2: Fins de Semana (Sat-Sun)
    schedule.profiles[1].name = '2:Fins de Semana'
    if weekend_values:
        we_vals = weekend_values[:24] if len(weekend_values) >= 24 else weekend_values + [0] * (24 - len(weekend_values))
    else:
        we_vals = wk_vals.copy()
    schedule.profiles[1].hourly_values = we_vals

    # Profile 3: Feriados
    schedule.profiles[2].name = '3:Feriados'
    if holiday_values:
        hol_vals = holiday_values[:24] if len(holiday_values) >= 24 else holiday_values + [0] * (24 - len(holiday_values))
    else:
        hol_vals = we_vals.copy()  # Default: feriados = fins de semana
    schedule.profiles[2].hourly_values = hol_vals

    # Set unknown values (min values for each profile)
    min_weekday = min(schedule.profiles[0].hourly_values)
    min_weekend = min(schedule.profiles[1].hourly_values)
    min_holiday = min(schedule.profiles[2].hourly_values)
    schedule.unknown_values = [min_weekday, min_weekend, min_holiday, 100, 100, 100, 100, 100]

    # Day mapping: 12 meses × 8 dias
    # índice = mês × 8 + dia
    schedule.day_mapping = [1] * 100  # Default all to profile 1

    # Set proper day mapping for all months
    schedule.set_weekday_profile(1)   # Mon-Fri = Profile 1
    schedule.set_weekend_profile(2)   # Sat-Sun = Profile 2
    schedule.set_holiday_profile(3)   # Holiday = Profile 3
    schedule.set_design_profile(1)    # Design = Profile 1

    return schedule


def create_office_schedule(name: str) -> HAPSchedule:
    """Create a typical office schedule (8h-18h operation).

    Returns a schedule with:
    - Profile 1: Dias de Semana (100% 8:00-18:00, 5% other times)
    - Profile 2: Fins de Semana (5% all day)
    - Profile 3: Feriados (5% all day)
    """
    # Office hours: 8:00-18:00 (hours 8-17)
    weekday_values = [5] * 24  # Base 5%
    for h in range(8, 18):
        weekday_values[h] = 100  # 100% during office hours

    weekend_values = [5] * 24  # 5% all day
    holiday_values = [5] * 24  # 5% all day

    return create_simple_schedule(name, weekday_values, weekend_values, holiday_values)


def create_24h_schedule(name: str) -> HAPSchedule:
    """Create a 24/7 schedule (100% always)."""
    always_on = [100] * 24
    return create_simple_schedule(name, always_on, always_on, always_on)


def create_residential_schedule(name: str) -> HAPSchedule:
    """Create a typical residential schedule."""
    # Weekdays: occupied morning and evening
    weekday_values = [30] * 24
    for h in range(6, 9):     # Morning
        weekday_values[h] = 100
    for h in range(18, 23):   # Evening
        weekday_values[h] = 100

    # Weekends: more occupied
    weekend_values = [50] * 24
    for h in range(8, 23):
        weekend_values[h] = 80

    return create_simple_schedule(name, weekday_values, weekend_values)


def create_commercial_schedule(name: str) -> HAPSchedule:
    """Create a commercial/retail schedule (10h-22h)."""
    # Weekdays: 10:00-22:00
    weekday_values = [5] * 24
    for h in range(10, 22):
        weekday_values[h] = 100

    # Weekends: shorter hours
    weekend_values = [5] * 24
    for h in range(10, 20):
        weekend_values[h] = 100

    return create_simple_schedule(name, weekday_values, weekend_values)


# =============================================================================
# MAIN (DEMO/TEST)
# =============================================================================

if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("HAP 5.1 Schedule Library v2.0")
        print("=============================")
        print()
        print("Usage:")
        print(f"  python {sys.argv[0]} <file.E3A>  - Read and display schedules")
        print(f"  python {sys.argv[0]} --demo      - Show demo schedule creation")
        print()
        print("Day Mapping Structure:")
        print("  - 12 meses × 8 dias = 96 valores (índices 0-95)")
        print("  - 4 valores Design (índices 96-99)")
        print("  - Fórmula: índice = mês × 8 + dia")
        print("  - Dias: 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun, 7=Holiday")
        sys.exit(0)

    if sys.argv[1] == '--demo':
        print("Creating demo office schedule...")
        schedule = create_office_schedule('Escritorio')
        print(f"Name: {schedule.name}")
        print()
        print("Profiles:")
        for i, p in enumerate(schedule.profiles[:3]):
            print(f"  {p.name}")
        print()
        print(f"Profile 1 (Dias de Semana):")
        print(f"  Hours 0-7:  {schedule.profiles[0].hourly_values[:8]}")
        print(f"  Hours 8-17: {schedule.profiles[0].hourly_values[8:18]}")
        print(f"  Hours 18-23: {schedule.profiles[0].hourly_values[18:]}")
        print()
        print("Assignments:")
        schedule.print_assignments()
    else:
        filepath = sys.argv[1]
        print(f"Loading schedules from: {filepath}")
        print()

        manager = ScheduleManager.from_e3a_file(filepath)
        manager.print_summary()

        # Detailed view of first non-Sample schedule
        if len(manager.schedules) > 1:
            sch = manager.schedules[1]
            print(f"\n{'='*60}")
            print(f"Detailed view: {sch.name.strip()}")
            print(f"{'='*60}")
            print()
            print("Profiles:")
            for i, p in enumerate(sch.profiles[:3]):
                if p.name.strip():
                    print(f"  {p.name.strip()}")
            print()
            print("Assignments:")
            sch.print_assignments()
