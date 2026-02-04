"""
Criar/Modificar Schedules em ficheiro HAP 5.1
=============================================

Este script permite criar novos schedules ou modificar existentes
num ficheiro HAP .E3A.

Usage:
    python criar_schedule.py <ficheiro.E3A> --criar "Nome" --tipo escritorio
    python criar_schedule.py <ficheiro.E3A> --listar
    python criar_schedule.py <ficheiro.E3A> --ver "Nome do Schedule"
"""

import argparse
import zipfile
import struct
import shutil
import os
from hap_schedule_library import (
    HAPSchedule, ScheduleManager, ScheduleProfile,
    create_simple_schedule, create_office_schedule,
    SCHEDULE_RECORD_SIZE
)


def criar_schedule_escritorio(name: str) -> HAPSchedule:
    """Cria schedule típico de escritório (8h-18h)."""
    # Dias de semana: 8:00-18:00 ocupado
    weekday = [5] * 24
    for h in range(8, 18):
        weekday[h] = 100

    # Fins de semana: 5% todo o dia
    weekend = [5] * 24

    return create_simple_schedule(name, weekday, weekend)


def criar_schedule_24h(name: str) -> HAPSchedule:
    """Cria schedule 24/7 (100% sempre)."""
    always_on = [100] * 24
    return create_simple_schedule(name, always_on, always_on, always_on)


def criar_schedule_residencial(name: str) -> HAPSchedule:
    """Cria schedule residencial típico."""
    # Dias de semana: ocupado manhã e noite
    weekday = [50] * 24
    for h in range(7, 9):    # Manhã
        weekday[h] = 100
    for h in range(18, 23):  # Noite
        weekday[h] = 100

    # Fins de semana: mais ocupado
    weekend = [50] * 24
    for h in range(8, 23):
        weekend[h] = 80

    return create_simple_schedule(name, weekday, weekend)


def criar_schedule_comercio(name: str) -> HAPSchedule:
    """Cria schedule de comércio (10h-22h)."""
    # Dias de semana
    weekday = [5] * 24
    for h in range(10, 22):
        weekday[h] = 100

    # Fins de semana: horário reduzido
    weekend = [5] * 24
    for h in range(10, 20):
        weekend[h] = 100

    return create_simple_schedule(name, weekday, weekend)


def adicionar_schedule_ao_ficheiro(filepath: str, schedule: HAPSchedule):
    """Adiciona um schedule a um ficheiro .E3A existente."""
    # Ler ficheiro
    with zipfile.ZipFile(filepath, 'r') as zf:
        files = {name: zf.read(name) for name in zf.namelist()}

    # Ler schedules existentes
    sch_data = bytearray(files['HAP51SCH.DAT'])
    num_schedules = len(sch_data) // SCHEDULE_RECORD_SIZE

    print(f"Schedules existentes: {num_schedules}")

    # Verificar se já existe schedule com mesmo nome
    for i in range(num_schedules):
        existing_name = sch_data[i*SCHEDULE_RECORD_SIZE:i*SCHEDULE_RECORD_SIZE+24].decode('latin-1').rstrip('\x00')
        if existing_name.strip() == schedule.name.strip():
            print(f"Aviso: Já existe schedule com nome '{schedule.name}'. A substituir...")
            # Substituir
            from hap_schedule_library import encode_schedule
            sch_data[i*SCHEDULE_RECORD_SIZE:(i+1)*SCHEDULE_RECORD_SIZE] = encode_schedule(schedule)
            files['HAP51SCH.DAT'] = bytes(sch_data)

            # Guardar
            with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
                for name, data in files.items():
                    zf.writestr(name, data)
            print(f"Schedule '{schedule.name}' substituído com sucesso!")
            return

    # Adicionar novo schedule
    from hap_schedule_library import encode_schedule
    sch_data.extend(encode_schedule(schedule))

    files['HAP51SCH.DAT'] = bytes(sch_data)

    # Guardar
    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)

    print(f"Schedule '{schedule.name}' adicionado com sucesso!")
    print(f"Total de schedules: {num_schedules + 1}")


def listar_schedules(filepath: str):
    """Lista todos os schedules num ficheiro."""
    manager = ScheduleManager.from_e3a_file(filepath)
    manager.print_summary()


def ver_schedule(filepath: str, name: str):
    """Mostra detalhes de um schedule específico."""
    manager = ScheduleManager.from_e3a_file(filepath)
    schedule = manager.get_schedule_by_name(name)

    if not schedule:
        print(f"Schedule '{name}' não encontrado.")
        print("Schedules disponíveis:")
        for s in manager.schedules:
            print(f"  - {s.name.strip()}")
        return

    print(f"=== {schedule.name.strip()} ===")
    print()

    # Mostrar profiles
    print("Profiles:")
    for i, p in enumerate(schedule.profiles):
        if p.name.strip():
            print(f"\n  {i+1}. {p.name.strip()}")
            print(f"      Valores horários:")
            for h in range(0, 24, 6):
                vals = p.hourly_values[h:h+6]
                hours = [f"H{h+j:02d}:{vals[j]:3d}%" for j in range(len(vals))]
                print(f"        {', '.join(hours)}")

    # Mostrar day mapping
    print("\nMapeamento de dias (por mês):")
    day_names = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Fer']
    months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

    # Verificar se todos os meses são iguais
    month_patterns = []
    for month in range(12):
        start = month * 8
        if start + 8 <= len(schedule.day_mapping):
            pattern = tuple(schedule.day_mapping[start:start+8])
            month_patterns.append(pattern)

    if len(set(month_patterns)) == 1:
        # Todos iguais
        print("  (Mesmo padrão para todos os meses)")
        pattern = month_patterns[0]
        for i, day in enumerate(day_names):
            print(f"    {day}: Profile {pattern[i]}")
    else:
        # Diferente por mês
        for month_idx, pattern in enumerate(month_patterns):
            formatted = ', '.join(f"{day_names[i]}={pattern[i]}" for i in range(8))
            print(f"  {months[month_idx]}: {formatted}")


def main():
    parser = argparse.ArgumentParser(description='Criar/Modificar Schedules HAP')
    parser.add_argument('ficheiro', help='Ficheiro .E3A')
    parser.add_argument('--listar', action='store_true', help='Listar schedules')
    parser.add_argument('--ver', type=str, help='Ver detalhes de um schedule')
    parser.add_argument('--criar', type=str, help='Nome do novo schedule')
    parser.add_argument('--tipo', type=str, choices=['escritorio', '24h', 'residencial', 'comercio'],
                        default='escritorio', help='Tipo de schedule')

    args = parser.parse_args()

    if not os.path.exists(args.ficheiro):
        print(f"Erro: Ficheiro não encontrado: {args.ficheiro}")
        return

    if args.listar:
        listar_schedules(args.ficheiro)
    elif args.ver:
        ver_schedule(args.ficheiro, args.ver)
    elif args.criar:
        # Criar schedule
        if args.tipo == 'escritorio':
            schedule = criar_schedule_escritorio(args.criar)
        elif args.tipo == '24h':
            schedule = criar_schedule_24h(args.criar)
        elif args.tipo == 'residencial':
            schedule = criar_schedule_residencial(args.criar)
        elif args.tipo == 'comercio':
            schedule = criar_schedule_comercio(args.criar)

        adicionar_schedule_ao_ficheiro(args.ficheiro, schedule)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
