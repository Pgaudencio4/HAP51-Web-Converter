"""
Validador e Corrector de ficheiros HAP 5.1 (.E3A)

Verifica e corrige problemas conhecidos:
1. Calendário dos schedules com valores inválidos (>8)
2. Default Space com schedule IDs != 0
3. Schedule IDs inválidos nos spaces

Usage:
    python validar_e3a.py <ficheiro.E3A> [--fix]

Exemplo:
    python validar_e3a.py MeuFicheiro.E3A          # Só validar
    python validar_e3a.py MeuFicheiro.E3A --fix    # Validar e corrigir
"""

import zipfile
import struct
import sys
import os
import shutil

def validate_e3a(path, fix=False):
    """Valida um ficheiro E3A e opcionalmente corrige erros."""

    errors = []
    warnings = []
    fixes_made = []

    print(f"\n{'='*60}")
    print(f"VALIDAÇÃO: {os.path.basename(path)}")
    print(f"{'='*60}")

    # Ler o ficheiro
    with zipfile.ZipFile(path, 'r') as z:
        files_content = {}
        for name in z.namelist():
            files_content[name] = bytearray(z.read(name))

    # =================================================================
    # 1. VERIFICAR SCHEDULES (calendário)
    # =================================================================
    print("\n[1] Verificar Schedules...")

    sch_data = files_content.get('HAP51SCH.DAT')
    if sch_data:
        num_schedules = len(sch_data) // 792
        print(f"    Schedules encontrados: {num_schedules}")

        schedules_with_bad_calendar = []

        for i in range(num_schedules):
            sch_start = i * 792
            sch_name = sch_data[sch_start:sch_start+40].rstrip(b' \x00').decode('latin1', errors='ignore')

            # Verificar calendário (bytes 576-791)
            bad_values = []
            for j in range(576, 792, 2):
                val = struct.unpack('<H', sch_data[sch_start+j:sch_start+j+2])[0]
                if val > 8 and val != 0:
                    bad_values.append((j, val))

            if bad_values:
                schedules_with_bad_calendar.append((i, sch_name, bad_values))
                errors.append(f"Schedule {i} '{sch_name}': calendário com valores inválidos ({bad_values[0][1]}, ...)")

                if fix:
                    # Corrigir: substituir valores >8 por 1
                    for j, val in bad_values:
                        sch_data[sch_start+j:sch_start+j+2] = struct.pack('<H', 1)
                    fixes_made.append(f"Schedule {i} '{sch_name}': calendário corrigido")

        if schedules_with_bad_calendar:
            print(f"    ERRO: {len(schedules_with_bad_calendar)} schedules com calendário inválido!")
        else:
            print(f"    OK: Todos os schedules têm calendário válido")

    # =================================================================
    # 2. VERIFICAR DEFAULT SPACE
    # =================================================================
    print("\n[2] Verificar Default Space...")

    spc_data = files_content.get('HAP51SPC.DAT')
    if spc_data:
        num_spaces = len(spc_data) // 682
        print(f"    Spaces encontrados: {num_spaces}")

        # Default Space é o primeiro (índice 0)
        ds = spc_data[0:682]
        ds_name = ds[0:24].rstrip(b' \x00').decode('latin1', errors='ignore')
        print(f"    Default Space: '{ds_name}'")

        schedule_offsets = {
            554: "Infiltration Sch 1",
            560: "Infiltration Sch 2",
            566: "Infiltration Sch 3",
            594: "People Schedule",
            616: "Light Schedule",
            660: "Equip Schedule",
        }

        ds_errors = []
        for offset, desc in schedule_offsets.items():
            val = struct.unpack('<H', ds[offset:offset+2])[0]
            if val != 0:
                ds_errors.append((offset, desc, val))
                errors.append(f"Default Space: {desc} (offset {offset}) = {val} (deve ser 0)")

                if fix:
                    spc_data[offset:offset+2] = struct.pack('<H', 0)
                    fixes_made.append(f"Default Space: offset {offset} corrigido para 0")

        if ds_errors:
            print(f"    ERRO: Default Space tem {len(ds_errors)} schedule IDs != 0")
        else:
            print(f"    OK: Default Space tem todos os schedule IDs = 0")

    # =================================================================
    # 3. VERIFICAR SCHEDULE IDs NOS SPACES
    # =================================================================
    print("\n[3] Verificar Schedule IDs nos Spaces...")

    if spc_data and sch_data:
        num_schedules = len(sch_data) // 792
        num_spaces = len(spc_data) // 682

        invalid_refs = []

        for i in range(1, num_spaces):  # Começar do 1 (ignorar Default Space)
            space = spc_data[i*682:(i+1)*682]
            space_name = space[0:24].rstrip(b' \x00').decode('latin1', errors='ignore')

            for offset, desc in schedule_offsets.items():
                val = struct.unpack('<H', space[offset:offset+2])[0]
                if val >= num_schedules and val != 0 and val != 65535:
                    invalid_refs.append((i, space_name, offset, desc, val))
                    errors.append(f"Space {i} '{space_name}': {desc} = {val} >= {num_schedules}")

        if invalid_refs:
            print(f"    ERRO: {len(invalid_refs)} referências a schedules inválidos!")
        else:
            print(f"    OK: Todos os schedule IDs são válidos")

    # =================================================================
    # 4. VERIFICAR ASSEMBLIES (opcional)
    # =================================================================
    print("\n[4] Verificar Assemblies...")

    wall_data = files_content.get('HAP51WAL.DAT')
    if wall_data:
        if len(wall_data) % 3187 != 0:
            warnings.append(f"HAP51WAL.DAT: tamanho {len(wall_data)} não é múltiplo de 3187")
            print(f"    AVISO: HAP51WAL.DAT tamanho irregular")
        else:
            num_walls = len(wall_data) // 3187
            print(f"    Walls: {num_walls} assemblies OK")

    roof_data = files_content.get('HAP51ROF.DAT')
    if roof_data:
        if len(roof_data) % 3187 != 0:
            warnings.append(f"HAP51ROF.DAT: tamanho {len(roof_data)} não é múltiplo de 3187")
            print(f"    AVISO: HAP51ROF.DAT tamanho irregular")
        else:
            num_roofs = len(roof_data) // 3187
            print(f"    Roofs: {num_roofs} assemblies OK")

    # =================================================================
    # RESULTADO
    # =================================================================
    print(f"\n{'='*60}")
    print("RESULTADO")
    print(f"{'='*60}")

    if errors:
        print(f"\nERROS: {len(errors)}")
        for e in errors[:10]:  # Mostrar só os primeiros 10
            print(f"  - {e}")
        if len(errors) > 10:
            print(f"  ... e mais {len(errors)-10} erros")
    else:
        print("\nNenhum erro encontrado!")

    if warnings:
        print(f"\nAVISOS: {len(warnings)}")
        for w in warnings:
            print(f"  - {w}")

    # =================================================================
    # GRAVAR CORRECÇÕES
    # =================================================================
    if fix and fixes_made:
        print(f"\nCORRECÇÕES APLICADAS: {len(fixes_made)}")
        for f in fixes_made[:10]:
            print(f"  - {f}")
        if len(fixes_made) > 10:
            print(f"  ... e mais {len(fixes_made)-10} correcções")

        # Actualizar dados
        files_content['HAP51SCH.DAT'] = bytes(sch_data)
        files_content['HAP51SPC.DAT'] = bytes(spc_data)

        # Criar backup
        backup_path = path + ".backup"
        if not os.path.exists(backup_path):
            shutil.copy(path, backup_path)
            print(f"\nBackup criado: {backup_path}")

        # Gravar ficheiro corrigido
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            for name, content in files_content.items():
                z.writestr(name, bytes(content))

        print(f"\nFicheiro corrigido gravado: {path}")

    return len(errors) == 0


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    path = sys.argv[1]
    fix = '--fix' in sys.argv

    if not os.path.exists(path):
        print(f"ERRO: Ficheiro não encontrado: {path}")
        sys.exit(1)

    if fix:
        print("\n*** MODO CORRECÇÃO ACTIVO ***")
        print("O ficheiro será modificado se houver erros!")

    valid = validate_e3a(path, fix=fix)

    sys.exit(0 if valid else 1)


if __name__ == '__main__':
    main()
