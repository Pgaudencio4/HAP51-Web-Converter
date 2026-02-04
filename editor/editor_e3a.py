"""
Editor E3A - Modifica ficheiros HAP existentes

Usa o mesmo formato do comparador:
- PREV (vazio) - preencher só o que queres alterar
- REF (preenchido) - valores actuais do E3A
- ? - indicador de alteração

Fluxo:
1. Extrair E3A para Excel com formato PREV/REF (usa template do comparador)
2. Utilizador preenche PREV (só os campos que quer alterar)
3. Aplicar alterações ao E3A original (mantém sistemas, schedules, etc.)

Usage:
    # Passo 1: Criar Excel de edição a partir do E3A
    python editor_e3a.py extrair <ficheiro.E3A> <output_editor.xlsx>

    # Passo 2: (Manual) Preencher coluna PREV no Excel

    # Passo 3: Aplicar alterações ao E3A
    python editor_e3a.py aplicar <ficheiro.E3A> <editor.xlsx> <output.E3A>
"""

import sys
import os
import struct
import zipfile
import shutil
import tempfile
import re
import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Adicionar paths para reutilizar módulos
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'extractor'))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'comparador'))

# Constantes
SPACE_RECORD_SIZE = 682

# Cores
PREV_FILL = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')  # Amarelo claro
REF_FILL = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')  # Verde claro
CHECK_FILL = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')  # Cinza
CHANGED_FILL = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')  # Amarelo

# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

def extract_space_name(raw_bytes):
    """Extrai nome do espaço correctamente"""
    try:
        name_str = raw_bytes.decode('latin-1')
        match = re.match(r'^([^\x00]+?)(\s{3,}|\x00)', name_str)
        if match:
            return match.group(1).strip()
        return name_str.split()[0] if name_str.split() else ''
    except:
        return ''


# =============================================================================
# EXTRACÇÃO PARA EXCEL DE EDIÇÃO (formato comparador)
# =============================================================================

def extract_for_editing(e3a_path, output_xlsx):
    """Extrai E3A para Excel com formato do comparador (PREV/REF)"""

    print(f"A extrair {e3a_path} para edição...")

    # Usar o extractor existente para extrair dados
    extractor_path = os.path.join(os.path.dirname(__file__), '..', 'extractor', 'hap_extractor.py')

    # Extrair para ficheiro temporário
    temp_xlsx = output_xlsx.replace('.xlsx', '_temp.xlsx')

    # Executar extractor via subprocess
    import subprocess
    result = subprocess.run(
        [sys.executable, extractor_path, e3a_path, temp_xlsx],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"Erro no extractor: {result.stderr}")
        return
    print(result.stdout)

    # Carregar template do comparador (mesmo formato do conversor)
    template_path = os.path.join(os.path.dirname(__file__), '..', 'comparador', 'Template_Comparacao.xlsx')

    if not os.path.exists(template_path):
        print(f"Template não encontrado: {template_path}")
        print("A criar template...")
        # Criar template se não existir
        criar_template_path = os.path.join(os.path.dirname(__file__), '..', 'comparador', 'criar_template_comparacao.py')
        spec2 = importlib.util.spec_from_file_location("criar_template", criar_template_path)
        criar_mod = importlib.util.module_from_spec(spec2)
        spec2.loader.exec_module(criar_mod)
        criar_mod.criar_template()

    # Carregar template
    wb_template = openpyxl.load_workbook(template_path)
    ws = wb_template['Comparacao']

    # Carregar dados extraídos
    wb_data = openpyxl.load_workbook(temp_xlsx)
    ws_data = wb_data['Espacos']

    # Ler headers do dados extraído (linha 3)
    data_headers = [ws_data.cell(3, col).value for col in range(1, 148)]

    # Actualizar headers do template (linha 3) - PREV fica vazio, REF tem nome
    max_col = 147 * 3
    for col in range(1, max_col + 1):
        cell = ws.cell(3, col)
        if cell.value:
            val = str(cell.value)
            val = re.sub(r'\([^)]+\)$', '', val).strip()
            col_type = (col - 1) % 3
            if col_type == 0:  # PREV
                cell.value = f"{val} (PREV)"
                cell.fill = PREV_FILL
            elif col_type == 1:  # REF
                cell.value = f"{val} (REF)"
                cell.fill = REF_FILL

    # Copiar dados para colunas PREV (coluna 1 de cada trio)
    # Formato: PREV (preenchido) | REF (vazio) | ?
    row_out = 4
    for row_in in range(4, ws_data.max_row + 1):
        space_name = ws_data.cell(row_in, 1).value
        if not space_name or str(space_name).strip() == '':
            continue

        out_col = 1
        for in_col in range(1, 148):
            value = ws_data.cell(row_in, in_col).value

            # Coluna PREV - valor actual (amarelo claro)
            c1 = ws.cell(row_out, out_col, value=value)
            c1.fill = PREV_FILL

            # Coluna REF - vazia (verde claro)
            c2 = ws.cell(row_out, out_col + 1)
            c2.value = None
            c2.fill = REF_FILL

            # Coluna ? - vazia (cinza)
            c3 = ws.cell(row_out, out_col + 2)
            c3.value = None
            c3.fill = CHECK_FILL
            c3.alignment = Alignment(horizontal='center')

            out_col += 3

        row_out += 1

    # Processar também Windows, Walls, Roofs se existirem
    for sheet_name in ['Windows', 'Walls', 'Roofs']:
        if sheet_name in wb_data.sheetnames and sheet_name in wb_template.sheetnames:
            process_simple_sheet(wb_template, wb_data, sheet_name)

    # Guardar
    wb_template.save(output_xlsx)

    # Remover temporário
    if os.path.exists(temp_xlsx):
        os.remove(temp_xlsx)

    print(f"\nFicheiro criado: {output_xlsx}")
    print(f"  - Coluna PREV (amarelo): preencher só o que queres alterar")
    print(f"  - Coluna REF (verde): valores actuais do E3A")
    print(f"  - Deixar PREV vazio = não altera o campo")


def process_simple_sheet(wb_template, wb_data, sheet_name):
    """Processa folhas simples (Windows, Walls, Roofs)"""
    ws_template = wb_template[sheet_name]
    ws_data = wb_data[sheet_name]

    # Actualizar headers (linha 3)
    for col in range(1, ws_template.max_column + 1):
        cell = ws_template.cell(3, col)
        if cell.value:
            val = str(cell.value)
            val = re.sub(r'\([^)]+\)$', '', val).strip()
            col_type = (col - 1) % 3
            if col_type == 0:
                cell.value = f"{val} (PREV)"
                cell.fill = PREV_FILL
            elif col_type == 1:
                cell.value = f"{val} (REF)"
                cell.fill = REF_FILL

    # Copiar dados
    # Formato: PREV (preenchido) | REF (vazio) | ?
    num_fields = ws_template.max_column // 3
    row_out = 4

    for row_in in range(4, ws_data.max_row + 1):
        name = ws_data.cell(row_in, 1).value
        if not name:
            continue

        out_col = 1
        for in_col in range(1, ws_data.max_column + 1):
            if out_col > ws_template.max_column:
                break

            value = ws_data.cell(row_in, in_col).value

            # PREV com valor
            c1 = ws_template.cell(row_out, out_col, value=value)
            c1.fill = PREV_FILL

            # REF vazio
            c2 = ws_template.cell(row_out, out_col + 1)
            c2.value = None
            c2.fill = REF_FILL

            # ? vazio
            c3 = ws_template.cell(row_out, out_col + 2)
            c3.value = None
            c3.fill = CHECK_FILL

            out_col += 3

        row_out += 1


# =============================================================================
# APLICAR ALTERAÇÕES AO E3A
# =============================================================================

def apply_changes(e3a_path, editor_xlsx, output_path):
    """Aplica alterações do Excel ao E3A"""

    print(f"A aplicar alterações de {editor_xlsx} a {e3a_path}...")

    # Carregar Excel de edição
    wb = openpyxl.load_workbook(editor_xlsx)
    ws = wb['Comparacao']

    # Ler E3A
    with zipfile.ZipFile(e3a_path, 'r') as zf:
        spc_data = bytearray(zf.read('HAP51SPC.DAT'))

    # Mapear nomes de espaços para offsets
    space_offsets = {}
    num_spaces = len(spc_data) // SPACE_RECORD_SIZE
    for i in range(num_spaces):
        offset = i * SPACE_RECORD_SIZE
        name = extract_space_name(spc_data[offset:offset+100])
        if name and not name.startswith('Default') and not name.startswith('Sample'):
            space_offsets[name] = offset

    # Ler headers para mapear colunas para campos
    # Formato: "Campo (PREV)" nas colunas 1, 4, 7, ...
    field_columns = {}  # {col_prev: field_name}
    for col in range(1, ws.max_column + 1, 3):
        header = ws.cell(3, col).value
        if header:
            field_name = re.sub(r'\s*\(PREV\)$', '', str(header)).strip()
            field_columns[col] = field_name

    # Mapeamento de campos para offsets no E3A
    # Formato: 'Nome Header': (offset, tipo, conversão_SI_para_IP)
    # Conversões: m² -> ft² (10.7639), m -> ft (3.28084), etc.
    FIELD_OFFSETS = {
        # GENERAL (cols 1-6)
        'Space Name': (0, 's24'),
        'Area': (24, 'f', 10.7639),  # m² -> ft²
        'Height': (28, 'f', 3.28084),  # m -> ft
        # 'Building Weight': (32, 'f', 0.204816),  # kg/m² -> lb/ft² (não no template)

        # INTERNALS - PEOPLE (cols 7-11)
        'Occupants': (580, 'f', 1),
        'Sensible': (586, 'f', 3.412),  # W -> BTU/hr
        'Latent': (590, 'f', 3.412),  # W -> BTU/hr

        # INTERNALS - LIGHTING (cols 12-16)
        'Lighting W/m2': (606, 'f', 0.0929),  # Converte para W e depois divide pela area
        'Ballast Multiplier': (610, 'f', 1),

        # INTERNALS - EQUIPMENT (cols 17-18)
        'Equipment W/m2': (656, 'f', 0.0929),  # W/m² -> W/ft²

        # INTERNALS - MISC (cols 19-22)
        'Misc Sensible': (632, 'f', 3.412),  # W -> BTU/hr
        'Misc Latent': (636, 'f', 3.412),  # W -> BTU/hr

        # INFILTRATION (cols 23-26)
        'ACH Cooling': (556, 'f', 1),  # ACH (sem conversão)
        'ACH Heating': (562, 'f', 1),
        'ACH Ventilation': (568, 'f', 1),

        # PARTITIONS - CEILING
        'Ceiling U-Value': (446, 'f', 0.1761),  # W/m²K -> BTU/(hr·ft²·°F)
        'Ceiling Area': (442, 'f', 10.7639),  # m² -> ft²

        # PARTITIONS - WALL
        'Part Wall U-Value': (472, 'f', 0.1761),
        'Part Wall Area': (468, 'f', 10.7639),
    }

    changes_count = 0

    # Processar cada linha de dados (a partir da linha 4)
    for row in range(4, ws.max_row + 1):
        # O nome do espaço está na coluna PREV do primeiro campo (coluna 1)
        space_name = ws.cell(row, 1).value
        if not space_name or space_name not in space_offsets:
            continue

        space_offset = space_offsets[space_name]

        # Verificar cada campo
        for col_prev, field_name in field_columns.items():
            prev_value = ws.cell(row, col_prev).value

            # Só processar se PREV tiver valor
            if prev_value is None or prev_value == '':
                continue

            # Encontrar offset e tipo do campo
            if field_name not in FIELD_OFFSETS:
                continue

            field_info = FIELD_OFFSETS[field_name]
            field_offset = field_info[0]
            field_type = field_info[1]
            conversion = field_info[2] if len(field_info) > 2 else 1

            abs_offset = space_offset + field_offset

            # Converter e escrever valor
            try:
                if field_type == 'f':
                    value = float(prev_value) * conversion
                    packed = struct.pack('<f', value)
                    spc_data[abs_offset:abs_offset+4] = packed
                elif field_type.startswith('s'):
                    str_len = int(field_type[1:])
                    encoded = str(prev_value).encode('latin-1')[:str_len-1]
                    encoded = encoded + b'\x00' * (str_len - len(encoded))
                    spc_data[abs_offset:abs_offset+str_len] = encoded

                print(f"  {space_name} / {field_name}: {prev_value}")
                changes_count += 1

            except Exception as e:
                print(f"  ERRO {space_name}/{field_name}: {e}")

    if changes_count == 0:
        print("\n  Nenhuma alteração encontrada na coluna PREV")
        return

    # Guardar E3A modificado
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(e3a_path, 'r') as zf:
            zf.extractall(tmpdir)

        spc_path = os.path.join(tmpdir, 'HAP51SPC.DAT')
        with open(spc_path, 'wb') as f:
            f.write(spc_data)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, tmpdir)
                    zf.write(file_path, arc_name)

    print(f"\n  {changes_count} alterações aplicadas")
    print(f"  Ficheiro criado: {output_path}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    comando = sys.argv[1].lower()

    if comando == 'extrair':
        if len(sys.argv) < 4:
            print("Uso: python editor_e3a.py extrair <ficheiro.E3A> <output.xlsx>")
            sys.exit(1)
        extract_for_editing(sys.argv[2], sys.argv[3])

    elif comando == 'aplicar':
        if len(sys.argv) < 5:
            print("Uso: python editor_e3a.py aplicar <original.E3A> <editor.xlsx> <output.E3A>")
            sys.exit(1)
        apply_changes(sys.argv[2], sys.argv[3], sys.argv[4])

    else:
        print(f"Comando desconhecido: {comando}")
        print("Comandos disponíveis: extrair, aplicar")
        sys.exit(1)


if __name__ == '__main__':
    main()
