"""
HAP 5.1 Web Application
=======================

Aplicacao web para converter entre Excel e ficheiros HAP 5.1 (.E3A).

Funcionalidades:
- Upload Excel -> Download E3A
- Upload E3A -> Download Excel

Executar:
    python app.py

Abrir no browser:
    http://localhost:5000
"""

import os
import tempfile
import shutil
import zipfile
import struct
from flask import Flask, render_template_string, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

# Importar bibliotecas do projecto
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

app = Flask(__name__)
app.secret_key = 'hap51_converter_secret_key'

# Configuracao
UPLOAD_FOLDER = tempfile.mkdtemp()
ALLOWED_EXTENSIONS_EXCEL = {'xlsx', 'xls'}
ALLOWED_EXTENSIONS_E3A = {'e3a'}
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODELO_RSECE = os.path.join(BASE_DIR, 'Modelo_RSECE.E3A')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# =============================================================================
# CONSTANTES E CONVERSOES (copiadas de excel_to_hap.py)
# =============================================================================

RECORD_SIZE = 682
SCHEDULE_RECORD_SIZE = 792
WALL_BLOCK_SIZE = 34
WALL_BLOCK_START = 72
ROOF_BLOCK_SIZE = 24
ROOF_BLOCK_START = 344

DIRECTION_CODES = {
    'N': 1, 'NNE': 2, 'NE': 3, 'ENE': 4,
    'E': 5, 'ESE': 6, 'SE': 7, 'SSE': 8,
    'S': 9, 'SSW': 10, 'SW': 11, 'WSW': 12,
    'W': 13, 'WNW': 14, 'NW': 15, 'NNW': 16
}

DIRECTION_NAMES = {v: k for k, v in DIRECTION_CODES.items()}

ACTIVITY_CODES = {
    'Seated at Rest': 0, 'Office Work': 3, 'Sedentary Work': 4,
    'Light Bench Work': 4, 'Medium Work': 5, 'Heavy Work': 6,
    'Dancing': 7, 'Athletics': 8,
}

FIXTURE_CODES = {
    'Recessed Unvented': 0, 'Vented to Return Air': 1,
    'Vented to Supply & Return': 2, 'Surface Mount/Pendant': 3,
}

FLOOR_TYPE_CODES = {
    'Floor Above Cond Space': 1, 'Floor Above Uncond Space': 2,
    'Slab Floor On Grade': 3, 'Slab Floor Below Grade': 4,
}

OA_UNIT_CODES = {'L/s': 1, 'L/s/m2': 2, 'L/s/person': 3, '%': 4}
OA_UNIT_NAMES = {v: k for k, v in OA_UNIT_CODES.items()}

OA_A = 0.00470356
OA_B = 2.71147770

import math

def m2_to_ft2(m2):
    return float(m2 or 0) * 10.7639

def ft2_to_m2(ft2):
    return float(ft2 or 0) / 10.7639

def m_to_ft(m):
    return float(m or 0) * 3.28084

def ft_to_m(ft):
    return float(ft or 0) / 3.28084

def kg_m2_to_lb_ft2(kg):
    return float(kg or 0) / 4.8824

def lb_ft2_to_kg_m2(lb):
    return float(lb or 0) * 4.8824

def u_si_to_ip(u):
    return float(u or 0) / 5.678

def u_ip_to_si(u):
    return float(u or 0) * 5.678

def w_to_btu(w):
    return float(w or 0) * 3.412

def btu_to_w(btu):
    return float(btu or 0) / 3.412

def w_m2_to_w_ft2(w):
    return float(w or 0) / 10.764

def w_ft2_to_w_m2(w):
    return float(w or 0) * 10.764

def encode_oa(value, unit_code):
    if not value or float(value) <= 0:
        return 0.0
    return math.log(float(value) / OA_A) / OA_B

def decode_oa(internal):
    if internal <= 0:
        return 0.0
    return OA_A * math.exp(OA_B * internal)

def safe_float(val, default=0.0):
    try:
        return float(val) if val else default
    except:
        return default

def safe_int(val, default=0):
    try:
        return int(val) if val else default
    except:
        return default

# =============================================================================
# FUNCOES DE CONVERSAO
# =============================================================================

def read_excel_spaces(excel_path):
    """Le espacos do Excel."""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb['Espacos']

    spaces = []
    for row in range(4, ws.max_row + 1):
        name = ws.cell(row=row, column=1).value
        if not name or str(name).strip() == '':
            continue

        space = {
            'name': str(name)[:24],
            'area': ws.cell(row=row, column=2).value,
            'height': ws.cell(row=row, column=3).value,
            'weight': ws.cell(row=row, column=4).value,
            'oa': ws.cell(row=row, column=5).value,
            'oa_unit': ws.cell(row=row, column=6).value,
            'occupancy': ws.cell(row=row, column=7).value,
            'activity': ws.cell(row=row, column=8).value,
            'sensible': ws.cell(row=row, column=9).value,
            'latent': ws.cell(row=row, column=10).value,
            'people_sch': ws.cell(row=row, column=11).value,
            'task_light': ws.cell(row=row, column=12).value,
            'general_light': ws.cell(row=row, column=13).value,
            'fixture': ws.cell(row=row, column=14).value,
            'ballast': ws.cell(row=row, column=15).value,
            'light_sch': ws.cell(row=row, column=16).value,
            'equipment': ws.cell(row=row, column=17).value,
            'equip_sch': ws.cell(row=row, column=18).value,
            'ach_clg': ws.cell(row=row, column=23).value,
            'ach_htg': ws.cell(row=row, column=24).value,
            'ach_energy': ws.cell(row=row, column=25).value,
        }

        # Walls
        space['walls'] = []
        for w in range(8):
            col = 51 + w * 9
            wall = {
                'exposure': ws.cell(row=row, column=col).value,
                'area': ws.cell(row=row, column=col+1).value,
                'type': ws.cell(row=row, column=col+2).value,
                'win1': ws.cell(row=row, column=col+3).value,
                'win1_qty': ws.cell(row=row, column=col+4).value,
                'win2': ws.cell(row=row, column=col+5).value,
                'win2_qty': ws.cell(row=row, column=col+6).value,
                'door': ws.cell(row=row, column=col+7).value,
                'door_qty': ws.cell(row=row, column=col+8).value,
            }
            space['walls'].append(wall)

        # Roofs
        space['roofs'] = []
        for r in range(4):
            col = 123 + r * 6
            roof = {
                'exposure': ws.cell(row=row, column=col).value,
                'area': ws.cell(row=row, column=col+1).value,
                'slope': ws.cell(row=row, column=col+2).value,
                'type': ws.cell(row=row, column=col+3).value,
            }
            space['roofs'].append(roof)

        spaces.append(space)

    # Ler tipos
    types = {'walls': {}, 'windows': {}, 'doors': {}, 'roofs': {}, 'schedules': {}}
    if 'Tipos' in wb.sheetnames:
        ws_tipos = wb['Tipos']
        for row in range(3, ws_tipos.max_row + 1):
            # Schedules (cols 13-14)
            id_val = ws_tipos.cell(row=row, column=13).value
            name = ws_tipos.cell(row=row, column=14).value
            if id_val and name:
                types['schedules'][str(name).strip()] = int(id_val)

    return spaces, types


def read_e3a_spaces(e3a_path):
    """Le espacos de um ficheiro E3A."""
    spaces = []
    schedules = {}

    with zipfile.ZipFile(e3a_path, 'r') as zf:
        # Ler espacos
        spc_data = zf.read('HAP51SPC.DAT')
        num_spaces = len(spc_data) // RECORD_SIZE

        for i in range(1, num_spaces):  # Skip record 0 (default)
            offset = i * RECORD_SIZE
            record = spc_data[offset:offset + RECORD_SIZE]

            name = record[0:24].rstrip(b'\x00').decode('latin-1', errors='ignore').strip()
            if not name:
                continue

            area_ft2 = struct.unpack_from('<f', record, 24)[0]
            height_ft = struct.unpack_from('<f', record, 28)[0]
            weight_lb = struct.unpack_from('<f', record, 32)[0]
            oa_internal = struct.unpack_from('<f', record, 46)[0]
            oa_unit = struct.unpack_from('<H', record, 50)[0]

            occupancy = struct.unpack_from('<f', record, 580)[0]
            activity = struct.unpack_from('<H', record, 584)[0]
            sensible_btu = struct.unpack_from('<f', record, 586)[0]
            latent_btu = struct.unpack_from('<f', record, 590)[0]
            people_sch = struct.unpack_from('<H', record, 594)[0]

            task_light = struct.unpack_from('<f', record, 600)[0]
            fixture = struct.unpack_from('<H', record, 604)[0]
            general_light = struct.unpack_from('<f', record, 606)[0]
            ballast = struct.unpack_from('<f', record, 610)[0]
            light_sch = struct.unpack_from('<H', record, 616)[0]

            equip_w_ft2 = struct.unpack_from('<f', record, 656)[0]
            equip_sch = struct.unpack_from('<H', record, 660)[0]

            ach_clg = struct.unpack_from('<f', record, 556)[0]
            ach_htg = struct.unpack_from('<f', record, 562)[0]
            ach_energy = struct.unpack_from('<f', record, 568)[0]

            # Floor (offset 492-541)
            floor_type = struct.unpack_from('<H', record, 492)[0]
            floor_area = struct.unpack_from('<f', record, 494)[0]
            floor_u = struct.unpack_from('<f', record, 498)[0]

            # Ceiling Partition (offset 440-465)
            ceil_area = struct.unpack_from('<f', record, 442)[0]
            ceil_u = struct.unpack_from('<f', record, 446)[0]

            # Wall Partition (offset 466-491)
            part_wall_area = struct.unpack_from('<f', record, 468)[0]
            part_wall_u = struct.unpack_from('<f', record, 472)[0]

            # Walls
            walls = []
            for w in range(8):
                wall_offset = WALL_BLOCK_START + w * WALL_BLOCK_SIZE
                exp_code = struct.unpack_from('<H', record, wall_offset)[0]
                wall_area = struct.unpack_from('<f', record, wall_offset + 2)[0]
                wall_type = struct.unpack_from('<H', record, wall_offset + 6)[0]
                win1_type = struct.unpack_from('<H', record, wall_offset + 8)[0]
                win2_type = struct.unpack_from('<H', record, wall_offset + 10)[0]
                win1_qty = struct.unpack_from('<H', record, wall_offset + 12)[0]
                win2_qty = struct.unpack_from('<H', record, wall_offset + 14)[0]
                door_type = struct.unpack_from('<H', record, wall_offset + 16)[0]
                door_qty = struct.unpack_from('<H', record, wall_offset + 18)[0]
                if exp_code > 0:
                    walls.append({
                        'exposure': DIRECTION_NAMES.get(exp_code, ''),
                        'area': ft2_to_m2(wall_area),
                        'type_id': wall_type,
                        'win1_id': win1_type,
                        'win1_qty': win1_qty if win1_qty > 0 else None,
                        'win2_id': win2_type,
                        'win2_qty': win2_qty if win2_qty > 0 else None,
                        'door_id': door_type,
                        'door_qty': door_qty if door_qty > 0 else None,
                    })

            # Roofs
            roofs = []
            for r in range(4):
                roof_offset = ROOF_BLOCK_START + r * ROOF_BLOCK_SIZE
                exp_code = struct.unpack_from('<H', record, roof_offset)[0]
                slope = struct.unpack_from('<H', record, roof_offset + 2)[0]
                roof_area = struct.unpack_from('<f', record, roof_offset + 4)[0]
                roof_type = struct.unpack_from('<H', record, roof_offset + 8)[0]
                sky_type = struct.unpack_from('<H', record, roof_offset + 10)[0]
                sky_qty = struct.unpack_from('<H', record, roof_offset + 12)[0]
                if exp_code > 0:
                    roofs.append({
                        'exposure': DIRECTION_NAMES.get(exp_code, ''),
                        'area': ft2_to_m2(roof_area),
                        'slope': slope,
                        'type_id': roof_type,
                        'sky_id': sky_type,
                        'sky_qty': sky_qty if sky_qty > 0 else None,
                    })

            space = {
                'name': name,
                'area': ft2_to_m2(area_ft2),
                'height': ft_to_m(height_ft),
                'weight': lb_ft2_to_kg_m2(weight_lb),
                'oa': decode_oa(oa_internal),
                'oa_unit': OA_UNIT_NAMES.get(oa_unit, 'L/s'),
                'occupancy': occupancy,
                'sensible': btu_to_w(sensible_btu),
                'latent': btu_to_w(latent_btu),
                'people_sch_id': people_sch,
                'task_light': task_light,
                'general_light': general_light,
                'ballast': ballast,
                'light_sch_id': light_sch,
                'equipment': w_ft2_to_w_m2(equip_w_ft2),
                'equip_sch_id': equip_sch,
                'ach_clg': ach_clg,
                'ach_htg': ach_htg,
                'ach_energy': ach_energy,
                'floor_area': ft2_to_m2(floor_area),
                'floor_u': u_ip_to_si(floor_u),
                'ceil_area': ft2_to_m2(ceil_area),
                'ceil_u': u_ip_to_si(ceil_u),
                'part_wall_area': ft2_to_m2(part_wall_area),
                'part_wall_u': u_ip_to_si(part_wall_u),
                'walls': walls,
                'roofs': roofs,
            }
            spaces.append(space)

        # Ler schedules
        if 'HAP51SCH.DAT' in zf.namelist():
            sch_data = zf.read('HAP51SCH.DAT')
            num_schedules = len(sch_data) // SCHEDULE_RECORD_SIZE
            for i in range(num_schedules):
                offset = i * SCHEDULE_RECORD_SIZE
                name = sch_data[offset:offset+24].rstrip(b'\x00').decode('latin-1', errors='ignore').strip()
                if name:
                    schedules[i] = name

        # Ler Windows (HAP51WIN.DAT - 555 bytes cada)
        windows = {}
        if 'HAP51WIN.DAT' in zf.namelist():
            win_data = zf.read('HAP51WIN.DAT')
            WIN_RECORD_SIZE = 555
            num_windows = len(win_data) // WIN_RECORD_SIZE
            for i in range(num_windows):
                offset = i * WIN_RECORD_SIZE
                name = win_data[offset:offset+255].rstrip(b'\x00').rstrip(b' ').decode('latin-1', errors='ignore').strip()
                if name:
                    height_ft = struct.unpack_from('<f', win_data, offset + 257)[0]
                    width_ft = struct.unpack_from('<f', win_data, offset + 261)[0]
                    u_value_ip = struct.unpack_from('<f', win_data, offset + 269)[0]
                    shgc = struct.unpack_from('<f', win_data, offset + 273)[0]
                    windows[i] = {
                        'name': name,
                        'u_value': u_ip_to_si(u_value_ip),
                        'shgc': shgc,
                        'height': ft_to_m(height_ft),
                        'width': ft_to_m(width_ft),
                    }

        # Ler Walls (HAP51WAL.DAT) - usar MDB para nomes
        walls_types = {}

        # Ler Roofs (HAP51ROF.DAT) - usar MDB para nomes
        roofs_types = {}

        # Tentar ler nomes do MDB (mais fiavel)
        try:
            import pyodbc
            # Extrair MDB temporariamente
            mdb_data = zf.read('HAP51INX.MDB')
            mdb_temp = tempfile.mktemp(suffix='.mdb')
            with open(mdb_temp, 'wb') as f:
                f.write(mdb_data)

            conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_temp};'
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            # Ler WallIndex
            try:
                cursor.execute("SELECT nIndex, szName, fOverallUValue, fOverallWeight, fThickness FROM WallIndex")
                for row in cursor.fetchall():
                    if row.szName:
                        walls_types[row.nIndex] = {
                            'name': row.szName.strip(),
                            'u_value': u_ip_to_si(row.fOverallUValue) if row.fOverallUValue else None,
                            'weight': lb_ft2_to_kg_m2(row.fOverallWeight) if row.fOverallWeight else None,
                            'thickness': ft_to_m(row.fThickness) if row.fThickness else None,
                        }
            except:
                pass

            # Ler RoofIndex
            try:
                cursor.execute("SELECT nIndex, szName, fOverallUValue, fOverallWeight, fThickness FROM RoofIndex")
                for row in cursor.fetchall():
                    if row.szName:
                        roofs_types[row.nIndex] = {
                            'name': row.szName.strip(),
                            'u_value': u_ip_to_si(row.fOverallUValue) if row.fOverallUValue else None,
                            'weight': lb_ft2_to_kg_m2(row.fOverallWeight) if row.fOverallWeight else None,
                            'thickness': ft_to_m(row.fThickness) if row.fThickness else None,
                        }
            except:
                pass

            conn.close()
            os.remove(mdb_temp)
        except:
            pass

    return spaces, schedules, windows, walls_types, roofs_types


def create_space_binary(space, types, template_record):
    """Cria registo binario de 682 bytes para um espaco."""
    data = bytearray(template_record)

    # Nome
    name_bytes = space['name'].encode('latin-1', errors='ignore')[:24].ljust(24, b'\x00')
    data[0:24] = name_bytes

    # Area, Altura, Peso
    struct.pack_into('<f', data, 24, m2_to_ft2(space.get('area')))
    struct.pack_into('<f', data, 28, m_to_ft(space.get('height')))
    struct.pack_into('<f', data, 32, kg_m2_to_lb_ft2(space.get('weight')))

    # OA
    oa_unit = OA_UNIT_CODES.get(space.get('oa_unit'), 3)
    oa_internal = encode_oa(space.get('oa'), oa_unit)
    struct.pack_into('<f', data, 46, oa_internal)
    struct.pack_into('<H', data, 50, oa_unit)

    # Walls
    for i in range(8):
        wall_start = WALL_BLOCK_START + i * WALL_BLOCK_SIZE
        for j in range(WALL_BLOCK_SIZE):
            data[wall_start + j] = 0

        if i < len(space.get('walls', [])):
            wall = space['walls'][i]
            exp = wall.get('exposure')
            if exp and exp in DIRECTION_CODES:
                struct.pack_into('<H', data, wall_start, DIRECTION_CODES[exp])
                struct.pack_into('<f', data, wall_start + 2, m2_to_ft2(wall.get('area')))

    # Roofs
    for i in range(4):
        roof_start = ROOF_BLOCK_START + i * ROOF_BLOCK_SIZE
        for j in range(ROOF_BLOCK_SIZE):
            data[roof_start + j] = 0

        if i < len(space.get('roofs', [])):
            roof = space['roofs'][i]
            exp = roof.get('exposure')
            if exp and exp in DIRECTION_CODES:
                struct.pack_into('<H', data, roof_start, DIRECTION_CODES[exp])
                struct.pack_into('<H', data, roof_start + 2, safe_int(roof.get('slope')))
                struct.pack_into('<f', data, roof_start + 4, m2_to_ft2(roof.get('area')))

    # Infiltration
    struct.pack_into('<H', data, 554, 2)
    struct.pack_into('<f', data, 556, safe_float(space.get('ach_clg')))
    struct.pack_into('<H', data, 560, 2)
    struct.pack_into('<f', data, 562, safe_float(space.get('ach_htg')))
    struct.pack_into('<H', data, 566, 2)
    struct.pack_into('<f', data, 568, safe_float(space.get('ach_energy')))

    # People
    struct.pack_into('<f', data, 580, safe_float(space.get('occupancy')))
    struct.pack_into('<H', data, 584, ACTIVITY_CODES.get(space.get('activity'), 3))
    struct.pack_into('<f', data, 586, w_to_btu(space.get('sensible')))
    struct.pack_into('<f', data, 590, w_to_btu(space.get('latent')))

    # Lighting
    struct.pack_into('<f', data, 600, safe_float(space.get('task_light')))
    struct.pack_into('<H', data, 604, FIXTURE_CODES.get(space.get('fixture'), 0))
    struct.pack_into('<f', data, 606, safe_float(space.get('general_light')))
    struct.pack_into('<f', data, 610, safe_float(space.get('ballast'), 1.0))

    # Equipment
    struct.pack_into('<f', data, 656, w_m2_to_w_ft2(space.get('equipment')))

    return bytes(data)


def excel_to_e3a(excel_path, output_path):
    """Converte Excel para E3A."""
    if not os.path.exists(MODELO_RSECE):
        raise FileNotFoundError(f"Modelo nao encontrado: {MODELO_RSECE}")

    spaces, types = read_excel_spaces(excel_path)
    if not spaces:
        raise ValueError("Nenhum espaco encontrado no Excel")

    temp_dir = tempfile.mkdtemp()
    try:
        # Extrair modelo
        with zipfile.ZipFile(MODELO_RSECE, 'r') as zf:
            zf.extractall(temp_dir)

        # Ler template e schedules
        spc_path = os.path.join(temp_dir, 'HAP51SPC.DAT')
        with open(spc_path, 'rb') as f:
            spc_data = f.read()

        template_record = spc_data[RECORD_SIZE:RECORD_SIZE*2]
        default_record = spc_data[0:RECORD_SIZE]

        # Ler schedules do modelo
        sch_path = os.path.join(temp_dir, 'HAP51SCH.DAT')
        if os.path.exists(sch_path):
            with open(sch_path, 'rb') as f:
                sch_data = f.read()
            num_schedules = len(sch_data) // SCHEDULE_RECORD_SIZE
            for i in range(num_schedules):
                offset = i * SCHEDULE_RECORD_SIZE
                name = sch_data[offset:offset+24].rstrip(b'\x00').decode('latin-1', errors='ignore').strip()
                if name and name not in types['schedules']:
                    types['schedules'][name] = i

        # Criar espacos
        new_spc_data = bytearray(default_record)
        for space in spaces:
            space_binary = create_space_binary(space, types, template_record)
            new_spc_data.extend(space_binary)

        with open(spc_path, 'wb') as f:
            f.write(bytes(new_spc_data))

        # Criar ZIP
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zf.write(file_path, arc_name)

        return len(spaces)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def e3a_to_excel(e3a_path, output_path):
    """Converte E3A para Excel usando o template completo HAP_Template_RSECE.xlsx."""
    spaces, schedules, windows, walls_types, roofs_types = read_e3a_spaces(e3a_path)
    if not spaces:
        raise ValueError("Nenhum espaco encontrado no E3A")

    # Usar o template como base
    template_path = os.path.join(BASE_DIR, 'HAP_Template_RSECE.xlsx')
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template nao encontrado: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    ws = wb['Espacos']

    # Limpar dados existentes (linhas 4 em diante)
    for row in range(4, ws.max_row + 1):
        for col in range(1, 148):
            ws.cell(row=row, column=col).value = None

    # Preencher dados dos espacos (a partir da linha 4)
    for row_idx, space in enumerate(spaces, 4):
        # GENERAL (cols 1-6)
        ws.cell(row=row_idx, column=1, value=space['name'])
        ws.cell(row=row_idx, column=2, value=round(space['area'], 2) if space['area'] else None)
        ws.cell(row=row_idx, column=3, value=round(space['height'], 2) if space['height'] else None)
        ws.cell(row=row_idx, column=4, value=round(space['weight'], 2) if space['weight'] else None)
        ws.cell(row=row_idx, column=5, value=round(space['oa'], 2) if space['oa'] else None)
        ws.cell(row=row_idx, column=6, value=space['oa_unit'])

        # PEOPLE (cols 7-11)
        ws.cell(row=row_idx, column=7, value=space.get('occupancy'))
        # col 8 = Activity Level (nao temos)
        ws.cell(row=row_idx, column=9, value=round(space['sensible'], 1) if space.get('sensible') else None)
        ws.cell(row=row_idx, column=10, value=round(space['latent'], 1) if space.get('latent') else None)
        ws.cell(row=row_idx, column=11, value=schedules.get(space.get('people_sch_id'), ''))

        # LIGHTING (cols 12-16)
        ws.cell(row=row_idx, column=12, value=space.get('task_light'))
        ws.cell(row=row_idx, column=13, value=space.get('general_light'))
        # col 14 = Fixture Type (nao temos)
        ws.cell(row=row_idx, column=15, value=space.get('ballast'))
        ws.cell(row=row_idx, column=16, value=schedules.get(space.get('light_sch_id'), ''))

        # EQUIPMENT (cols 17-18)
        ws.cell(row=row_idx, column=17, value=round(space['equipment'], 2) if space.get('equipment') else None)
        ws.cell(row=row_idx, column=18, value=schedules.get(space.get('equip_sch_id'), ''))

        # MISC (cols 19-22) - nao temos dados

        # INFILTRATION (cols 23-25)
        ws.cell(row=row_idx, column=23, value=space.get('ach_clg'))
        ws.cell(row=row_idx, column=24, value=space.get('ach_htg'))
        ws.cell(row=row_idx, column=25, value=space.get('ach_energy'))

        # FLOORS (cols 26-38)
        floor_area = space.get('floor_area')
        floor_u = space.get('floor_u')
        ws.cell(row=row_idx, column=27, value=round(floor_area, 2) if floor_area is not None and floor_area > 0 else None)
        ws.cell(row=row_idx, column=28, value=round(floor_u, 3) if floor_u is not None and floor_u > 0 else None)

        # PARTITIONS - CEILING (cols 39-44)
        ceil_area = space.get('ceil_area')
        ceil_u = space.get('ceil_u')
        ws.cell(row=row_idx, column=39, value=round(ceil_area, 2) if ceil_area is not None and ceil_area > 0 else None)
        ws.cell(row=row_idx, column=40, value=round(ceil_u, 3) if ceil_u is not None and ceil_u > 0 else None)

        # PARTITIONS - WALL (cols 45-50)
        part_wall_area = space.get('part_wall_area')
        part_wall_u = space.get('part_wall_u')
        ws.cell(row=row_idx, column=45, value=round(part_wall_area, 2) if part_wall_area is not None and part_wall_area > 0 else None)
        ws.cell(row=row_idx, column=46, value=round(part_wall_u, 3) if part_wall_u is not None and part_wall_u > 0 else None)

        # WALLS (cols 51-122: 8 walls x 9 cols)
        for w_idx, wall in enumerate(space.get('walls', [])):
            col_start = 51 + w_idx * 9
            ws.cell(row=row_idx, column=col_start, value=wall.get('exposure'))
            ws.cell(row=row_idx, column=col_start + 1, value=round(wall['area'], 2) if wall.get('area') else None)
            ws.cell(row=row_idx, column=col_start + 2, value=wall.get('type_name', ''))
            ws.cell(row=row_idx, column=col_start + 3, value=wall.get('win1_name', ''))
            ws.cell(row=row_idx, column=col_start + 4, value=wall.get('win1_qty'))
            ws.cell(row=row_idx, column=col_start + 5, value=wall.get('win2_name', ''))
            ws.cell(row=row_idx, column=col_start + 6, value=wall.get('win2_qty'))
            ws.cell(row=row_idx, column=col_start + 7, value=wall.get('door_name', ''))
            ws.cell(row=row_idx, column=col_start + 8, value=wall.get('door_qty'))

        # ROOFS (cols 123-146: 4 roofs x 6 cols)
        for r_idx, roof in enumerate(space.get('roofs', [])):
            col_start = 123 + r_idx * 6
            ws.cell(row=row_idx, column=col_start, value=roof.get('exposure'))
            ws.cell(row=row_idx, column=col_start + 1, value=round(roof['area'], 2) if roof.get('area') else None)
            ws.cell(row=row_idx, column=col_start + 2, value=roof.get('slope'))
            ws.cell(row=row_idx, column=col_start + 3, value=roof.get('type_name', ''))
            ws.cell(row=row_idx, column=col_start + 4, value=roof.get('sky_name', ''))
            ws.cell(row=row_idx, column=col_start + 5, value=roof.get('sky_qty'))

    # Actualizar sheet Tipos com os schedules do E3A
    if 'Tipos' in wb.sheetnames:
        ws_tipos = wb['Tipos']
        # Limpar schedules existentes (cols 13-14)
        for row in range(3, ws_tipos.max_row + 1):
            ws_tipos.cell(row=row, column=13).value = None
            ws_tipos.cell(row=row, column=14).value = None
        # Adicionar schedules do E3A
        for i, (sch_id, sch_name) in enumerate(sorted(schedules.items()), 3):
            ws_tipos.cell(row=i, column=13, value=sch_id)
            ws_tipos.cell(row=i, column=14, value=sch_name)

    # Preencher sheet Windows
    if 'Windows' in wb.sheetnames and windows:
        ws_win = wb['Windows']
        # Limpar dados existentes (linha 4 em diante)
        for row in range(4, ws_win.max_row + 1):
            for col in range(1, 6):
                ws_win.cell(row=row, column=col).value = None
        # Preencher com dados do E3A
        for i, (win_id, win_data) in enumerate(sorted(windows.items()), 4):
            ws_win.cell(row=i, column=1, value=win_data['name'])
            ws_win.cell(row=i, column=2, value=round(win_data['u_value'], 3) if win_data['u_value'] else None)
            ws_win.cell(row=i, column=3, value=round(win_data['shgc'], 3) if win_data['shgc'] else None)
            ws_win.cell(row=i, column=4, value=round(win_data['height'], 3) if win_data['height'] else None)
            ws_win.cell(row=i, column=5, value=round(win_data['width'], 3) if win_data['width'] else None)

    # Preencher sheet Walls
    if 'Walls' in wb.sheetnames and walls_types:
        ws_wal = wb['Walls']
        # Limpar dados existentes (linha 4 em diante)
        for row in range(4, ws_wal.max_row + 1):
            for col in range(1, 5):
                ws_wal.cell(row=row, column=col).value = None
        # Preencher com dados do E3A
        for i, (wal_id, wal_data) in enumerate(sorted(walls_types.items()), 4):
            ws_wal.cell(row=i, column=1, value=wal_data['name'])
            # U-Value, Peso, Espessura - se disponivel
            if 'u_value' in wal_data:
                ws_wal.cell(row=i, column=2, value=round(wal_data['u_value'], 3))
            if 'weight' in wal_data:
                ws_wal.cell(row=i, column=3, value=round(wal_data['weight'], 2))
            if 'thickness' in wal_data:
                ws_wal.cell(row=i, column=4, value=round(wal_data['thickness'], 3))

    # Preencher sheet Roofs
    if 'Roofs' in wb.sheetnames and roofs_types:
        ws_rof = wb['Roofs']
        # Limpar dados existentes (linha 4 em diante)
        for row in range(4, ws_rof.max_row + 1):
            for col in range(1, 5):
                ws_rof.cell(row=row, column=col).value = None
        # Preencher com dados do E3A
        for i, (rof_id, rof_data) in enumerate(sorted(roofs_types.items()), 4):
            ws_rof.cell(row=i, column=1, value=rof_data['name'])
            # U-Value, Peso, Espessura - se disponivel
            if 'u_value' in rof_data:
                ws_rof.cell(row=i, column=2, value=round(rof_data['u_value'], 3))
            if 'weight' in rof_data:
                ws_rof.cell(row=i, column=3, value=round(rof_data['weight'], 2))
            if 'thickness' in rof_data:
                ws_rof.cell(row=i, column=4, value=round(rof_data['thickness'], 3))

    wb.save(output_path)
    return len(spaces)


# =============================================================================
# ROTAS WEB
# =============================================================================

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HAP 5.1 Converter</title>
    <style>
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            min-height: 100vh;
            padding: 20px;
            color: #fff;
        }
        .container {
            max-width: 900px;
            margin: 0 auto;
        }
        header {
            text-align: center;
            margin-bottom: 40px;
        }
        h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            background: linear-gradient(90deg, #00d2ff, #3a7bd5);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .subtitle {
            color: #888;
            font-size: 1.1em;
        }
        .cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 30px;
            margin-bottom: 30px;
        }
        .card {
            background: rgba(255, 255, 255, 0.05);
            border-radius: 20px;
            padding: 30px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            transition: transform 0.3s, box-shadow 0.3s;
        }
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3);
        }
        .card h2 {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 20px;
            font-size: 1.4em;
        }
        .card.excel h2 { color: #00d26a; }
        .card.e3a h2 { color: #ff6b6b; }
        .icon {
            font-size: 1.5em;
        }
        .description {
            color: #aaa;
            margin-bottom: 25px;
            line-height: 1.6;
        }
        .upload-area {
            border: 2px dashed rgba(255, 255, 255, 0.2);
            border-radius: 15px;
            padding: 30px;
            text-align: center;
            margin-bottom: 20px;
            transition: border-color 0.3s, background 0.3s;
            cursor: pointer;
        }
        .upload-area:hover {
            border-color: rgba(255, 255, 255, 0.4);
            background: rgba(255, 255, 255, 0.02);
        }
        .upload-area.dragover {
            border-color: #00d2ff;
            background: rgba(0, 210, 255, 0.1);
        }
        .upload-icon {
            font-size: 3em;
            margin-bottom: 15px;
            opacity: 0.5;
        }
        .upload-text {
            color: #888;
        }
        .upload-text strong {
            color: #00d2ff;
        }
        input[type="file"] {
            display: none;
        }
        .file-name {
            margin-top: 10px;
            padding: 10px;
            background: rgba(0, 210, 255, 0.1);
            border-radius: 8px;
            color: #00d2ff;
            display: none;
        }
        .file-name.show {
            display: block;
        }
        button {
            width: 100%;
            padding: 15px 30px;
            font-size: 1.1em;
            font-weight: 600;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            transition: all 0.3s;
        }
        .card.excel button {
            background: linear-gradient(90deg, #00d26a, #00b359);
            color: white;
        }
        .card.e3a button {
            background: linear-gradient(90deg, #ff6b6b, #ee5a5a);
            color: white;
        }
        button:hover {
            transform: scale(1.02);
            box-shadow: 0 5px 20px rgba(0, 0, 0, 0.3);
        }
        button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none;
        }
        .messages {
            margin-bottom: 30px;
        }
        .alert {
            padding: 15px 20px;
            border-radius: 10px;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .alert.success {
            background: rgba(0, 210, 106, 0.2);
            border: 1px solid #00d26a;
            color: #00d26a;
        }
        .alert.error {
            background: rgba(255, 107, 107, 0.2);
            border: 1px solid #ff6b6b;
            color: #ff6b6b;
        }
        footer {
            text-align: center;
            color: #666;
            padding: 20px;
        }
        footer a {
            color: #00d2ff;
            text-decoration: none;
        }
        .info-box {
            background: rgba(0, 210, 255, 0.1);
            border: 1px solid rgba(0, 210, 255, 0.3);
            border-radius: 10px;
            padding: 20px;
            margin-top: 30px;
        }
        .info-box h3 {
            color: #00d2ff;
            margin-bottom: 10px;
        }
        .info-box ul {
            margin-left: 20px;
            color: #aaa;
        }
        .info-box li {
            margin: 5px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>HAP 5.1 Converter</h1>
            <p class="subtitle">Converta facilmente entre Excel e ficheiros HAP</p>
        </header>

        <div class="messages">
            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                        <div class="alert {{ category }}">
                            {% if category == 'success' %}&#10004;{% else %}&#10006;{% endif %}
                            {{ message }}
                        </div>
                    {% endfor %}
                {% endif %}
            {% endwith %}
        </div>

        <div class="cards">
            <div class="card excel">
                <h2><span class="icon">&#128196;</span> Excel → E3A</h2>
                <p class="description">
                    Carrega um ficheiro Excel (.xlsx) com os dados dos espacos e
                    obtem um ficheiro HAP 5.1 (.E3A) pronto a usar.
                </p>
                <form action="/excel-to-e3a" method="post" enctype="multipart/form-data" id="form-excel">
                    <div class="upload-area" onclick="document.getElementById('excel-file').click()">
                        <div class="upload-icon">&#128194;</div>
                        <p class="upload-text">Clica ou arrasta o ficheiro <strong>.xlsx</strong></p>
                        <input type="file" name="file" id="excel-file" accept=".xlsx,.xls" required>
                        <div class="file-name" id="excel-name"></div>
                    </div>
                    <button type="submit" id="btn-excel" disabled>Converter para E3A</button>
                </form>
            </div>

            <div class="card e3a">
                <h2><span class="icon">&#128230;</span> E3A → Excel</h2>
                <p class="description">
                    Carrega um ficheiro HAP 5.1 (.E3A) e exporta os dados dos
                    espacos para um ficheiro Excel (.xlsx).
                </p>
                <form action="/e3a-to-excel" method="post" enctype="multipart/form-data" id="form-e3a">
                    <div class="upload-area" onclick="document.getElementById('e3a-file').click()">
                        <div class="upload-icon">&#128230;</div>
                        <p class="upload-text">Clica ou arrasta o ficheiro <strong>.E3A</strong></p>
                        <input type="file" name="file" id="e3a-file" accept=".e3a,.E3A" required>
                        <div class="file-name" id="e3a-name"></div>
                    </div>
                    <button type="submit" id="btn-e3a" disabled>Converter para Excel</button>
                </form>
            </div>
        </div>

        <div class="info-box">
            <h3>Informacoes</h3>
            <ul>
                <li>Usa o template <strong>HAP_Template_RSECE.xlsx</strong> para criar novos projectos</li>
                <li>O modelo base inclui <strong>82 schedules RSECE</strong> pre-configurados</li>
                <li>Suporta todas as 27 tipologias do Anexo XV do RSECE</li>
                <li>Tamanho maximo: 50MB</li>
            </ul>
        </div>

        <footer>
            <p>HAP 5.1 Tools &copy; 2026 | <a href="/download-template">Descarregar Template</a></p>
        </footer>
    </div>

    <script>
        // File input handlers
        document.getElementById('excel-file').addEventListener('change', function() {
            const fileName = this.files[0]?.name || '';
            const nameDiv = document.getElementById('excel-name');
            const btn = document.getElementById('btn-excel');
            if (fileName) {
                nameDiv.textContent = fileName;
                nameDiv.classList.add('show');
                btn.disabled = false;
            } else {
                nameDiv.classList.remove('show');
                btn.disabled = true;
            }
        });

        document.getElementById('e3a-file').addEventListener('change', function() {
            const fileName = this.files[0]?.name || '';
            const nameDiv = document.getElementById('e3a-name');
            const btn = document.getElementById('btn-e3a');
            if (fileName) {
                nameDiv.textContent = fileName;
                nameDiv.classList.add('show');
                btn.disabled = false;
            } else {
                nameDiv.classList.remove('show');
                btn.disabled = true;
            }
        });

        // Drag and drop
        document.querySelectorAll('.upload-area').forEach(area => {
            area.addEventListener('dragover', e => {
                e.preventDefault();
                area.classList.add('dragover');
            });
            area.addEventListener('dragleave', e => {
                area.classList.remove('dragover');
            });
            area.addEventListener('drop', e => {
                e.preventDefault();
                area.classList.remove('dragover');
                const input = area.querySelector('input[type="file"]');
                input.files = e.dataTransfer.files;
                input.dispatchEvent(new Event('change'));
            });
        });
    </script>
</body>
</html>
'''

def allowed_file(filename, extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in extensions


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/excel-to-e3a', methods=['POST'])
def convert_excel_to_e3a():
    if 'file' not in request.files:
        flash('Nenhum ficheiro seleccionado', 'error')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('Nenhum ficheiro seleccionado', 'error')
        return redirect(url_for('index'))

    if not allowed_file(file.filename, ALLOWED_EXTENSIONS_EXCEL):
        flash('Tipo de ficheiro invalido. Use .xlsx ou .xls', 'error')
        return redirect(url_for('index'))

    try:
        # Guardar ficheiro temporario
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)

        # Converter
        output_filename = filename.rsplit('.', 1)[0] + '.E3A'
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        num_spaces = excel_to_e3a(input_path, output_path)

        # Limpar input
        os.remove(input_path)

        # Enviar ficheiro
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/octet-stream'
        )

    except Exception as e:
        flash(f'Erro na conversao: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/e3a-to-excel', methods=['POST'])
def convert_e3a_to_excel():
    if 'file' not in request.files:
        flash('Nenhum ficheiro seleccionado', 'error')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('Nenhum ficheiro seleccionado', 'error')
        return redirect(url_for('index'))

    if not allowed_file(file.filename, ALLOWED_EXTENSIONS_E3A):
        flash('Tipo de ficheiro invalido. Use .E3A', 'error')
        return redirect(url_for('index'))

    try:
        # Guardar ficheiro temporario
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)

        # Converter
        output_filename = filename.rsplit('.', 1)[0] + '_export.xlsx'
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        num_spaces = e3a_to_excel(input_path, output_path)

        # Limpar input
        os.remove(input_path)

        # Enviar ficheiro
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        flash(f'Erro na conversao: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/download-template')
def download_template():
    template_path = os.path.join(BASE_DIR, 'HAP_Template_RSECE.xlsx')
    if os.path.exists(template_path):
        return send_file(
            template_path,
            as_attachment=True,
            download_name='HAP_Template_RSECE.xlsx'
        )
    else:
        flash('Template nao encontrado', 'error')
        return redirect(url_for('index'))


if __name__ == '__main__':
    print("=" * 60)
    print("HAP 5.1 Web Converter")
    print("=" * 60)
    print()
    print("Abrir no browser: http://localhost:5000")
    print()
    print("Funcionalidades:")
    print("  - Excel -> E3A (usando Modelo_RSECE.E3A)")
    print("  - E3A -> Excel (exportacao)")
    print("  - Download do template")
    print()
    print("Ctrl+C para parar")
    print("=" * 60)

    app.run(debug=False, host='0.0.0.0', port=5000)
