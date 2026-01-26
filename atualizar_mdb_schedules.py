"""
Atualizar MDB com nomes dos Schedules RSECE
============================================

Atualiza a tabela ScheduleIndex no HAP51INX.MDB para que
os nomes dos schedules apareçam corretamente no HAP.

Usage:
    python atualizar_mdb_schedules.py <ficheiro.E3A>
"""

import zipfile
import struct
import sys
import os
import tempfile
import shutil

# Para acesso ao MDB no Windows
try:
    import pyodbc
    HAS_PYODBC = True
except ImportError:
    HAS_PYODBC = False

from hap_schedule_library import SCHEDULE_RECORD_SIZE


def ler_nomes_schedules_do_dat(e3a_path: str) -> list:
    """Lê os nomes dos schedules do HAP51SCH.DAT."""
    with zipfile.ZipFile(e3a_path, 'r') as zf:
        sch_data = zf.read('HAP51SCH.DAT')

    nomes = []
    num_schedules = len(sch_data) // SCHEDULE_RECORD_SIZE

    for i in range(num_schedules):
        offset = i * SCHEDULE_RECORD_SIZE
        nome = sch_data[offset:offset+24].decode('latin-1').rstrip('\x00')
        nomes.append(nome)

    return nomes


def atualizar_mdb_schedules(e3a_path: str):
    """Atualiza a tabela ScheduleIndex no MDB."""
    if not HAS_PYODBC:
        print("ERRO: pyodbc não está instalado.")
        print("Instala com: pip install pyodbc")
        print()
        print("Alternativa: atualização manual via script SQL")
        gerar_sql_alternativo(e3a_path)
        return

    # Ler nomes do DAT
    nomes = ler_nomes_schedules_do_dat(e3a_path)
    print(f"Schedules encontrados no DAT: {len(nomes)}")

    # Extrair MDB para pasta temporária
    temp_dir = tempfile.mkdtemp()
    mdb_temp_path = os.path.join(temp_dir, 'HAP51INX.MDB')

    try:
        with zipfile.ZipFile(e3a_path, 'r') as zf:
            files = {name: zf.read(name) for name in zf.namelist()}

        # Escrever MDB temporariamente
        with open(mdb_temp_path, 'wb') as f:
            f.write(files['HAP51INX.MDB'])

        # Conectar ao MDB
        conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_temp_path};'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        # Limpar tabela existente
        cursor.execute("DELETE FROM ScheduleIndex")

        # Inserir novos schedules
        for i, nome in enumerate(nomes):
            # nIndex é 1-based
            # nScheduleType: 0 = Fractional, 1 = On/Off
            cursor.execute(
                "INSERT INTO ScheduleIndex (nIndex, szName, nScheduleType) VALUES (?, ?, ?)",
                i + 1, nome, 0
            )
            print(f"  {i+1}. {nome}")

        conn.commit()
        cursor.close()
        conn.close()

        # Ler MDB atualizado
        with open(mdb_temp_path, 'rb') as f:
            files['HAP51INX.MDB'] = f.read()

        # Guardar E3A atualizado
        with zipfile.ZipFile(e3a_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, data in files.items():
                zf.writestr(name, data)

        print()
        print(f"MDB atualizado com sucesso!")
        print(f"Ficheiro: {e3a_path}")

    finally:
        # Limpar ficheiros temporários
        shutil.rmtree(temp_dir, ignore_errors=True)


def gerar_sql_alternativo(e3a_path: str):
    """Gera script SQL para atualização manual."""
    nomes = ler_nomes_schedules_do_dat(e3a_path)

    sql_file = e3a_path.replace('.E3A', '_schedules.sql')

    with open(sql_file, 'w', encoding='utf-8') as f:
        f.write("-- Script SQL para atualizar ScheduleIndex\n")
        f.write("-- Executar manualmente no Access ou via outra ferramenta\n")
        f.write("\n")
        f.write("DELETE FROM ScheduleIndex;\n")
        f.write("\n")

        for i, nome in enumerate(nomes):
            nome_escaped = nome.replace("'", "''")
            f.write(f"INSERT INTO ScheduleIndex (nIndex, szName, nScheduleType) VALUES ({i+1}, '{nome_escaped}', 0);\n")

    print(f"Script SQL gerado: {sql_file}")
    print()
    print("Para atualizar manualmente:")
    print("1. Extrair HAP51INX.MDB do ficheiro E3A (é um ZIP)")
    print("2. Abrir o MDB no Access")
    print("3. Executar o script SQL")
    print("4. Guardar e colocar de volta no ZIP")


def main():
    if len(sys.argv) < 2:
        print("Usage: python atualizar_mdb_schedules.py <ficheiro.E3A>")
        return

    e3a_path = sys.argv[1]

    if not os.path.exists(e3a_path):
        print(f"ERRO: Ficheiro não encontrado: {e3a_path}")
        return

    print("=" * 60)
    print("ATUALIZAR MDB COM NOMES DOS SCHEDULES")
    print("=" * 60)
    print()

    atualizar_mdb_schedules(e3a_path)


if __name__ == '__main__':
    main()
