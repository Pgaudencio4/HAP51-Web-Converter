# HAP 5.1 - Ferramentas de Automatizacao

Conjunto de ferramentas Python para automatizar a criacao de ficheiros HAP 5.1 (.E3A) a partir de templates Excel, com perfis RSECE pre-configurados.

---

## Aplicacao Web (NOVO!)

Para uma experiencia mais simples, use a aplicacao web:

```bash
python app.py
```

Abrir no browser: **http://localhost:5000**

Funcionalidades:
- Upload Excel -> Download E3A
- Upload E3A -> Download Excel
- Download do template

---

## Indice

1. [Aplicacao Web](#aplicacao-web-novo)
2. [Estrutura da Pasta](#estrutura-da-pasta)
3. [Workflow Principal](#workflow-principal)
4. [Ficheiros Principais](#ficheiros-principais)
5. [Scripts Utilitarios](#scripts-utilitarios)
6. [Template Excel - Detalhes](#template-excel---detalhes)
7. [Perfis RSECE](#perfis-rsece)
8. [Requisitos](#requisitos)
9. [Especificacoes Tecnicas](#especificacoes-tecnicas)

---

## Estrutura da Pasta

```
HAPPXXXX/
|
|-- FICHEIROS PRINCIPAIS
|   |-- Modelo_RSECE.E3A              # Modelo base com 82 schedules RSECE
|   |-- HAP_Template_RSECE.xlsx       # Template Excel com dropdowns RSECE
|   |-- README.md                     # Este ficheiro
|   |-- GUIA_RAPIDO.txt               # Referencia rapida (1 pagina)
|
|-- APLICACAO WEB
|   |-- app.py                        # [NOVO] Aplicacao web Flask
|
|-- SCRIPTS PRINCIPAIS
|   |-- excel_to_hap.py               # Converter Excel -> HAP (linha de comando)
|   |-- hap_to_excel.py               # Exportar HAP -> Excel
|   |-- criar_perfis_rsece.py         # Criar modelo com perfis RSECE
|   |-- criar_schedule.py             # Listar/criar schedules
|   |-- adicionar_dropdowns_rsece.py  # Adicionar dropdowns a Excel existente
|
|-- BIBLIOTECAS
|   |-- hap_library.py                # Biblioteca principal (estruturas binarias)
|   |-- hap_schedule_library.py       # Biblioteca de schedules
|
|-- OUTROS SCRIPTS
|   |-- criar_excel_template.py       # Gerador do template Excel
|   |-- atualizar_mdb_schedules.py    # Actualizar nomes no MDB
|
|-- _exemplos/                        # Ficheiros de exemplo
|   |-- HAP_Exemplo_5Espacos.xlsx     # Template Excel original
|   |-- Vale Formoso - Porto.E3A      # Exemplo de projecto real
|
|-- _documentacao/                    # Documentacao tecnica detalhada
|   |-- HAP_FILE_SPECIFICATION.md     # Especificacao completa do formato
|   |-- HAP_COMPLETE_FIELD_MAP.md     # Mapa de campos do registo Space
|
|-- _desenvolvimento/                 # Scripts de debug/desenvolvimento
|-- _testes/                          # Ficheiros de teste
```

---

## Workflow Principal

### Passo 1: Preparar o Excel

Copiar `HAP_Template_RSECE.xlsx` para um novo ficheiro ou usar directamente.

### Passo 2: Preencher Dados

Na sheet **"Espacos"** (linha 4 em diante):
- Preencher dados dos espacos (nome, area, altura, etc.)
- Seleccionar schedules nos **dropdowns** (colunas K, P, R)
- Adicionar paredes, janelas e coberturas conforme necessario

Na sheet **"Tipos"** (opcional):
- Mapear nomes de Wall Types, Window Types, etc. para IDs

Nas sheets **"Windows"**, **"Walls"**, **"Roofs"** (opcional):
- Definir tipos de vaos, paredes e coberturas customizados

### Passo 3: Gerar o Ficheiro HAP

```bash
python excel_to_hap.py <excel.xlsx> Modelo_RSECE.E3A <output.E3A>
```

**Exemplo:**
```bash
python excel_to_hap.py MeuEdificio.xlsx Modelo_RSECE.E3A MeuEdificio.E3A
```

### Passo 4: Abrir no HAP 5.1

O ficheiro `.E3A` pode ser aberto directamente no Carrier HAP 5.1.

---

## Ficheiros Principais

### Modelo_RSECE.E3A

Modelo HAP com **82 schedules** pre-configurados segundo o RSECE (Anexo XV):

- 1 x Sample Schedule (default do HAP)
- 27 tipologias x 3 tipos = 81 schedules RSECE

**Tipologias incluidas:**

| Comercio | Hotelaria/Lazer | Servicos | Outros |
|----------|-----------------|----------|--------|
| Hipermercado | Hotel 4-5 Estrelas | Escritorio | Escola |
| Venda Grosso | Hotel 1-3 Estrelas | Banco Sede | Universidade |
| Supermercado | Cinema Teatro | Banco Filial | Saude Sem Intern |
| Centro Comercial | Discoteca | Comunicacoes | Saude Com Intern |
| Pequena Loja | Bingo Clube Social | Biblioteca | Prisao |
| Restaurante | Clube Desp Piscina | Museu Galeria | Tribunal Camara |
| Pastelaria | Clube Desportivo | | |
| Pronto-a-Comer | | | |

### HAP_Template_RSECE.xlsx

Template Excel com:

- **Sheet "Espacos"**: Dados principais dos espacos (147 colunas)
- **Sheet "Tipos"**: Mapeamento nome -> ID para tipos
- **Sheet "Windows"**: Definicao de tipos de vaos
- **Sheet "Walls"**: Definicao de tipos de paredes
- **Sheet "Roofs"**: Definicao de tipos de coberturas
- **Sheet "Schedules_RSECE"**: Lista de todos os schedules disponiveis

**Dropdowns automaticos** nas colunas:
- Col K (11): People Schedule
- Col P (16): Light Schedule
- Col R (18): Equipment Schedule

---

## Scripts Utilitarios

### excel_to_hap.py - Conversor Principal

```bash
python excel_to_hap.py <input.xlsx> <modelo.E3A> <output.E3A>
```

**Exemplo:**
```bash
python excel_to_hap.py Edificio_Teste.xlsx Modelo_RSECE.E3A Edificio.E3A
```

#### O que o conversor faz (passo a passo):

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  Excel Template │ --> │    Conversor    │ --> │  Ficheiro .E3A  │
│                 │     │    (Python)     │     │                 │
│ - Espacos       │     │                 │     │ - HAP51SPC.DAT  │
│ - Windows       │     │  + Modelo Base  │     │ - HAP51WIN.DAT  │
│ - Walls         │     │    (schedules)  │     │ - HAP51WAL.DAT  │
│ - Roofs         │     │                 │     │ - HAP51ROF.DAT  │
└─────────────────┘     └─────────────────┘     │ - HAP51INX.MDB  │
                                                └─────────────────┘
```

1. **Le o Excel:**
   - Sheet "Espacos": dados dos espacos (147 colunas)
   - Sheet "Windows": tipos de janela (nome, U-value, SHGC, dimensoes)
   - Sheet "Walls": tipos de parede (nome, U-value, peso, espessura)
   - Sheet "Roofs": tipos de cobertura (nome, U-value, peso, espessura)

2. **Cria tipos binarios:**
   - `HAP51WIN.DAT`: 555 bytes por window
   - `HAP51WAL.DAT`: copia template e modifica nome/U-value
   - `HAP51ROF.DAT`: copia template e modifica nome/U-value

3. **Cria espacos binarios:**
   - `HAP51SPC.DAT`: 682 bytes por espaco
   - Escreve schedules nos offsets correctos (594, 616, 660)
   - Associa paredes, janelas e coberturas

4. **Actualiza o MDB (HAP51INX.MDB):**
   - SpaceIndex: lista de espacos
   - WindowIndex, WallIndex, RoofIndex: tipos criados
   - Links: Space_Schedule_Links, Space_Wall_Links, etc.

5. **Gera o .E3A** (ficheiro ZIP com tudo)

#### Como os nomes dos schedules sao convertidos para IDs:

O conversor le automaticamente os schedules do modelo base e cria um mapeamento nome -> ID:

```python
# Exemplo interno do conversor:
types['schedules'] = {
    'Sample Schedule': 0,
    'Hipermercado Ocup': 1,
    'Hipermercado Ilum': 2,
    'Hipermercado Equip': 3,
    ...
    'Escritorio Ocup': 46,
    'Escritorio Ilum': 47,
    'Escritorio Equip': 48,
    ...
}
```

Quando o Excel tem `People Schedule = "Escritorio Ocup"`, o conversor:
1. Procura "Escritorio Ocup" no dicionario -> encontra ID 46
2. Escreve o valor 46 no offset 594 do registo binario do espaco

#### Funcionalidades completas:

- Le espacos do Excel (sheet "Espacos")
- Le e cria tipos de Windows/Walls/Roofs das respectivas sheets
- Le schedules do modelo base automaticamente
- Cria registos binarios de 682 bytes por espaco
- Actualiza MDB (SpaceIndex, WindowIndex, WallIndex, RoofIndex, links)
- Gera ficheiro E3A completo e funcional

### hap_to_excel.py - Exportador

```bash
python hap_to_excel.py <input.E3A> <output.xlsx>
```

Exporta dados de um ficheiro HAP existente para Excel.

### criar_schedule.py - Utilitario de Schedules

```bash
# Listar todos os schedules
python criar_schedule.py <ficheiro.E3A> --listar

# Ver detalhes de um schedule
python criar_schedule.py <ficheiro.E3A> --ver "Nome do Schedule"

# Criar novo schedule
python criar_schedule.py <ficheiro.E3A> --criar "Nome" --tipo escritorio
```

### criar_perfis_rsece.py - Gerador do Modelo RSECE

```bash
python criar_perfis_rsece.py <modelo_base.E3A> <output.E3A>
```

Cria o modelo com todos os 82 perfis RSECE. Ja foi usado para criar `Modelo_RSECE.E3A`.

### adicionar_dropdowns_rsece.py - Adicionar Dropdowns

```bash
python adicionar_dropdowns_rsece.py <input.xlsx> [output.xlsx]
```

Adiciona dropdowns RSECE a um Excel existente que nao os tenha.

---

## Template Excel - Detalhes

### Sheet "Espacos" - Colunas Principais

| Colunas | Categoria | Campos |
|---------|-----------|--------|
| 1-6 | GENERAL | Nome, Area, Altura, Peso, OA valor, OA unidade |
| 7-11 | PEOPLE | Ocupacao, Actividade, Sens W/pes, Lat W/pes, **Schedule** |
| 12-16 | LIGHTING | Task W, General W, Fixture, Ballast, **Schedule** |
| 17-18 | EQUIPMENT | W/m2, **Schedule** |
| 19-22 | MISC | Sens W, Lat W, Sens Sch, Lat Sch |
| 23-26 | INFILTRATION | Metodo, ACH (Clg, Htg, Energy) |
| 27-39 | FLOORS | Tipo, Area, U-value, etc. |
| 40-51 | PARTITIONS | Ceiling e Wall partitions |
| 52-123 | WALLS | 8 paredes x 9 campos cada |
| 124-147 | ROOFS | 4 coberturas x 6 campos cada |

### Estrutura das Paredes (9 campos por parede)

1. Exposure (N, NE, E, SE, S, SW, W, NW)
2. Area (m2)
3. Wall Type (nome)
4. Window 1 Type (nome)
5. Window 1 Qty
6. Window 2 Type (nome)
7. Window 2 Qty
8. Door Type (nome)
9. Door Qty

### Estrutura das Coberturas (6 campos por cobertura)

1. Exposure (N, NE, E, SE, S, SW, W, NW)
2. Area (m2)
3. Slope (graus)
4. Roof Type (nome)
5. Skylight Type (nome)
6. Skylight Qty

### Sheet "Windows"

| Coluna | Campo | Unidade |
|--------|-------|---------|
| A | Nome | - |
| B | U-Value | W/m2K |
| C | SHGC | 0-1 |
| D | Altura | m |
| E | Largura | m |

### Sheet "Walls"

| Coluna | Campo | Unidade |
|--------|-------|---------|
| A | Nome | - |
| B | U-Value | W/m2K |
| C | Peso | kg/m2 |
| D | Espessura | m |

### Sheet "Roofs"

| Coluna | Campo | Unidade |
|--------|-------|---------|
| A | Nome | - |
| B | U-Value | W/m2K |
| C | Peso | kg/m2 |
| D | Espessura | m |

---

## Perfis RSECE

Os perfis seguem o **Anexo XV do RSECE** (Diario da Republica, 4 Abril 2006).

### Estrutura de Cada Perfil

Cada tipologia tem **3 perfis**:
- **Ocup** - % de ocupacao por hora
- **Ilum** - % de iluminacao por hora
- **Equip** - % de equipamento por hora

Cada perfil tem **3 variantes**:
- Segunda a Sexta
- Sabados
- Domingos e Feriados

### Nomenclatura dos Schedules

Formato: `<Tipologia> <Tipo>`

Exemplos:
- `Escritorio Ocup`
- `Escritorio Ilum`
- `Escritorio Equip`
- `Hotel 4-5 Estrelas Ocup`
- `Saude Com Intern Ilum`

### Lista Completa de Tipologias (27)

1. Hipermercado
2. Venda Grosso
3. Supermercado
4. Centro Comercial
5. Pequena Loja
6. Restaurante
7. Pastelaria
8. Pronto-a-Comer
9. Hotel 4-5 Estrelas
10. Hotel 1-3 Estrelas
11. Cinema Teatro
12. Discoteca
13. Bingo Clube Social
14. Clube Desp Piscina
15. Clube Desportivo
16. Tribunal Camara
17. Prisao
18. Escola
19. Universidade
20. Escritorio
21. Banco Sede
22. Banco Filial
23. Comunicacoes
24. Biblioteca
25. Museu Galeria
26. Saude Sem Intern
27. Saude Com Intern

---

## Requisitos

### Python

- Python 3.8 ou superior

### Bibliotecas

```bash
pip install openpyxl flask
```

**Opcional** (para actualizar MDB directamente):
```bash
pip install pyodbc
```

Nota: pyodbc requer o Microsoft Access Database Engine instalado.

---

## Especificacoes Tecnicas

### Formato do Ficheiro .E3A

O ficheiro `.E3A` e um arquivo ZIP contendo:

| Ficheiro | Descricao | Tamanho Registo |
|----------|-----------|-----------------|
| HAP51SPC.DAT | Espacos | 682 bytes |
| HAP51SCH.DAT | Schedules | 792 bytes |
| HAP51WAL.DAT | Paredes | ~128 bytes |
| HAP51WIN.DAT | Janelas | 555 bytes |
| HAP51DOR.DAT | Portas | variavel |
| HAP51ROF.DAT | Coberturas | variavel |
| HAP51INX.MDB | Base de dados Access (indices) | - |

### Encoding

- Strings: Latin-1 (ISO-8859-1)
- Numeros: Little-endian IEEE 754 (float 32-bit)

### Conversoes de Unidades

| De (SI) | Para (IP) | Formula |
|---------|-----------|---------|
| m2 | ft2 | x 10.7639 |
| m | ft | x 3.28084 |
| kg/m2 | lb/ft2 | / 4.8824 |
| W/m2K | BTU/hr.ft2.F | / 5.678 |
| W | BTU/hr | x 3.412 |
| W/m2 | W/ft2 | / 10.764 |
| C | F | x 1.8 + 32 |

### Estrutura Binaria do Registo Space (682 bytes)

| Offset | Bytes | Campo |
|--------|-------|-------|
| 0-23 | 24 | Nome |
| 24-27 | 4 | Area (ft2) |
| 28-31 | 4 | Altura (ft) |
| 32-35 | 4 | Peso (lb/ft2) |
| 46-49 | 4 | OA (encoded) |
| 50-51 | 2 | OA Unit |
| 72-343 | 272 | 8 Walls (34 bytes cada) |
| 344-439 | 96 | 4 Roofs (24 bytes cada) |
| 440-465 | 26 | Ceiling Partition |
| 466-491 | 26 | Wall Partition |
| 492-541 | 50 | Floor |
| 554-571 | 18 | Infiltration |
| 580-595 | 16 | People |
| **594** | **2** | **People Schedule ID** |
| 600-617 | 18 | Lighting |
| **616** | **2** | **Light Schedule ID** |
| 632-645 | 14 | Misc |
| 656-661 | 6 | Equipment |
| **660** | **2** | **Equipment Schedule ID** |

### IMPORTANTE: Offsets dos Schedules nos Espacos

Os schedules sao associados aos espacos atraves de **indices** escritos em offsets especificos do registo binario de 682 bytes:

```
┌─────────────────────────────────────────────────────────────┐
│                    REGISTO SPACE (682 bytes)                │
├─────────────────────────────────────────────────────────────┤
│  Offset 594 (2 bytes) = People Schedule ID                  │
│  Offset 616 (2 bytes) = Light Schedule ID    <-- NAO e 614! │
│  Offset 660 (2 bytes) = Equipment Schedule ID               │
└─────────────────────────────────────────────────────────────┘
```

**ATENCAO:** O offset do Light Schedule e **616**, nao 614! Esta foi uma descoberta critica durante o desenvolvimento. O offset 614 tem outro proposito.

### Como os Schedule IDs funcionam

1. Os schedules sao armazenados no ficheiro `HAP51SCH.DAT` (792 bytes cada)
2. O primeiro schedule (index 0) e o "Sample Schedule" default
3. Os schedules RSECE comecam no index 1

**Exemplo para tipologia "Escritorio":**
- Index 46 = "Escritorio Ocup" (People)
- Index 47 = "Escritorio Ilum" (Light)
- Index 48 = "Escritorio Equip" (Equipment)

**No registo binario do espaco:**
```
Offset 594: 46 (0x2E 0x00) -> People usa "Escritorio Ocup"
Offset 616: 47 (0x2F 0x00) -> Light usa "Escritorio Ilum"
Offset 660: 48 (0x30 0x00) -> Equipment usa "Escritorio Equip"
```

### Estrutura de cada Wall Block (34 bytes)

Cada espaco pode ter ate 8 paredes (offsets 72-343):

| Offset Relativo | Bytes | Campo |
|-----------------|-------|-------|
| +0 | 2 | Direction (1=N, 5=E, 9=S, 13=W, etc.) |
| +2 | 4 | Area (ft2, float) |
| +6 | 2 | Wall Type ID |
| +8 | 2 | Window 1 Type ID |
| +10 | 2 | Window 1 Quantity |
| +12 | 2 | Window 2 Type ID |
| +14 | 2 | Window 2 Quantity |
| +16 | 2 | Door Type ID |
| +18 | 2 | Door Quantity |

**ATENCAO:** A quantidade da janela (Window Qty) esta no offset **+10**, nao +12!

### Estrutura Binaria do Registo Window (555 bytes)

| Offset | Bytes | Campo |
|--------|-------|-------|
| 0-254 | 255 | Nome |
| 257-260 | 4 | Altura (ft) |
| 261-264 | 4 | Largura (ft) |
| 269-272 | 4 | U-Value (BTU/hr.ft2.F) |
| 273-276 | 4 | SHGC |

---

## Exemplos

Ver pasta `_exemplos/` para ficheiros de referencia:

- `HAP_Exemplo_5Espacos.xlsx` - Template Excel original
- `Vale Formoso - Porto.E3A` - Exemplo de projecto real

### Exemplo Basico

```bash
# 1. Copiar template
cp HAP_Template_RSECE.xlsx MeuEdificio.xlsx

# 2. Preencher dados no Excel (abrir e editar)

# 3. Gerar ficheiro HAP
python excel_to_hap.py MeuEdificio.xlsx Modelo_RSECE.E3A MeuEdificio.E3A

# 4. Abrir no HAP 5.1
```

### Listar Schedules Disponiveis

```bash
python criar_schedule.py Modelo_RSECE.E3A --listar
```

---

## Contacto e Suporte

Desenvolvido para automatizacao de projectos RSECE com Carrier HAP 5.1.

Para documentacao tecnica adicional, ver pasta `_documentacao/`.
