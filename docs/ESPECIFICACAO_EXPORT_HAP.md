# ESPECIFICACAO TECNICA: Exportacao de Dados para HAP 5.1

## Objectivo

Este documento define as especificacoes tecnicas para **gerar codigo** que exporte dados de uma aplicacao para o formato Excel compativel com o conversor HAP 5.1.

O codigo gerado deve criar um ficheiro Excel que:
1. Siga exactamente a estrutura do template `HAP_Template_RSECE.xlsx`
2. Seja compativel com o conversor `excel_to_hap.py`
3. Produza ficheiros `.E3A` validos para o Carrier HAP 5.1

---

## FICHEIROS DE REFERENCIA

```
Pasta base: C:\Users\pedro\Downloads\Programas2\HAPPXXXX\

Template Excel:  HAP_Template_RSECE.xlsx   (estrutura a replicar)
Conversor:       excel_to_hap.py           (le o Excel gerado)
Modelo HAP:      Modelo_RSECE.E3A          (82 schedules RSECE)
```

---

## ESTRUTURA DO EXCEL A GERAR

### Sheets Obrigatorias

O Excel gerado DEVE ter exactamente estas 7 sheets:

```javascript
const SHEETS_OBRIGATORIAS = [
  'Espacos',         // Dados dos espacos (principal)
  'Windows',         // Tipos de janelas
  'Walls',           // Tipos de paredes
  'Roofs',           // Tipos de coberturas
  'Tipos',           // Mapeamento nome -> ID (pode ficar vazio)
  'Legenda',         // Descricao campos (pode ficar vazio)
  'Schedules_RSECE'  // Lista de schedules (pode ficar vazio)
];
```

---

## SHEET: Espacos

### Estrutura de Linhas

```
Linha 1: Categorias (merged cells)
Linha 2: Subcategorias
Linha 3: Headers das colunas (nomes dos campos)
Linha 4+: DADOS (uma linha por espaco)
```

### Total: 147 Colunas

O codigo DEVE gerar exactamente 147 colunas na ordem especificada abaixo.

### LINHA 1 - Categorias (onde colocar cada uma)

```javascript
const LINHA1_CATEGORIAS = {
  1: 'GENERAL',      // Colunas 1-6
  7: 'INTERNALS',    // Colunas 7-22
  23: 'INFILTRATION', // Colunas 23-26
  27: 'FLOORS',      // Colunas 27-39
  40: 'PARTITIONS',  // Colunas 40-51
  52: 'WALLS',       // Colunas 52-123
  124: 'ROOFS'       // Colunas 124-147
};
// Restantes colunas ficam vazias (null)
```

### LINHA 2 - Subcategorias (onde colocar cada uma)

```javascript
const LINHA2_SUBCATEGORIAS = {
  7: 'PEOPLE',       // Colunas 7-11
  12: 'LIGHTING',    // Colunas 12-16
  17: 'EQUIPMENT',   // Colunas 17-18
  19: 'MISC',        // Colunas 19-22
  40: 'CEILING',     // Colunas 40-45
  46: 'WALL',        // Colunas 46-51
  52: 'WALL 1',      // Colunas 52-60
  61: 'WALL 2',      // Colunas 61-69
  70: 'WALL 3',      // Colunas 70-78
  79: 'WALL 4',      // Colunas 79-87
  88: 'WALL 5',      // Colunas 88-96
  97: 'WALL 6',      // Colunas 97-105
  106: 'WALL 7',     // Colunas 106-114
  115: 'WALL 8',     // Colunas 115-123
  124: 'ROOF 1',     // Colunas 124-129
  130: 'ROOF 2',     // Colunas 130-135
  136: 'ROOF 3',     // Colunas 136-141
  142: 'ROOF 4'      // Colunas 142-147
};
// Restantes colunas ficam vazias (null)
```

### LINHA 3 - Headers (TODAS as 147 colunas)

```javascript
const LINHA3_HEADERS = [
  // GENERAL (1-6)
  'Space Name', 'Floor Area\n(m2)', 'Avg Ceiling Ht\n(m)', 'Building Wt\n(kg/m2)', 'Outdoor Air\n(valor)', 'OA Unit',
  // PEOPLE (7-11)
  'Occupancy\n(people)', 'Activity Level', 'Sensible\n(W/person)', 'Latent\n(W/person)', 'Schedule',
  // LIGHTING (12-16)
  'Task Lighting\n(W)', 'General Ltg\n(W)', 'Fixture Type', 'Ballast Mult', 'Schedule',
  // EQUIPMENT (17-18)
  'Equipment\n(W/m2)', 'Schedule',
  // MISC (19-22)
  'Sensible\n(W)', 'Latent\n(W)', 'Sens Sch', 'Lat Sch',
  // INFILTRATION (23-26)
  'Infil Method', 'Design Clg\n(ACH)', 'Design Htg\n(ACH)', 'Energy\n(ACH)',
  // FLOORS (27-39)
  'Floor Type', 'Floor Area\n(m2)', 'U-Value\n(W/m2K)', 'Exp Perim\n(m)', 'Edge R\n(m2K/W)', 'Depth\n(m)',
  'Bsmt Wall U\n(W/m2K)', 'Wall Ins R\n(m2K/W)', 'Ins Depth\n(m)', 'Unc Max\n(C)', 'Out Max\n(C)', 'Unc Min\n(C)', 'Out Min\n(C)',
  // PARTITIONS CEILING (40-45)
  'Area\n(m2)', 'U-Value\n(W/m2K)', 'Unc Max\n(C)', 'Out Max\n(C)', 'Unc Min\n(C)', 'Out Min\n(C)',
  // PARTITIONS WALL (46-51)
  'Area\n(m2)', 'U-Value\n(W/m2K)', 'Unc Max\n(C)', 'Out Max\n(C)', 'Unc Min\n(C)', 'Out Min\n(C)',
  // WALL 1 (52-60)
  'Exposure', 'Gross Area\n(m2)', 'Wall Type', 'Window 1', 'Win1 Qty', 'Window 2', 'Win2 Qty', 'Door', 'Door Qty',
  // WALL 2 (61-69)
  'Exposure', 'Gross Area\n(m2)', 'Wall Type', 'Window 1', 'Win1 Qty', 'Window 2', 'Win2 Qty', 'Door', 'Door Qty',
  // WALL 3 (70-78)
  'Exposure', 'Gross Area\n(m2)', 'Wall Type', 'Window 1', 'Win1 Qty', 'Window 2', 'Win2 Qty', 'Door', 'Door Qty',
  // WALL 4 (79-87)
  'Exposure', 'Gross Area\n(m2)', 'Wall Type', 'Window 1', 'Win1 Qty', 'Window 2', 'Win2 Qty', 'Door', 'Door Qty',
  // WALL 5 (88-96)
  'Exposure', 'Gross Area\n(m2)', 'Wall Type', 'Window 1', 'Win1 Qty', 'Window 2', 'Win2 Qty', 'Door', 'Door Qty',
  // WALL 6 (97-105)
  'Exposure', 'Gross Area\n(m2)', 'Wall Type', 'Window 1', 'Win1 Qty', 'Window 2', 'Win2 Qty', 'Door', 'Door Qty',
  // WALL 7 (106-114)
  'Exposure', 'Gross Area\n(m2)', 'Wall Type', 'Window 1', 'Win1 Qty', 'Window 2', 'Win2 Qty', 'Door', 'Door Qty',
  // WALL 8 (115-123)
  'Exposure', 'Gross Area\n(m2)', 'Wall Type', 'Window 1', 'Win1 Qty', 'Window 2', 'Win2 Qty', 'Door', 'Door Qty',
  // ROOF 1 (124-129)
  'Exposure', 'Gross Area\n(m2)', 'Slope\n(deg)', 'Roof Type', 'Skylight', 'Sky Qty',
  // ROOF 2 (130-135)
  'Exposure', 'Gross Area\n(m2)', 'Slope\n(deg)', 'Roof Type', 'Skylight', 'Sky Qty',
  // ROOF 3 (136-141)
  'Exposure', 'Gross Area\n(m2)', 'Slope\n(deg)', 'Roof Type', 'Skylight', 'Sky Qty',
  // ROOF 4 (142-147)
  'Exposure', 'Gross Area\n(m2)', 'Slope\n(deg)', 'Roof Type', 'Skylight', 'Sky Qty'
];
// Total: 147 headers
```

### Funcao para gerar as 3 linhas de header

```javascript
function gerarHeadersEspacos() {
  // Linha 1 - Categorias (147 colunas, maioria vazia)
  const linha1 = new Array(147).fill(null);
  linha1[0] = 'GENERAL';
  linha1[6] = 'INTERNALS';
  linha1[22] = 'INFILTRATION';
  linha1[26] = 'FLOORS';
  linha1[39] = 'PARTITIONS';
  linha1[51] = 'WALLS';
  linha1[123] = 'ROOFS';

  // Linha 2 - Subcategorias (147 colunas, maioria vazia)
  const linha2 = new Array(147).fill(null);
  linha2[6] = 'PEOPLE';
  linha2[11] = 'LIGHTING';
  linha2[16] = 'EQUIPMENT';
  linha2[18] = 'MISC';
  linha2[39] = 'CEILING';
  linha2[45] = 'WALL';
  linha2[51] = 'WALL 1';
  linha2[60] = 'WALL 2';
  linha2[69] = 'WALL 3';
  linha2[78] = 'WALL 4';
  linha2[87] = 'WALL 5';
  linha2[96] = 'WALL 6';
  linha2[105] = 'WALL 7';
  linha2[114] = 'WALL 8';
  linha2[123] = 'ROOF 1';
  linha2[129] = 'ROOF 2';
  linha2[135] = 'ROOF 3';
  linha2[141] = 'ROOF 4';

  // Linha 3 - Headers (usar LINHA3_HEADERS acima)
  const linha3 = LINHA3_HEADERS;

  return { linha1, linha2, linha3 };
}
```

---

## MAPEAMENTO DE COLUNAS - SHEET ESPACOS

### GENERAL (Colunas 1-6)

```javascript
const GENERAL = {
  1:  { nome: 'Space Name',      tipo: 'string',  unidade: null,    obrigatorio: true,  maxLength: 24 },
  2:  { nome: 'Floor Area',      tipo: 'number',  unidade: 'm2',    obrigatorio: true },
  3:  { nome: 'Avg Ceiling Ht',  tipo: 'number',  unidade: 'm',     obrigatorio: true },
  4:  { nome: 'Building Wt',     tipo: 'number',  unidade: 'kg/m2', obrigatorio: false, default: 200 },
  5:  { nome: 'Outdoor Air',     tipo: 'number',  unidade: 'valor', obrigatorio: true },
  6:  { nome: 'OA Unit',         tipo: 'string',  unidade: null,    obrigatorio: true,  valores: ['L/s', 'L/s/m2', 'L/s/person', '%'] }
};
```

### INTERNALS - PEOPLE (Colunas 7-11)

```javascript
const PEOPLE = {
  7:  { nome: 'Occupancy',       tipo: 'number',  unidade: 'pessoas',   obrigatorio: false },
  8:  { nome: 'Activity Level',  tipo: 'string',  unidade: null,        obrigatorio: false, valores: ACTIVITY_LEVELS },
  9:  { nome: 'Sensible',        tipo: 'number',  unidade: 'W/pessoa',  obrigatorio: false },
  10: { nome: 'Latent',          tipo: 'number',  unidade: 'W/pessoa',  obrigatorio: false },
  11: { nome: 'Schedule',        tipo: 'string',  unidade: null,        obrigatorio: true,  valores: SCHEDULES_OCUP }
};

const ACTIVITY_LEVELS = [
  'Seated at Rest',
  'Office Work',
  'Sedentary Work',
  'Light Bench Work',
  'Medium Work',
  'Heavy Work',
  'Dancing',
  'Athletics'
];
```

### INTERNALS - LIGHTING (Colunas 12-16)

```javascript
const LIGHTING = {
  12: { nome: 'Task Lighting',   tipo: 'number',  unidade: 'W',    obrigatorio: false },
  13: { nome: 'General Ltg',     tipo: 'number',  unidade: 'W',    obrigatorio: true },
  14: { nome: 'Fixture Type',    tipo: 'string',  unidade: null,   obrigatorio: false, valores: FIXTURE_TYPES },
  15: { nome: 'Ballast Mult',    tipo: 'number',  unidade: null,   obrigatorio: false, default: 1.0 },
  16: { nome: 'Schedule',        tipo: 'string',  unidade: null,   obrigatorio: true,  valores: SCHEDULES_ILUM }
};

const FIXTURE_TYPES = [
  'Recessed Unvented',
  'Vented to Return Air',
  'Vented to Supply & Return',
  'Surface Mount/Pendant'
];
```

### INTERNALS - EQUIPMENT (Colunas 17-18)

```javascript
const EQUIPMENT = {
  17: { nome: 'Equipment',       tipo: 'number',  unidade: 'W/m2', obrigatorio: false },
  18: { nome: 'Schedule',        tipo: 'string',  unidade: null,   obrigatorio: true,  valores: SCHEDULES_EQUIP }
};
```

### INTERNALS - MISC (Colunas 19-22)

```javascript
const MISC = {
  19: { nome: 'Sensible',        tipo: 'number',  unidade: 'W',    obrigatorio: false },
  20: { nome: 'Latent',          tipo: 'number',  unidade: 'W',    obrigatorio: false },
  21: { nome: 'Sens Sch',        tipo: 'string',  unidade: null,   obrigatorio: false },
  22: { nome: 'Lat Sch',         tipo: 'string',  unidade: null,   obrigatorio: false }
};
```

### INFILTRATION (Colunas 23-26)

```javascript
const INFILTRATION = {
  23: { nome: 'Infil Method',    tipo: 'string',  unidade: null,   obrigatorio: false, valores: ['Air Change'] },
  24: { nome: 'Design Clg',      tipo: 'number',  unidade: 'ACH',  obrigatorio: false },
  25: { nome: 'Design Htg',      tipo: 'number',  unidade: 'ACH',  obrigatorio: false },
  26: { nome: 'Energy',          tipo: 'number',  unidade: 'ACH',  obrigatorio: false }
};
```

### FLOORS (Colunas 27-39)

```javascript
const FLOORS = {
  27: { nome: 'Floor Type',      tipo: 'string',  unidade: null,     obrigatorio: false, valores: FLOOR_TYPES },
  28: { nome: 'Floor Area',      tipo: 'number',  unidade: 'm2',     obrigatorio: false },
  29: { nome: 'U-Value',         tipo: 'number',  unidade: 'W/m2K',  obrigatorio: false },
  30: { nome: 'Exp Perim',       tipo: 'number',  unidade: 'm',      obrigatorio: false },
  31: { nome: 'Edge R',          tipo: 'number',  unidade: 'm2K/W',  obrigatorio: false },
  32: { nome: 'Depth',           tipo: 'number',  unidade: 'm',      obrigatorio: false },
  33: { nome: 'Bsmt Wall U',     tipo: 'number',  unidade: 'W/m2K',  obrigatorio: false },
  34: { nome: 'Wall Ins R',      tipo: 'number',  unidade: 'm2K/W',  obrigatorio: false },
  35: { nome: 'Ins Depth',       tipo: 'number',  unidade: 'm',      obrigatorio: false },
  36: { nome: 'Unc Max',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  37: { nome: 'Out Max',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  38: { nome: 'Unc Min',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  39: { nome: 'Out Min',         tipo: 'number',  unidade: 'C',      obrigatorio: false }
};

const FLOOR_TYPES = [
  'Floor Above Cond Space',
  'Floor Above Uncond Space',
  'Slab Floor On Grade',
  'Slab Floor Below Grade'
];
```

### PARTITIONS - CEILING (Colunas 40-45)

```javascript
const PARTITIONS_CEILING = {
  40: { nome: 'Area',            tipo: 'number',  unidade: 'm2',     obrigatorio: false },
  41: { nome: 'U-Value',         tipo: 'number',  unidade: 'W/m2K',  obrigatorio: false },
  42: { nome: 'Unc Max',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  43: { nome: 'Out Max',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  44: { nome: 'Unc Min',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  45: { nome: 'Out Min',         tipo: 'number',  unidade: 'C',      obrigatorio: false }
};
```

### PARTITIONS - WALL (Colunas 46-51)

```javascript
const PARTITIONS_WALL = {
  46: { nome: 'Area',            tipo: 'number',  unidade: 'm2',     obrigatorio: false },
  47: { nome: 'U-Value',         tipo: 'number',  unidade: 'W/m2K',  obrigatorio: false },
  48: { nome: 'Unc Max',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  49: { nome: 'Out Max',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  50: { nome: 'Unc Min',         tipo: 'number',  unidade: 'C',      obrigatorio: false },
  51: { nome: 'Out Min',         tipo: 'number',  unidade: 'C',      obrigatorio: false }
};
```

### WALLS - 8 Paredes (Colunas 52-123)

Cada parede ocupa **9 colunas**. Total: 8 paredes x 9 colunas = 72 colunas.

```javascript
const WALL_COLUMNS_PER_WALL = 9;
const NUM_WALLS = 8;

// Estrutura de cada parede (offsets relativos)
const WALL_STRUCTURE = {
  0: { nome: 'Exposure',    tipo: 'string',  unidade: null,  valores: EXPOSURES },
  1: { nome: 'Gross Area',  tipo: 'number',  unidade: 'm2',  nota: 'AREA, nao comprimento!' },
  2: { nome: 'Wall Type',   tipo: 'string',  unidade: null,  nota: 'Nome da sheet Walls' },
  3: { nome: 'Window 1',    tipo: 'string',  unidade: null,  nota: 'Nome da sheet Windows' },
  4: { nome: 'Win1 Qty',    tipo: 'number',  unidade: null },
  5: { nome: 'Window 2',    tipo: 'string',  unidade: null },
  6: { nome: 'Win2 Qty',    tipo: 'number',  unidade: null },
  7: { nome: 'Door',        tipo: 'string',  unidade: null },
  8: { nome: 'Door Qty',    tipo: 'number',  unidade: null }
};

// Colunas de cada parede
const WALL_COLUMNS = {
  1: { start: 52,  end: 60 },   // Wall 1: colunas 52-60
  2: { start: 61,  end: 69 },   // Wall 2: colunas 61-69
  3: { start: 70,  end: 78 },   // Wall 3: colunas 70-78
  4: { start: 79,  end: 87 },   // Wall 4: colunas 79-87
  5: { start: 88,  end: 96 },   // Wall 5: colunas 88-96
  6: { start: 97,  end: 105 },  // Wall 6: colunas 97-105
  7: { start: 106, end: 114 },  // Wall 7: colunas 106-114
  8: { start: 115, end: 123 }   // Wall 8: colunas 115-123
};

const EXPOSURES = [
  'N', 'NNE', 'NE', 'ENE',
  'E', 'ESE', 'SE', 'SSE',
  'S', 'SSW', 'SW', 'WSW',
  'W', 'WNW', 'NW', 'NNW'
];
```

**IMPORTANTE - Calculo de Gross Area:**
```javascript
// Gross Area = comprimento da parede × pe direito
// NAO e o comprimento da parede!
function calcularGrossArea(comprimentoParede, peDireito) {
  return comprimentoParede * peDireito;  // resultado em m2
}
```

### ROOFS - 4 Coberturas (Colunas 124-147)

Cada cobertura ocupa **6 colunas**. Total: 4 coberturas x 6 colunas = 24 colunas.

```javascript
const ROOF_COLUMNS_PER_ROOF = 6;
const NUM_ROOFS = 4;

// Estrutura de cada cobertura (offsets relativos)
const ROOF_STRUCTURE = {
  0: { nome: 'Exposure',    tipo: 'string',  unidade: null,    valores: [...EXPOSURES, 'H', 'HORIZ'] },
  1: { nome: 'Gross Area',  tipo: 'number',  unidade: 'm2' },
  2: { nome: 'Slope',       tipo: 'number',  unidade: 'graus', nota: '0 para horizontal' },
  3: { nome: 'Roof Type',   tipo: 'string',  unidade: null,    nota: 'Nome da sheet Roofs' },
  4: { nome: 'Skylight',    tipo: 'string',  unidade: null,    nota: 'Nome da sheet Windows' },
  5: { nome: 'Sky Qty',     tipo: 'number',  unidade: null }
};

// Colunas de cada cobertura
const ROOF_COLUMNS = {
  1: { start: 124, end: 129 },  // Roof 1: colunas 124-129
  2: { start: 130, end: 135 },  // Roof 2: colunas 130-135
  3: { start: 136, end: 141 },  // Roof 3: colunas 136-141
  4: { start: 142, end: 147 }   // Roof 4: colunas 142-147
};
```

---

## SHEET: Windows

### Estrutura

```
Linha 1-2: Headers de categoria (opcional, pode deixar vazio)
Linha 3:   Headers das colunas
Linha 4+:  DADOS (uma janela por linha)
```

### Colunas (5 colunas)

```javascript
const WINDOWS_COLUMNS = {
  1: { nome: 'Nome',     tipo: 'string',  unidade: null,    obrigatorio: true,  nota: 'Nome unico, usado em Espacos' },
  2: { nome: 'U-Value',  tipo: 'number',  unidade: 'W/m2K', obrigatorio: true },
  3: { nome: 'SHGC',     tipo: 'number',  unidade: '0-1',   obrigatorio: true,  min: 0, max: 1 },
  4: { nome: 'Altura',   tipo: 'number',  unidade: 'm',     obrigatorio: true },
  5: { nome: 'Largura',  tipo: 'number',  unidade: 'm',     obrigatorio: true }
};
```

---

## SHEET: Walls

### Estrutura

```
Linha 1-2: Headers de categoria (opcional, pode deixar vazio)
Linha 3:   Headers das colunas
Linha 4+:  DADOS (um tipo por linha)
```

### Colunas (4 colunas)

```javascript
const WALLS_COLUMNS = {
  1: { nome: 'Nome',      tipo: 'string',  unidade: null,    obrigatorio: true,  nota: 'Nome unico, usado em Espacos' },
  2: { nome: 'U-Value',   tipo: 'number',  unidade: 'W/m2K', obrigatorio: true },
  3: { nome: 'Peso',      tipo: 'number',  unidade: 'kg/m2', obrigatorio: false, default: 200 },
  4: { nome: 'Espessura', tipo: 'number',  unidade: 'm',     obrigatorio: false, default: 0.3 }
};
```

---

## SHEET: Roofs

### Estrutura

```
Linha 1-2: Headers de categoria (opcional, pode deixar vazio)
Linha 3:   Headers das colunas
Linha 4+:  DADOS (um tipo por linha)
```

### Colunas (4 colunas)

```javascript
const ROOFS_COLUMNS = {
  1: { nome: 'Nome',      tipo: 'string',  unidade: null,    obrigatorio: true,  nota: 'Nome unico, usado em Espacos' },
  2: { nome: 'U-Value',   tipo: 'number',  unidade: 'W/m2K', obrigatorio: true },
  3: { nome: 'Peso',      tipo: 'number',  unidade: 'kg/m2', obrigatorio: false, default: 300 },
  4: { nome: 'Espessura', tipo: 'number',  unidade: 'm',     obrigatorio: false, default: 0.35 }
};
```

---

## SCHEDULES RSECE - VALORES VALIDOS

O conversor so aceita schedules que existam no modelo `Modelo_RSECE.E3A`.

### Lista Completa (82 schedules)

```javascript
const SCHEDULES_RSECE = {
  // Default
  'Sample Schedule': 0,

  // Comercio
  'Hipermercado Ocup': 1, 'Hipermercado Ilum': 2, 'Hipermercado Equip': 3,
  'Venda Grosso Ocup': 4, 'Venda Grosso Ilum': 5, 'Venda Grosso Equip': 6,
  'Supermercado Ocup': 7, 'Supermercado Ilum': 8, 'Supermercado Equip': 9,
  'Centro Comercial Ocup': 10, 'Centro Comercial Ilum': 11, 'Centro Comercial Equip': 12,
  'Pequena Loja Ocup': 13, 'Pequena Loja Ilum': 14, 'Pequena Loja Equip': 15,
  'Restaurante Ocup': 16, 'Restaurante Ilum': 17, 'Restaurante Equip': 18,
  'Pastelaria Ocup': 19, 'Pastelaria Ilum': 20, 'Pastelaria Equip': 21,
  'Pronto-a-Comer Ocup': 22, 'Pronto-a-Comer Ilum': 23, 'Pronto-a-Comer Equip': 24,

  // Hotelaria
  'Hotel 4-5 Estrelas Ocup': 25, 'Hotel 4-5 Estrelas Ilum': 26, 'Hotel 4-5 Estrelas Equip': 27,
  'Hotel 1-3 Estrelas Ocup': 28, 'Hotel 1-3 Estrelas Ilum': 29, 'Hotel 1-3 Estrelas Equip': 30,

  // Lazer
  'Cinema Teatro Ocup': 31, 'Cinema Teatro Ilum': 32, 'Cinema Teatro Equip': 33,
  'Discoteca Ocup': 34, 'Discoteca Ilum': 35, 'Discoteca Equip': 36,
  'Bingo Clube Social Ocup': 37, 'Bingo Clube Social Ilum': 38, 'Bingo Clube Social Equip': 39,
  'Clube Desp Piscina Ocup': 40, 'Clube Desp Piscina Ilum': 41, 'Clube Desp Piscina Equip': 42,
  'Clube Desportivo Ocup': 43, 'Clube Desportivo Ilum': 44, 'Clube Desportivo Equip': 45,

  // Servicos
  'Escritorio Ocup': 46, 'Escritorio Ilum': 47, 'Escritorio Equip': 48,
  'Banco Sede Ocup': 49, 'Banco Sede Ilum': 50, 'Banco Sede Equip': 51,
  'Banco Filial Ocup': 52, 'Banco Filial Ilum': 53, 'Banco Filial Equip': 54,
  'Comunicacoes Ocup': 55, 'Comunicacoes Ilum': 56, 'Comunicacoes Equip': 57,
  'Biblioteca Ocup': 58, 'Biblioteca Ilum': 59, 'Biblioteca Equip': 60,
  'Museu Galeria Ocup': 61, 'Museu Galeria Ilum': 62, 'Museu Galeria Equip': 63,

  // Publico
  'Tribunal Camara Ocup': 64, 'Tribunal Camara Ilum': 65, 'Tribunal Camara Equip': 66,
  'Prisao Ocup': 67, 'Prisao Ilum': 68, 'Prisao Equip': 69,

  // Educacao
  'Escola Ocup': 70, 'Escola Ilum': 71, 'Escola Equip': 72,
  'Universidade Ocup': 73, 'Universidade Ilum': 74, 'Universidade Equip': 75,

  // Saude
  'Saude Sem Intern Ocup': 76, 'Saude Sem Intern Ilum': 77, 'Saude Sem Intern Equip': 78,
  'Saude Com Intern Ocup': 79, 'Saude Com Intern Ilum': 80, 'Saude Com Intern Equip': 81
};

// Schedules por tipo (para facilitar seleccao)
const SCHEDULES_OCUP = Object.keys(SCHEDULES_RSECE).filter(s => s.endsWith(' Ocup') || s === 'Sample Schedule');
const SCHEDULES_ILUM = Object.keys(SCHEDULES_RSECE).filter(s => s.endsWith(' Ilum') || s === 'Sample Schedule');
const SCHEDULES_EQUIP = Object.keys(SCHEDULES_RSECE).filter(s => s.endsWith(' Equip') || s === 'Sample Schedule');
```

### Regras de Uso

```javascript
// CORRECTO: usar sufixo apropriado
{ peopleSchedule: 'Escritorio Ocup' }   // Ocupacao -> Ocup
{ lightSchedule: 'Escritorio Ilum' }    // Iluminacao -> Ilum
{ equipSchedule: 'Escritorio Equip' }   // Equipamento -> Equip

// ERRADO: nomes que NAO existem
{ peopleSchedule: 'Escritorios' }       // NAO EXISTE
{ peopleSchedule: 'Escritorio' }        // NAO EXISTE (falta sufixo)
{ peopleSchedule: 'Comercio Ocup' }     // NAO EXISTE
{ peopleSchedule: 'Habitacao Ocup' }    // NAO EXISTE
```

---

## VALIDACOES OBRIGATORIAS

O codigo de exportacao DEVE validar:

### 1. Campos Obrigatorios

```javascript
function validarEspaco(espaco) {
  const erros = [];

  // Campos obrigatorios
  if (!espaco.nome || espaco.nome.trim() === '') erros.push('Nome obrigatorio');
  if (espaco.nome && espaco.nome.length > 24) erros.push('Nome max 24 caracteres');
  if (!espaco.area || espaco.area <= 0) erros.push('Area obrigatoria e > 0');
  if (!espaco.altura || espaco.altura <= 0) erros.push('Altura obrigatoria e > 0');
  if (!espaco.outdoorAir) erros.push('Outdoor Air obrigatorio');
  if (!espaco.oaUnit) erros.push('OA Unit obrigatorio');

  return erros;
}
```

### 2. Valores Validos

```javascript
function validarValores(espaco) {
  const erros = [];

  // OA Unit
  if (espaco.oaUnit && !['L/s', 'L/s/m2', 'L/s/person', '%'].includes(espaco.oaUnit)) {
    erros.push(`OA Unit invalido: ${espaco.oaUnit}`);
  }

  // Schedules
  if (espaco.peopleSchedule && !SCHEDULES_RSECE[espaco.peopleSchedule]) {
    erros.push(`People Schedule invalido: ${espaco.peopleSchedule}`);
  }
  if (espaco.lightSchedule && !SCHEDULES_RSECE[espaco.lightSchedule]) {
    erros.push(`Light Schedule invalido: ${espaco.lightSchedule}`);
  }
  if (espaco.equipSchedule && !SCHEDULES_RSECE[espaco.equipSchedule]) {
    erros.push(`Equipment Schedule invalido: ${espaco.equipSchedule}`);
  }

  // Exposures
  espaco.paredes?.forEach((parede, i) => {
    if (parede.exposure && !EXPOSURES.includes(parede.exposure)) {
      erros.push(`Wall ${i+1} Exposure invalido: ${parede.exposure}`);
    }
  });

  return erros;
}
```

### 3. Referencias Cruzadas

```javascript
function validarReferencias(espacos, windows, walls, roofs) {
  const erros = [];

  const windowNames = new Set(windows.map(w => w.nome));
  const wallNames = new Set(walls.map(w => w.nome));
  const roofNames = new Set(roofs.map(r => r.nome));

  espacos.forEach(espaco => {
    espaco.paredes?.forEach((parede, i) => {
      if (parede.wallType && !wallNames.has(parede.wallType)) {
        erros.push(`${espaco.nome}: Wall Type '${parede.wallType}' nao existe na sheet Walls`);
      }
      if (parede.window1 && !windowNames.has(parede.window1)) {
        erros.push(`${espaco.nome}: Window '${parede.window1}' nao existe na sheet Windows`);
      }
    });

    espaco.coberturas?.forEach((cob, i) => {
      if (cob.roofType && !roofNames.has(cob.roofType)) {
        erros.push(`${espaco.nome}: Roof Type '${cob.roofType}' nao existe na sheet Roofs`);
      }
    });
  });

  return erros;
}
```

---

## EXEMPLO DE CODIGO DE EXPORTACAO

### Estrutura de Dados de Entrada

```javascript
const dadosParaExportar = {
  espacos: [
    {
      nome: 'Escritorio_01',
      area: 50,                      // m2
      altura: 2.8,                   // m
      peso: 200,                     // kg/m2
      outdoorAir: 250,               // valor
      oaUnit: 'L/s',                 // unidade

      // People
      occupancy: 10,
      activityLevel: 'Office Work',
      sensible: 70,                  // W/pessoa
      latent: 45,                    // W/pessoa
      peopleSchedule: 'Escritorio Ocup',

      // Lighting
      taskLighting: 0,               // W
      generalLighting: 500,          // W
      fixtureType: 'Recessed Unvented',
      ballastMult: 1.0,
      lightSchedule: 'Escritorio Ilum',

      // Equipment
      equipment: 15,                 // W/m2
      equipSchedule: 'Escritorio Equip',

      // Infiltration
      infilMethod: 'Air Change',
      achClg: 0.3,
      achHtg: 0.3,
      achEnergy: 0.3,

      // Paredes (array de 0-8 paredes)
      paredes: [
        {
          exposure: 'S',
          grossArea: 14,             // m2 (5m comprimento x 2.8m altura)
          wallType: 'Parede Exterior',
          window1: 'Janela_Duplo_01',
          win1Qty: 2,
          window2: null,
          win2Qty: 0,
          door: null,
          doorQty: 0
        }
      ],

      // Coberturas (array de 0-4 coberturas)
      coberturas: [
        {
          exposure: 'H',
          grossArea: 50,             // m2
          slope: 0,                  // graus
          roofType: 'Cobertura Plana',
          skylight: null,
          skyQty: 0
        }
      ]
    }
  ],

  windows: [
    { nome: 'Janela_Duplo_01', uValue: 2.8, shgc: 0.70, altura: 1.5, largura: 1.2 }
  ],

  walls: [
    { nome: 'Parede Exterior', uValue: 0.50, peso: 250, espessura: 0.35 }
  ],

  roofs: [
    { nome: 'Cobertura Plana', uValue: 0.40, peso: 350, espessura: 0.40 }
  ]
};
```

### Funcao de Exportacao (Pseudocodigo)

```javascript
function exportarParaExcel(dados) {
  const workbook = criarWorkbook();

  // 1. Criar sheet Espacos
  const sheetEspacos = workbook.addSheet('Espacos');

  // Linha 1: Categorias
  sheetEspacos.setRow(1, gerarCategorias());

  // Linha 2: Subcategorias
  sheetEspacos.setRow(2, gerarSubcategorias());

  // Linha 3: Headers
  sheetEspacos.setRow(3, gerarHeaders());

  // Linha 4+: Dados
  dados.espacos.forEach((espaco, index) => {
    const row = 4 + index;
    const rowData = espacoParaRow(espaco);  // Converte para array de 147 valores
    sheetEspacos.setRow(row, rowData);
  });

  // 2. Criar sheet Windows
  const sheetWindows = workbook.addSheet('Windows');
  sheetWindows.setRow(3, ['Nome', 'U-Value', 'SHGC', 'Altura', 'Largura']);
  dados.windows.forEach((win, index) => {
    sheetWindows.setRow(4 + index, [win.nome, win.uValue, win.shgc, win.altura, win.largura]);
  });

  // 3. Criar sheet Walls
  const sheetWalls = workbook.addSheet('Walls');
  sheetWalls.setRow(3, ['Nome', 'U-Value', 'Peso', 'Espessura']);
  dados.walls.forEach((wall, index) => {
    sheetWalls.setRow(4 + index, [wall.nome, wall.uValue, wall.peso, wall.espessura]);
  });

  // 4. Criar sheet Roofs
  const sheetRoofs = workbook.addSheet('Roofs');
  sheetRoofs.setRow(3, ['Nome', 'U-Value', 'Peso', 'Espessura']);
  dados.roofs.forEach((roof, index) => {
    sheetRoofs.setRow(4 + index, [roof.nome, roof.uValue, roof.peso, roof.espessura]);
  });

  // 5. Criar sheets vazias obrigatorias
  workbook.addSheet('Tipos');
  workbook.addSheet('Legenda');
  workbook.addSheet('Schedules_RSECE');

  return workbook;
}

function espacoParaRow(espaco) {
  const row = new Array(147).fill(null);

  // GENERAL (1-6)
  row[0] = espaco.nome;
  row[1] = espaco.area;
  row[2] = espaco.altura;
  row[3] = espaco.peso || 200;
  row[4] = espaco.outdoorAir;
  row[5] = espaco.oaUnit;

  // PEOPLE (7-11)
  row[6] = espaco.occupancy;
  row[7] = espaco.activityLevel;
  row[8] = espaco.sensible;
  row[9] = espaco.latent;
  row[10] = espaco.peopleSchedule;

  // LIGHTING (12-16)
  row[11] = espaco.taskLighting;
  row[12] = espaco.generalLighting;
  row[13] = espaco.fixtureType;
  row[14] = espaco.ballastMult || 1.0;
  row[15] = espaco.lightSchedule;

  // EQUIPMENT (17-18)
  row[16] = espaco.equipment;
  row[17] = espaco.equipSchedule;

  // MISC (19-22)
  row[18] = espaco.miscSensible;
  row[19] = espaco.miscLatent;
  row[20] = espaco.miscSensSch;
  row[21] = espaco.miscLatSch;

  // INFILTRATION (23-26)
  row[22] = espaco.infilMethod;
  row[23] = espaco.achClg;
  row[24] = espaco.achHtg;
  row[25] = espaco.achEnergy;

  // FLOORS (27-39) - 13 colunas
  row[26] = espaco.floorType;
  row[27] = espaco.floorArea;
  row[28] = espaco.floorU;
  row[29] = espaco.floorPerim;
  row[30] = espaco.floorEdgeR;
  row[31] = espaco.floorDepth;
  row[32] = espaco.bsmtWallU;
  row[33] = espaco.wallInsR;
  row[34] = espaco.insDepth;
  row[35] = espaco.floorUncMax;
  row[36] = espaco.floorOutMax;
  row[37] = espaco.floorUncMin;
  row[38] = espaco.floorOutMin;

  // PARTITIONS CEILING (40-45) - 6 colunas
  row[39] = espaco.ceilArea;
  row[40] = espaco.ceilU;
  row[41] = espaco.ceilUncMax;
  row[42] = espaco.ceilOutMax;
  row[43] = espaco.ceilUncMin;
  row[44] = espaco.ceilOutMin;

  // PARTITIONS WALL (46-51) - 6 colunas
  row[45] = espaco.wallPartArea;
  row[46] = espaco.wallPartU;
  row[47] = espaco.wallUncMax;
  row[48] = espaco.wallOutMax;
  row[49] = espaco.wallUncMin;
  row[50] = espaco.wallOutMin;

  // WALLS (52-123) - 8 paredes x 9 colunas
  for (let w = 0; w < 8; w++) {
    const baseCol = 51 + (w * 9);  // 51, 60, 69, 78, 87, 96, 105, 114
    const parede = espaco.paredes?.[w] || {};

    row[baseCol + 0] = parede.exposure;
    row[baseCol + 1] = parede.grossArea;
    row[baseCol + 2] = parede.wallType;
    row[baseCol + 3] = parede.window1;
    row[baseCol + 4] = parede.win1Qty;
    row[baseCol + 5] = parede.window2;
    row[baseCol + 6] = parede.win2Qty;
    row[baseCol + 7] = parede.door;
    row[baseCol + 8] = parede.doorQty;
  }

  // ROOFS (124-147) - 4 coberturas x 6 colunas
  for (let r = 0; r < 4; r++) {
    const baseCol = 123 + (r * 6);  // 123, 129, 135, 141
    const cobertura = espaco.coberturas?.[r] || {};

    row[baseCol + 0] = cobertura.exposure;
    row[baseCol + 1] = cobertura.grossArea;
    row[baseCol + 2] = cobertura.slope;
    row[baseCol + 3] = cobertura.roofType;
    row[baseCol + 4] = cobertura.skylight;
    row[baseCol + 5] = cobertura.skyQty;
  }

  return row;
}
```

---

## RESUMO - CHECKLIST PARA O CODIGO

O codigo de exportacao DEVE:

- [ ] Criar Excel com exactamente 7 sheets: Espacos, Windows, Walls, Roofs, Tipos, Legenda, Schedules_RSECE
- [ ] Sheet Espacos: 3 linhas de header + dados a partir da linha 4
- [ ] Sheet Espacos: exactamente 147 colunas na ordem especificada
- [ ] Sheets Windows/Walls/Roofs: headers na linha 3, dados a partir da linha 4
- [ ] Validar campos obrigatorios: nome, area, altura, outdoorAir, oaUnit
- [ ] Validar nome do espaco max 24 caracteres
- [ ] Validar schedules existem na lista SCHEDULES_RSECE
- [ ] Validar referencias: window/wall/roof types existem nas respectivas sheets
- [ ] Calcular Gross Area = comprimento × pe direito (NAO usar comprimento directamente)
- [ ] Usar valores exactos para campos enumerados (OA Unit, Exposure, Activity Level, etc.)

---

## CONVERSAO FINAL

Apos gerar o Excel, converter para HAP:

```bash
cd C:\Users\pedro\Downloads\Programas2\HAPPXXXX
python excel_to_hap.py <excel_gerado>.xlsx Modelo_RSECE.E3A <output>.E3A
```

O ficheiro `Modelo_RSECE.E3A` e OBRIGATORIO pois contem os schedules RSECE.
