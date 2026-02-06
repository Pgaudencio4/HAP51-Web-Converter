# Editor E3A

Modifica ficheiros HAP existentes sem perder sistemas, schedules, etc.

## Conceito

```
E3A Original (com sistemas AVAC, schedules, etc.)
       │
       ▼ Passo 1: Extrair
Excel Editor:
       ├── ESPAÇO    - Nome do espaço
       ├── CAMPO     - Nome do campo
       ├── PREV      - VAZIO (preencher só o que queres alterar)
       ├── REF       - Valor actual do E3A
       └── UNIDADE   - Unidade SI
       │
       ▼ Passo 2: (Manual) Preencher coluna PREV
       │
       ▼ Passo 3: Aplicar
E3A Modificado (mantém TUDO, só altera campos preenchidos em PREV)
```

## Uso

### Passo 1: Extrair E3A para Excel de edição
```bash
python editor_e3a.py extrair MeuProjecto.E3A MeuProjecto_EDITOR.xlsx
```

### Passo 2: Editar o Excel
1. Abrir `MeuProjecto_EDITOR.xlsx`
2. Ver valores actuais na coluna **REF**
3. Preencher coluna **PREV** APENAS com valores que queres alterar
4. Deixar PREV vazio = não altera o campo

### Passo 3: Aplicar alterações ao E3A
```bash
python editor_e3a.py aplicar MeuProjecto.E3A MeuProjecto_EDITOR.xlsx MeuProjecto_Novo.E3A
```

## Campos Suportados

| Campo | Unidade | Descrição |
|-------|---------|-----------|
| Space Name | - | Nome do espaço |
| Area | m² | Área do espaço |
| Height | m | Pé-direito |
| Floor Number | - | Número do piso |
| Multiplier | - | Multiplicador |
| Occupants | - | Número de ocupantes |
| People Sensible | W | Carga sensível pessoas |
| People Latent | W | Carga latente pessoas |
| Lighting W/m2 | W/m² | Iluminação |
| Ballast Multiplier | - | Multiplicador balastro |
| Equipment W/m2 | W/m² | Equipamento |
| ACH Heating | ACH | Infiltração aquecimento |
| ACH Cooling | ACH | Infiltração arrefecimento |
| ACH Ventilation | ACH | Ventilação |
| Floor Area | m² | Área do piso |
| Ceiling Area | m² | Área do tecto |
| Ceiling U-Value | W/m²K | U-value tecto |
| PartWall Area | m² | Área parede partição |
| PartWall U-Value | W/m²K | U-value parede partição |

## Vantagens

- **Não perdes sistemas AVAC** - Só altera os campos que preenches
- **Não perdes schedules** - Ficam intactos
- **Não perdes resultados** - Mantém tudo o resto
- **Simples de usar** - Só preenches o que queres mudar
- **Verificável** - Vês o valor actual (REF) ao lado

## Exemplo

Se queres alterar a área do espaço "Sala1" de 50 m² para 60 m²:

| ESPAÇO | CAMPO | PREV | REF | UNIDADE |
|--------|-------|------|-----|---------|
| Sala1 | Area | **60** | 50 | m² |

Todos os outros campos de Sala1 ficam inalterados.

---

## Calcular Valores REF (Referência RECS)

Para preencher automaticamente as colunas REF com valores de referência RECS:

```bash
python calcular_ref.py MeuProjecto_EDITOR.xlsx MeuProjecto_EDITOR_REF.xlsx
```

### O que calcula?

#### 1. Iluminação REF
```
Potência REF (W) = DPI_ref × (Iluminância / 100) × Área
```
- **DPI_ref** = 2.5 W/m²/100lux (valor RECS)
- **Iluminância** = determinada pelo tipo de espaço

| Tipo de Espaço      | Lux | W/m² REF |
|---------------------|-----|----------|
| Escritório/Area     | 500 |    12.5  |
| Copa/Cozinha        | 300 |     7.5  |
| IS/WC               | 200 |     5.0  |
| Zonas técnicas      | 200 |     5.0  |
| Escadas             | 150 |    3.75  |
| Circulação/Corredor | 100 |     2.5  |
| Estacionamento      |  75 |   1.875  |

#### 2. Caudais Ar Novo REF
```
Caudal REF (L/s) = ROUNDUP(MAX(Qocup, Qedif) / Eficácia / 3.6, 0)
```
- **Qocup** = ocupantes × 24 m³/h/pessoa (actividade sedentária)
- **Qedif** = área × 3 m³/h/m² (sem poluentes)
- **Eficácia** = 0.8 (espaços com ocupação) ou 1.0 (espaços sem ocupação)

**NOTA**: REF é o caudal mínimo regulamentar. PREV pode ser maior (ex: cozinhas com extracção).

Valores RECS por Actividade Metabólica:
| Actividade  | met  | Caudal/pessoa |
|-------------|------|---------------|
| Sedentária  | 1.2  | 24 m³/h       |
| Moderada    | 1.75 | 35 m³/h       |
| Alta        | 5.0  | 98 m³/h       |

Valores RECS por Carga Poluente:
| Tipo | Caudal/m² |
|------|-----------|
| Sem poluentes | 3 m³/h/m² |
| Com poluentes | 5 m³/h/m² |

Eficácia de Ventilação REF:
| Tipo Espaço | Eficácia |
|-------------|----------|
| COM ocupação (Area, Escritorio, Sala, Copa, Quarto, Gab) | 0.8 |
| SEM ocupação (Escadas, Circulação, IS, Zonas técnicas, Estacionamento) | 1.0 |

### Fórmulas Excel inseridas

O script insere fórmulas (não valores fixos) para rastreabilidade:

| Coluna | Fórmula |
|--------|---------|
| OA REF (com ocupação) | `=ROUNDUP(MAX(S{row}*24, D{row}*3)/0.8/3.6, 0)` |
| OA REF (sem ocupação) | `=ROUNDUP(MAX(S{row}*24, D{row}*3)/1.0/3.6, 0)` |
| Task Ltg REF | `0` |
| Gen Ltg REF | `=ROUND(2.5*(lux/100)*D{row}, 0)` |
