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
