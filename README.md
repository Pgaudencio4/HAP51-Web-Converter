# HAP 5.1 Excel Converter

Conversor de ficheiros Excel para HAP 5.1 (.E3A).

## Estrutura do Projecto

```
HAPPXXXX/
├── excel_to_hap.py      # Conversor principal Excel -> HAP
├── validar_e3a.py       # Validador e corrector de ficheiros E3A
├── hap_to_excel.py      # Conversor HAP -> Excel (extracção)
├── hap_library.py       # Biblioteca de funções HAP
├── app.py               # Interface web (Flask)
│
├── templates/           # Ficheiros base para conversão
│   ├── Template_Limpo_RSECE.E3A    # Template E3A limpo (usar este!)
│   ├── Modelo_RSECE.E3A            # Modelo RSECE
│   └── HAP_Template_RSECE.xlsx     # Template Excel
│
├── docs/                # Documentação
│   ├── FORMATO_HAP51.md            # Especificação do formato
│   ├── ERRO9_SUBSCRIPT_OUT_OF_RANGE.md  # Resolução do Erro 9
│   └── ESPECIFICACAO_EXPORT_HAP.md # Especificação de exportação
│
└── old/                 # Ficheiros antigos/testes (pode apagar)
```

## Uso Rápido

### Converter Excel para HAP:

```bash
python excel_to_hap.py <input.xlsx> <modelo_base.E3A> <output.E3A>

# Exemplo:
python excel_to_hap.py MeusDados.xlsx templates/Template_Limpo_RSECE.E3A Output.E3A
```

### Validar ficheiro E3A:

```bash
# Só validar
python validar_e3a.py MeuFicheiro.E3A

# Validar e corrigir automaticamente
python validar_e3a.py MeuFicheiro.E3A --fix
```

## Aplicação Web

Para uma experiência mais simples, use a aplicação web:

```bash
python app.py
```

Abrir no browser: **http://localhost:5000**

## Problemas Conhecidos e Soluções

### Erro 9 "Subscript out of range"

**Causa:** Calendário dos schedules com valor 100 (inválido) em vez de 1-8.

**Solução:**
```bash
python validar_e3a.py MeuFicheiro.E3A --fix
```

Ver `docs/ERRO9_SUBSCRIPT_OUT_OF_RANGE.md` para detalhes.

### HAP mostra espaços errados (de outro edifício)

**Causa:** MDB (HAP51INX.MDB) não foi actualizado correctamente.

**Solução:** O conversor actualizado já corrige este problema. Certifique-se de usar a versão mais recente do `excel_to_hap.py`.

## Formato do Ficheiro Excel

O ficheiro Excel deve ter as seguintes folhas:

1. **Spaces** - Lista de espaços com propriedades
2. **Walls** - Definições de paredes (opcional)
3. **Windows** - Definições de janelas (opcional)
4. **Roofs** - Definições de coberturas (opcional)

Ver `templates/HAP_Template_RSECE.xlsx` como exemplo.

## Requisitos

- Python 3.8+
- openpyxl
- pyodbc (para actualizar MDB)

```bash
pip install openpyxl pyodbc
```

## Ficheiros de Referência

| Ficheiro | Descrição |
|----------|-----------|
| `templates/Template_Limpo_RSECE.E3A` | Template base limpo (recomendado) |
| `templates/HAP_Template_RSECE.xlsx` | Template Excel com estrutura correcta |

---

**Última actualização:** 2026-02-04
