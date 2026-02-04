# HAP 5.1 Tools

Conjunto de ferramentas para trabalhar com ficheiros HAP 5.1 (Carrier):
- **Conversor**: Excel → E3A (criar projectos HAP a partir de Excel)
- **Extractor**: E3A → Excel (exportar projectos HAP para Excel)
- **Comparador**: Comparar dois ficheiros E3A lado a lado
- **Editor**: Modificar E3A existente (mantém sistemas, schedules, etc.)

---

## 📁 Estrutura do Projecto

```
HAPPXXXX/
│
├── conversor/                    ← CONVERTER Excel para E3A
│   ├── excel_to_hap.py           Script principal de conversão
│   ├── hap_library.py            Biblioteca de funções HAP
│   ├── hap_schedule_library.py   Biblioteca de schedules
│   ├── validar_e3a.py            Validador de ficheiros E3A
│   ├── validar_excel_hap.py      Validador de Excel antes de converter
│   └── templates/
│       ├── HAP_Template_RSECE.xlsx   ⭐ FOLHA MODELO (preencher esta!)
│       └── Modelo_RSECE.E3A          E3A base para conversão
│
├── extractor/                    ← EXTRAIR E3A para Excel
│   ├── hap_extractor.py          Script principal de extracção
│   └── hap_to_excel.py           Versão alternativa
│
├── comparador/                   ← COMPARAR dois E3A
│   ├── comparar_com_template.py  Script principal de comparação
│   ├── criar_template_v7.py      Cria template de comparação formatado
│   ├── Template_Comparacao_v7.xlsx   Template formatado
│   ├── comparar_excels.py        Comparador simples
│   └── comparar_lado_a_lado.py   Comparador lado a lado (antigo)
│
├── editor/                       ← EDITAR E3A existente (novo!)
│   ├── editor_e3a.py             Script principal de edição
│   └── README.md                 Documentação do editor
│
├── exemplos/                     ← Ficheiros de exemplo
│   ├── Malhoa22.E3A              Exemplo de E3A completo
│   ├── Malhoa22_Final.xlsx       Exemplo de Excel preenchido
│   └── ...
│
├── docs/                         ← Documentação técnica
│   ├── HAP_FILE_SPECIFICATION.md Especificação do formato E3A
│   ├── HAP_COMPLETE_FIELD_MAP.md Mapeamento dos 147 campos
│   └── ...
│
├── _arquivo/                     ← Ficheiros antigos (backup)
│
├── app.py                        Interface web (Flask) - opcional
└── README.md                     Este ficheiro
```

---

## 🔄 1. CONVERSOR (Excel → E3A)

### Para que serve?
Criar um ficheiro HAP (.E3A) a partir de um Excel preenchido com os dados do projecto.

### Como usar?

#### Passo 1: Preencher a folha modelo
```
conversor/templates/HAP_Template_RSECE.xlsx
```
Esta folha tem todas as colunas necessárias. Preenche os espaços na folha "Espacos".

#### Passo 2: Executar o conversor
```bash
cd conversor
python excel_to_hap.py <teu_excel.xlsx> templates/Modelo_RSECE.E3A <output.E3A>
```

**Exemplo:**
```bash
python excel_to_hap.py MeuProjecto.xlsx templates/Modelo_RSECE.E3A MeuProjecto.E3A
```

#### Passo 3: Validar o ficheiro (opcional)
```bash
python validar_e3a.py MeuProjecto.E3A --fix
```

### Campos suportados (147 campos)
- **GENERAL**: Nome, Tipo, Área, Pé-direito, Piso, Multiplicador
- **INTERNALS**: People, Lighting, Equipment, Misc (com schedules)
- **INFILTRATION**: ACH Heating/Cooling/Ventilation
- **FLOORS**: Edge R, Length, Parcel (4 pisos)
- **PARTITIONS**: Ceiling e Wall (U-value, Área, Temperatura)
- **WALLS**: 8 paredes com Assembly, Orientação, Área, Janelas, Sombreamento
- **ROOFS**: 4 coberturas com Assembly, Orientação, Área, Skylights

---

## 📤 2. EXTRACTOR (E3A → Excel)

### Para que serve?
Exportar os dados de um ficheiro HAP (.E3A) para Excel, para análise ou edição.

### Como usar?
```bash
cd extractor
python hap_extractor.py <ficheiro.E3A> <output.xlsx>
```

**Exemplo:**
```bash
python hap_extractor.py MeuProjecto.E3A MeuProjecto_Extraido.xlsx
```

### O que extrai?
O Excel gerado tem 4 folhas:
- **Espacos**: Todos os 147 campos de cada espaço
- **Windows**: Nome, U-Value, SHGC, Altura, Largura
- **Walls**: Nome, U-Value, Espessura, Massa
- **Roofs**: Nome, U-Value, Espessura, Massa

---

## ⚖️ 3. COMPARADOR (E3A vs E3A)

### Para que serve?
Comparar dois ficheiros E3A (ex: versão Previsto vs Referência) e ver as diferenças.

### Como usar?

#### Passo 1: Extrair ambos os E3A
```bash
cd extractor
python hap_extractor.py Projecto_Prev.E3A Prev_extraido.xlsx
python hap_extractor.py Projecto_Ref.E3A Ref_extraido.xlsx
```

#### Passo 2: Executar a comparação
```bash
cd ../comparador
python comparar_com_template.py Template_Comparacao_v7.xlsx ../Prev_extraido.xlsx ../Ref_extraido.xlsx Comparacao.xlsx
```

### Resultado
Excel com comparação lado a lado:
- **Verde (OK)**: Valores iguais
- **Vermelho (DIFF)**: Valores diferentes
- **F1/F2**: Valor só existe num dos ficheiros

Inclui comparação de:
- Espacos (147 campos × 3 colunas)
- Windows (Nome, U-Value, SHGC, Dimensões)
- Walls (Nome, U-Value, Espessura, Massa)
- Roofs (Nome, U-Value, Espessura, Massa)

---

## ✏️ 4. EDITOR (Modificar E3A existente)

### Para que serve?
Modificar campos de um E3A existente **sem perder sistemas AVAC, schedules, etc.**

### Como usar?

#### Passo 1: Extrair E3A para Excel de edição
```bash
cd editor
python editor_e3a.py extrair MeuProjecto.E3A MeuProjecto_EDITOR.xlsx
```

#### Passo 2: Editar o Excel
1. Abrir `MeuProjecto_EDITOR.xlsx`
2. A coluna **REF** mostra os valores actuais do E3A
3. Preencher a coluna **PREV** apenas com os valores que queres alterar
4. Deixar **PREV vazio** = campo não é alterado

#### Passo 3: Aplicar alterações
```bash
python editor_e3a.py aplicar MeuProjecto.E3A MeuProjecto_EDITOR.xlsx MeuProjecto_Novo.E3A
```

### Vantagens
- ✅ **Mantém sistemas AVAC** intactos
- ✅ **Mantém schedules** intactos
- ✅ **Mantém resultados** de simulações
- ✅ Só altera o que preenches em PREV

### Exemplo
Para alterar a área do espaço "Sala1" de 50 m² para 60 m²:

| ESPAÇO | CAMPO | PREV | REF |
|--------|-------|------|-----|
| Sala1 | Area | **60** | 50 |

---

## 🌐 Interface Web (Opcional)

Para uma interface gráfica simples:
```bash
python app.py
```
Abrir no browser: **http://localhost:5000**

---

## 📋 Requisitos

```bash
pip install openpyxl pyodbc flask
```

- Python 3.8+
- openpyxl (manipulação de Excel)
- pyodbc (actualização de MDB - só para conversor)
- flask (interface web - opcional)

---

## ❓ Problemas Comuns

### Erro 9 "Subscript out of range"
O HAP não abre o ficheiro E3A.

**Solução:**
```bash
cd conversor
python validar_e3a.py MeuFicheiro.E3A --fix
```

### HAP mostra espaços de outro projecto
O MDB interno não foi actualizado correctamente.

**Solução:** Usar a versão mais recente do conversor que já corrige este problema automaticamente.

---

## 📊 Fluxos de Trabalho

### Criar E3A novo (Conversor)
```
Excel Modelo → conversor/excel_to_hap.py → E3A novo
```

### Exportar E3A para análise (Extractor)
```
E3A → extractor/hap_extractor.py → Excel com dados
```

### Comparar dois E3A (Comparador)
```
E3A₁ → extractor → Excel₁ ─┐
                           ├→ comparador → Excel comparação
E3A₂ → extractor → Excel₂ ─┘
```

### Modificar E3A existente (Editor) ⭐ RECOMENDADO
```
E3A original → editor extrair → Excel PREV/REF
                                      │
                          (preencher PREV)
                                      │
                                      ▼
E3A original + Excel → editor aplicar → E3A modificado
                                        (mantém sistemas!)
```

---

**Última actualização:** 2026-02-04
