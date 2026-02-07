# HAP 5.1 Tools - Contexto do Projecto

Ferramentas Python para automatizar ficheiros Carrier HAP 5.1 (.E3A) para certificação energética portuguesa (RSECE/SCE).

## Estrutura do projecto

```
conversor/       Excel → E3A (criar projectos HAP)
extractor/       E3A → Excel (exportar dados)
editor/          Modificar E3A existente (mantém sistemas AVAC)
comparador/      Comparar dois E3A lado a lado
iee/             Calcular IEE e Classe Energética (A+ a F)
adaptador/       Converter formato HAP 5.2 → standard
docs/            Documentação técnica detalhada
exemplos/        Ficheiros de exemplo (Malhoa22, PWC, etc.)
_arquivo/        Scripts antigos - IGNORAR, não usar
```

## Skills disponíveis

- `/hap-converter` — Excel → E3A (valida + converte + valida)
- `/hap-extract` — E3A → Excel
- `/hap-edit` — Editar E3A existente (extrair → editar PREV → aplicar)
- `/hap-compare` — Comparar 2 E3A (extrai ambos + compara)
- `/hap-iee` — CSV HAP → Excel IEE → Classe Energética
- `/hap-adapt` — Excel HAP 5.2 → formato standard
- `/hap-validate` — Validar E3A ou Excel
- `/hap-help` — Lista de todas as skills

## Comandos principais

```bash
# Converter
python conversor/excel_to_hap.py <input.xlsx> conversor/templates/Modelo_RSECE.E3A <output.E3A>

# Extrair
python extractor/hap_extractor.py <input.E3A> [output.xlsx]

# Editor (2 fases)
python editor/editor_e3a.py extrair <ficheiro.E3A> <editor.xlsx>
python editor/editor_e3a.py aplicar <original.E3A> <editor.xlsx> <output.E3A>

# Comparar (3 passos)
python extractor/hap_extractor.py <e3a_1> <temp1.xlsx>
python extractor/hap_extractor.py <e3a_2> <temp2.xlsx>
python comparador/comparar_com_template.py comparador/Template_Comparacao_v7.xlsx <temp1.xlsx> <temp2.xlsx> <output.xlsx>

# IEE
python iee/iee_completo_v3.py <pasta_prev> <pasta_ref> [output.xlsx]

# Adaptar HAP 5.2
python adaptador/adapter_hap52.py <input_hap52.xlsx> [output.xlsx]

# Validar
python validar_e3a.py <ficheiro.E3A> [--fix]
python conversor/validar_excel_hap.py <ficheiro.xlsx>
```

## Ficheiros importantes

- `conversor/templates/HAP_Template_RSECE.xlsx` — Template Excel (147 colunas)
- `conversor/templates/Modelo_RSECE.E3A` — Modelo base com 82 schedules RSECE
- `comparador/Template_Comparacao_v7.xlsx` — Template de comparação formatado
- `conversor/hap_library.py` — Biblioteca de estruturas binárias
- `conversor/hap_schedule_library.py` — Biblioteca de schedules

## Formato E3A (referência rápida)

Ficheiros .E3A são ZIPs contendo:

| Ficheiro | Conteúdo | Tamanho registo |
|----------|----------|-----------------|
| HAP51SPC.DAT | Espaços | 682 bytes |
| HAP51SCH.DAT | Schedules | 792 bytes |
| HAP51WAL.DAT | Paredes (assemblies) | 3187 bytes |
| HAP51ROF.DAT | Coberturas (assemblies) | 3187 bytes |
| HAP51WIN.DAT | Janelas | 555 bytes |
| HAP51DOR.DAT | Portas | variável |
| HAP51INX.MDB | Base de dados Access (índices e links) | - |

### Offsets críticos no Space (682 bytes)

```
0-23     Nome (Latin-1, null-terminated)
24-27    Área (ft², float)
28-31    Pé-direito (ft, float)
32-35    Peso (lb/ft², float)
46-49    OA valor (encoded, float) — fórmula piecewise!
50-51    OA unit (uint16: 1=L/s, 2=L/s/m², 3=L/s/person, 4=%)
72-343   8 Wall blocks (34 bytes cada)
344-439  4 Roof blocks (24 bytes cada)
440-465  Partition 1 (Ceiling)
466-491  Partition 2 (Wall)
492-541  Floor
554-571  Infiltration
580-595  People (schedule ID @ 594-595)
600-617  Lighting (schedule ID @ 616-617, NÃO 614!)
632-645  Misc loads
656-661  Equipment (schedule ID @ 660-661)
```

### Wall block (34 bytes)

```
+0   direction code (uint16: N=1, E=5, S=9, W=13, H=17)
+2   gross area (ft², float)
+6   wall type ID (uint16)
+8   window 1 type ID
+10  window 2 type ID (NÃO reservado!)
+12  window 1 quantity
+14  window 2 quantity
+16  door type ID
+18  door quantity
```

### Window record (555 bytes)

```
0-254    Nome (Latin-1)
257-260  Altura (ft, float) — NÃO 24!
261-264  Largura (ft, float) — NÃO 28!
269-272  U-Value (BTU/hr·ft²·°F, float)
273-276  SHGC (float) — NÃO 433!
```

### Fórmula OA (piecewise fast_exp2)

```
Y0 = 512 CFM em L/s = 241.637...
Encode: v=user/Y0, t=log2(v), if t<0: int=t/4+4, else int=t/2+4
Decode: if int<4: t=(int-4)*4, else t=(int-4)*2, user=Y0*2^t
```

## Conversões de unidades

| SI → Imperial | Fórmula |
|---------------|---------|
| m² → ft² | × 10.7639 |
| m → ft | × 3.28084 |
| kg/m² → lb/ft² | ÷ 4.8824 |
| W/m²K → BTU/hr·ft²·°F | ÷ 5.678 |
| W → BTU/hr | × 3.412 |
| W/m² → W/ft² | ÷ 10.764 |
| °C → °F | × 1.8 + 32 |

## Schedules RSECE

82 schedules: 1 Sample + 27 tipologias × 3 perfis (Ocup, Ilum, Equip).
Tipologias: Hipermercado, Supermercado, Centro Comercial, Pequena Loja, Restaurante, Pastelaria, Pronto-a-Comer, Venda Grosso, Hotel 4-5*, Hotel 1-3*, Cinema Teatro, Discoteca, Bingo Clube Social, Clube Desp Piscina, Clube Desportivo, Escritorio, Banco Sede, Banco Filial, Comunicacoes, Biblioteca, Museu Galeria, Tribunal Camara, Prisao, Escola, Universidade, Saude Sem Intern, Saude Com Intern.

## Erros comuns

### Erro 9 "Subscript out of range"
Causa: Schedules com calendário inválido (valor 100 em vez de 1-8).
Solução: `python validar_e3a.py ficheiro.E3A --fix`

### HAP mostra espaços de outro projecto
Causa: MDB interno com índices antigos.
Solução: Reconverter com versão actual do conversor.

### Valores OA incorrectos
Causa: Fórmula OA antiga (exponencial simples) vs actual (piecewise fast_exp2).
Solução: Usar versão actual dos scripts.

## 9 bugs históricos (todos corrigidos)

1. Wall block offsets desalinhados no conversor
2. Lighting schedule offset 614→616 na library
3. Wall block offsets errados no extractor
4. Window SHGC offset 433→273 no extractor
5. Window SHGC offset 433→273 no editor
6. Infiltration offsets 492→554 na library
7. Thermostat/partition overlap na library
8. Missing data_only=True no editor
9. Fórmula OA antiga em editor/library/extractor

## Dependências

```
pip install openpyxl        # obrigatório
pip install pyodbc          # opcional, para actualizar MDB
pip install flask           # opcional, interface web
```
