---
description: "Calcular IEE e Classe Energética SCE. Usa quando o utilizador quer calcular iee, classe energética, certificação energética, processar csv do hap, energia primária, riee."
---

# Calcular IEE e Classe Energética (SCE)

Calcula o IEE (Indicador de Eficiência Energética) e a Classe Energética (A+ a F) a partir dos CSV exportados do HAP.

## Workflow

### 1. Pedir ao utilizador os dados

Pergunta ao utilizador:
- **Pasta com CSV PREV** (ficheiros CSV exportados do HAP para o cenário Previsto)
- **Pasta com CSV REF** (ficheiros CSV exportados do HAP para o cenário Referência)
- **Nome do Excel de output** (opcional)

Estrutura esperada das pastas:
```
PREV/
├── HAP51_Monthly_1_Sistema1.csv
├── HAP51_Monthly_2_Sistema2.csv
└── ...
REF/
├── HAP51_Monthly_1_Sistema1.csv
└── ...
```

### 2. Executar o cálculo

```bash
python iee/iee_completo_v3.py "<pasta_prev>" "<pasta_ref>" "<output.xlsx>"
```

### 3. Explicar ao utilizador o que preencher

O Excel gerado tem 17 folhas. O utilizador precisa de preencher manualmente:

**Obrigatório:**
1. **Simulacao** (folha 5) → Preencher EER e COP de cada sistema (células amarelas)
2. **EnergiaPrimaria** (folha 14) → Preencher **Área Útil (m²)**

**Se aplicável:**
3. IluminacaoENU → Iluminação exterior/não regulada
4. AQS → Águas quentes sanitárias
5. PV → Fotovoltaico
6. EquipamentosExtra → Equipamentos não simulados
7. Elevadores → Dados dos elevadores (cálculo RECS)
8. VentilacaoExtra → Ventilação não simulada
9. Bombagem → Bombas não simuladas

**Automático (não mexer):**
- DetalhePREV/REF, MensalPREV/REF → Dados brutos dos CSV
- Desagregacao → Resume todos os consumos
- IEE → Indicadores calculados
- Classe → Resultado final (A+ a F)

### 4. Reportar resultado

Mostra:
- Quantos sistemas foram processados (PREV e REF)
- Confirma que o Excel foi criado
- Recorda os passos de preenchimento manual

## Fórmulas de cálculo (referência)

```
IEEprev,s = Energia Primária PREV (Tipo S) / Área Útil
IEEref,s  = Energia Primária REF (Tipo S) / Área Útil
IEEren    = Energia Renovável × Fpu / Área Útil
RIEE      = (IEEprev,s - IEEren) / IEEref,s
```

Escala: A+ (≤0.25), A (0.26-0.50), B (0.51-0.75), B- (0.76-1.00), C (1.01-1.50), D (1.51-2.00), E (2.01-2.50), F (>2.50)
