# RECS - Edifício de Referência

## O que muda entre o Edifício Previsto e o Edifício de Referência

No RECS (Regulamento de Desempenho Energético dos Edifícios de Comércio e Serviços), o **edifício de referência** é um edifício virtual com:

1. **Mesma geometria** do edifício real (previsto)
2. **Mesmos perfis de utilização** (ocupação, iluminação, equipamentos)
3. **Mesma localização** (zona climática)
4. **Soluções de referência** para envolvente e sistemas técnicos

### Elementos que MUDAM para o Edifício de Referência:

| Elemento | Previsto | Referência |
|----------|----------|------------|
| **Paredes Exteriores** | U real do projeto | Uref (por zona climática) |
| **Coberturas** | U real do projeto | Uref (por zona climática) |
| **Pavimentos** | U real do projeto | Uref (por zona climática) |
| **Envidraçados** | U e g reais | Uref e gTref |
| **Sistemas AVAC** | EER/COP reais | EER/COP de referência |
| **Iluminação** | DPI real | DPI de referência |

### Elementos que NÃO mudam:

- Geometria do edifício
- Áreas de espaços
- Perfis de ocupação
- Perfis de iluminação (horários)
- Perfis de equipamentos

---

## Valores de Referência - Envolvente Térmica

### Coeficientes de Transmissão Térmica de Referência (Uref)

**Fonte: Portaria 349-D/2013 e Portaria 138-I/2021**

#### Edifícios de Comércio e Serviços (RECS)

| Elemento | I1 | I2 | I3 | Unidade |
|----------|-----|-----|-----|---------|
| **Paredes Exteriores (verticais)** | 0.90 | 0.50 | 0.40 | W/(m².°C) |
| **Coberturas (horizontais)** | 0.60 | 0.40 | 0.35 | W/(m².°C) |
| **Pavimentos** | 0.60 | 0.40 | 0.35 | W/(m².°C) |
| **Envidraçados** | 4.30 | 3.30 | 2.80 | W/(m².°C) |

#### Coeficientes Máximos Admissíveis (Umáx)

| Elemento | I1 | I2 | I3 | Unidade |
|----------|-----|-----|-----|---------|
| **Paredes Exteriores** | 1.75 | 1.60 | 1.45 | W/(m².°C) |
| **Coberturas** | 1.25 | 1.00 | 0.90 | W/(m².°C) |
| **Pavimentos** | 1.25 | 1.00 | 0.90 | W/(m².°C) |
| **Envidraçados** | - | - | - | W/(m².°C) |

### Fator Solar de Referência (gTref)

| Zona Verão | gTref |
|------------|-------|
| V1 | 0.56 |
| V2 | 0.56 |
| V3 | 0.50 |

---

## Valores de Referência - Sistemas AVAC

### Eficiência de Referência (EER e COP)

**Fonte: Guia SCE - Indicadores de Desempenho RECS**

Para equipamentos de **bomba de calor/chiller de compressão** com permuta exterior de ar, **classe energética B**:

| Parâmetro | Valor de Referência |
|-----------|---------------------|
| **COP** (aquecimento) | 3.0 |
| **EER** (arrefecimento) | 2.9 |

Estes valores são aplicados quando:
- Não existe sistema de climatização no edifício
- Para o cálculo do edifício de referência

### Caldeiras (Aquecimento)

| Tipo | Rendimento Referência |
|------|----------------------|
| Caldeira a gás | 0.93 (93%) |
| Caldeira a gasóleo | 0.90 (90%) |

---

## Valores de Referência - Iluminação

### Densidade de Potência de Iluminação (DPI)

**Fonte: Portaria 349-D/2013**

| Parâmetro | Valor de Referência | Máximo |
|-----------|---------------------|--------|
| **DPI** | 2.5 W/m²/100lux | 3.8 W/m²/100lux |

A DPI é expressa em W/m² por cada 100 lux de iluminância.

---

## Zonas Climáticas de Portugal

### Zonas de Inverno (I1, I2, I3)

Baseadas nos graus-dias de aquecimento (GD) a 18°C:

| Zona | GD (°C.dia) | Regiões Típicas |
|------|-------------|-----------------|
| **I1** | GD ≤ 1300 | Litoral Sul, Algarve |
| **I2** | 1300 < GD ≤ 1800 | Litoral Centro e Norte |
| **I3** | GD > 1800 | Interior, Serra |

### Zonas de Verão (V1, V2, V3)

Baseadas na temperatura média exterior no verão:

| Zona | θext,v (°C) | Regiões Típicas |
|------|-------------|-----------------|
| **V1** | θext ≤ 20 | Litoral Norte |
| **V2** | 20 < θext ≤ 22 | Litoral Centro/Sul |
| **V3** | θext > 22 | Interior, Alentejo |

---

## Fórmulas de Cálculo do IEE

### IEE Previsto (IEEpr)

```
IEEpr = (Energia Primária PREV - Eren) / Área Útil
```

Onde:
- Energia Primária PREV = Σ(Consumo Final × Fpu)
- Fpu = 2.5 (eletricidade)
- Eren = Energia renovável produzida no local

### IEE Referência (IEEref)

```
IEEref = (Energia Primária REF) / Área Útil
```

Onde:
- Energia Primária REF = calculada com soluções de referência

### RIEE (Rácio de Eficiência)

```
RIEE = (IEEpr,s - IEEren) / IEEref,s
```

Onde:
- IEEpr,s = IEE previsto para consumos Tipo S
- IEEren = contribuição de energia renovável
- IEEref,s = IEE de referência para consumos Tipo S

### Consumos Tipo S e Tipo T

| Tipo S (conta para classificação) | Tipo T (não conta) |
|-----------------------------------|-------------------|
| AVAC (aquecimento e arrefecimento) | Equipamentos |
| AQS (águas quentes sanitárias) | Processos industriais |
| Iluminação interior | |
| Elevadores e escadas rolantes | |
| Ventilação | |

---

## Escala de Classes Energéticas

| Classe | RIEE |
|--------|------|
| **A+** | RIEE ≤ 0.25 |
| **A** | 0.25 < RIEE ≤ 0.50 |
| **B** | 0.50 < RIEE ≤ 0.75 |
| **B-** | 0.75 < RIEE ≤ 1.00 |
| **C** | 1.00 < RIEE ≤ 1.50 |
| **D** | 1.50 < RIEE ≤ 2.00 |
| **E** | 2.00 < RIEE ≤ 2.50 |
| **F** | RIEE > 2.50 |

---

## Legislação Relevante

1. **Decreto-Lei 101-D/2020** - Regime jurídico do SCE
2. **Portaria 349-D/2013** - Requisitos de qualidade térmica da envolvente
3. **Portaria 138-I/2021** - Atualização dos requisitos mínimos
4. **Despacho 15793-K/2013** - Parâmetros de cálculo

---

## Fontes

- [Guia SCE - Avaliação de Requisitos RECS](https://www.sce.pt/wp-content/uploads/2020/04/5.3-Guia-SCE-%E2%80%93-Avalia%C3%A7%C3%A3o-de-Requisitos-RECS_V1-1.pdf)
- [Guia SCE - Indicadores de Desempenho RECS](https://www.sce.pt/wp-content/uploads/2020/04/5.4-Guia-SCE-%E2%80%93-Indicadores-de-desempenho-RECS_V1-1.pdf)
- [Guia SCE - Conceitos e Definições RECS](https://www.sce.pt/wp-content/uploads/2020/04/5.1-Guia-SCE-Conceitos-e-Defini%C3%A7%C3%B5es-RECS_V1-1.pdf)
- [SCE - Legislação](https://www.sce.pt/legislacao/)
- [DGEG - Edifícios](https://www.dgeg.gov.pt/pt/areas-setoriais/energia/eficiencia-energetica/edificios)

---

**Última atualização:** 2026-02-05
