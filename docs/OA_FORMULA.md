# OA (Outdoor Air) Encoding Formula - HAP 5.1

**Data:** 2026-02-05
**Status:** DEFINITIVO - Formula exacta, erro 0.0 L/s

---

## Resumo

O HAP 5.1 codifica o valor de Outdoor Air (OA) nos ficheiros E3A usando uma
funcao **nao-linear** baseada em `fast_exp2`, uma aproximacao linear por partes
de 2^t. Esta formula foi descoberta por engenharia reversa apos calibracao
com 43 pontos medidos directamente no HAP 5.1.

**IMPORTANTE:** Versoes anteriores da documentacao tinham formulas ERRADAS
(lineares ou exponenciais simples). Apenas esta formula e correcta.

---

## Localizacao no Ficheiro Binario

No registo de espaco (682 bytes, ficheiro `HAP51SPC.DAT`):

| Offset | Size | Type   | Campo           |
|--------|------|--------|-----------------|
| 44-45  | 2    | bytes  | OA Auxiliary    |
| 46-49  | 4    | float  | OA Internal (x) |
| 50-51  | 2    | uint16 | OA Unit Code    |

**Unit Codes:**
- `1` = L/s (caudal total)
- `2` = L/s/m2 (caudal por area)
- `3` = L/s/person (caudal por pessoa)
- `4` = % (percentagem do ar insuflado)

Para unit codes 1, 2 e 3: usa a formula fast_exp2 descrita abaixo.
Para unit code 4 (%): conversao linear simples `valor = interno * 28.5714`.

---

## A Formula

### Decode (Ler): interno (x) -> valor em L/s

```
y = Y0 * fast_exp2(k * (x - 4))
```

Onde:
- **Y0** = 512 * (28.316846592 / 60) = **241.637090918 L/s** (= 512 CFM)
- **k = 4** quando x < 4 (equivale a base 16 por unidade de x)
- **k = 2** quando x >= 4 (equivale a base 4 por unidade de x)

### Encode (Escrever): valor em L/s -> interno (x)

```
t = fast_log2(y / Y0)
x = t/4 + 4    se t < 0  (y < Y0, ou seja, < 241.6 L/s)
x = t/2 + 4    se t >= 0 (y >= Y0, ou seja, >= 241.6 L/s)
```

---

## fast_exp2 e fast_log2

### fast_exp2(t) - Aproximacao linear de 2^t

```
fast_exp2(t) = 2^floor(t) * (1 + frac(t))
```

Onde `frac(t) = t - floor(t)` e a parte fracionaria.

Em vez de calcular `2^t` exactamente (exponencial verdadeira), o HAP interpola
**linearmente** entre potencias inteiras consecutivas de 2:

```
Entre t=0 e t=1:  fast_exp2(t) = 1 * (1 + t)     -> de 1.0 a 2.0
Entre t=1 e t=2:  fast_exp2(t) = 2 * (1 + t-1)   -> de 2.0 a 4.0
Entre t=2 e t=3:  fast_exp2(t) = 4 * (1 + t-2)   -> de 4.0 a 8.0
```

Isto e uma tecnica classica da era VB6/C para evitar chamadas a `Exp()` lento.

### fast_log2(v) - Inversa de fast_exp2

```
n = floor(log2(v))
f = v / 2^n - 1
fast_log2(v) = n + f
```

---

## Porque dois regimes (k=4 e k=2)?

O ponto de viragem e em **x = 4**, que corresponde a **Y0 = 241.6 L/s = 512 CFM**.

- Para **x < 4** (valores pequenos, < 512 CFM): k=4, o valor duplica a cada
  Dx=0.25, dando uma resolucao mais fina na gama baixa de caudais.
- Para **x >= 4** (valores grandes, >= 512 CFM): k=2, o valor duplica a cada
  Dx=0.5, cobrindo uma gama maior com menos resolucao.

A formula e **continua em x=4**: ambos os ramos dao y = Y0.

---

## Constantes Fisicas

A constante Y0 vem da conversao exacta de CFM para L/s:

```
1 ft3 = 28.316846592 litros  (exacto, definicao SI)
1 CFM = 28.316846592 / 60 = 0.471947443... L/s
512 CFM = 512 * 0.471947443 = 241.637090918... L/s
```

O valor 512 = 2^9 e uma potencia de 2, o que faz sentido com a tecnica fast_exp2:
e o ponto de referencia natural da funcao.

---

## Implementacao em Python

```python
import math

_OA_Y0 = 512.0 * (28.316846592 / 60.0)  # 241.637090918... L/s

def _fast_exp2(t):
    """Aproximacao linear por partes de 2^t."""
    n = math.floor(t)
    f = t - n
    return (2.0 ** n) * (1.0 + f)

def _fast_log2(v):
    """Inversa de _fast_exp2."""
    n = math.floor(math.log2(v))
    f = v / (2.0 ** n) - 1.0
    if f < 0:
        n -= 1
        f = v / (2.0 ** n) - 1.0
    if f >= 1.0:
        n += 1
        f = v / (2.0 ** n) - 1.0
    return n + f

def decode_oa(x):
    """Decode: interno (x) -> L/s"""
    if x <= 0:
        return 0.0
    k = 4.0 if x < 4.0 else 2.0
    return _OA_Y0 * _fast_exp2(k * (x - 4.0))

def encode_oa(value_ls):
    """Encode: L/s -> interno (x)"""
    if value_ls <= 0:
        return 0.0
    t = _fast_log2(value_ls / _OA_Y0)
    if t < 0:
        return t / 4.0 + 4.0
    else:
        return t / 2.0 + 4.0
```

---

## Tabela de Referencia

| Interno (x) | L/s        | CFM       | Notas                     |
|-------------|------------|-----------|---------------------------|
| 2.0         | 0.9        | 1.9       | Minimo pratico            |
| 2.5         | 3.8        | 8.0       |                           |
| 3.0         | 15.1       | 32.0      |                           |
| 3.5         | 60.4       | 128.0     |                           |
| 3.75        | 120.8      | 256.0     |                           |
| 4.0         | 241.6      | 512.0     | **Ponto de viragem (Y0)** |
| 4.5         | 483.3      | 1024.0    |                           |
| 5.0         | 966.5      | 2048.0    |                           |
| 5.5         | 1933.1     | 4096.0    |                           |
| 6.0         | 3866.2     | 8192.0    |                           |

---

## Validacao

Testado contra 43 pontos de calibracao medidos directamente no HAP 5.1:

- **43/43 pontos**: decode correcto (dentro de 0.1 L/s de arredondamento)
- **Round-trip** (encode->decode): erro 0.000000 L/s para todos os valores
- **Extrapolacao**: funciona para qualquer valor positivo (sem limites de gama)

Os 43 pontos de calibracao estao em `C:\Users\pedro\Downloads\OA_Calibracao_43pts.csv`.

### Amostra de verificacao

| Valor Desejado (L/s) | Interno (x) | HAP Mostra (L/s) | Erro |
|-----------------------|-------------|-------------------|------|
| 5                     | 2.581       | 5.0               | 0.0  |
| 50                    | 3.414       | 50.0              | 0.0  |
| 92                    | 3.631       | 92.0              | 0.0  |
| 235                   | 3.986       | 235.0             | 0.0  |
| 472                   | 4.477       | 472.0             | 0.0  |
| 527                   | 4.545       | 527.0             | 0.0  |
| 1000                  | 5.017       | 1000.0            | 0.0  |
| 1500                  | 5.276       | 1500.0            | 0.0  |
| 2000                  | 5.517       | 2000.0            | 0.0  |
| 5000                  | 6.147       | 5000.0            | 0.0  |
| 10000                 | 6.647       | 10000.0           | 0.0  |

---

## Historico

1. **Versao 1 (2026-01):** Formula linear `y = A*x + B` - ERRADA (~30% erro)
2. **Versao 2 (2026-02-04):** Formula exponencial `y = A*exp(B*x)` - ERRADA (~10% erro)
3. **Versao 3 (2026-02-05):** PCHIP interpolation com 43 pontos - FUNCIONAVA (<0.5% erro)
   mas requeria tabela de lookup. Usada como solucao temporaria.
4. **Versao 4 (2026-02-05):** Formula exacta `fast_exp2` - CORRECTA (0% erro).
   Descoberta por analise dos 43 pontos de calibracao. Formula closed-form,
   sem tabelas, funciona para qualquer valor.

---

## Ficheiros Actualizados

A formula esta implementada em 4 ficheiros:

| Ficheiro | Funcao | Direccao |
|----------|--------|----------|
| `conversor/excel_to_hap.py` | `encode_oa()` | L/s -> interno |
| `conversor/hap_library.py` | `_encode_oa()` + `_decode_oa()` | Ambas |
| `editor/editor_e3a.py` | `encode_oa()` | L/s -> interno |
| `extractor/hap_extractor.py` | `decode_oa()` | interno -> L/s |

---

## Nota sobre VB6 e fast_exp2

A funcao `fast_exp2` e uma tecnica de optimizacao comum em software dos anos 90-2000.
Em vez de calcular `2^t` usando `Exp(t * Log(2))` (operacao de floating-point lenta),
interpola-se linearmente entre potencias inteiras de 2. O erro maximo desta
aproximacao e ~6% (no meio do intervalo, onde `t = n + 0.5`), mas a funcao e
muito mais rapida e suficientemente precisa para aplicacoes AVAC.

O HAP 5.1 e escrito em Visual Basic 6 (VB6) e as DLLs estao em
`C:\E20-II\HAP51\CODE\` (SSNSpaceObj13.dll, HAPDataHandler19.4.dll).
A formula foi descoberta por reverse engineering dos dados, nao por
descompilacao do codigo.
