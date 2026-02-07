---
description: "Mostrar ajuda e lista de skills HAP disponíveis. Usa quando o utilizador pergunta que comandos existem, o que pode fazer, ajuda hap, help, lista de skills."
---

# Skills HAP Disponíveis

Mostra ao utilizador esta lista:

## Conversão e Extracção

| Comando | O que faz |
|---------|-----------|
| `/hap-converter` | Converter Excel → E3A (valida antes e depois) |
| `/hap-extract` | Extrair E3A → Excel (147 campos, 4 folhas) |
| `/hap-adapt` | Converter formato HAP 5.2 → formato standard |

## Edição e Comparação

| Comando | O que faz |
|---------|-----------|
| `/hap-edit` | Editar E3A existente sem perder sistemas AVAC |
| `/hap-compare` | Comparar 2 E3A lado a lado (verde=igual, vermelho=diferente) |

## Certificação Energética

| Comando | O que faz |
|---------|-----------|
| `/hap-iee` | Calcular IEE e Classe Energética (A+ a F) a partir de CSV |

## Validação e Debug

| Comando | O que faz |
|---------|-----------|
| `/hap-validate` | Validar E3A ou Excel (com correcção automática) |

O agente **hap-debug** é activado automaticamente quando descreves um problema com o HAP (ex: "erro 9", "valores errados", "e3a não abre").

## Workflow típico

```
1. Preencher Excel template    (conversor/templates/HAP_Template_RSECE.xlsx)
2. /hap-converter              → cria E3A
3. Simular no HAP 5.1          (manual)
4. Exportar CSV do HAP          (manual)
5. /hap-iee                    → Classe Energética
6. /hap-compare                → Comparar PREV vs REF
```
