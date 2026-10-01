# Auditor de Exames — PCMSO x Benner

Aplicativo Streamlit para conferir se os exames previstos no PCMSO estão cadastrados no Benner com os campos ocupacionais correspondentes.

## Arquivos de entrada

1. Relatório do Benner (`.xlsx`)
2. PCMSO (`.docx`)

## Conferências

- Cargo PCMSO x cargo do Benner
- Exame PCMSO x exame do Benner
- Admissional
- Periódico
- Mudança de Função
- Retorno ao Trabalho
- Demissional, quando a coluna existir no relatório do Benner
- Periodicidade conforme preenchida no PCMSO
- Exames compostos

## Regra de Mudança de Risco / Retorno

O PCMSO usa uma coluna combinada `Mud. Risco Ocupacional | Retorno`. Quando ela está preenchida, o app exige simultaneamente no Benner:

- `M. Função = S`
- `R. Trabalho = S`

## Exames compostos

O app separa automaticamente exames que aparecem juntos no PCMSO e normalmente separados no Benner, incluindo TGO + TGP, grupos de sorologias, grupos de metais na urina e combinações com alternativas OR.

## Status

- `OK`: requisito do PCMSO encontrado no Benner
- `FALTA`: requisito previsto no PCMSO não encontrado/conferido no Benner
- `NÃO CONFERIDO — COLUNA AUSENTE`: coluna necessária não existe no relatório enviado
- `NÃO CONFERIDO — EXAME NÃO MAPEADO`: exame do PCMSO precisa de mapeamento adicional
- `EXTRA NO BENNER`: exame/quadrinho marcado no Benner que não está previsto para aquele cargo no PCMSO

## Streamlit Cloud

Suba `app.py` e `requirements.txt` em um repositório GitHub e selecione `app.py` como arquivo principal do aplicativo.
