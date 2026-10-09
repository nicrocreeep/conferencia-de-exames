
import io
import re
import unicodedata
from collections import defaultdict
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from docx import Document


st.set_page_config(
    page_title="Auditor PCMSO × Benner",
    page_icon="🩺",
    layout="wide",
)


# =============================================================================
# NORMALIZAÇÃO
# =============================================================================

ROMAN_WORDS = {
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"
}

ROLE_FILLERS = {
    "DE", "DA", "DO", "DAS", "DOS", "EM", "E", "A", "O"
}

ROLE_TOKEN_REPLACEMENTS = {
    "TEC": "TECNICO",
    "TECN": "TECNICO",
    "AUX": "AUXILIAR",
    "ENG": "ENGENHEIRO",
    "SUP": "SUPERVISOR",
    "COORD": "COORDENADOR",
    "ADM": "ADMINISTRATIVO",
    "OP": "OPERADOR",
    "OPER": "OPERADOR",
    "SEG": "SEGURANCA",
    "SEGUR": "SEGURANCA",
    "RH": "RECURSOS HUMANOS",
    "MOV": "MOVIMENTACAO",
    "ORCAMENTISTA": "ORCAMENTO",
    "ORCAMENTOS": "ORCAMENTO",
    "ANDAIMES": "ANDAIME",
    "MATERIAIS": "MATERIAL",
    "CARGAS": "CARGA",
    "SERVICOS": "SERVICO",
    "OBRAS": "OBRA",
    "OPERADORES": "OPERADOR",
    "FAXINEIRA": "FAXINEIRO",
}

# Equivalências observadas entre nomenclaturas do PCMSO e do Benner.
ROLE_EQUIVALENCES = {
    "OPERADOR MAQUINA SOLDA AUTOMATICA": {
        "OPERADOR MAQUINA SOLDAR",
        "OPERADOR DE MAQUINA DE SOLDAR",
    },
    "ENCARREGADO ISOLAMENTO": {
        "ENCARREGADO ISOLAMENTO TERMICO",
    },
    "MOTORISTA CARRETA": {
        "CARRETEIRO",
        "MOTORISTA",
    },
    "COORDENADOR SEGURANCA TRABALHO": {
        "COORDENADOR SEGURANCA",
    },
    "ELETRICISTA CONTROLE FORCA": {
        "ELETRICISTA FORCA CONTROLE",
    },
    "ENGENHEIRO PRODUCAO": {
        "ENGENHEIRO PRODUCAO TRAINEE",
    },
    "DIRETOR": {
        "DIRETOR COMERCIAL",
        "DIRETOR OPERACIONAL",
    },
    "FUNILEIRO": {
        "FUNILEIRO BANCADA",
        "FUNILEIRO MONTADOR",
    },
}


def strip_accents(text: str) -> str:
    text = "" if text is None else str(text)
    return (
        unicodedata.normalize("NFKD", text)
        .encode("ascii", "ignore")
        .decode("ascii")
    )


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", str(text)).strip()


def normalize_role(text: str) -> str:
    text = strip_accents(text).upper()
    text = text.replace("&", " E ")
    text = re.sub(r"[/(),.;:_-]+", " ", text)

    tokens = []
    for token in text.split():
        token = ROLE_TOKEN_REPLACEMENTS.get(token, token)
        if token in ROMAN_WORDS:
            continue
        tokens.append(token)

    tokens = [t for t in tokens if t not in ROLE_FILLERS]

    result = " ".join(tokens)
    result = re.sub(r"\bSOLDAR\b", "SOLDA", result)
    result = re.sub(r"\bORCAMENTISTA\b", "ORCAMENTO", result)
    result = re.sub(r"\bORCAMENTOS\b", "ORCAMENTO", result)
    return normalize_spaces(result)


def expand_role_names(text: str) -> List[str]:
    """
    Converte linhas agregadas do PCMSO em candidatos individuais.

    Quando a linha contém vários cargos separados por vírgula/barra, a
    linha inteira não é usada como um cargo único. Isso evita falsos
    "cargos ausentes".
    """
    text = "" if text is None else str(text).strip()
    if not text:
        return []

    before_paren = re.split(r"\(", text, maxsplit=1)[0].strip(" ,")

    # Detecta separadores no nível superior.
    has_top_level_separator = False
    depth = 0
    for ch in before_paren:
        if ch == "(":
            depth += 1
        elif ch == ")" and depth > 0:
            depth -= 1
        elif depth == 0 and ch in ",/;":
            has_top_level_separator = True
            break

    if has_top_level_separator:
        candidates = []
    elif before_paren and before_paren != text:
        # Ex.: "Soldador (lista enorme de processos)" -> "Soldador".
        candidates = [before_paren]
    else:
        candidates = [text]

    parts = []
    current = []
    depth = 0

    for ch in before_paren:
        if ch == "(":
            depth += 1
        elif ch == ")" and depth > 0:
            depth -= 1

        if depth == 0 and ch in ",/;":
            part = "".join(current).strip(" ,")
            if part:
                parts.append(part)
            current = []
        else:
            current.append(ch)

    part = "".join(current).strip(" ,")
    if part:
        parts.append(part)

    candidates.extend(parts)

    if parts:
        candidates.append(parts[0])

    # Adiciona uma versão sem níveis como Junior/Pleno/Senior/Trainee.
    for candidate in list(candidates):
        cleaned = re.sub(
            r"\s+(?:JUNIOR|PLENO|SENIOR|TRAINEE)\s*$",
            "",
            candidate,
            flags=re.IGNORECASE,
        ).strip(" ,")
        if cleaned:
            candidates.append(cleaned)

    result = []
    seen = set()

    for candidate in candidates:
        norm = normalize_role(candidate)
        if not norm or norm in ROMAN_WORDS:
            continue

        # Descarta artefatos de listas como
        # "CALDEIREIRO CALDEIREIRO I" que não representam um cargo real.
        words = norm.split()
        if (
            len(words) >= 2
            and words[0] == words[1]
        ):
            continue
        if (
            len(words) >= 3
            and words[0] in words[1:]
            and words.count(words[0]) >= 2
        ):
            continue

        variants = {norm}

        for canonical, aliases in ROLE_EQUIVALENCES.items():
            if norm == canonical or norm in aliases:
                variants.add(canonical)
                variants.update(aliases)

        for variant in variants:
            v = normalize_role(variant)
            if v and v not in ROMAN_WORDS and v not in seen:
                seen.add(v)
                result.append(v)

    return result


# =============================================================================
# LEITURA DO RELATÓRIO BENNER# =============================================================================
# LEITURA DO RELATÓRIO BENNER
# =============================================================================

BENNER_ASO_ALIASES = {
    "admissional": [
        "admissional",
        "admissao",
    ],
    "periodico": [
        "periodico",
        "periódico",
    ],
    "mudanca_funcao": [
        "m. funcao",
        "m. função",
        "mudanca de funcao",
        "mudança de função",
        "mud. funcao",
    ],
    "retorno": [
        "r. trabalho",
        "retorno ao trabalho",
        "retorno",
        "r trabalho",
    ],
    "demissional": [
        "demissional",
        "demissao",
        "demissão",
    ],
}


def find_column(columns, aliases: List[str]) -> Optional[str]:
    normalized = {
        str(c): normalize_role(str(c))
        for c in columns
    }

    for original, norm in normalized.items():
        for alias in aliases:
            alias_norm = normalize_role(alias)
            if norm == alias_norm:
                return original

    # Segundo passe: procura trecho exato da forma normalizada.
    for original, norm in normalized.items():
        for alias in aliases:
            alias_norm = normalize_role(alias)
            if alias_norm and alias_norm in norm:
                return original

    return None


@st.cache_data(show_spinner=False)
def read_benner(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    # Remove linhas de cabeçalho repetido do próprio relatório.
    if "Cargo" in df.columns:
        df = df[
            normalize_series(df["Cargo"]).ne("CARGO")
        ].copy()

    # Descobre as colunas principais.
    project_col = find_column(df.columns, ["Projeto"])
    role_col = find_column(df.columns, ["Cargo"])
    exam_col = find_column(df.columns, ["Exame"])

    if role_col is None or exam_col is None:
        raise ValueError(
            "O relatório do Benner precisa possuir as colunas Cargo e Exame."
        )

    rename = {
        role_col: "Cargo",
        exam_col: "Exame",
    }
    if project_col is not None:
        rename[project_col] = "Projeto"

    df = df.rename(columns=rename)

    # Padroniza as quatro colunas principais.
    for col in ["Cargo", "Exame"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    if "Projeto" not in df.columns:
        df["Projeto"] = ""

    aso_columns = {}
    for canonical, aliases in BENNER_ASO_ALIASES.items():
        col = find_column(df.columns, aliases)
        if col:
            aso_columns[canonical] = col

    if "admissional" not in aso_columns or "periodico" not in aso_columns:
        raise ValueError(
            "Não encontrei as colunas Admissional e/ou Periodico no relatório do Benner."
        )

    # Renomeia os ASOs detectados para nomes internos.
    rename_aso = {
        col: f"ASO_{canonical}"
        for canonical, col in aso_columns.items()
    }
    df = df.rename(columns=rename_aso)

    # Garante as 5 colunas internas; a demissional pode não existir no relatório.
    for canonical in BENNER_ASO_ALIASES:
        col = f"ASO_{canonical}"
        if col not in df.columns:
            df[col] = pd.NA

    return df


def normalize_series(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .map(strip_accents)
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def is_marked(value) -> bool:
    if pd.isna(value):
        return False
    value = strip_accents(str(value)).upper().strip()
    return value in {
        "S", "SIM", "X", "1", "TRUE", "VERDADEIRO", "OK"
    }


def get_projects(df: pd.DataFrame) -> List[str]:
    values = [
        str(x).strip()
        for x in df["Projeto"].dropna().unique()
        if str(x).strip() and str(x).strip().upper() != "PROJETO"
    ]
    return sorted(set(values))


# =============================================================================
# LEITURA DO PCMSO DOCX
# =============================================================================


def semantic_exam_headers(header_values: List[str]):
    """
    Normaliza os quatro campos lógicos dos quadros de exames do PCMSO.
    """
    groups = []

    for idx, value in enumerate(header_values[3:], start=3):
        label = normalize_spaces(value)
        norm = strip_accents(label).upper()

        if "ADMISSA" in norm:
            canonical = "Admissao"
        elif "MUD" in norm and "RISCO" in norm and "RETORNO" in norm:
            canonical = "Mudanca_Risco_Retorno"
        elif "PERIODIC" in norm:
            canonical = "Periodicidade"
        elif "DEMISS" in norm:
            canonical = "Demissao"
        else:
            canonical = label

        if groups and groups[-1][0] == canonical:
            groups[-1][1].append(idx)
        else:
            groups.append([canonical, [idx]])

    return groups


@st.cache_data(show_spinner=False)
def read_pcmso(file_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    doc = Document(io.BytesIO(file_bytes))

    matrix_rows = []
    role_catalog = []

    for table_index, table in enumerate(doc.tables):
        header_idx = None
        header_values = None

        for row_index, row in enumerate(table.rows):
            values = [
                cell.text.strip().replace("\n", " ")
                for cell in row.cells
            ]
            normalized_values = [strip_accents(v).strip().upper() for v in values]
            if (
                "ADMISSAO" in normalized_values
                and "PERIODICIDADE" in normalized_values
                and "DEMISSAO" in normalized_values
            ):
                header_idx = row_index
                header_values = values
                break

        if header_idx is None:
            continue

        # O modelo do PCMSO atual traz GHE, setor e funções no segundo registro.
        if len(table.rows) < 2 or len(table.rows[1].cells) < 3:
            continue

        ghe = table.rows[1].cells[0].text.strip().replace("\n", " ")
        setor = table.rows[1].cells[1].text.strip().replace("\n", " ")

        functions = [
            line.strip()
            for line in table.rows[1].cells[2].text.splitlines()
            if line.strip()
        ]

        for function_line in functions:
            for individual_role in expand_role_names(function_line):
                role_catalog.append({
                    "Cargo PCMSO": individual_role,
                    "Cargo original no PCMSO": function_line,
                    "GHE": ghe,
                    "Setor": setor,
                    "Tabela Word": table_index + 1,
                })

        header_groups = semantic_exam_headers(header_values)
        semantic_labels = [g[0] for g in header_groups]

        # O modelo esperado possui exatamente quatro campos lógicos.
        if len(semantic_labels) != 4:
            continue

        for row in table.rows[header_idx + 1:]:
            values = [
                cell.text.strip().replace("\n", " ")
                for cell in row.cells
            ]
            if not values:
                continue

            exam = values[0].strip()
            if not exam:
                continue

            if normalize_role(exam) in {
                "EXAMES", "PERIODICIDADE"
            }:
                continue

            record = {
                "GHE": ghe,
                "Setor": setor,
                "Exame PCMSO": exam,
            }

            for label, indexes in header_groups:
                idx = indexes[0]
                record[label] = (
                    values[idx].strip()
                    if idx < len(values)
                    else ""
                )

            # Só mantemos exames que possuem pelo menos um campo preenchido.
            if any(
                str(record.get(label, "")).strip()
                for label in semantic_labels
            ):
                matrix_rows.append(record)

    matrix = pd.DataFrame(matrix_rows)
    roles = pd.DataFrame(role_catalog)

    if matrix.empty:
        raise ValueError(
            "Não encontrei os quadros de exames do PCMSO no arquivo Word."
        )

    # Remove duplicidades.
    matrix = matrix.drop_duplicates().reset_index(drop=True)
    roles = roles.drop_duplicates().reset_index(drop=True)

    return matrix, roles, semantic_labels


# =============================================================================
# MATCHING DE CARGOS
# =============================================================================


def build_pcmso_role_index(roles_df: pd.DataFrame):
    exact = defaultdict(list)

    for _, row in roles_df.iterrows():
        role = str(row["Cargo PCMSO"]).strip()
        norm = normalize_role(role)
        if norm:
            exact[norm].append(row.to_dict())

    return exact


# =============================================================================
# MATCHING DE EXAMES
# =============================================================================


def normalize_exam(text: str) -> str:
    text = strip_accents("" if text is None else str(text)).upper()
    text = text.replace("–", "-").replace("—", "-")
    text = text.replace("C/", "COM ")
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return normalize_spaces(text)


EXAM_EXPLICIT_ALIASES = {
    # Nomenclaturas exatamente observadas no relatório do Benner.
    "AUDIOMETRIA TONAL OCUPACIONAL": "AUDIOMETRIA",
    "AVAL PSICOLOGICA PSICOSSOCIAL": "AVAL_PSI_PSICOSSOCIAL",
    "AVALIACAO PSICOLOGICO": "AVAL_PSI_PSICOSSOCIAL",
    "AVALIACAO PSICOLOGICA": "AVAL_PSI_PSICOSSOCIAL",
    "AVALIACAO DA ACUIDADE VISUAL": "ACUIDADE_VISUAL",
    "CADMIO URINA": "CADMIO",
    "CHUMBO URINA": "CHUMBO",
    "CROMO URINA": "CROMO",
    "ECG ELETROCARDIOGRAMA": "ECG",
    "EEG ELETROENCEFALOGRAMA ROTINA": "EEG",
    "ESPIROMETRIA": "ESPIROMETRIA",
    "EXAME CLINICO": "CLINICO",
    "GAMA GT": "GAMA_GT",
    "GLICEMIA": "GLICEMIA",
    "HEMOGRAMA COMPLETO COM PLAQUETA": "HEMOGRAMA",
    "MANGANES URINARIO": "MANGANES",
    "RAIO X COLUNA CERVICAL": "RX_CERVICAL",
    "RAIO X COLUNA LOMBO SACRA": "RX_LOMBO",
    "RAIO X DE QUADRIL": "RX_QUADRIL",
    "RAIO X DO TORAX PA OIT": "RX_TORAX",
    "TESTE DE ROMBERG": "ROMBERG",
    "TGO TRANSAMINASE OXALACETICA": "TGO",
    "TGP TRANSAMINASE PIRUVICA": "TGP",
    "TIPAGEM SANGUE E FATOR RH": "TIPAGEM_SANG",
    "CONTAGEM DE RETICULOCITOS": "RETICULOCITOS",
    "CONTAGEM DE RETICULOCITOS PERCENTUAL E ABSOLUTA": "RETICULOCITOS",
    "ACIDO TRANS TRANS MUCONICO URINARIO": "ACIDO_TRANS_TRANS_MUCONICO",
}


def exam_key(text: str) -> str:
    n = normalize_exam(text)

    # Primeiro reconhece os dois exames adicionados ao PCMSO da Braskem.
    # O Benner pode exibir um código numérico antes do nome do exame; por isso
    # usamos palavras-chave e não dependemos de igualdade exata.
    if "RETICULOCITO" in n:
        return "RETICULOCITOS"
    if "MUCONICO" in n and "TRANS" in n and "ACIDO" in n:
        return "ACIDO_TRANS_TRANS_MUCONICO"

    # Depois aplica aliases explícitos do Benner/PCMSO.
    if n in EXAM_EXPLICIT_ALIASES:
        return EXAM_EXPLICIT_ALIASES[n]

    # Mais específicos primeiro.
    if "ACIDO METIL HIPURICO" in n or "ACIDO METILHIPURICO" in n:
        return "ACIDO_METIL_HIPURICO"
    if "ACIDO HIPURICO" in n:
        return "ACIDO_HIPURICO"

    if "ANTI HBSAG" in n:
        return "ANTI_HBSAG"
    if re.search(r"\bANTI HBC\b", n):
        return "ANTI_HBC"
    if re.search(r"\bANTI HCV\b", n):
        return "ANTI_HCV"
    if re.search(r"\bANTI HBS\b", n):
        return "ANTI_HBS"

    if "P AMINOFENOL" in n:
        return "P_AMINOFENOL"
    if "METAHEMOGLOBINA" in n:
        return "METAHEMOGLOBINA"

    if "TGO" in n and "TGP" in n:
        return "TGO_TGP"
    if re.search(r"\bTGO\b", n):
        return "TGO"
    if re.search(r"\bTGP\b", n):
        return "TGP"

    if "CADMIO" in n:
        return "CADMIO"
    if "MANGANES" in n:
        return "MANGANES"
    if "CROMO" in n:
        return "CROMO"
    if "CHUMBO" in n:
        return "CHUMBO"

    if "URINA" in n and "ACETONA" in n:
        return "URINA_ACETONA"

    if "CLINICO" in n and "PSICOLOG" not in n:
        return "CLINICO"
    if "AUDIOMETRIA" in n:
        return "AUDIOMETRIA"
    if "PSICOLOGICA" in n and "PSICOSSOCIAL" in n:
        return "AVAL_PSI_PSICOSSOCIAL"
    if "ACUIDADE" in n and "VISUAL" in n:
        return "ACUIDADE_VISUAL"
    if "HEMOGRAMA" in n:
        return "HEMOGRAMA"
    if "TIPAGEM" in n and "SANG" in n:
        return "TIPAGEM_SANG"
    if "ELETROCARDIOGRAMA" in n or re.search(r"\bECG\b", n):
        return "ECG"
    if "ELETROENCEFALOGRAMA" in n or re.search(r"\bEEG\b", n):
        return "EEG"
    if "ESPIROMETRIA" in n:
        return "ESPIROMETRIA"
    if "GAMA GT" in n:
        return "GAMA_GT"
    if "GLICEMIA" in n:
        return "GLICEMIA"
    if "ROMBERG" in n:
        return "ROMBERG"
    if "ISHIHARA" in n:
        return "ISHIHARA"

    if "RAIO X" in n or "RAIO X" in n.replace("-", " "):
        if "TORAX" in n:
            return "RX_TORAX"
        if "QUADRIL" in n:
            return "RX_QUADRIL"
        if "CERVICAL" in n:
            return "RX_CERVICAL"
        if "LOMBO" in n:
            return "RX_LOMBO"

    return ""


PCMSO_GROUPS = {
    "TGO_TGP": {
        "label": "TGO + TGP",
        "mode": "AND",
        "components": ["TGO", "TGP"],
    },
    "URINA_ACIDOS": {
        "label": "URINA: ÁCIDO HIPÚRICO E METIL HIPÚRICO",
        "mode": "AND",
        "components": ["ACIDO_HIPURICO", "ACIDO_METIL_HIPURICO"],
    },
    "SOROLOGIAS_HEPATITE": {
        "label": "ANTI HBS – ANTI HBC – ANTI HCV - HBsAg",
        "mode": "AND",
        "components": [
            "ANTI_HBS",
            "ANTI_HBC",
            "ANTI_HCV",
            "ANTI_HBSAG",
        ],
    },
    "URINA_METAIS": {
        "label": "URINA (CADMIO, MANGANES, CROMO, CHUMBO)",
        "mode": "AND",
        "components": [
            "CADMIO",
            "MANGANES",
            "CROMO",
            "CHUMBO",
        ],
    },
    "P_AMINOFENOL_OU_META": {
        "label": "P- AMINOFENOL OU METAHEMOGLOBINA SÉRIA",
        "mode": "OR",
        "components": ["P_AMINOFENOL", "METAHEMOGLOBINA"],
    },
    "CHUMBO_OU": {
        "label": "CHUMBO SÉRICO OU URINARIO",
        "mode": "OR",
        "components": ["CHUMBO"],
    },
}


def requirement_for_exam(pcmso_exam: str):
    key = exam_key(pcmso_exam)

    # Detecta os grupos compostos pelo nome original.
    n = normalize_exam(pcmso_exam)

    # Alguns GHEs agrupam os três exames em uma única linha do PCMSO:
    # GAMA GT + TGO + TGP. Nesses casos, o Benner pode ter três cadastros
    # separados, e todos devem ser considerados parte do requisito composto.
    if "GAMA GT" in n and "TGO" in n and "TGP" in n:
        return "AND", ["GAMA_GT", "TGO", "TGP"]

    # Outros GHEs trazem apenas TGO + TGP; nesses casos GAMA GT não é exigido
    # por essa linha e continua sendo validado separadamente, se previsto.
    if "TGO" in n and "TGP" in n:
        return "AND", ["TGO", "TGP"]

    if (
        "URINA" in n
        and "ACIDO HIPURICO" in n
        and "METIL HIPURICO" in n
    ):
        return "AND", ["ACIDO_HIPURICO", "ACIDO_METIL_HIPURICO"]

    if (
        "ANTI HBS" in n
        and "ANTI HBC" in n
        and "ANTI HCV" in n
        and "HBSAG" in n
    ):
        return "AND", ["ANTI_HBS", "ANTI_HBC", "ANTI_HCV", "ANTI_HBSAG"]

    if (
        "URINA" in n
        and "CADMIO" in n
        and "MANGANES" in n
        and "CROMO" in n
        and "CHUMBO" in n
    ):
        return "AND", ["CADMIO", "MANGANES", "CROMO", "CHUMBO"]

    if "P AMINOFENOL" in n and "METAHEMOGLOBINA" in n:
        return "OR", ["P_AMINOFENOL", "METAHEMOGLOBINA"]

    if "CHUMBO" in n and "SERICO" in n and "URINARIO" in n:
        return "OR", ["CHUMBO"]

    if key:
        return "SINGLE", [key]

    return "UNKNOWN", []


def pcmso_allows_benner_exam(
    cargo: str,
    exam_key_benner: str,
    aso_key: str,
    matched_role_groups,
    exams_by_group,
) -> Tuple[bool, bool]:
    """
    Confere diretamente no PCMSO se o exame do Benner:
      1) existe entre os exames previstos para o grupo do cargo;
      2) está previsto para o ASO analisado.

    O segundo retorno diferencia "exame não previsto" de
    "exame previsto, mas quadrinho não previsto para aquele ASO".
    """
    found_in_pcmso = False
    allowed_for_aso = False

    field_map = {
        "admissional": "Admissao",
        "periodico": "Periodicidade",
        "mudanca_funcao": "Mudanca_Risco_Retorno",
        "demissional": "Demissao",
    }
    pcmso_field = field_map.get(aso_key, "")

    for group_key in matched_role_groups.get(cargo, set()):
        group_rows = exams_by_group.get(
            (str(group_key[0]), str(group_key[1])),
            [],
        )

        for exam_row in group_rows:
            mode, components = requirement_for_exam(
                str(exam_row["Exame PCMSO"]).strip()
            )

            if exam_key_benner not in components:
                continue

            found_in_pcmso = True

            if pcmso_field and pcms_is_required(
                exam_row.get(pcmso_field, "")
            ):
                allowed_for_aso = True

    return found_in_pcmso, allowed_for_aso


def build_benner_exam_map(benner_df: pd.DataFrame):
    """
    Mapa:
      cargo normalizado -> exame canonical -> lista de linhas do Benner.
    """
    result = defaultdict(lambda: defaultdict(list))

    for _, row in benner_df.iterrows():
        cargo = str(row["Cargo"]).strip()
        exame = str(row["Exame"]).strip()

        if not cargo or not exame:
            continue

        key = exam_key(exame)
        if not key:
            continue

        result[normalize_role(cargo)][key].append(row.to_dict())

    return result


def get_component_rows(
    benner_exam_map,
    cargo: str,
    component: str,
):
    role_key = normalize_role(cargo)
    return benner_exam_map.get(role_key, {}).get(component, [])


def component_display(
    benner_exam_map,
    cargo: str,
    component_keys: List[str],
):
    texts = []
    for key in component_keys:
        rows = get_component_rows(benner_exam_map, cargo, key)
        if rows:
            for row in rows:
                text = str(row.get("Exame", "")).strip()
                if text and text not in texts:
                    texts.append(text)
    return " + ".join(texts)


def check_requirement(
    pcmso_exam: str,
    cargo: str,
    benner_exam_map,
    required_columns: Dict[str, bool],
    pcmso_values: Dict[str, str],
):
    mode, components = requirement_for_exam(pcmso_exam)

    if mode == "UNKNOWN":
        return {
            "known": False,
            "matched": False,
            "components": [],
            "display": "",
            "reason": "EXAME PCMSO NÃO MAPEADO",
        }

    result = {
        "known": True,
        "matched": True,
        "components": [],
        "display": "",
        "reason": "",
    }

    # Cada componente é conferido nos ASOs posteriormente.
    for component in components:
        rows = get_component_rows(
            benner_exam_map,
            cargo,
            component,
        )
        result["components"].append(
            {
                "key": component,
                "rows": rows,
            }
        )

    result["display"] = component_display(
        benner_exam_map,
        cargo,
        components,
    )

    # Presença do exame em si.
    present = [bool(x["rows"]) for x in result["components"]]

    if mode == "AND":
        result["matched"] = all(present)
    elif mode == "OR":
        result["matched"] = any(present)
    else:
        result["matched"] = bool(present and present[0])

    if not result["matched"]:
        result["reason"] = "EXAME NÃO LOCALIZADO NO BENNER"

    return result


# =============================================================================
# AUDITORIA
# =============================================================================


PCMSO_LOGICAL_ASOS = {
    "admissional": "Admissao",
    "mudanca_funcao": "Mudanca_Risco_Retorno",
    "periodico": "Periodicidade",
    "demissional": "Demissao",
}


def pcms_is_required(value) -> bool:
    return bool(str(value or "").strip())


def audit_exams(
    benner_df: pd.DataFrame,
    pcmso_matrix: pd.DataFrame,
    pcmso_roles: pd.DataFrame,
):
    benner_exam_map = build_benner_exam_map(benner_df)

    role_index = build_pcmso_role_index(pcmso_roles)
    role_norms = sorted(role_index.keys())

    # Agrupa os exames por GHE/Setor e aplica aos cargos do respectivo GHE.
    group_info = {}
    for _, row in pcmso_roles.iterrows():
        key = (
            str(row["GHE"]),
            str(row["Setor"]),
            str(row["Cargo original no PCMSO"]),
        )
        group_info[key] = {
            "GHE": row["GHE"],
            "Setor": row["Setor"],
            "Cargo original no PCMSO": row["Cargo original no PCMSO"],
        }

    # Determina o conjunto de exames do GHE.
    exams_by_group = defaultdict(list)
    for _, row in pcmso_matrix.iterrows():
        key = (
            str(row["GHE"]),
            str(row["Setor"]),
        )
        exams_by_group[key].append(row.to_dict())

    # Cargos do sistema.
    benner_roles = sorted(
        {
            str(x).strip()
            for x in benner_df["Cargo"].dropna().unique()
            if str(x).strip()
        }
    )

    detailed_rows = []
    pending_rows = []
    extra_rows = []
    extra_cargo_rows = []
    role_status_rows = []
    missing_pcmso_roles = []

    matched_role_groups = {}

    for cargo in benner_roles:
        matches, role_conf = match_benner_role_to_pcmso_role(
            cargo,
            role_index,
            role_norms,
        )

        # Mantém os grupos únicos.
        group_keys = {
            (
                item["GHE"],
                item["Setor"],
                item["Cargo original no PCMSO"],
            )
            for item in matches
        }

        if not group_keys:
            continue

        matched_role_groups[cargo] = group_keys

        overall_missing = False

        # Um cargo do Benner pode cair em mais de um grupo somente quando
        # existem equivalências explícitas. Consolidamos todos os exames.
        group_exam_rows = []
        for group_key in group_keys:
            ghe = group_key[0]
            setor = group_key[1]
            group_exam_rows.extend(
                exams_by_group.get((str(ghe), str(setor)), [])
            )

        # Dedupe por exame dentro dos grupos correspondentes.
        exam_map = {}
        for exam_row in group_exam_rows:
            key = (
                normalize_exam(str(exam_row["Exame PCMSO"])),
                str(exam_row["Exame PCMSO"]),
            )
            exam_map[key] = exam_row

        for exam_row in exam_map.values():
            pcmso_exam = str(exam_row["Exame PCMSO"]).strip()

            fields = {
                "admissional": str(exam_row.get("Admissao", "") or "").strip(),
                "periodico": str(exam_row.get("Periodicidade", "") or "").strip(),
                "mudanca_funcao": str(
                    exam_row.get("Mudanca_Risco_Retorno", "") or ""
                ).strip(),
                "demissional": str(exam_row.get("Demissao", "") or "").strip(),
            }

            requirement = check_requirement(
                pcmso_exam,
                cargo,
                benner_exam_map,
                {},
                fields,
            )

            for aso_key, pcmso_field in PCMSO_LOGICAL_ASOS.items():
                pc_required = pcms_is_required(fields[aso_key])

                # O PCMSO agrupa Mudança de Risco Ocupacional e Retorno.
                # No relatório do Benner, R. Trabalho está saindo com valores
                # numéricos/códigos inconsistentes. Por isso ele é ignorado por
                # completo: somente M. Função participa da conferência.
                if aso_key == "mudanca_funcao":
                    benner_change_available = not benner_df["ASO_mudanca_funcao"].isna().all()

                    change_values = [
                        row.get("ASO_mudanca_funcao")
                        for comp in requirement["components"]
                        for row in comp["rows"]
                    ]

                    change_marked = any(is_marked(v) for v in change_values)

                    if pc_required:
                        if not benner_change_available:
                            status = "NÃO CONFERIDO — COLUNA AUSENTE"
                        elif change_marked:
                            status = "OK"
                        else:
                            status = "FALTA"
                            overall_missing = True
                    else:
                        # Nunca gerar EXTRA por causa da coluna R. Trabalho.
                        status = "EXTRA NO BENNER" if change_marked else "OK"

                    benner_display = (
                        "Mud. Função="
                        + ("S" if change_marked else "N")
                        + " | R. Trabalho=IGNORADO"
                    )

                else:
                    column = f"ASO_{aso_key}"
                    column_available = not benner_df[column].isna().all()

                    values = [
                        row.get(column)
                        for comp in requirement["components"]
                        for row in comp["rows"]
                    ]

                    if pc_required:
                        if not column_available:
                            status = "NÃO CONFERIDO — COLUNA AUSENTE"
                        else:
                            # Para AND, o exame/grupo precisa estar marcado
                            # em cada componente. Para OR, basta uma alternativa.
                            if requirement["known"]:
                                if (
                                    len(requirement["components"]) == 1
                                    and not requirement["components"][0]["rows"]
                                ):
                                    status = "FALTA"
                                else:
                                    component_ok = []
                                    for comp in requirement["components"]:
                                        comp_rows = comp["rows"]
                                        marked = any(
                                            is_marked(row.get(column))
                                            for row in comp_rows
                                        )
                                        component_ok.append(marked)

                                    mode, _ = requirement_for_exam(pcmso_exam)

                                    if mode == "AND":
                                        ok = all(component_ok)
                                    elif mode == "OR":
                                        ok = any(component_ok)
                                    else:
                                        ok = any(component_ok)

                                    status = "OK" if ok else "FALTA"

                                    if status == "FALTA":
                                        overall_missing = True
                            else:
                                status = "NÃO CONFERIDO — EXAME NÃO MAPEADO"
                    else:
                        if not column_available:
                            status = "OK"
                        elif any(is_marked(v) for v in values):
                            status = "EXTRA NO BENNER"
                        else:
                            status = "OK"

                    marked_text = "S" if any(is_marked(v) for v in values) else "N"
                    benner_display = marked_text

                detail = {
                    "Cargo": cargo,
                    "GHE": " | ".join(
                        sorted({str(k[0]) for k in group_keys})
                    ),
                    "Setor": " | ".join(
                        sorted({str(k[1]) for k in group_keys})
                    ),
                    "Confiança cargo": round(float(role_conf), 1),
                    "Exame PCMSO": pcmso_exam,
                    "Admissão PCMSO": fields["admissional"],
                    "Mud. Risco / Retorno PCMSO": fields["mudanca_funcao"],
                    "Periodicidade PCMSO": fields["periodico"],
                    "Demissão PCMSO": fields["demissional"],
                    "Exame no Benner": requirement["display"],
                    "ASO conferido": pcmso_field,
                    "Benner": benner_display,
                    "Status": status,
                }
                detailed_rows.append(detail)

                if status.startswith("FALTA") or status.startswith("NÃO CONFERIDO"):
                    pending_rows.append({
                        **detail,
                        "Problema": status,
                    })

        role_status_rows.append({
            "Cargo": cargo,
            "GHE/Setor": " | ".join(
                sorted({f"{k[0]} — {k[1]}" for k in group_keys})
            ),
            "Confiança": round(float(role_conf), 1),
            "Status": "FALHAS" if overall_missing else "ATENDIDO",
        })

    # Cargos existentes no PCMSO que não foram localizados no Benner.
    benner_role_norms = set()
    benner_roles_by_norm = defaultdict(list)
    for cargo in benner_roles:
        for alias in expand_role_names(cargo):
            benner_role_norms.add(alias)
            benner_roles_by_norm[alias].append(cargo)

    for _, row in pcmso_roles.drop_duplicates(
        subset=["Cargo PCMSO", "GHE", "Setor"]
    ).iterrows():
        role = str(row["Cargo PCMSO"]).strip()
        role_norm = normalize_role(role)
        if not role_norm:
            continue

        covered = False
        for alias in expand_role_names(role):
            if alias in benner_role_norms:
                covered = True
                break

        if not covered:
            results = process.extract(
                role_norm,
                list(benner_role_norms),
                scorer=fuzz.token_set_ratio,
                limit=1,
            )
            suggestion = ""
            if results and results[0][1] >= 92:
                suggestion = str(results[0][0])

            missing_pcmso_roles.append({
                "Cargo PCMSO": role,
                "GHE": row["GHE"],
                "Setor": row["Setor"],
                "Possível correspondente": suggestion,
                "Status": "CARGO DO PCMSO NÃO ENCONTRADO NO BENNER",
            })

    # CARGOS EXTRAS:
    # Cargos existentes no relatório do Benner que não encontraram nenhum
    # grupo/cargo correspondente no PCMSO ficam em uma aba separada.
    # Todos os exames cadastrados para esses cargos são listados ali e não
    # poluem a aba de extras de exames dos cargos que têm correspondência.
    extra_cargo_names = set(benner_roles) - set(matched_role_groups.keys())
    if extra_cargo_names:
        unmatched_df = benner_df[benner_df["Cargo"].astype(str).str.strip().isin(extra_cargo_names)]
        for _, row in unmatched_df.iterrows():
            extra_cargo_rows.append({
                "Projeto": str(row.get("Projeto", "") or "").strip(),
                "Cargo no Benner": str(row.get("Cargo", "") or "").strip(),
                "Exame no Benner": str(row.get("Exame", "") or "").strip(),
                "Admissional": str(row.get("ASO_admissional", "") or "").strip(),
                "Demissional": str(row.get("ASO_demissional", "") or "").strip(),
                "Periódico": str(row.get("ASO_periodico", "") or "").strip(),
                "Mudança de Função": str(row.get("ASO_mudanca_funcao", "") or "").strip(),
                "Status": "CARGO DO BENNER NÃO LOCALIZADO NO PCMSO",
            })

    # Extras em cargos que possuem correspondência no PCMSO:
    # exame não previsto para o cargo ou quadrinho marcado onde o PCMSO está em branco.
    required_by_role = defaultdict(dict)

    for cargo, group_keys in matched_role_groups.items():
        combined_exam_rows = []
        for key in group_keys:
            combined_exam_rows.extend(
                exams_by_group.get((str(key[0]), str(key[1])), [])
            )

        for exam_row in combined_exam_rows:
            exam = str(exam_row["Exame PCMSO"]).strip()
            mode, components = requirement_for_exam(exam)
            for component in components:
                requirement_entry = required_by_role[cargo].setdefault(
                    component,
                    {
                        "admissional": False,
                        "periodico": False,
                        "mudanca_funcao": False,
                        "demissional": False,
                        "or_groups": [],
                    },
                )

                if (
                    mode == "OR"
                ):
                    requirement_entry["or_groups"].append(
                        {
                            "exam": exam,
                            "components": components,
                        }
                    )
                else:
                    requirement_entry["admissional"] |= pcms_is_required(
                        exam_row.get("Admissao", "")
                    )
                    requirement_entry["periodico"] |= pcms_is_required(
                        exam_row.get("Periodicidade", "")
                    )
                    requirement_entry["mudanca_funcao"] |= pcms_is_required(
                        exam_row.get(
                            "Mudanca_Risco_Retorno",
                            "",
                        )
                    )
                    requirement_entry["demissional"] |= pcms_is_required(
                        exam_row.get("Demissao", "")
                    )

    for cargo in benner_roles:
        # Cargos sem correspondência no PCMSO já estão em "Cargos extras".
        # Não duplicar todos os seus exames na tabela de extras por exame.
        if cargo not in matched_role_groups:
            continue

        role_key = normalize_role(cargo)
        cargo_map = benner_exam_map.get(role_key, {})

        for exam_key_benner, rows in cargo_map.items():
            for row in rows:
                exam_name = str(row.get("Exame", "")).strip()

                # Exames fora do vocabulário conhecido continuam sendo
                # reportados como extras.
                if not exam_key_benner:
                    extra_rows.append({
                        "Cargo": cargo,
                        "Exame no Benner": exam_name,
                        "ASO": "Exame",
                        "Valor": "",
                        "Motivo": "EXAME NÃO MAPEADO / NÃO IDENTIFICADO NO VOCABULÁRIO",
                    })
                    continue

                # O cruzamento é feito diretamente contra os exames do grupo
                # do cargo no PCMSO. Isso evita falsos extras causados por
                # nomenclaturas compostas ou pequenas diferenças de cargo.
                for aso_key, label in [
                    ("admissional", "Admissional"),
                    ("periodico", "Periódico"),
                    ("mudanca_funcao", "Mudança de Função"),
                    ("demissional", "Demissional"),
                ]:
                    value = row.get(f"ASO_{aso_key}")

                    if not is_marked(value):
                        continue

                    found_in_pcmso, allowed_for_aso = pcmso_allows_benner_exam(
                        cargo,
                        exam_key_benner,
                        aso_key,
                        matched_role_groups,
                        exams_by_group,
                    )

                    if not found_in_pcmso:
                        extra_rows.append({
                            "Cargo": cargo,
                            "Exame no Benner": exam_name,
                            "ASO": label,
                            "Valor": str(value),
                            "Motivo": "EXAME NÃO PREVISTO NO PCMSO PARA O CARGO",
                        })
                    elif not allowed_for_aso:
                        extra_rows.append({
                            "Cargo": cargo,
                            "Exame no Benner": exam_name,
                            "ASO": label,
                            "Valor": str(value),
                            "Motivo": "QUADRINHO MARCADO NO BENNER, MAS EM BRANCO NO PCMSO",
                        })

    detailed_df = pd.DataFrame(detailed_rows)
    pending_df = pd.DataFrame(pending_rows)
    extras_df = pd.DataFrame(extra_rows)
    extra_cargos_df = pd.DataFrame(
        extra_cargo_rows,
        columns=[
            "Projeto",
            "Cargo no Benner",
            "Exame no Benner",
            "Admissional",
            "Demissional",
            "Periódico",
            "Mudança de Função",
            "Status",
        ],
    )
    role_status_df = pd.DataFrame(role_status_rows)
    missing_roles_df = pd.DataFrame(missing_pcmso_roles)

    summary = {
        "cargos_benner": len(benner_roles),
        "cargos_pcmso": int(pcmso_roles["Cargo PCMSO"].nunique()),
        "cargos_pcmso_nao_encontrados": len(missing_roles_df),
        "linhas_exames": len(detailed_df),
        "pendencias": int(
            (detailed_df["Status"] == "FALTA").sum()
        ) if not detailed_df.empty else 0,
        "nao_conferidos": int(
            detailed_df["Status"].astype(str).str.startswith("NÃO CONFERIDO").sum()
        ) if not detailed_df.empty else 0,
        "extras_benner": len(extras_df),
        "cargos_extras": int(len(extra_cargo_names)),
        "linhas_cargos_extras": int(len(extra_cargos_df)),
        "atendidos": int(
            (role_status_df["Status"] == "ATENDIDO").sum()
        ) if not role_status_df.empty else 0,
        "faltas_cargo": int(
            (role_status_df["Status"] == "FALHAS").sum()
        ) if not role_status_df.empty else 0,
        "demissional_disponivel": not benner_df["ASO_demissional"].isna().all(),
    }

    return (
        summary,
        role_status_df,
        detailed_df,
        pending_df,
        extras_df,
        extra_cargos_df,
        missing_roles_df,
    )


def match_benner_role_to_pcmso_role(
    cargo: str,
    role_index,
    role_norms,
):
    aliases = expand_role_names(cargo)
    matches = {}

    for alias in aliases:
        for item in role_index.get(alias, []):
            key = (
                item["GHE"],
                item["Setor"],
                item["Cargo original no PCMSO"],
            )
            matches[key] = item

    if matches:
        return list(matches.values()), 100.0

    best_score = 0.0
    for alias in aliases:
        results = process.extract(
            alias,
            role_norms,
            scorer=fuzz.token_set_ratio,
            limit=5,
        )

        for candidate, score, _ in results:
            score2 = fuzz.ratio(alias, candidate)
            final_score = max(float(score), float(score2))
            best_score = max(best_score, final_score)

            alias_first = alias.split()[0] if alias.split() else ""
            candidate_first = candidate.split()[0] if candidate.split() else ""

            if (
                final_score >= 92
                and alias_first
                and candidate_first
                and alias_first == candidate_first
            ):
                for item in role_index.get(candidate, []):
                    key = (
                        item["GHE"],
                        item["Setor"],
                        item["Cargo original no PCMSO"],
                    )
                    matches[key] = item

    return list(matches.values()), best_score


# =============================================================================
# EXPORTAÇÃO
# =============================================================================


def build_excel(
    summary,
    role_status_df,
    detailed_df,
    pending_df,
    extras_df,
    extra_cargos_df,
    missing_roles_df,
):
    output = io.BytesIO()

    summary_df = pd.DataFrame([{
        "Cargos no Benner": summary["cargos_benner"],
        "Cargos no PCMSO": summary["cargos_pcmso"],
        "Cargos do PCMSO não encontrados no Benner": summary[
            "cargos_pcmso_nao_encontrados"
        ],
        "Linhas de exames auditadas": summary["linhas_exames"],
        "Faltas": summary["pendencias"],
        "Não conferidos": summary["nao_conferidos"],
        "Exames extras em cargos correspondentes": summary["extras_benner"],
        "Cargos extras no Benner (sem correspondência no PCMSO)": summary["cargos_extras"],
        "Linhas listadas em Cargos extras": summary["linhas_cargos_extras"],
        "Cargos atendidos": summary["atendidos"],
        "Cargos com falhas": summary["faltas_cargo"],
        "Coluna Demissional disponível": (
            "SIM" if summary["demissional_disponivel"] else "NÃO"
        ),
    }])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="Resumo", index=False)
        role_status_df.to_excel(writer, sheet_name="Cargos", index=False)
        detailed_df.to_excel(writer, sheet_name="Detalhado", index=False)
        pending_df.to_excel(writer, sheet_name="Pendências", index=False)
        extras_df.to_excel(writer, sheet_name="Extras Benner", index=False)
        extra_cargos_df.to_excel(writer, sheet_name="Cargos extras", index=False)
        missing_roles_df.to_excel(
            writer,
            sheet_name="Cargos ausentes",
            index=False,
        )

        workbook = writer.book

        header_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#17365D",
            "font_color": "white",
            "border": 1,
        })
        ok_fmt = workbook.add_format({"bg_color": "#E2F0D9"})
        fail_fmt = workbook.add_format({"bg_color": "#FCE4D6"})
        warn_fmt = workbook.add_format({"bg_color": "#FFF2CC"})

        sheets = {
            "Resumo": summary_df,
            "Cargos": role_status_df,
            "Detalhado": detailed_df,
            "Pendências": pending_df,
            "Extras Benner": extras_df,
            "Cargos extras": extra_cargos_df,
            "Cargos ausentes": missing_roles_df,
        }

        for sheet_name, frame in sheets.items():
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)

            if len(frame.columns) > 0:
                ws.autofilter(
                    0,
                    0,
                    max(1, len(frame)),
                    len(frame.columns) - 1,
                )

            for idx, col in enumerate(frame.columns):
                width = min(
                    max(len(str(col)) + 2, 14),
                    45,
                )

                if not frame.empty:
                    sample = frame[col].astype(str).head(100)
                    if not sample.empty:
                        width = min(
                            max(width, int(sample.map(len).max()) + 2),
                            60,
                        )

                ws.set_column(idx, idx, width)
                ws.write(0, idx, col, header_fmt)

        if not detailed_df.empty and "Status" in detailed_df.columns:
            ws = writer.sheets["Detalhado"]
            status_col = detailed_df.columns.get_loc("Status")

            for row_idx, value in enumerate(
                detailed_df["Status"].astype(str),
                start=1,
            ):
                if value == "OK":
                    ws.write(row_idx, status_col, value, ok_fmt)
                elif value.startswith("FALTA"):
                    ws.write(row_idx, status_col, value, fail_fmt)
                elif value.startswith("NÃO CONFERIDO"):
                    ws.write(row_idx, status_col, value, warn_fmt)
                elif value == "EXTRA NO BENNER":
                    ws.write(row_idx, status_col, value, warn_fmt)

        if not pending_df.empty and "Problema" in pending_df.columns:
            ws = writer.sheets["Pendências"]
            problem_col = pending_df.columns.get_loc("Problema")

            for row_idx, value in enumerate(
                pending_df["Problema"].astype(str),
                start=1,
            ):
                if value.startswith("FALTA"):
                    ws.write(row_idx, problem_col, value, fail_fmt)
                else:
                    ws.write(row_idx, problem_col, value, warn_fmt)

    output.seek(0)
    return output.getvalue()


# =============================================================================
# INTERFACE
# =============================================================================

st.title("🩺 Auditor de Exames — PCMSO × Benner")
st.markdown(
    "O aplicativo compara os exames previstos no **PCMSO** com os exames "
    "cadastrados no **Benner**, inclusive os quadrinhos de Admissional, "
    "Periódico, Mudança de Função, Retorno ao Trabalho e Demissional."
)

with st.sidebar:
    st.header("Arquivos")

    benner_file = st.file_uploader(
        "1. Relatório do Benner",
        type=["xlsx", "xls"],
        key="benner",
    )

    pcmso_file = st.file_uploader(
        "2. PCMSO",
        type=["docx"],
        key="pcmso",
    )

if not benner_file or not pcmso_file:
    st.info(
        "Envie o relatório do Benner (.xlsx) e o PCMSO (.docx) para iniciar."
    )
    st.stop()

try:
    benner = read_benner(benner_file.getvalue())
    pcmso_matrix, pcmso_roles, logical_headers = read_pcmso(
        pcmso_file.getvalue()
    )
except Exception as exc:
    st.error(f"Não foi possível ler os arquivos: {exc}")
    st.stop()

# Se houver mais de um projeto, permite selecionar.
projects = get_projects(benner)
selected_project = None

if len(projects) > 1:
    selected_project = st.sidebar.selectbox(
        "Projeto do relatório",
        ["Todos"] + projects,
    )

if selected_project and selected_project != "Todos":
    benner = benner[
        benner["Projeto"].astype(str).str.strip() == selected_project
    ].copy()

if benner.empty:
    st.warning("Nenhuma linha do Benner ficou disponível para o projeto selecionado.")
    st.stop()

with st.spinner("Conferindo exames do PCMSO com o Benner..."):
    (
        summary,
        role_status_df,
        detailed_df,
        pending_df,
        extras_df,
        extra_cargos_df,
        missing_roles_df,
    ) = audit_exams(
        benner,
        pcmso_matrix,
        pcmso_roles,
    )

# Avisos importantes.
if not summary["demissional_disponivel"]:
    st.warning(
        "O relatório do Benner enviado não possui coluna de Demissional. "
        "As exigências de Demissão do PCMSO ficam como "
        "'NÃO CONFERIDO — COLUNA AUSENTE'."
    )

m1, m2, m3, m4 = st.columns(4)
m1.metric("Cargos no Benner", summary["cargos_benner"])
m2.metric("Cargos no PCMSO", summary["cargos_pcmso"])
m3.metric("Cargos atendidos", summary["atendidos"])
m4.metric("Cargos extras", summary["cargos_extras"])

m5, m6, m7 = st.columns(3)
m5.metric("Faltas", summary["pendencias"])
m6.metric("Não conferidos", summary["nao_conferidos"])
m7.metric("Exames extras em cargos correspondentes", summary["extras_benner"])

st.divider()

if summary["pendencias"] == 0:
    st.success(
        "Nenhuma falta foi encontrada nos campos que puderam ser conferidos."
    )
else:
    st.error(
        f"Foram encontradas {summary['pendencias']} faltas para conferência."
    )

if summary["nao_conferidos"] > 0:
    st.warning(
        f"{summary['nao_conferidos']} conferências ficaram como 'NÃO CONFERIDO' "
        "porque o relatório do Benner não trouxe todos os campos necessários."
    )

# As informações mais importantes ficam na primeira aba. Os cargos que não
# possuem correspondência no PCMSO ficam separados para não poluir a auditoria.
tab_auditoria, tab_cargos_extras, tab_exames_extras, tab_cargos_ausentes = st.tabs([
    "Conferência PCMSO × Benner",
    "Cargos extras",
    "Exames extras em cargos correspondentes",
    "Cargos do PCMSO ausentes no Benner",
])

with tab_auditoria:
    st.subheader("🔎 Detalhado — exames previstos no PCMSO × Benner")
    if not detailed_df.empty:
        cargos = st.multiselect(
            "Filtrar cargo",
            sorted(detailed_df["Cargo"].unique()),
            key="filtro_cargo_auditoria",
        )

        status_filter = st.multiselect(
            "Filtrar status",
            [
                "OK",
                "FALTA",
                "NÃO CONFERIDO — COLUNA AUSENTE",
                "NÃO CONFERIDO — EXAME NÃO MAPEADO",
                "EXTRA NO BENNER",
            ],
            default=["FALTA", "NÃO CONFERIDO — COLUNA AUSENTE"],
            key="filtro_status_auditoria",
        )

        view = detailed_df.copy()
        if cargos:
            view = view[view["Cargo"].isin(cargos)]
        if status_filter:
            view = view[view["Status"].isin(status_filter)]

        st.dataframe(view, use_container_width=True, hide_index=True)
    else:
        st.info("Não há linhas para exibir na auditoria.")

    st.subheader("📋 Status por cargo")
    st.dataframe(role_status_df, use_container_width=True, hide_index=True)

    with st.expander("⚠️ Pendências / não conferidos"):
        if pending_df.empty:
            st.success("Nenhuma pendência.")
        else:
            st.dataframe(pending_df, use_container_width=True, hide_index=True)

with tab_cargos_extras:
    st.subheader("Cargos cadastrados no Benner sem correspondência no PCMSO")
    st.caption(
        "Esta aba é informativa: lista os exames cadastrados para cargos que não "
        "foram localizados no PCMSO. Eles não entram na contagem de faltas nem "
        "na lista de exames extras dos cargos correspondentes."
    )
    if extra_cargos_df.empty:
        st.success("Nenhum cargo extra foi identificado.")
    else:
        st.dataframe(extra_cargos_df, use_container_width=True, hide_index=True)

with tab_exames_extras:
    st.subheader("Exames extras em cargos que correspondem ao PCMSO")
    st.caption(
        "Aqui aparecem somente exames marcados no Benner que não estão previstos "
        "para o cargo correspondente no PCMSO, ou marcados em um ASO em que o "
        "PCMSO não prevê o exame."
    )
    if extras_df.empty:
        st.success("Nenhum exame extra foi identificado em cargos correspondentes.")
    else:
        st.dataframe(extras_df, use_container_width=True, hide_index=True)

with tab_cargos_ausentes:
    st.subheader("Cargos previstos no PCMSO não encontrados no Benner")
    if missing_roles_df.empty:
        st.success("Todos os cargos do PCMSO tiveram correspondência no Benner.")
    else:
        st.dataframe(missing_roles_df, use_container_width=True, hide_index=True)

excel_bytes = build_excel(
    summary,
    role_status_df,
    detailed_df,
    pending_df,
    extras_df,
    extra_cargos_df,
    missing_roles_df,
)

st.download_button(
    "⬇️ Baixar auditoria em Excel",
    data=excel_bytes,
    file_name="auditoria_pcmso_x_benner.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)

st.caption(
    "Regra especial: quando o PCMSO marca "
    "'Mud. Risco Ocupacional | Retorno', o aplicativo confere "
    "principalmente o campo 'Mudança de Função'. O campo 'Retorno ao Trabalho' "
    "fica apenas como informação no relatório e não gera falta nesta versão. "
    "Exames compostos, como TGO + TGP, são conferidos por seus componentes."
)
