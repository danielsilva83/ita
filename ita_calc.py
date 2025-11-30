import pandas as pd
import numpy as np

def padronizar_grr(df):
    """
    Padroniza a coluna GRR para o formato 'GRR' + números.
    """
    # Localiza a coluna que contém 'grr' no nome (insensível a maiúsculas) - logic simplified to assume "GRR" exists or find it
    # In the notebook, it checks for "GRR" specifically.
    
    if "GRR" not in df.columns:
        # Try to find a column that looks like GRR
        found = False
        for col in df.columns:
            if "GRR" in str(col).upper():
                df.rename(columns={col: "GRR"}, inplace=True)
                found = True
                break
        if not found:
             raise ValueError("Nenhuma coluna contendo 'GRR' foi encontrada no DataFrame.")

    # Limpa e padroniza os valores
    df["GRR"] = (
        df["GRR"]
        .astype(str)
        .str.extract(r"(\d+)")   # extrai apenas os números
        .dropna()                # remove valores nulos
        [0]                      # pega a primeira coluna do extract
        .astype(int)             # garante que é numérico
        .astype(str)             # converte novamente para string
        .radd("GRR")             # adiciona o prefixo 'GRR' à esquerda
    )

    return df

def normalize(series: pd.Series) -> pd.Series:
    s = series.astype(float).fillna(0.0)
    rng = s.max() - s.min()
    return (s - s.min()) / (rng + 1e-9)

def aplicar_regra_renda(
    df: pd.DataFrame,
    classe_col: str = "classe-da-renda",
    nota_col: str = "nota-da-renda",
    out_col: str = "pontuacao-renda",
) -> pd.DataFrame:
    """
    Aplica a lógica da fórmula Excel para pontuação de renda.
    """
    out = df.copy()

    # Normaliza tipos/valores
    if classe_col not in out.columns or nota_col not in out.columns:
         out[out_col] = 0
         return out

    classe = out[classe_col].astype(str).str.strip().str.upper()
    nota = pd.to_numeric(out[nota_col], errors="coerce")

    conditions = [
        (classe == "A") & (nota > 25),
        (classe == "A") & (nota.between(11, 25, inclusive="both")),
        (classe == "A") & (nota.between(1, 10, inclusive="both")),
        (classe == "B") & (nota >= 11),
        (classe == "B") & (nota.between(0, 10, inclusive="both")),
        (classe == "C") & (nota >= 11),
        (classe == "C") & (nota.between(0, 10, inclusive="both")),  # 0 a 10
        (classe == "C") & (nota < 0),
    ]
    choices = [30, 25, 20, 15, 10, 8, 5, 2]

    out[out_col] = np.select(conditions, choices, default=0).astype(int)
    return out

def _to_yesno(s: pd.Series) -> pd.Series:
    """Normaliza strings para 'SIM', 'NÃO' ou NaN."""
    x = s.astype(str).str.strip().str.upper()
    x = x.str.replace(r'\bNAO\b', 'NÃO', regex=True)
    x = x.replace({'': np.nan, 'NAN': np.nan, 'NONE': np.nan})
    x = x.infer_objects(copy=False)
    return x

def _to_bool_aval2024(s: pd.Series) -> pd.Series:
    if s.dtype == bool:
        return s
    if np.issubdtype(s.dtype, np.number):
        return s.fillna(0).astype(int).astype(bool)
    xs = _to_yesno(s)
    return xs.map({'SIM': True, 'NÃO': False}).fillna(False)

def aplicar_indicador_acomp_adesao(
    df: pd.DataFrame,
    status_col: str = "A/O ESTUDANTE ATENDE AOS CRITÉRIOS? (Sim ou Não)",
    avaliacao2024_col: str = "esteve-na-avaliacao-2024",
    chi_col: str = "ch_media_esperada",
    out_score_col: str = "indicador-acomp-adesao",
    out_class_col: str = "classificacao-acomp-adesao",
) -> pd.DataFrame:
    out = df.copy()

    if status_col not in out.columns:
        # Create column with NaN if missing to avoid error, or raise? 
        # Notebook raised KeyError, but for app robustness we might want to handle it gracefully or let it fail.
        # Let's try to be robust but warn. For now, strict as per notebook.
        # Actually, let's just add it as NaN if missing to allow partial data? 
        # No, the logic depends on it. Let's assume it might be missing and treat as blank.
        out[status_col] = np.nan
    
    status = _to_yesno(out[status_col])

    if avaliacao2024_col not in out.columns:
        esteve2024 = pd.Series(False, index=out.index)
    else:
        esteve2024 = _to_bool_aval2024(out[avaliacao2024_col])

    if chi_col in out.columns:
        chi = pd.to_numeric(out[chi_col], errors="coerce")
        ingressante = chi < 24
    else:
        ingressante = pd.Series(False, index=out.index)

    conditions = [
        ingressante,  # 0
        (~esteve2024) & (status == "SIM"),   # 1
        (~esteve2024) & (status == "NÃO"),   # 2
        (~esteve2024) & (status.isna()),     # 3 (em branco)
        (esteve2024) & (status == "SIM"),    # 4
        (esteve2024) & (status == "NÃO"),    # 5
        (esteve2024) & (status.isna()),      # 6 (em branco)
    ]
    scores = [0.0, 0.2, 0.6, 0.8, 0.1, 1.0, 0.9]
    classes = [
        "Sem risco / não pontua",
        "Estável",
        "Em alerta",
        "Prioridade de inserção",
        "Estável",
        "Crítico / penalização máxima",
        "Prioridade de convocação",
    ]

    out[out_score_col] = np.select(conditions, scores, default=np.nan).astype(float)
    out[out_class_col] = np.select(conditions, classes, default="Regra não classificada").astype(str)
    return out

def calcular_ita_final(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cols = ["nota_final", "pontuacao-renda", "indicador-acomp-adesao"]

    for col in cols:
        if col not in df.columns:
            df[col] = np.nan
        df[col] = pd.to_numeric(df[col], errors="coerce")

    pesos = {"nota_final": 6, "pontuacao-renda": 3, "indicador-acomp-adesao": 1}
    soma_pesos = sum(pesos.values())

    df["ITA"] = (
        (df["nota_final"] * pesos["nota_final"]) +
        (df["pontuacao-renda"] * pesos["pontuacao-renda"]) +
        (df["indicador-acomp-adesao"] * pesos["indicador-acomp-adesao"])
    ) / soma_pesos

    conditions = [
        df["ITA"] < 0.3,
        df["ITA"].between(0.3, 0.6, inclusive="both"),
        df["ITA"] > 0.6
    ]
    classes = ["baixo risco", "risco moderado", "risco alto"]

    df["classificacao_ita"] = np.select(conditions, classes, default="não classificado")
    df = df.sort_values(by="ITA", ascending=False).reset_index(drop=True)
    return df

def calculate_ita(file_path, file_path_crite, file_paht_form):
    pd.set_option('future.no_silent_downcasting', True)
    
    # Load Data
    try:
        df = pd.read_excel(file_path, sheet_name="PLANILHA COMPLETA")
        df_social = pd.read_excel(file_path_crite, sheet_name="Serviço Social")
        df_psicologia = pd.read_excel(file_path_crite, sheet_name="Psicologia")
        df_geral = pd.read_excel(file_path_crite, sheet_name="Geral")
        df_form_ = pd.read_excel(file_paht_form, sheet_name="Sheet1")
    except Exception as e:
        raise ValueError(f"Erro ao ler planilhas: {e}")

    df_form = padronizar_grr(df_form_)

    # 2. Risco Aprovação
    df["porcentagem_aprov_num"] = (pd.to_numeric(
        df["porcentagem-aprovacao"].replace(['#REF!', 'x'], np.nan),
        errors="coerce"
    ).fillna(0.0))
    df["risco_aprovacao"] = (((100 - df["porcentagem_aprov_num"]) / 100) ** 2)

    # 3. Risco Cancelamento
    df["fracao_disciplinas_canceladas"] = (
        df['qtd-matricula-cancelada'].replace(['#REF!', 'x'], np.nan).astype(float) /
        df['qtd-matriculada'].replace(['#REF!', 'x'], np.nan).astype(float)
    ).round(2)

    df["risco_cancelamento"] = (np.where(
        df["fracao_disciplinas_canceladas"] >= 0.5,
        1.0,
        df["fracao_disciplinas_canceladas"].clip(lower=0)
    )).round(2)

    # 4. Risco Reprovação por Frequência
    df["qtd-rep-frequencia"] = (pd.to_numeric(
        df["qtd-rep-frequencia"].replace(['#REF!', 'x'], np.nan), errors="coerce"
    )).round(2)
    df["qtd-matriculada"] = (pd.to_numeric(
        df["qtd-matriculada"].replace(['#REF!', 'x'], np.nan), errors="coerce"
    )).round(2)

    df["risco_rep_freq"] = (np.where(
        ((df["qtd-matriculada"] <= 2) & (df["qtd-rep-frequencia"] > 0)) |
        ((df["qtd-matriculada"] == 3) & (df["qtd-rep-frequencia"] > 1)) |
        ((df["qtd-matriculada"] == 4) & (df["qtd-rep-frequencia"] > 1)) |
        ((df["qtd-matriculada"] >= 5) & (df["qtd-rep-frequencia"] > 2)),
        1, 0
    )).round(2)

    # 5. Risco Histórico de Reprovação por Frequência
    df["hist_freq_pct"] = (pd.to_numeric(
        df["porcentagem-historica-de-reprovacao-frequencia"].replace(['#REF!', 'x'], np.nan),
        errors="coerce"
    ).fillna(0.0))
    df["risco_hist_freq"] = (df["hist_freq_pct"].clip(0, 1))

    # 6. Risco Carga Horária Integralizada
    df["TEMPO UFPR - SEM"] = pd.to_numeric(df["TEMPO UFPR - SEM"], errors="coerce").fillna(1).clip(lower=1)
    df["ch-integralizada"] = pd.to_numeric(df["ch-integralizada"], errors="coerce").fillna(0).clip(0, 100)
    
    df["ch_media_esperada"] = 8 * df["TEMPO UFPR - SEM"]
    
    cond_boost = (df["ch_media_esperada"] > 110) & (df["porcentagem_aprov_num"] < 50)
    df["ajuste_tempo"] = np.where(cond_boost, 5 * df["TEMPO UFPR - SEM"], df["TEMPO UFPR - SEM"])
    
    df["risco_ch_integralizada"] = (
        1 - (df["ch-integralizada"] / df["ch_media_esperada"])
    ) * df["ajuste_tempo"]
    
    df["risco_ch_integralizada"] = df["risco_ch_integralizada"].clip(lower=0)
    
    mask_non = ~cond_boost
    if mask_non.any():
        min_val = df.loc[mask_non, "risco_ch_integralizada"].quantile(0.05)
        max_val = df.loc[mask_non, "risco_ch_integralizada"].quantile(0.95)
        span = (max_val - min_val) if (max_val - min_val) != 0 else 1.0
    
        df.loc[mask_non, "risco_ch_integralizada"] = (
            (df.loc[mask_non, "risco_ch_integralizada"] - min_val) / span
        ).clip(0, 1)
    
    df.loc[cond_boost, "risco_ch_integralizada"] = 1.00
    df.loc[mask_non, "risco_ch_integralizada"] = (df.loc[mask_non, "risco_ch_integralizada"] * 0.25)
    df["risco_ch_integralizada"] = df["risco_ch_integralizada"].round(2)

    # 7. Risco Histórico de Avaliações
    df["risco_historico"] = (np.where(
        (df["apareceu-na-avaliacao-semestre-anterior?"] == 1) |
        (df["apareceu-na-avaliacao-semestre-anterior?"] == "0"),
        1.0, 0.0
    )).round(2)

    # 8. Risco Carga Horária cursada
    df["ch_cursada"] = pd.to_numeric(
        df["TEMPO UFPR - SEM"].replace(['#REF!', 'x'], np.nan),
        errors="coerce"
    ).fillna(0)
    
    df["risco_ch_cursada"] = (np.select(
        [
            df["ch_cursada"] >= 300,
            (df["ch_cursada"] <= 299) & (df["ch_cursada"] >= 200),
            (df["ch_cursada"] <= 199) & (df["ch_cursada"] >= 100),
            (df["ch_cursada"] <= 99) & (df["ch_cursada"] >= 0)
        ],
        [0, 0.33, 0.66, 1.0],
        default=np.nan
    )).round(2)

    # 9. Pesos e Nota Final
    pesos = {
        "aprovacao": 4,
        "rep_freq": 1.5,
        "hist_freq": 0.5,
        "ch_integralizada": 3,
        "historico": 0.5,
        "ch_cursada": 0.5,
    }
    
    for chave, peso in pesos.items():
        df[f"peso_{chave}"] = peso
        df[f"nota_parcial_{chave}"] = (df[f"risco_{chave}"] * peso).round(2)
    
    df["nota_final"] = (df[[f"nota_parcial_{k}" for k in pesos.keys()]].sum(axis=1)).round(2)

    # 10. Ordenação e Merge
    colunas_id = ["GRR", "NOME","SETOR","proafe", "curso", "ano-ingresso" ,"TEMPO UFPR - SEM","IRA SEM","CPF","renda-per-capta","classe-da-renda","nota-da-renda","E-MAIL PESSOAL","E-MAIL INSTITUCIONAL","TELEFONE","MOTIVO","planilha_andre"]
    
    # Ensure columns exist
    for col in colunas_id:
        if col not in df.columns:
            df[col] = np.nan

    blocos = [
        (["porcentagem-aprovacao"], "risco_aprovacao", "peso_aprovacao", "nota_parcial_aprovacao"),
        (["qtd-matriculada","qtd-reprovacao-por-nota","qtd-matricula-cancelada","PORT 5 - CAN",
          "qtd-rep-frequencia","PORT 5 - FREQ","BAIXA MAT"], "risco_rep_freq", "peso_rep_freq", "nota_parcial_rep_freq"),
        (["porcentagem-historica-de-reprovacao-frequencia"], "risco_hist_freq", "peso_hist_freq", "nota_parcial_hist_freq"),
        (["ch-integralizada","TEMPO UFPR - SEM", "ch_media_esperada","CH REC SEM","CH ABAIXO", "CH MTO ABAIXO"], "risco_ch_integralizada", "peso_ch_integralizada", "nota_parcial_ch_integralizada"),
        (["apareceu-na-avaliacao-semestre-anterior?"], "risco_historico", "peso_historico", "nota_parcial_historico"),
        (["responsavel","TEMPO UFPR - SEM","CH MAT TOTAL","% Rep Freq 2024-2","% Rep Freq 2024-1","% Rep Freq 2023 -2","Editais 2023","AVALIAÇÃO 2024","recebeu-probem-ano-anterior?",], "risco_ch_cursada", "peso_ch_cursada", "nota_parcial_ch_cursada"),
    ]

    ordered_cols = colunas_id.copy()
    for origs, risco, peso, nota in blocos:
        # Only add columns that exist
        valid_origs = [c for c in origs if c in df.columns]
        ordered_cols.extend(valid_origs + [risco, peso, nota])

    ordered_cols = list(dict.fromkeys(ordered_cols))
    ordered_cols.append("nota_final")
    
    # Filter only existing columns for the final view
    existing_cols = [c for c in ordered_cols if c in df.columns]
    df_detalhado = df[existing_cols].sort_values("nota_final", ascending=False)

    # -------- 1. Remover duplicatas dentro de cada df ----------
    df_social_unique = df_social.drop_duplicates(subset="GRR", keep="last")
    df_psicologia_unique = df_psicologia.drop_duplicates(subset="GRR", keep="last")
    df_geral_unique = df_geral.drop_duplicates(subset="GRR", keep="last")

    # -------- 2. Unir as três bases de Serviço Social, Psicologia e Geral ----------
    df_servicos_unificado = (
        df_social_unique
            .merge(df_psicologia_unique, on="GRR", how="outer")
            .merge(df_geral_unique, on="GRR", how="outer")
    )

    # -------- 3. Agora sim, fazer o merge com df_detalhado ----------
    df_merged = df_detalhado.merge(df_servicos_unificado, on="GRR", how="left")

    df_renda = aplicar_regra_renda(df_merged)
    df_acomp = aplicar_indicador_acomp_adesao(df_renda)
    df_ITA = calcular_ita_final(df_acomp)
    
    df_merged_final = df_ITA.merge(df_form, on="GRR", how="left")
    
    return df_merged_final
