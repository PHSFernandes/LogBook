# -*- coding: utf-8 -*-
"""
Motor de processamento de dados e geração de PDF (ReportLab) e Excel (OpenPyXL)
para os 5 relatórios da Plataforma Logbook - Egas Moniz.
"""

import io
import re
import zipfile
from datetime import datetime, date
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

from rules_data import (
    ALL_SPECIES,
    LIVESTOCK_SPECIES,
    COMPANION_SPECIES,
    EQUINE_SPECIES,
    parse_rule_blocks,
)

# ---------------------------------------------------------------------------
# EXTRAÇÃO DE METADADOS DO NOME DO FICHEIRO E NORMALIZAÇÃO DOS CSVs
# ---------------------------------------------------------------------------

def format_subject_pt(subject: str) -> str:
    """Converte '4th Year Logbook' -> '4º Ano', '5th Year Specialization Logbook' -> '5º Ano Specialization'."""
    s = str(subject or "").strip()
    s_no_lb = re.sub(r"\s*Logbook\s*$", "", s, flags=re.IGNORECASE).strip()
    m = re.match(r"^(\d+)(?:st|nd|rd|th)\s+Year(?:\s+(.*))?$", s_no_lb, flags=re.IGNORECASE)
    if m:
        num = m.group(1)
        extra = (m.group(2) or "").strip()
        return f"{num}º Ano {extra}".strip()
    return s_no_lb or s


def format_course_year_label(curso: str, subjects=None) -> str:
    """Retorna 'MIMV - 4º Ano' quando houver ano(s) selecionado(s), ou apenas 'MIMV' caso contrário."""
    if not subjects:
        return str(curso).strip()
    if isinstance(subjects, str):
        subj_list = [subjects] if subjects.strip() else []
    else:
        subj_list = [str(x).strip() for x in subjects if str(x).strip()]
    if not subj_list:
        return str(curso).strip()
    formatted = [format_subject_pt(x) for x in subj_list]
    return f"{curso} - {', '.join(formatted)}"


def extract_date_from_filename(filename: str):
    """
    Extrai a data no formato dd_mm_aaaa (ou dd-mm-aaaa como contingência) do nome do ficheiro.
    Retorna (date_obj | None, str_dd_mm_yyyy | None).
    """
    m = re.search(r"(\d{2})[_-](\d{2})[_-](\d{4})", str(filename))
    if not m:
        return None, None
    day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        dt_obj = date(year, month, day)
        return dt_obj, dt_obj.strftime("%d/%m/%Y")
    except ValueError:
        return None, None


def build_source_note(df_subset: pd.DataFrame) -> str:
    """
    Gera a nota padrão de rastreabilidade do ficheiro e data de extração:
    'Report gerado com dados obtidos do LogBook em dd/mm/aaaa (Ficheiro: ...).'
    """
    if df_subset is None or df_subset.empty or "Source_File" not in df_subset.columns:
        return "Report gerado com dados obtidos do LogBook."

    pairs = (
        df_subset[["Source_File", "Extraction_Date_Str"]]
        .drop_duplicates()
        .to_records(index=False)
    )
    entries = []
    dates_found = []
    for fname, dstr in pairs:
        if dstr and str(dstr) != "None" and str(dstr) != "nan":
            dates_found.append(str(dstr))
            entries.append(f"{fname} ({dstr})")
        else:
            entries.append(f"{fname}")

    unique_dates = sorted(set(dates_found), key=lambda s: datetime.strptime(s, "%d/%m/%Y"))
    files_str = "; ".join(entries)

    if len(unique_dates) == 1:
        return f"Report gerado com dados obtidos do LogBook em {unique_dates[0]} (Ficheiro: {files_str})."
    elif len(unique_dates) > 1:
        dates_joined = ", ".join(unique_dates[:-1]) + f" e {unique_dates[-1]}"
        return f"Report gerado com dados obtidos do LogBook em {dates_joined} (Ficheiros: {files_str})."
    else:
        return f"Report gerado com dados obtidos do LogBook (Ficheiro: {files_str})."


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).replace("\ufeff", "").strip().strip('"') for c in df.columns]
    return df


def parse_dt_series(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce", utc=True).dt.tz_localize(None)


def prepare_clinical_activities(df_raw: pd.DataFrame, curso_tag: str, source_filename: str = "") -> pd.DataFrame:
    """
    Normaliza clinical-activities:
    - Deduplica por 'Código do Log' (mantendo o registro mais recente)
    - Filtra apenas Estado == 'APPROVED' ou Atividade == 'direct_attribution' ('Approved como DA')
    - Explode alunos (Criado por quando STUDENT; Nome do estudante separado por vírgula quando TEACHER/direct_attribution)
    """
    df = clean_columns(df_raw)
    ext_dt, ext_str = extract_date_from_filename(source_filename)
    df["Curso_Tag"] = curso_tag
    df["Source_File"] = source_filename
    df["Extraction_Date"] = ext_dt
    df["Extraction_Date_Str"] = ext_str

    if "Atualização da Aprovação" in df.columns:
        df["_upd"] = parse_dt_series(df["Atualização da Aprovação"])
        df = df.sort_values("_upd", ascending=False)

    df = df.drop_duplicates(subset=["Código do Log"], keep="first").copy()

    estado = df["Estado do Log"].fillna("").astype(str).str.strip().str.upper()
    ativ = df["Atividade"].fillna("").astype(str).str.strip()

    is_approved = estado == "APPROVED"
    is_da = ativ.str.lower() == "direct_attribution"

    df = df[is_approved | is_da].copy()
    df["Status_Report"] = "Approved"
    df.loc[is_da & (~is_approved), "Status_Report"] = "Approved como DA"

    df["Data_Atividade_Dt"] = parse_dt_series(df["Data da Atividade"]).dt.date

    def extract_students(row):
        nome_est = str(row.get("Nome do estudante", "") or "").strip()
        if nome_est and nome_est.lower() != "nan":
            parts = [p.strip() for p in nome_est.split(",") if p.strip()]
            parts = [p for p in parts if p.lower() != "student"]
            if parts:
                return parts
        criado = str(row.get("Criado por", "") or "").strip()
        perfil = str(row.get("Perfil", "") or "").strip().upper()
        if criado and criado.lower() != "nan" and perfil == "STUDENT":
            return [criado]
        return []

    df["Lista_Alunos"] = df.apply(extract_students, axis=1)
    df = df[df["Lista_Alunos"].map(len) > 0].copy()
    df_exp = df.explode("Lista_Alunos").rename(columns={"Lista_Alunos": "Aluno"})

    def parse_species_cell(val):
        s = str(val or "").strip()
        if not s or s.lower() == "nan":
            return []
        return [x.strip() for x in s.split(",") if x.strip()]

    df_exp["Lista_Especies"] = df_exp["Espécie"].apply(parse_species_cell)
    return df_exp


def prepare_approval_control(df_raw: pd.DataFrame, curso_tag: str, source_filename: str = "") -> pd.DataFrame:
    """
    Normaliza approval-control eliminando falsos positivos e dupla contagem:
    1. Exige Estado do Log == 'AWAITING_APPROVAL' (exclui APPROVED, REJECTED, FILLING).
    2. Avalia o fluxo por Código do Log:
       - Se existir linha MONITOR com Aprovado vazio (NaN): pendência exclusiva do MONITOR.
       - Se NÃO existir MONITOR (sem_monitor) OU se o MONITOR já tiver avaliado (Aprovado == 1.0 ou 0.0,
         equivalente a conform_monitor / non_conform_monitor) E o TEACHER tiver Aprovado vazio (NaN):
         pendência exclusiva do TEACHER.
    """
    df = clean_columns(df_raw)
    ext_dt, ext_str = extract_date_from_filename(source_filename)
    df["Curso_Tag"] = curso_tag
    df["Source_File"] = source_filename
    df["Extraction_Date"] = ext_dt
    df["Extraction_Date_Str"] = ext_str

    estado = df["Estado do Log"].fillna("").astype(str).str.strip().str.upper()
    df_aw = df[estado == "AWAITING_APPROVAL"].copy()
    if df_aw.empty:
        return pd.DataFrame(columns=list(df.columns) + ["Perfil_Norm", "Nome_Aprovador_Clean", "Solicitado_Dt"])

    df_aw["Perfil_Norm"] = df_aw["Perfil do Aprovador"].fillna("").astype(str).str.strip().str.upper()
    df_aw["Nome_Aprovador_Clean"] = df_aw["Nome do Aprovador"].fillna("Sem Nome Atribuído").astype(str).str.strip()
    df_aw["Solicitado_Dt"] = parse_dt_series(df_aw["Solicitado para Aprovação em"])

    # Logs onde há monitor atribuído e ainda sem parecer (Aprovado isna)
    mon_rows = df_aw[df_aw["Perfil_Norm"] == "MONITOR"].copy()
    mon_pending_rows = mon_rows[mon_rows["Aprovado"].isna()].copy()
    logs_blocked_by_monitor = set(mon_pending_rows["Código do Log"].dropna().unique())

    # Logs pendentes com docente: Estado == AWAITING_APPROVAL, Aprovado docente vazio,
    # e NÃO bloqueados na fila do monitor (ou seja: sem monitor OU monitor já validou 1.0/0.0)
    tea_rows = df_aw[df_aw["Perfil_Norm"] == "TEACHER"].copy()
    tea_pending_rows = tea_rows[
        tea_rows["Aprovado"].isna() & (~tea_rows["Código do Log"].isin(logs_blocked_by_monitor))
    ].copy()

    return pd.concat([mon_pending_rows, tea_pending_rows], ignore_index=True)


def prepare_subcompetences(df_raw: pd.DataFrame, curso_tag: str, source_filename: str = "") -> pd.DataFrame:
    """Normaliza sub-competences-student-subject."""
    df = clean_columns(df_raw)
    ext_dt, ext_str = extract_date_from_filename(source_filename)
    df["Curso_Tag"] = curso_tag
    df["Source_File"] = source_filename
    df["Extraction_Date"] = ext_dt
    df["Extraction_Date_Str"] = ext_str

    df["Nome_Clean"] = df["Nome"].fillna("").astype(str).str.strip()
    df["Codigo_Clean"] = df["Código da Subcompetência"].fillna("").astype(str).str.strip()
    df["Especie_Clean"] = df["Espécie"].fillna("").astype(str).str.strip()
    df["Subject_Clean"] = df["Subject"].fillna("").astype(str).str.strip()
    df["Origem_Clean"] = df["Origem da atribuição"].fillna("").astype(str).str.strip()
    df["Data_Atividade_Dt"] = parse_dt_series(df["Data da atividade"]).dt.date
    return df


# ---------------------------------------------------------------------------
# LÓGICA DE CÁLCULO DOS 5 REPORTS
# ---------------------------------------------------------------------------

def build_report_1_uc(
    df_act_exp: pd.DataFrame,
    uc_name: str,
    dt_start: date,
    dt_end: date,
    selected_species: list,
    selected_activities: list,
    subject_filter: list = None,
) -> tuple[pd.DataFrame, dict]:
    """
    Gera matriz Alunos x Espécies + Totais (Total Geral, Espécie Pecuária, Animais de Cia, Equinos)
    para uma UC no período e filtros configurados.
    """
    df = df_act_exp.copy()
    df = df[(df["Data_Atividade_Dt"] >= dt_start) & (df["Data_Atividade_Dt"] <= dt_end)]

    if subject_filter:
        df = df[df["Subject"].isin(subject_filter)]

    if selected_activities:
        df = df[df["Atividade"].isin(selected_activities)]

    df_sp = df.explode("Lista_Especies").rename(columns={"Lista_Especies": "Especie_Item"})
    df_sp = df_sp[df_sp["Especie_Item"].notna() & (df_sp["Especie_Item"] != "")]

    if selected_species:
        df_sp = df_sp[df_sp["Especie_Item"].isin(selected_species)]

    if subject_filter:
        all_students = sorted(df_act_exp[df_act_exp["Subject"].isin(subject_filter)]["Aluno"].unique())
    else:
        all_students = sorted(df_sp["Aluno"].unique())

    cols_species = [sp for sp in ALL_SPECIES if (not selected_species or sp in selected_species)]
    if not cols_species:
        cols_species = list(ALL_SPECIES)

    if df_sp.empty:
        piv = pd.DataFrame(0, index=all_students, columns=cols_species)
    else:
        piv = pd.crosstab(df_sp["Aluno"], df_sp["Especie_Item"])
        piv = piv.reindex(index=all_students, columns=cols_species, fill_value=0)

    piv.index.name = "Alunos"
    res = piv.reset_index()

    livestock_cols = [c for c in LIVESTOCK_SPECIES if c in res.columns]
    companion_cols = [c for c in COMPANION_SPECIES if c in res.columns]
    equine_cols = [c for c in EQUINE_SPECIES if c in res.columns]

    res["Totais"] = res[cols_species].sum(axis=1)
    res["Espécie Pecuária"] = res[livestock_cols].sum(axis=1) if livestock_cols else 0
    res["Animais de Cia"] = res[companion_cols].sum(axis=1) if companion_cols else 0
    res["Equinos"] = res[equine_cols].sum(axis=1) if equine_cols else 0

    totals_row = {"Alunos": "TOTAL GERAL"}
    for c in cols_species + ["Totais", "Espécie Pecuária", "Animais de Cia", "Equinos"]:
        totals_row[c] = int(res[c].sum())
    res = pd.concat([res, pd.DataFrame([totals_row])], ignore_index=True)

    curso_tag_val = str(df_act_exp["Curso_Tag"].iloc[0]) if not df_act_exp.empty else ""
    meta = {
        "uc_name": uc_name,
        "curso_header": format_course_year_label(curso_tag_val, subject_filter),
        "dt_start": dt_start.strftime("%d/%m/%Y"),
        "dt_end": dt_end.strftime("%d/%m/%Y"),
        "activities": ", ".join(selected_activities) if selected_activities else "Todas",
        "species": ", ".join(selected_species) if selected_species else "Todas",
        "logs_approved": int(df[df["Status_Report"] == "Approved"]["Código do Log"].nunique()) if not df.empty else 0,
        "logs_da": int(df[df["Status_Report"] == "Approved como DA"]["Código do Log"].nunique()) if not df.empty else 0,
        "source_note": build_source_note(df_act_exp),
    }
    return res, meta


def build_report_2_3_delays(
    df_app_all: pd.DataFrame,
    perfil: str,
    ref_date: date,
    grace_days: int = 3,
) -> tuple[pd.DataFrame, str]:
    """
    Calcula pendentes com 0 dias (todos AWAITING_APPROVAL sem decisão)
    e com carência (> grace_days), além do maior atraso em dias à data de referência.
    Retorna (DataFrame, source_note).
    """
    source_note = build_source_note(df_app_all)
    df = df_app_all[df_app_all["Perfil_Norm"] == perfil].copy()
    if df.empty:
        return pd.DataFrame(columns=[
            "Nome", "Cursos", "Pendentes (0 dias)", f"Em Atraso (> {grace_days} dias)", "Maior Atraso (dias)"
        ]), source_note

    ref_ts = pd.Timestamp(datetime.combine(ref_date, datetime.max.time()))
    df["Dias_Atraso"] = (ref_ts - df["Solicitado_Dt"]).dt.days.clip(lower=0).fillna(0).astype(int)

    df_u = df.drop_duplicates(subset=["Nome_Aprovador_Clean", "Código do Log"]).copy()
    df_u["Atraso_Carencia"] = (df_u["Dias_Atraso"] > grace_days).astype(int)

    agg = df_u.groupby("Nome_Aprovador_Clean").agg(
        Cursos=("Curso_Tag", lambda x: ", ".join(sorted(set(x)))),
        Pendentes_0=("Código do Log", "count"),
        Atraso_Carencia=("Atraso_Carencia", "sum"),
        Maior_Atraso=("Dias_Atraso", "max"),
    ).reset_index()

    agg = agg.rename(columns={
        "Nome_Aprovador_Clean": "Nome",
        "Pendentes_0": "Pendentes (0 dias)",
        "Atraso_Carencia": f"Em Atraso (> {grace_days} dias)",
        "Maior_Atraso": "Maior Atraso (dias)",
    })

    agg = agg.sort_values(
        by=[f"Em Atraso (> {grace_days} dias)", "Pendentes (0 dias)", "Maior Atraso (dias)", "Nome"],
        ascending=[False, False, False, True],
    ).reset_index(drop=True)

    total_row = pd.DataFrame([{
        "Nome": "TOTAL GERAL",
        "Cursos": "-",
        "Pendentes (0 dias)": int(agg["Pendentes (0 dias)"].sum()),
        f"Em Atraso (> {grace_days} dias)": int(agg[f"Em Atraso (> {grace_days} dias)"].sum()),
        "Maior Atraso (dias)": int(agg["Maior Atraso (dias)"].max()) if not agg.empty else 0,
    }])
    return pd.concat([agg, total_row], ignore_index=True), source_note


def evaluate_student_competences(
    df_sub_filtered: pd.DataFrame,
    students: list,
    rules_df: pd.DataFrame,
) -> dict:
    """
    Avalia cada aluno contra todas as regras ativas do catálogo.
    """
    active_rules = rules_df[rules_df["Estado"] == "Ativa"].copy()
    parsed_rules = []
    for _, r in active_rules.iterrows():
        blocks = parse_rule_blocks(r["Regra Espécies"], r["Regra Mínimo"])
        total_min = sum(b["min"] for b in blocks)
        parsed_rules.append({
            "code": str(r["Código"]).strip(),
            "desc": str(r["Subcompetência"]).strip(),
            "sp_rule": str(r["Regra Espécies"]).strip(),
            "min_rule": str(r["Regra Mínimo"]).strip(),
            "blocks": blocks,
            "total_min": total_min,
        })

    if df_sub_filtered.empty:
        counts_map = {}
    else:
        grp = df_sub_filtered.groupby(["Nome_Clean", "Codigo_Clean", "Especie_Clean"]).size()
        counts_map = grp.to_dict()

    results_by_student = {}
    for st in students:
        rows = []
        for pr in parsed_rules:
            code = pr["code"]
            sp_counts = {sp: counts_map.get((st, code, sp), 0) for sp in ALL_SPECIES}
            total_obt = sum(sp_counts.values())

            block_statuses = []
            all_blocks_ok = True
            for b in pr["blocks"]:
                b_obt = sum(sp_counts.get(sp, 0) for sp in b["species"])
                b_req = b["min"]
                if b_obt < b_req:
                    all_blocks_ok = False
                block_statuses.append(f"{b['label']}: {b_obt}/{b_req}")

            status = "Atingido" if all_blocks_ok else "Abaixo do Mínimo"
            row_dict = {
                "Código": code,
                "Subcompetência": pr["desc"],
                "Regra Espécies": pr["sp_rule"],
                "Mínimo": pr["min_rule"],
                "Obtido": total_obt,
                "Progresso Blocos": " | ".join(block_statuses),
                "Status": status,
            }
            for sp in ALL_SPECIES:
                row_dict[sp] = sp_counts[sp]
            rows.append(row_dict)

        results_by_student[st] = pd.DataFrame(rows)
    return results_by_student


def build_report_5_managerial(
    results_by_student: dict,
    curso: str,
    subject: str,
) -> pd.DataFrame:
    """
    Consolida para todos os alunos do curso/ano quais competências estão abaixo do mínimo.
    """
    total_students = len(results_by_student)
    if total_students == 0:
        return pd.DataFrame()

    all_dfs = []
    for st, df_st in results_by_student.items():
        tmp = df_st[["Código", "Subcompetência", "Regra Espécies", "Mínimo", "Obtido", "Status"]].copy()
        tmp["Aluno"] = st
        all_dfs.append(tmp)

    concat_df = pd.concat(all_dfs, ignore_index=True)
    below = concat_df[concat_df["Status"] == "Abaixo do Mínimo"]

    summary = concat_df.groupby(["Código", "Subcompetência", "Regra Espécies", "Mínimo"], sort=False).agg(
        Media_Obtida=("Obtido", "mean"),
    ).reset_index()

    below_counts = below.groupby("Código").agg(
        Alunos_Abaixo=("Aluno", "count"),
        Lista_Alunos=("Aluno", lambda x: ", ".join(sorted(x))),
    ).reset_index()

    merged = summary.merge(below_counts, on="Código", how="inner")
    merged.insert(0, "Curso", curso)
    merged.insert(1, "Ano / Subject", subject)
    merged["Total Alunos"] = total_students
    merged["% Abaixo do Mínimo"] = (merged["Alunos_Abaixo"] / total_students * 100).round(1)
    merged["Media_Obtida"] = merged["Media_Obtida"].round(1)

    merged = merged.rename(columns={
        "Alunos_Abaixo": "Alunos Abaixo do Mínimo",
        "Media_Obtida": "Média Obtida",
        "Lista_Alunos": "Estudantes Pendentes",
    })

    merged = merged.sort_values(
        by=["Alunos Abaixo do Mínimo", "Código"], ascending=[False, True]
    ).reset_index(drop=True)
    return merged


# ---------------------------------------------------------------------------
# EXPORTAÇÃO EXCEL (.XLSX)
# ---------------------------------------------------------------------------

def _style_worksheet(ws, title: str, subtitle: str, source_note: str = "", green_cols=None, highlight_red_col=None):
    green_cols = green_cols or set()
    navy_fill = PatternFill("solid", fgColor="1B365D")
    green_hdr_fill = PatternFill("solid", fgColor="1E6B38")
    green_cell_fill = PatternFill("solid", fgColor="E6F4EA")
    zebra_fill = PatternFill("solid", fgColor="F5F7FA")
    total_fill = PatternFill("solid", fgColor="D9E2EC")
    red_font = Font(name="Calibri", size=10, bold=True, color="B00020")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=10, color="111111")
    bold_font = Font(name="Calibri", size=10, bold=True, color="111111")
    thin_side = Side(border_style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws.insert_rows(1, 4)
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1B365D")
    ws["A2"] = subtitle
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="475569")
    ws["A3"] = source_note
    ws["A3"].font = Font(name="Calibri", size=9.5, bold=True, color="1E3A8A")

    header_row = 5
    max_col = ws.max_column
    max_row = ws.max_row

    for c in range(1, max_col + 1):
        cell = ws.cell(header_row, c)
        col_name = str(cell.value or "")
        cell.fill = green_hdr_fill if col_name in green_cols else navy_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[header_row].height = 28

    for r in range(header_row + 1, max_row + 1):
        first_val = str(ws.cell(r, 1).value or "")
        is_total = "TOTAL GERAL" in first_val.upper()
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            col_name = str(ws.cell(header_row, c).value or "")
            cell.border = border
            if is_total:
                cell.fill = total_fill
                cell.font = bold_font
            else:
                if col_name in green_cols:
                    cell.fill = green_cell_fill
                elif r % 2 == 1:
                    cell.fill = zebra_fill
                cell.font = body_font

            if highlight_red_col and col_name == highlight_red_col and str(cell.value) == "Abaixo do Mínimo":
                cell.font = red_font

            if c > 2 and isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        max_len = 10
        for r in range(header_row, max_row + 1):
            val = str(ws.cell(r, c).value or "")
            max_len = max(max_len, min(len(val) + 3, 48))
        ws.column_dimensions[col_letter].width = max_len

    ws.freeze_panes = f"A{header_row + 1}"


def to_excel_bytes(sheets_dict: dict, green_cols=None, highlight_red_col=None) -> bytes:
    """
    sheets_dict: {sheet_name: (df, title, subtitle, source_note)}
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, item in sheets_dict.items():
        if len(item) == 4:
            df, title, subtitle, source_note = item
        else:
            df, title, subtitle = item
            source_note = ""
        safe_name = str(sheet_name)[:31].replace("/", "-").replace(":", "-")
        ws = wb.create_sheet(title=safe_name)
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append(list(row))
        _style_worksheet(
            ws, title, subtitle, source_note=source_note,
            green_cols=green_cols, highlight_red_col=highlight_red_col
        )

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# EXPORTAÇÃO PDF (REPORTLAB)
# ---------------------------------------------------------------------------

def _get_pdf_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="InstHeader", fontName="Helvetica-Bold", fontSize=13,
        textColor=colors.HexColor("#1B365D"), spaceAfter=2
    ))
    styles.add(ParagraphStyle(
        name="InstSub", fontName="Helvetica", fontSize=8.5,
        textColor=colors.HexColor("#475569"), spaceAfter=2
    ))
    styles.add(ParagraphStyle(
        name="SourceNote", fontName="Helvetica-Bold", fontSize=8,
        textColor=colors.HexColor("#1E3A8A"), spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        name="CellTxt", fontName="Helvetica", fontSize=7.5, leading=9,
        textColor=colors.HexColor("#111111")
    ))
    styles.add(ParagraphStyle(
        name="CellBold", fontName="Helvetica-Bold", fontSize=7.5, leading=9,
        textColor=colors.HexColor("#111111")
    ))
    styles.add(ParagraphStyle(
        name="CellRed", fontName="Helvetica-Bold", fontSize=7.5, leading=9,
        textColor=colors.HexColor("#B00020")
    ))
    styles.add(ParagraphStyle(
        name="HdrTxt", fontName="Helvetica-Bold", fontSize=7.2, leading=8.5,
        textColor=colors.white, alignment=TA_CENTER
    ))
    styles.add(ParagraphStyle(
        name="FootNote", fontName="Helvetica-Oblique", fontSize=7.5,
        textColor=colors.HexColor("#475569"), spaceBefore=6
    ))
    return styles


def _add_page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.HexColor("#64748B"))
    dt_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    canvas.drawString(12 * mm, 8 * mm, f"Egas Moniz School of Health & Science — Emitido em {dt_str}")
    canvas.drawRightString(doc.pagesize[0] - 12 * mm, 8 * mm, f"Página {doc.page}")
    canvas.restoreState()


def generate_pdf_report_1(df: pd.DataFrame, meta: dict, curso: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm, topMargin=12 * mm, bottomMargin=14 * mm
    )
    styles = _get_pdf_styles()
    story = []

    hdr_r1 = meta.get("curso_header") or curso
    story.append(Paragraph(f"RELATÓRIO DE ATIVIDADES CLÍNICAS POR UC — {hdr_r1}", styles["InstHeader"]))
    story.append(Paragraph(
        f"<b>UC de referência:</b> {meta['uc_name']} &nbsp;|&nbsp; "
        f"<b>Período:</b> {meta['dt_start']} a {meta['dt_end']} &nbsp;|&nbsp; "
        f"<b>Atividades:</b> {meta['activities']}",
        styles["InstSub"]
    ))
    if meta.get("source_note"):
        story.append(Paragraph(meta["source_note"], styles["SourceNote"]))

    green_set = set(LIVESTOCK_SPECIES + ["Espécie Pecuária"])
    cols = list(df.columns)
    header_row = [Paragraph(c, styles["HdrTxt"]) for c in cols]
    table_data = [header_row]

    for _, r in df.iterrows():
        is_tot = str(r["Alunos"]) == "TOTAL GERAL"
        st_use = styles["CellBold"] if is_tot else styles["CellTxt"]
        row_cells = [Paragraph(str(r["Alunos"]), st_use)]
        for c in cols[1:]:
            row_cells.append(str(int(r[c])))
        table_data.append(row_cells)

    avail_w = landscape(A4)[0] - 20 * mm
    name_w = 46 * mm
    other_w = (avail_w - name_w) / max(1, len(cols) - 1)
    col_widths = [name_w] + [other_w] * (len(cols) - 1)

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    ts = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B365D")),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (1, 1), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F8FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D9E2EC")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for idx, c in enumerate(cols):
        if c in green_set:
            ts.append(("BACKGROUND", (idx, 0), (idx, 0), colors.HexColor("#1E6B38")))
            ts.append(("BACKGROUND", (idx, 1), (idx, -2), colors.HexColor("#E6F4EA")))

    t.setStyle(TableStyle(ts))
    story.append(t)
    story.append(Paragraph(
        "Nota: Contabilizados apenas registos com estado 'Approved' e atribuições diretas ('Approved como DA'). "
        "Colunas a verde destacam espécies pecuárias e respetivo somatório.",
        styles["FootNote"]
    ))
    doc.build(story, onFirstPage=_add_page_number, onLaterPages=_add_page_number)
    return buf.getvalue()


def generate_pdf_report_2_3(df: pd.DataFrame, titulo: str, ref_date: date, grace_days: int, source_note: str = "") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=portrait(A4),
        leftMargin=14 * mm, rightMargin=14 * mm, topMargin=14 * mm, bottomMargin=15 * mm
    )
    styles = _get_pdf_styles()
    story = []

    story.append(Paragraph(titulo, styles["InstHeader"]))
    story.append(Paragraph(
        f"<b>Data de referência do cálculo:</b> {ref_date.strftime('%d/%m/%Y')} &nbsp;|&nbsp; "
        f"<b>Carência aplicada:</b> {grace_days} dias (todos os cursos carregados)",
        styles["InstSub"]
    ))
    if source_note:
        story.append(Paragraph(source_note, styles["SourceNote"]))

    cols = list(df.columns)
    table_data = [[Paragraph(c, styles["HdrTxt"]) for c in cols]]
    for _, r in df.iterrows():
        is_tot = str(r["Nome"]) == "TOTAL GERAL"
        st_use = styles["CellBold"] if is_tot else styles["CellTxt"]
        table_data.append([
            Paragraph(str(r[cols[0]]), st_use),
            Paragraph(str(r[cols[1]]), st_use),
            str(r[cols[2]]),
            str(r[cols[3]]),
            str(r[cols[4]]),
        ])

    avail_w = portrait(A4)[0] - 28 * mm
    col_widths = [avail_w * 0.36, avail_w * 0.16, avail_w * 0.16, avail_w * 0.16, avail_w * 0.16]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B365D")),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (1, 1), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F8FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D9E2EC")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t)
    story.append(Paragraph(
        f"Observação: A coluna 'Pendentes (0 dias)' reflete todos os logs em estado AWAITING_APPROVAL sem decisão registada. "
        f"A coluna 'Em Atraso (> {grace_days} dias)' aplica o período de carência de {grace_days} dias após a solicitação de aprovação. "
        f"O 'Maior Atraso (dias)' corresponde ao caso mais antigo pendente de cada avaliador.",
        styles["FootNote"]
    ))
    doc.build(story, onFirstPage=_add_page_number, onLaterPages=_add_page_number)
    return buf.getvalue()


def generate_pdf_report_4_student(
    student_name: str,
    df_st: pd.DataFrame,
    curso: str,
    subject: str,
    only_below: bool,
    source_note: str = "",
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=9 * mm, rightMargin=9 * mm, topMargin=11 * mm, bottomMargin=13 * mm
    )
    styles = _get_pdf_styles()
    story = []

    df_view = df_st[df_st["Status"] == "Abaixo do Mínimo"].copy() if only_below else df_st.copy()
    modo_txt = "Apenas competências abaixo do mínimo" if only_below else "Todas as competências avaliadas"

    curso_ano_4 = format_course_year_label(curso, subject)
    story.append(Paragraph(f"RELATÓRIO INDIVIDUAL DE COMPETÊNCIAS — {curso_ano_4}", styles["InstHeader"]))
    story.append(Paragraph(
        f"<b>Estudante:</b> {student_name} &nbsp;|&nbsp; "
        f"<b>Ano / Subject:</b> {subject} &nbsp;|&nbsp; "
        f"<b>Filtro:</b> {modo_txt}",
        styles["InstSub"]
    ))
    if source_note:
        story.append(Paragraph(source_note, styles["SourceNote"]))

    if df_view.empty:
        story.append(Paragraph("Nenhuma subcompetência pendente para os filtros selecionados.", styles["CellBold"]))
    else:
        sp_short = [
            ("Equine", "Eq"), ("Dog", "Dg"), ("Cat", "Ct"), ("Bovine", "Bv"),
            ("Ovine", "Ov"), ("Goat", "Gt"), ("Swine", "Sw"), ("Poultry", "Pl"),
            ("Rabbit", "Rb"), ("Exotic - Bird", "ExB"), ("Exotic - Rodents", "ExR"),
            ("Exotic - Others", "ExO"), ("Fish", "Fs")
        ]
        headers = ["Cód.", "Subcompetência", "Regra", "Mín.", "Obt.", "Blocos"] + [s[1] for s in sp_short] + ["Estado"]
        table_data = [[Paragraph(h, styles["HdrTxt"]) for h in headers]]

        for _, r in df_view.iterrows():
            is_low = r["Status"] == "Abaixo do Mínimo"
            st_state = styles["CellRed"] if is_low else styles["CellTxt"]
            row_cells = [
                Paragraph(str(r["Código"]), styles["CellBold"]),
                Paragraph(str(r["Subcompetência"]), styles["CellTxt"]),
                Paragraph(str(r["Regra Espécies"]), styles["CellTxt"]),
                str(r["Mínimo"]),
                str(r["Obtido"]),
                Paragraph(str(r["Progresso Blocos"]), styles["CellTxt"]),
            ]
            for full_sp, _ in sp_short:
                val = int(r.get(full_sp, 0))
                row_cells.append(str(val) if val > 0 else "0")
            row_cells.append(Paragraph("Abaixo" if is_low else "OK", st_state))
            table_data.append(row_cells)

        avail_w = landscape(A4)[0] - 18 * mm
        w_cod = 14 * mm
        w_desc = 64 * mm
        w_reg = 28 * mm
        w_min = 12 * mm
        w_obt = 11 * mm
        w_blk = 34 * mm
        w_est = 14 * mm
        rem_w = avail_w - (w_cod + w_desc + w_reg + w_min + w_obt + w_blk + w_est)
        w_sp = rem_w / len(sp_short)
        col_widths = [w_cod, w_desc, w_reg, w_min, w_obt, w_blk] + [w_sp] * len(sp_short) + [w_est]

        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B365D")),
            ("ALIGN", (3, 1), (4, -1), "CENTER"),
            ("ALIGN", (6, 1), (-2, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (3, 1), (-2, -1), 6.8),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ]))
        story.append(t)

    doc.build(story, onFirstPage=_add_page_number, onLaterPages=_add_page_number)
    return buf.getvalue()


def generate_pdf_report_5(df: pd.DataFrame, curso: str, subject: str, source_note: str = "") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=11 * mm, rightMargin=11 * mm, topMargin=12 * mm, bottomMargin=14 * mm
    )
    styles = _get_pdf_styles()
    story = []

    curso_ano_5 = format_course_year_label(curso, subject)
    story.append(Paragraph(f"RELATÓRIO GERENCIAL — COMPETÊNCIAS ABAIXO DO MÍNIMO ({curso_ano_5})", styles["InstHeader"]))
    story.append(Paragraph(
        f"<b>Curso:</b> {curso} &nbsp;|&nbsp; <b>Ano / Subject:</b> {subject}",
        styles["InstSub"]
    ))
    if source_note:
        story.append(Paragraph(source_note, styles["SourceNote"]))

    cols = ["Código", "Subcompetência", "Regra Espécies", "Mínimo", "Alunos Abaixo do Mínimo", "Total Alunos", "% Abaixo do Mínimo", "Média Obtida"]
    table_data = [[Paragraph(c, styles["HdrTxt"]) for c in cols]]

    for _, r in df.iterrows():
        table_data.append([
            Paragraph(str(r["Código"]), styles["CellBold"]),
            Paragraph(str(r["Subcompetência"]), styles["CellTxt"]),
            Paragraph(str(r["Regra Espécies"]), styles["CellTxt"]),
            str(r["Mínimo"]),
            str(r["Alunos Abaixo do Mínimo"]),
            str(r["Total Alunos"]),
            f"{r['% Abaixo do Mínimo']:.1f}%",
            f"{r['Média Obtida']:.1f}",
        ])

    avail_w = landscape(A4)[0] - 22 * mm
    col_widths = [
        avail_w * 0.08, avail_w * 0.36, avail_w * 0.16, avail_w * 0.08,
        avail_w * 0.08, avail_w * 0.08, avail_w * 0.08, avail_w * 0.08
    ]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B365D")),
        ("ALIGN", (3, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (3, 1), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#CBD5E1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t)
    doc.build(story, onFirstPage=_add_page_number, onLaterPages=_add_page_number)
    return buf.getvalue()
