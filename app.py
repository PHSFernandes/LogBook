# -*- coding: utf-8 -*-
"""
Plataforma de Emissão de Reports — Logbook Clínico (MIMV, LEV, LCB)
Egas Moniz School of Health & Science
"""

import io
import os
import json
import zipfile
from datetime import date, timedelta
import pandas as pd
import streamlit as st

from rules_data import (
    ALL_SPECIES,
    LIVESTOCK_SPECIES,
    get_default_rules_list,
)
from report_generators import (
    prepare_clinical_activities,
    prepare_approval_control,
    prepare_subcompetences,
    build_report_1_uc,
    build_report_2_3_delays,
    evaluate_student_competences,
    build_report_5_managerial,
    to_excel_bytes,
    generate_pdf_report_1,
    generate_pdf_report_2_3,
    generate_pdf_report_4_student,
    generate_pdf_report_5,
)

st.set_page_config(
    page_title="Report Logbook — Egas Moniz",
    page_icon="📋",
    layout="wide",
)

UC_CONFIG_FILE = "uc_config.json"
RULES_CONFIG_FILE = "competences_rules_custom.json"

DEFAULT_UC_ROWS = [
    {
        "Curso": "MIMV",
        "UC / Disciplina": "Clínica de Animais de Companhia I",
        "Ano / Subject": "5th Year Logbook",
        "Data Início": date(2026, 2, 1),
        "Data Fim": date(2026, 5, 31),
        "Espécies": "Dog, Cat, Rabbit, Exotic - Bird, Exotic - Rodents, Exotic - Others",
        "Atividades": "Clinical Case, direct_attribution",
    },
    {
        "Curso": "MIMV",
        "UC / Disciplina": "Clínica de Espécies Pecuárias e Equinos",
        "Ano / Subject": "5th Year Logbook",
        "Data Início": date(2026, 2, 1),
        "Data Fim": date(2026, 5, 31),
        "Espécies": "Equine, Bovine, Ovine, Goat, Swine, Poultry",
        "Atividades": "Clinical Case, Herd Case, direct_attribution",
    },
]


def load_uc_config() -> pd.DataFrame:
    if os.path.exists(UC_CONFIG_FILE):
        try:
            with open(UC_CONFIG_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
            df = pd.DataFrame(raw)
            df["Data Início"] = pd.to_datetime(df["Data Início"]).dt.date
            df["Data Fim"] = pd.to_datetime(df["Data Fim"]).dt.date
            return df
        except Exception:
            pass
    return pd.DataFrame(DEFAULT_UC_ROWS)


def save_uc_config(df: pd.DataFrame):
    out = df.copy()
    out["Data Início"] = out["Data Início"].astype(str)
    out["Data Fim"] = out["Data Fim"].astype(str)
    with open(UC_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(out.to_dict(orient="records"), f, ensure_ascii=False, indent=2)


def load_rules_config() -> pd.DataFrame:
    if os.path.exists(RULES_CONFIG_FILE):
        try:
            with open(RULES_CONFIG_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
            return pd.DataFrame(raw)
        except Exception:
            pass
    return pd.DataFrame(get_default_rules_list())


def save_rules_config(df: pd.DataFrame):
    with open(RULES_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, indent=2)


def detect_course_from_filename(filename: str) -> str:
    up = filename.upper()
    for c in ("MIMV", "LEV", "LCB"):
        if c in up:
            return c
    return "MIMV"


# ---------------------------------------------------------------------------
# SIDEBAR — CARREGAMENTO DE FICHEIROS CSV POR CURSO
# ---------------------------------------------------------------------------
st.sidebar.title("📂 Base de Dados (CSVs)")
st.sidebar.caption(
    "Carregue os ficheiros CSV exportados da plataforma. "
    "Se o nome do ficheiro contiver **MIMV**, **LEV** ou **LCB**, o curso será identificado automaticamente."
)

uploaded_files = st.sidebar.file_uploader(
    "Selecione os CSVs (approval-control, clinical-activities, sub-competences)",
    type=["csv"],
    accept_multiple_files=True,
)

# Armazenamento em memória dos DataFrames carregados
act_list = []
app_list = []
sub_list = []
file_summary = []

if uploaded_files:
    for up in uploaded_files:
        fname = up.name
        detected_course = detect_course_from_filename(fname)
        course_choice = st.sidebar.selectbox(
            f"Curso de: {fname}",
            options=["MIMV", "LEV", "LCB"],
            index=["MIMV", "LEV", "LCB"].index(detected_course),
            key=f"crs_{fname}",
        )
        try:
            raw_df = pd.read_csv(up, encoding="utf-8-sig")
            cols_clean = [str(c).replace("\ufeff", "").strip().strip('"') for c in raw_df.columns]

            if "Perfil do Aprovador" in cols_clean:
                norm_df = prepare_approval_control(raw_df, course_choice)
                app_list.append(norm_df)
                file_summary.append((fname, course_choice, "Approval Control", len(norm_df)))
            elif "ID Case Instance" in cols_clean or "Data da Atividade" in cols_clean:
                norm_df = prepare_clinical_activities(raw_df, course_choice)
                act_list.append(norm_df)
                file_summary.append((fname, course_choice, "Clinical Activities", len(norm_df)))
            elif "Código da Subcompetência" in cols_clean:
                norm_df = prepare_subcompetences(raw_df, course_choice)
                sub_list.append(norm_df)
                file_summary.append((fname, course_choice, "Sub-Competences", len(norm_df)))
            else:
                st.sidebar.warning(f"Estrutura não reconhecida em: {fname}")
        except Exception as e:
            st.sidebar.error(f"Erro ao ler {fname}: {e}")

df_act_all = pd.concat(act_list, ignore_index=True) if act_list else pd.DataFrame()
df_app_all = pd.concat(app_list, ignore_index=True) if app_list else pd.DataFrame()
df_sub_all = pd.concat(sub_list, ignore_index=True) if sub_list else pd.DataFrame()

if file_summary:
    st.sidebar.markdown("---")
    st.sidebar.subheader("Ficheiros Ativos")
    for fn, crs, tipo, cnt in file_summary:
        st.sidebar.markdown(f"- **{crs}** | {tipo} (`{cnt}` reg.)")

# ---------------------------------------------------------------------------
# PÁGINA PRINCIPAL E AVISOS METODOLÓGICOS
# ---------------------------------------------------------------------------
st.title("📋 Plataforma de Emissão de Reports — Logbook Clínico")

with st.expander("📌 Notas Metodológicas e Exceções Configuradas (Ler antes da emissão)", expanded=True):
    st.markdown(
        """
        - **Report 1 (Unidades Curriculares):** Contabiliza exclusivamente registos com estado **Approved** e atribuições diretas (**Approved como DA**), desdobrando os estudantes quando registados em grupo. Colunas destacadas a verde representam **Espécies Pecuárias** (`Swine`, `Ovine`, `Bovine`, `Goat`, `Poultry`).
        - **Reports 2 e 3 (Monitores e Professores em Atraso):** Considera todos os cursos carregados (`MIMV`, `LEV`, `LCB`) onde `Estado do Log == AWAITING_APPROVAL` e o avaliador ainda não registou decisão. Apresenta duas métricas em paralelo: **Pendentes (0 dias)** (espelhando o dashboard da plataforma) e **Em Atraso (> 3 dias de carência)**, além do **Maior Atraso (dias)** à data de emissão.
        - **Reports 4 e 5 (Competências e Mínimos):** Aplicam-se as regras oficiais do catálogo a todos os anos curriculares, com as seguintes exceções definidas pela coordenação:
          1. **`2-US02` (*Echographic examination training on a phantom — Skill LAB*):** Desconsiderada da avaliação.
          2. **`3-Su12` a `3-Su17` (*Suturas / QE*):** Desconsideradas até definição específica de registo.
          3. **`3-Su01` (*Preparing and sterilizing surgical instruments and equipment*):** Mantida sem mínimo definido até segunda ordem (não gera pendência).
          4. **Registos curriculares (`UC01`–`UC63`) e dissecções (`3-Diss1`–`3-Diss7`):** Não entram na matriz de subcompetências clínicas.
        """
    )

tabs = st.tabs([
    "1️⃣ Report por UC",
    "2️⃣ Monitores em Atraso",
    "3️⃣ Professores em Atraso",
    "4️⃣ Competências por Aluno",
    "5️⃣ Gerencial (Abaixo do Mínimo)",
    "⚙️ Parâmetros (UCs & Mínimos)",
])

# ===========================================================================
# TAB 1: REPORT PARA UNIDADES CURRICULARES
# ===========================================================================
with tabs[0]:
    st.subheader("1️⃣ Report por Unidade Curricular (Período, Espécies e Atividades)")
    if df_act_all.empty:
        st.info("Carregue pelo menos um ficheiro **clinical-activities** na barra lateral para gerar este report.")
    else:
        uc_df = load_uc_config()
        cursos_disp = sorted(df_act_all["Curso_Tag"].unique().tolist())
        curso_r1 = st.selectbox("Curso", cursos_disp, key="r1_curso")

        df_act_curso = df_act_all[df_act_all["Curso_Tag"] == curso_r1].copy()
        uc_curso = uc_df[uc_df["Curso"] == curso_r1].reset_index(drop=True)

        use_preset = st.checkbox("Selecionar configuração a partir da Tabela de UCs", value=not uc_curso.empty)

        all_activities_avail = sorted(df_act_curso["Atividade"].dropna().unique().tolist())
        all_subjects_avail = sorted(df_act_curso["Subject"].dropna().unique().tolist())

        if use_preset and not uc_curso.empty:
            uc_idx = st.selectbox(
                "Unidade Curricular configurada",
                options=range(len(uc_curso)),
                format_func=lambda i: f"{uc_curso.loc[i, 'UC / Disciplina']} ({uc_curso.loc[i, 'Data Início']} a {uc_curso.loc[i, 'Data Fim']})",
            )
            row_uc = uc_curso.loc[uc_idx]
            uc_name = str(row_uc["UC / Disciplina"])
            dt_ini = row_uc["Data Início"]
            dt_fim = row_uc["Data Fim"]
            sp_default = [s.strip() for s in str(row_uc["Espécies"]).split(",") if s.strip() in ALL_SPECIES]
            act_default = [a.strip() for a in str(row_uc["Atividades"]).split(",") if a.strip() in all_activities_avail]
            subj_default = [str(row_uc["Ano / Subject"]).strip()] if str(row_uc["Ano / Subject"]).strip() in all_subjects_avail else []
        else:
            uc_name = st.text_input("Nome da Unidade Curricular", value="Unidade Curricular")
            c1, c2 = st.columns(2)
            dt_ini = c1.date_input("Data de Início", value=date(2026, 2, 1), key="r1_ini")
            dt_fim = c2.date_input("Data de Fim", value=date.today(), key="r1_fim")
            sp_default = list(ALL_SPECIES)
            act_default = all_activities_avail
            subj_default = []

        col_f1, col_f2, col_f3 = st.columns(3)
        sel_subjects = col_f1.multiselect("Ano / Subject (opcional)", all_subjects_avail, default=subj_default, key="r1_subj")
        sel_species = col_f2.multiselect("Espécies a contabilizar", ALL_SPECIES, default=sp_default, key="r1_sp")
        sel_activities = col_f3.multiselect("Atividades a considerar", all_activities_avail, default=act_default, key="r1_act")

        col_exp1, col_exp2 = st.columns(2)
        want_pdf_1 = col_exp1.checkbox("Exportar PDF", value=True, key="r1_pdf")
        want_xlsx_1 = col_exp2.checkbox("Exportar Excel (.xlsx)", value=False, key="r1_xlsx")

        if st.button("Gerar Report 1 (UC)", type="primary"):
            res_uc, meta_uc = build_report_1_uc(
                df_act_curso,
                uc_name=uc_name,
                dt_start=dt_ini,
                dt_end=dt_fim,
                selected_species=sel_species,
                selected_activities=sel_activities,
                subject_filter=sel_subjects,
            )
            st.markdown(
                f"**Resumo da Extração:** `{meta_uc['logs_approved']}` logs *Approved* | "
                f"`{meta_uc['logs_da']}` registos *Approved como DA*"
            )
            st.dataframe(res_uc, use_container_width=True, hide_index=True)

            safe_uc = uc_name.replace(" ", "_")[:30]
            if want_pdf_1:
                pdf_bytes = generate_pdf_report_1(res_uc, meta_uc, curso_r1)
                st.download_button(
                    "⬇️ Baixar PDF — Report UC",
                    data=pdf_bytes,
                    file_name=f"Report_UC_{curso_r1}_{safe_uc}.pdf",
                    mime="application/pdf",
                )
            if want_xlsx_1:
                green_cols = set(LIVESTOCK_SPECIES + ["Espécie Pecuária"])
                xlsx_bytes = to_excel_bytes(
                    {
                        "Report UC": (
                            res_uc,
                            f"RELATÓRIO DE ATIVIDADES CLÍNICAS POR UC — {curso_r1}",
                            f"UC: {uc_name} | Período: {meta_uc['dt_start']} a {meta_uc['dt_end']} | Atividades: {meta_uc['activities']}",
                        )
                    },
                    green_cols=green_cols,
                )
                st.download_button(
                    "⬇️ Baixar Excel (.xlsx) — Report UC",
                    data=xlsx_bytes,
                    file_name=f"Report_UC_{curso_r1}_{safe_uc}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ===========================================================================
# TAB 2: MONITORES DO HVUEM EM ATRASO
# ===========================================================================
with tabs[1]:
    st.subheader("2️⃣ Report de Monitores do HVUEM em Atraso (Todos os Cursos)")
    if df_app_all.empty:
        st.info("Carregue pelo menos um ficheiro **approval-control** na barra lateral.")
    else:
        c1, c2 = st.columns(2)
        ref_dt_2 = c1.date_input("Data de referência para contagem de atraso", value=date.today(), key="r2_dt")
        grace_2 = c2.number_input("Dias de carência", min_value=0, max_value=30, value=3, key="r2_grace")

        ce1, ce2 = st.columns(2)
        want_pdf_2 = ce1.checkbox("Exportar PDF", value=True, key="r2_pdf")
        want_xlsx_2 = ce2.checkbox("Exportar Excel (.xlsx)", value=False, key="r2_xlsx")

        if st.button("Gerar Report 2 (Monitores)", type="primary"):
            df_mon = build_report_2_3_delays(df_app_all, perfil="MONITOR", ref_date=ref_dt_2, grace_days=int(grace_2))
            st.dataframe(df_mon, use_container_width=True, hide_index=True)

            if want_pdf_2:
                pdf_b2 = generate_pdf_report_2_3(
                    df_mon,
                    "REPORT DE MONITORES DO HVUEM EM ATRASO",
                    ref_dt_2,
                    int(grace_2),
                )
                st.download_button(
                    "⬇️ Baixar PDF — Monitores em Atraso",
                    data=pdf_b2,
                    file_name=f"Report_Monitores_Atraso_{ref_dt_2.strftime('%d-%m-%Y')}.pdf",
                    mime="application/pdf",
                )
            if want_xlsx_2:
                xlsx_b2 = to_excel_bytes({
                    "Monitores em Atraso": (
                        df_mon,
                        "REPORT DE MONITORES DO HVUEM EM ATRASO",
                        f"Data de referência: {ref_dt_2.strftime('%d/%m/%Y')} | Carência: {int(grace_2)} dias (Todos os cursos)",
                    )
                })
                st.download_button(
                    "⬇️ Baixar Excel (.xlsx) — Monitores em Atraso",
                    data=xlsx_b2,
                    file_name=f"Report_Monitores_Atraso_{ref_dt_2.strftime('%d-%m-%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ===========================================================================
# TAB 3: PROFESSORES EM ATRASO
# ===========================================================================
with tabs[2]:
    st.subheader("3️⃣ Report de Professores em Atraso (Todos os Cursos — Com e Sem Monitor Somados)")
    if df_app_all.empty:
        st.info("Carregue pelo menos um ficheiro **approval-control** na barra lateral.")
    else:
        c1, c2 = st.columns(2)
        ref_dt_3 = c1.date_input("Data de referência para contagem de atraso", value=date.today(), key="r3_dt")
        grace_3 = c2.number_input("Dias de carência", min_value=0, max_value=30, value=3, key="r3_grace")

        ce1, ce2 = st.columns(2)
        want_pdf_3 = ce1.checkbox("Exportar PDF", value=True, key="r3_pdf")
        want_xlsx_3 = ce2.checkbox("Exportar Excel (.xlsx)", value=False, key="r3_xlsx")

        if st.button("Gerar Report 3 (Professores)", type="primary"):
            df_prof = build_report_2_3_delays(df_app_all, perfil="TEACHER", ref_date=ref_dt_3, grace_days=int(grace_3))
            st.dataframe(df_prof, use_container_width=True, hide_index=True)

            if want_pdf_3:
                pdf_b3 = generate_pdf_report_2_3(
                    df_prof,
                    "REPORT DE PROFESSORES EM ATRASO (LOGBOOK)",
                    ref_dt_3,
                    int(grace_3),
                )
                st.download_button(
                    "⬇️ Baixar PDF — Professores em Atraso",
                    data=pdf_b3,
                    file_name=f"Report_Professores_Atraso_{ref_dt_3.strftime('%d-%m-%Y')}.pdf",
                    mime="application/pdf",
                )
            if want_xlsx_3:
                xlsx_b3 = to_excel_bytes({
                    "Professores em Atraso": (
                        df_prof,
                        "REPORT DE PROFESSORES EM ATRASO (LOGBOOK)",
                        f"Data de referência: {ref_dt_3.strftime('%d/%m/%Y')} | Carência: {int(grace_3)} dias (Todos os cursos)",
                    )
                })
                st.download_button(
                    "⬇️ Baixar Excel (.xlsx) — Professores em Atraso",
                    data=xlsx_b3,
                    file_name=f"Report_Professores_Atraso_{ref_dt_3.strftime('%d-%m-%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ===========================================================================
# TAB 4: COMPETÊNCIAS ADQUIRIDAS POR ALUNO
# ===========================================================================
with tabs[3]:
    st.subheader("4️⃣ Número de Competências Adquiridas por Aluno (Curso / Ano)")
    if df_sub_all.empty:
        st.info("Carregue pelo menos um ficheiro **sub-competences-student-subject** na barra lateral.")
    else:
        rules_df = load_rules_config()
        cursos_4 = sorted(df_sub_all["Curso_Tag"].unique().tolist())
        c1, c2, c3 = st.columns(3)
        curso_4 = c1.selectbox("Curso", cursos_4, key="r4_curso")
        df_sub_c4 = df_sub_all[df_sub_all["Curso_Tag"] == curso_4]
        subjects_4 = sorted(df_sub_c4["Subject_Clean"].unique().tolist())
        subj_4 = c2.selectbox("Ano / Subject", subjects_4, key="r4_subj")
        origens_4 = sorted(df_sub_c4["Origem_Clean"].unique().tolist())
        sel_orig_4 = c3.multiselect("Origem da atribuição", origens_4, default=origens_4, key="r4_orig")

        df_sub_s4 = df_sub_c4[
            (df_sub_c4["Subject_Clean"] == subj_4) & (df_sub_c4["Origem_Clean"].isin(sel_orig_4))
        ]
        students_4 = sorted(df_sub_c4[df_sub_c4["Subject_Clean"] == subj_4]["Nome_Clean"].unique().tolist())

        modo_4 = st.radio(
            "Conteúdo do relatório por aluno:",
            options=["Apenas competências que ainda NÃO atingiram o mínimo", "Todas as competências (obtidas e não obtidas)"],
            horizontal=True,
            key="r4_modo",
        )
        only_below_4 = modo_4.startswith("Apenas")

        sel_students_4 = st.multiselect(
            "Selecione os alunos (vazio = todos os alunos do ano selecionado)",
            options=students_4,
            default=students_4[:1] if students_4 else [],
            key="r4_students",
        )
        target_students_4 = sel_students_4 if sel_students_4 else students_4

        ce1, ce2 = st.columns(2)
        want_pdf_4 = ce1.checkbox("Exportar PDF (ZIP se múltiplos alunos)", value=True, key="r4_pdf")
        want_xlsx_4 = ce2.checkbox("Exportar Excel (.xlsx — 1 aba por aluno)", value=False, key="r4_xlsx")

        if st.button("Gerar Report 4 (Por Aluno)", type="primary"):
            if not target_students_4:
                st.warning("Nenhum estudante encontrado para os filtros aplicados.")
            else:
                eval_dict = evaluate_student_competences(df_sub_s4, target_students_4, rules_df)
                first_st = target_students_4[0]
                df_preview = eval_dict[first_st]
                if only_below_4:
                    df_preview = df_preview[df_preview["Status"] == "Abaixo do Mínimo"]
                st.caption(f"Pré-visualização: **{first_st}** ({len(df_preview)} linhas)")
                st.dataframe(df_preview, use_container_width=True, hide_index=True)

                if want_pdf_4:
                    if len(target_students_4) == 1:
                        st_single = target_students_4[0]
                        pdf_single = generate_pdf_report_4_student(
                            st_single, eval_dict[st_single], curso_4, subj_4, only_below_4
                        )
                        st.download_button(
                            f"⬇️ Baixar PDF — {st_single}",
                            data=pdf_single,
                            file_name=f"Competencias_{curso_4}_{st_single.replace(' ', '_')}.pdf",
                            mime="application/pdf",
                        )
                    else:
                        zip_buf = io.BytesIO()
                        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                            for st_name in target_students_4:
                                pdf_st = generate_pdf_report_4_student(
                                    st_name, eval_dict[st_name], curso_4, subj_4, only_below_4
                                )
                                safe_st = st_name.replace(" ", "_").replace("/", "-")
                                zf.writestr(f"{curso_4}_{subj_4}_{safe_st}.pdf", pdf_st)
                        st.download_button(
                            f"⬇️ Baixar ZIP com {len(target_students_4)} PDFs Individuais",
                            data=zip_buf.getvalue(),
                            file_name=f"Competencias_{curso_4}_{subj_4.replace(' ', '_')}.zip",
                            mime="application/zip",
                        )

                if want_xlsx_4:
                    sheets = {}
                    for st_name in target_students_4:
                        df_s = eval_dict[st_name]
                        if only_below_4:
                            df_s = df_s[df_s["Status"] == "Abaixo do Mínimo"]
                        sheets[st_name] = (
                            df_s,
                            f"COMPETÊNCIAS INDIVIDUAIS — {st_name}",
                            f"Curso: {curso_4} | Ano: {subj_4} | Filtro: {modo_4}",
                        )
                    xlsx_b4 = to_excel_bytes(sheets, highlight_red_col="Status")
                    st.download_button(
                        "⬇️ Baixar Excel (.xlsx) — Competências por Aluno",
                        data=xlsx_b4,
                        file_name=f"Competencias_Alunos_{curso_4}_{subj_4.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

# ===========================================================================
# TAB 5: REPORT GERENCIAL (COMPETÊNCIAS ABAIXO DO MÍNIMO)
# ===========================================================================
with tabs[4]:
    st.subheader("5️⃣ Report Gerencial — Competências Abaixo do Mínimo por Curso / Ano")
    if df_sub_all.empty:
        st.info("Carregue pelo menos um ficheiro **sub-competences-student-subject** na barra lateral.")
    else:
        rules_df = load_rules_config()
        cursos_5 = sorted(df_sub_all["Curso_Tag"].unique().tolist())
        c1, c2 = st.columns(2)
        curso_5 = c1.selectbox("Curso", cursos_5, key="r5_curso")
        df_sub_c5 = df_sub_all[df_sub_all["Curso_Tag"] == curso_5]
        subjects_5 = sorted(df_sub_c5["Subject_Clean"].unique().tolist())
        subj_5 = c2.selectbox("Ano / Subject", subjects_5, key="r5_subj")

        ce1, ce2 = st.columns(2)
        want_pdf_5 = ce1.checkbox("Exportar PDF", value=True, key="r5_pdf")
        want_xlsx_5 = ce2.checkbox("Exportar Excel (.xlsx)", value=False, key="r5_xlsx")

        if st.button("Gerar Report 5 (Gerencial)", type="primary"):
            df_sub_s5 = df_sub_c5[df_sub_c5["Subject_Clean"] == subj_5]
            students_5 = sorted(df_sub_s5["Nome_Clean"].unique().tolist())
            eval_dict_5 = evaluate_student_competences(df_sub_s5, students_5, rules_df)
            df_ger = build_report_5_managerial(eval_dict_5, curso_5, subj_5)

            st.dataframe(df_ger, use_container_width=True, hide_index=True)

            if want_pdf_5 and not df_ger.empty:
                pdf_b5 = generate_pdf_report_5(df_ger, curso_5, subj_5)
                st.download_button(
                    "⬇️ Baixar PDF — Report Gerencial",
                    data=pdf_b5,
                    file_name=f"Report_Gerencial_Competencias_{curso_5}_{subj_5.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                )
            if want_xlsx_5 and not df_ger.empty:
                xlsx_b5 = to_excel_bytes({
                    "Gerencial Abaixo Mínimo": (
                        df_ger,
                        f"RELATÓRIO GERENCIAL — COMPETÊNCIAS ABAIXO DO MÍNIMO ({curso_5})",
                        f"Curso: {curso_5} | Ano / Subject: {subj_5} | Total de alunos avaliados: {len(students_5)}",
                    )
                })
                st.download_button(
                    "⬇️ Baixar Excel (.xlsx) — Report Gerencial",
                    data=xlsx_b5,
                    file_name=f"Report_Gerencial_Competencias_{curso_5}_{subj_5.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ===========================================================================
# TAB 6: PARÂMETROS EDITÁVEIS (TABELA DE UCs E CATÁLOGO DE MÍNIMOS)
# ===========================================================================
with tabs[5]:
    st.subheader("⚙️ Parametrização de Unidades Curriculares e Mínimos de Competências")

    st.markdown("### A. Tabela de Unidades Curriculares (Report 1)")
    st.caption(
        "Configure aqui as UCs com período de início/fim, espécies e atividades. "
        "Separe múltiplas espécies ou atividades por vírgula."
    )
    uc_current = load_uc_config()
    uc_edited = st.data_editor(
        uc_current,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Curso": st.column_config.SelectboxColumn("Curso", options=["MIMV", "LEV", "LCB"], required=True),
            "Data Início": st.column_config.DateColumn("Data Início", format="DD/MM/YYYY"),
            "Data Fim": st.column_config.DateColumn("Data Fim", format="DD/MM/YYYY"),
        },
        key="editor_uc",
    )
    if st.button("💾 Guardar Tabela de UCs"):
        save_uc_config(uc_edited)
        st.success("Configuração de UCs guardada com sucesso.")

    st.markdown("---")
    st.markdown("### B. Catálogo Oficial de Competências e Mínimos (Reports 4 e 5)")
    st.caption(
        "Use '/' para espécies alternativas (OU) e '+' para blocos cumulativos (E). "
        "Ex.: `D/C + R/ExB/ExRo/ExO/F` com mínimo `3+3`. "
        "Altere o Estado para `Ativa` quando definir o mínimo de `3-Su01`."
    )
    rules_current = load_rules_config()
    rules_edited = st.data_editor(
        rules_current,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Estado": st.column_config.SelectboxColumn(
                "Estado",
                options=["Ativa", "Desconsiderada", "Sem mínimo definido"],
                required=True,
            ),
        },
        key="editor_rules",
    )
    c_r1, c_r2 = st.columns(2)
    if c_r1.button("💾 Guardar Regras de Competências"):
        save_rules_config(rules_edited)
        st.success("Catálogo de competências atualizado.")
    if c_r2.button("🔄 Repor Catálogo Oficial (Competencias.pdf)"):
        default_df = pd.DataFrame(get_default_rules_list())
        save_rules_config(default_df)
        st.success("Catálogo oficial reposto.")
