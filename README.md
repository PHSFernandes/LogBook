# Plataforma de Emissão de Reports — Logbook Clínico (Egas Moniz)

Aplicação Streamlit para emissão automatizada dos 5 relatórios de acompanhamento do Logbook Clínico (`MIMV`, `LEV`, `LCB`) em **PDF** e **Excel (.xlsx)**.

## Ficheiros do Repositório
- `app.py`: Interface principal em Streamlit, filtros, editores de parâmetros e downloads sob demanda.
- `rules_data.py`: Catálogo oficial de subcompetências (`Competencias.pdf`), parser de regras combinadas (`/`, `+`) e exceções configuradas.
- `report_generators.py`: Tratamento dos CSVs (`approval-control`, `clinical-activities`, `sub-competences-student-subject`) e motores de geração PDF (ReportLab) e Excel (OpenPyXL).
- `requirements.txt`: Dependências Python para Streamlit Cloud.
- `packages.txt`: Fontes de sistema (`fonts-dejavu-core`) para renderização PDF em Linux.
