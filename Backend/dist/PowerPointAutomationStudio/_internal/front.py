import os
import shutil
import tempfile
import uuid
from pathlib import Path
import pandas as pd
import streamlit as st
from pptx import Presentation

from main import (
    PREVIEW_DIR,
    check_thinkcell_available,
    create_real_powerpoint_preview_images,
    normalize_mapping_schema,
    read_hidden_mappings_from_presentation,
    replace_charts_with_matching_excel_data,
    update_thinkcell_charts_with_com,
)


# ------------------------------------------------------------
# Page Config
# ------------------------------------------------------------

st.set_page_config(
    page_title="PowerPoint Automation Studio",
    page_icon="📊",
    layout="wide",
)


# ------------------------------------------------------------
# Styling
# ------------------------------------------------------------

st.markdown(
    """
    <style>
        .main-title {
            font-size: 2.4rem;
            font-weight: 800;
            margin-bottom: 0.2rem;
        }

        .sub-title {
            color: #6b7280;
            font-size: 1.05rem;
            margin-bottom: 1.5rem;
        }

        .section-card {
            padding: 1.2rem 1.4rem;
            border-radius: 18px;
            border: 1px solid #e5e7eb;
            background: #ffffff;
            box-shadow: 0 4px 18px rgba(0,0,0,0.04);
            margin-bottom: 1rem;
        }

        .metric-card {
            padding: 1rem;
            border-radius: 16px;
            border: 1px solid #e5e7eb;
            background: #f9fafb;
            text-align: center;
        }

        .small-muted {
            color: #6b7280;
            font-size: 0.9rem;
        }

        .success-box {
            padding: 1rem;
            border-radius: 14px;
            background: #ecfdf5;
            border: 1px solid #a7f3d0;
            color: #065f46;
            margin-bottom: 1rem;
        }

        .warning-box {
            padding: 1rem;
            border-radius: 14px;
            background: #fffbeb;
            border: 1px solid #fde68a;
            color: #92400e;
            margin-bottom: 1rem;
        }

        .placeholder-box {
            padding: 2rem;
            border-radius: 18px;
            border: 2px dashed #d1d5db;
            background: #f9fafb;
            color: #6b7280;
            text-align: center;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Header
# ------------------------------------------------------------

st.markdown('<div class="main-title">📊 PowerPoint Automation Studio</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">PowerPoint aktualisieren, Folien prüfen und Excel-Daten später zusammenführen.</div>',
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Helper
# ------------------------------------------------------------

def save_uploaded_file(uploaded_file, target_path: Path):
    target_path.parent.mkdir(parents=True, exist_ok=True)
    with target_path.open("wb") as buffer:
        shutil.copyfileobj(uploaded_file, buffer)


def generate_powerpoint_streamlit(excel_upload, ppt_upload) -> dict:
    job_id = str(uuid.uuid4())
    temp_dir = Path(tempfile.mkdtemp(prefix="ppt_automation_"))

    excel_path = temp_dir / f"{job_id}_{excel_upload.name}"
    template_path = temp_dir / f"{job_id}_{ppt_upload.name}"
    output_path = temp_dir / f"{job_id}_Aktualisierte_Praesentation.pptx"

    save_uploaded_file(excel_upload, excel_path)
    save_uploaded_file(ppt_upload, template_path)

    presentation = Presentation(template_path)

    raw_mappings = read_hidden_mappings_from_presentation(presentation)
    mappings = [normalize_mapping_schema(mapping) for mapping in raw_mappings]

    updated_native_charts = 0
    updated_thinkcell_charts = 0

    if mappings:
        mode_used = "thinkcell_mapping"

        if os.name != "nt":
            availability = check_thinkcell_available()
            raise RuntimeError(
                "Diese PowerPoint enthält think-cell-Mappings. "
                f"think-cell kann hier aber nicht aktualisiert werden: {availability.get('reason')}"
            )

        updated_thinkcell_charts = update_thinkcell_charts_with_com(
            ppt_template_path=template_path,
            excel_path=excel_path,
            output_path=output_path,
            mappings=mappings,
        )

    else:
        mode_used = "native_fallback"

        updated_native_charts = replace_charts_with_matching_excel_data(
            presentation=presentation,
            excel_path=excel_path,
        )

        presentation.save(output_path)

    preview_folder = PREVIEW_DIR / job_id

    preview_images = []
    preview_error = None

    try:
        preview_images = create_real_powerpoint_preview_images(
            pptx_path=output_path,
            preview_folder=preview_folder,
            job_id=job_id,
        )
    except Exception as error:
        preview_error = str(error)

    return {
        "job_id": job_id,
        "mode": mode_used,
        "mappings_found": len(mappings),
        "updated_native_charts": updated_native_charts,
        "updated_thinkcell_charts": updated_thinkcell_charts,
        "output_path": output_path,
        "preview_images": preview_images,
        "preview_error": preview_error,
    }


def reset_powerpoint_result():
    st.session_state.generation_result = None


# ------------------------------------------------------------
# Session State
# ------------------------------------------------------------

if "generation_result" not in st.session_state:
    st.session_state.generation_result = None

if "selected_slide_index" not in st.session_state:
    st.session_state.selected_slide_index = 0


# ------------------------------------------------------------
# Tabs
# ------------------------------------------------------------

tab_powerpoint, tab_excel = st.tabs(
    [
        "📊 PowerPoint erstellen",
        "📁 Excel zusammenführen",
    ]
)


# ============================================================
# TAB 1: POWERPOINT ERSTELLEN
# ============================================================

with tab_powerpoint:
    left_col, right_col = st.columns([0.35, 0.65], gap="large")

    with left_col:
        st.markdown("### Dateien hochladen")

        excel_file = st.file_uploader(
            "Excel-Datei",
            type=["xlsx", "xlsm", "xls"],
            key="ppt_excel_upload",
            on_change=reset_powerpoint_result,
        )

        ppt_file = st.file_uploader(
            "PowerPoint-Vorlage",
            type=["pptx"],
            key="ppt_template_upload",
            on_change=reset_powerpoint_result,
        )

        st.divider()

        generate_disabled = not excel_file or not ppt_file

        if st.button(
            "🚀 PowerPoint generieren",
            type="primary",
            use_container_width=True,
            disabled=generate_disabled,
        ):
            try:
                with st.spinner("PowerPoint wird aktualisiert..."):
                    st.session_state.generation_result = generate_powerpoint_streamlit(
                        excel_upload=excel_file,
                        ppt_upload=ppt_file,
                    )
                    st.session_state.selected_slide_index = 0

                st.success("PowerPoint wurde erfolgreich erstellt.")

            except Exception as error:
                st.session_state.generation_result = None
                st.error(f"Fehler beim Erstellen der PowerPoint: {error}")

        if generate_disabled:
            st.info("Bitte Excel-Datei und PowerPoint-Datei hochladen.")

    with right_col:
        result = st.session_state.generation_result

        if not result:
            st.markdown(
                """
                <div class="placeholder-box">
                    <h3>PowerPoint-Vorschau</h3>
                    <p>Nach dem Generieren kannst du hier die Folien durchgehen.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

        else:
            st.markdown(
                """
                <div class="success-box">
                    <strong>PowerPoint ist bereit.</strong><br>
                    Prüfe die Folien und lade danach die aktualisierte Datei herunter.
                </div>
                """,
                unsafe_allow_html=True,
            )

            metric_col_1, metric_col_2, metric_col_3, metric_col_4 = st.columns(4)

            total_updates = (
                result["updated_native_charts"]
                + result["updated_thinkcell_charts"]
            )

            with metric_col_1:
                st.metric("Modus", result["mode"])

            with metric_col_2:
                st.metric("Mappings", result["mappings_found"])

            with metric_col_3:
                st.metric("Native Charts", result["updated_native_charts"])

            with metric_col_4:
                st.metric("Updates gesamt", total_updates)

            st.divider()

            preview_images = result.get("preview_images", [])

            if preview_images:
                slide_count = len(preview_images)

                st.markdown("### Folien-Vorschau")

                slide_options = list(range(slide_count))

                selected_slide = st.select_slider(
                    "Folie auswählen",
                    options=slide_options,
                    value=min(st.session_state.selected_slide_index, slide_count - 1),
                    format_func=lambda index: f"Folie {index + 1}",
                )

                st.session_state.selected_slide_index = selected_slide

                nav_col_1, nav_col_2, nav_col_3 = st.columns([1, 2, 1])

                with nav_col_1:
                    if st.button(
                        "⬅️ Vorherige",
                        use_container_width=True,
                        disabled=selected_slide <= 0,
                    ):
                        st.session_state.selected_slide_index = max(0, selected_slide - 1)
                        st.rerun()

                with nav_col_2:
                    st.markdown(
                        f"<p style='text-align:center;'>Folie {selected_slide + 1} von {slide_count}</p>",
                        unsafe_allow_html=True,
                    )

                with nav_col_3:
                    if st.button(
                        "Nächste ➡️",
                        use_container_width=True,
                        disabled=selected_slide >= slide_count - 1,
                    ):
                        st.session_state.selected_slide_index = min(slide_count - 1, selected_slide + 1)
                        st.rerun()

                st.image(
                    str(preview_images[selected_slide]),
                    use_container_width=True,
                )

            else:
                preview_error = result.get("preview_error")

                if preview_error:
                    st.warning(
                        "PowerPoint wurde erstellt, aber die Vorschau konnte nicht erzeugt werden. "
                        f"Grund: {preview_error}"
                    )
                else:
                    st.info("Keine Vorschau-Bilder verfügbar.")

            st.divider()

            output_path = result["output_path"]

            with output_path.open("rb") as file:
                st.download_button(
                    label="⬇️ PowerPoint herunterladen",
                    data=file,
                    file_name="Aktualisierte_Praesentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True,
                )


# ============================================================
# TAB 2: EXCEL ZUSAMMENFÜHREN
# ============================================================

with tab_excel:
    left_col, right_col = st.columns([0.35, 0.65], gap="large")

    with left_col:
        st.markdown("### Excel-Dateien hochladen")

        excel_files = st.file_uploader(
            "Excel-Dateien für Zusammenführung",
            type=["xlsx", "xlsm", "xls"],
            accept_multiple_files=True,
            key="merge_excel_uploads",
        )

        st.divider()

        merge_type = st.selectbox(
            "Art der Zusammenführung",
            options=[
                "Noch nicht aktiv",
                "Mehrere Dateien untereinander anhängen",
                "Mehrere Sheets zusammenführen",
                "Kundenspezifische Logik",
            ],
        )

        output_name = st.text_input(
            "Name der späteren Output-Datei",
            value="Zusammengefuehrte_Daten.xlsx",
        )

        st.button(
            "🔧 Excel zusammenführen",
            type="primary",
            use_container_width=True,
            disabled=True,
        )

        st.caption("Backend für Excel-Zusammenführung wird später ergänzt.")

        with st.expander("ℹ️ Backend-Idee für später"):
            st.write(
                """
                Später kann das Backend hier fest codierte Regeln ausführen:

                - bestimmte Dateien erkennen
                - bestimmte Sheets lesen
                - Spaltennamen vereinheitlichen
                - Daten untereinander oder nebeneinander zusammenführen
                - Duplikate entfernen
                - finale Excel-Datei erzeugen
                - Ergebnis hier als Tabelle anzeigen und herunterladen
                """
            )

    with right_col:
        st.markdown("### Excel-Outcome Vorschau")

        if not excel_files:
            st.markdown(
                """
                <div class="placeholder-box">
                    <h3>Excel-Vorschau</h3>
                    <p>Lade Excel-Dateien hoch. Danach wird hier eine Vorschau angezeigt.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

        else:
            st.success(f"{len(excel_files)} Excel-Datei(en) hochgeladen.")

            selected_file = st.selectbox(
                "Datei für Vorschau auswählen",
                options=excel_files,
                format_func=lambda file: file.name,
            )

            try:
                selected_file.seek(0)
                excel_reader = pd.ExcelFile(selected_file)

                selected_sheet = st.selectbox(
                    "Sheet auswählen",
                    options=excel_reader.sheet_names,
                )

                selected_file.seek(0)
                preview_df = pd.read_excel(
                    selected_file,
                    sheet_name=selected_sheet,
                    nrows=100,
                )

                st.dataframe(
                    preview_df,
                    use_container_width=True,
                    height=420,
                )

                st.markdown(
                    """
                    <div class="warning-box">
                        <strong>Hinweis:</strong><br>
                        Das ist aktuell nur eine Frontend-Vorschau. 
                        Die echte Zusammenführungslogik kommt später ins Backend.
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            except Exception as error:
                st.error(f"Excel-Vorschau konnte nicht geladen werden: {error}")