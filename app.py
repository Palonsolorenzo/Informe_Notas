import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

from data_processor import (
    get_excel_sheets,
    preprocess_dataframe,
    enrich_data,
    get_global_metrics,
    COL_ASIGNATURA,
    COL_MATRICULADOS,
    COL_CO_TOTAL,
    COL_CE_TOTAL,
    COL_TOTAL_ACUMULADO,
    COL_PCT_CO_TOTAL,
    COL_PCT_CE_TOTAL,
    COL_PCT_TOTAL_ACUMULADO,
    COL_PCT_CO_EV_CONT,
    COL_PCT_CO_EV_FIN,
    COL_PCT_CE_EV_CONT,
    COL_PCT_CE_EV_FIN,
)
from excel_generator import generate_styled_excel

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Dashboard de Rendimiento Académico",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados
st.markdown("""
    <style>
    .degree-badge {
        background-color: #E2EFDA;
        color: #1E4D2B;
        font-size: 1.15rem;
        font-weight: 700;
        padding: 0.4rem 0.8rem;
        border-radius: 6px;
        display: inline-block;
        margin-bottom: 0.8rem;
        border: 1px solid #C6E0B4;
    }
    .main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1F4E78;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #555555;
        margin-bottom: 1.2rem;
    }
    .stDownloadButton > button {
        background-color: #1E4D2B;
        color: white;
        font-weight: bold;
        border-radius: 6px;
        padding: 0.6rem 1.2rem;
        border: none;
    }
    .stDownloadButton > button:hover {
        background-color: #14351D;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    # Sidebar
    st.sidebar.image("https://img.icons8.com/color/96/000000/graduation-cap.png", width=70)
    st.sidebar.title("Panel de Control")
    st.sidebar.markdown("---")
    
    sample_file_path = "examenes_20261001_165233.xlsx"
    has_sample = os.path.exists(sample_file_path)
    
    file_option = st.sidebar.radio(
        "Origen de los Datos:",
        ["Usar archivo de muestra", "Subir propio archivo Excel"] if has_sample else ["Subir propio archivo Excel"]
    )
    
    uploaded_file = None
    if file_option == "Subir propio archivo Excel":
        uploaded_file = st.sidebar.file_uploader(
            "Cargar archivo Excel en bruto (.xlsx)",
            type=["xlsx", "xls"]
        )
        if not uploaded_file:
            st.info("👈 Por favor, sube un archivo Excel en el menú lateral para generar el dashboard e informe.")
            st.stop()
    else:
        uploaded_file = sample_file_path

    # Obtener pestañas (titulaciones)
    sheets = get_excel_sheets(uploaded_file)
    selected_sheet = sheets[0]
    
    if len(sheets) > 1:
        selected_sheet = st.sidebar.selectbox(
            "Seleccionar Titulación (Pestaña Excel):",
            options=sheets,
            index=0
        )
    else:
        st.sidebar.markdown(f"**Titulación detectada:** `{selected_sheet}`")

    # Cargar y procesar datos de la pestaña seleccionada
    try:
        raw_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        df_clean = preprocess_dataframe(raw_df)
        df_enriched = enrich_data(df_clean)
        metrics = get_global_metrics(df_enriched)
    except Exception as e:
        st.error(f"Error al leer la pestaña '{selected_sheet}': {e}")
        st.stop()

    degree_name = selected_sheet

    # Filtros en Sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Filtros de Asignaturas")
    search_query = st.sidebar.text_input("Buscar por nombre de asignatura:")
    
    min_matr = int(df_enriched[COL_MATRICULADOS].min())
    max_matr = int(df_enriched[COL_MATRICULADOS].max())
    
    if max_matr > min_matr:
        selected_matr = st.sidebar.slider(
            "Nº mínimo de matriculados:",
            min_value=min_matr,
            max_value=max_matr,
            value=min_matr
        )
    else:
        selected_matr = min_matr

    # Aplicar filtros
    filtered_df = df_enriched[df_enriched[COL_MATRICULADOS] >= selected_matr]
    if search_query:
        filtered_df = filtered_df[filtered_df[COL_ASIGNATURA].str.contains(search_query, case=False, na=False)]

    # Encabezado Principal con Titulación
    st.markdown(f'<div class="degree-badge">🎓 TITULACIÓN: {degree_name.upper()}</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-title">📊 Dashboard Estadístico de Rendimiento Académico</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-title">Análisis detallado para <b>{degree_name}</b> en Convocatoria Ordinaria (CO) y Extraordinaria (CE)</div>', unsafe_allow_html=True)

    # Tarjetas KPI Globales
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    
    with kpi1:
        st.metric("Total Asignaturas", f"{metrics['num_asignaturas']}")
    with kpi2:
        st.metric("Alumnos Matriculados", f"{metrics['tot_matriculados']:,}".replace(",", "."))
    with kpi3:
        st.metric(
            "Aprobados CO",
            f"{metrics['tot_co_total']:,}".replace(",", "."),
            delta=f"{metrics['pct_co_total']*100:.1f}% tasa global"
        )
    with kpi4:
        st.metric(
            "Aprobados CE",
            f"{metrics['tot_ce_total']:,}".replace(",", "."),
            delta=f"{metrics['pct_ce_total']*100:.1f}% tasa global"
        )
    with kpi5:
        st.metric(
            "Total Aprobados",
            f"{metrics['tot_acumulado']:,}".replace(",", "."),
            delta=f"{metrics['pct_acumulado']*100:.1f}% tasa global"
        )

    st.markdown("---")

    # Generar Excel maquetado con el nombre de la titulación
    excel_io = generate_styled_excel(df_enriched, degree_name=degree_name)
    
    # Botón prominente de descarga en la barra lateral
    st.sidebar.markdown("### 📥 Exportación")
    st.sidebar.download_button(
        label=f"📗 Descargar Informe Excel ({degree_name[:15]}...)",
        data=excel_io.getvalue(),
        file_name=f"Informe_Notas_{degree_name.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Pestañas principales (sin pestaña de tabla de datos)
    tab_dash, tab_download = st.tabs([
        "📈 Dashboard Gráfico",
        "📥 Informe Excel Maquetado"
    ])

    with tab_dash:
        st.subheader(f"Visualización Comparativa de Resultados - {degree_name}")
        
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            st.markdown("#### Tasa de Éxito (%) por Convocatoria")
            df_plot = filtered_df.copy()
            df_plot["Nombre_Corto"] = df_plot[COL_ASIGNATURA].apply(lambda x: x if len(x) <= 35 else x[:32] + "...")
            
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(
                x=df_plot["Nombre_Corto"],
                y=df_plot[COL_PCT_CO_TOTAL] * 100,
                name="Aprobados CO (%)",
                marker_color="#2F5597",
                hovertemplate="%{x}<br>Aprobados CO: %{y:.1f}%<extra></extra>"
            ))
            fig_bar.add_trace(go.Bar(
                x=df_plot["Nombre_Corto"],
                y=df_plot[COL_PCT_CE_TOTAL] * 100,
                name="Aprobados CE (%)",
                marker_color="#70AD47",
                hovertemplate="%{x}<br>Aprobados CE: %{y:.1f}%<extra></extra>"
            ))
            fig_bar.update_layout(
                barmode="group",
                xaxis_title="Asignatura",
                yaxis_title="Porcentaje (%)",
                yaxis=dict(range=[0, 105]),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                height=450,
                margin=dict(l=20, r=20, t=30, b=120)
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_right:
            st.markdown("#### Impacto de Evaluación Continua vs Final (CO)")
            fig_scat = px.scatter(
                filtered_df,
                x=COL_PCT_CO_EV_CONT,
                y=COL_PCT_CO_EV_FIN,
                size=COL_MATRICULADOS,
                hover_name=COL_ASIGNATURA,
                color=COL_PCT_TOTAL_ACUMULADO,
                color_continuous_scale="Viridis",
                labels={
                    COL_PCT_CO_EV_CONT: "% Superan Ev. Continua CO",
                    COL_PCT_CO_EV_FIN: "% Superan Ev. Final CO",
                    COL_PCT_TOTAL_ACUMULADO: "% Total Aprobados"
                }
            )
            fig_scat.update_traces(marker=dict(sizemode='area', sizeref=2.*max(filtered_df[COL_MATRICULADOS])/(40.**2), sizemin=6))
            fig_scat.update_layout(
                xaxis_tickformat=".0%",
                yaxis_tickformat=".0%",
                height=450
            )
            st.plotly_chart(fig_scat, use_container_width=True)

        st.markdown("---")
        
        # Fila de Ranking Top/Bottom
        r_col1, r_col2 = st.columns(2)
        
        with r_col1:
            st.markdown(f"#### 🏆 Top 5 Asignaturas con Mayor Tasa de Aprobados ({degree_name})")
            top_5 = df_enriched.nlargest(5, COL_PCT_TOTAL_ACUMULADO)
            fig_top = px.bar(
                top_5,
                y=COL_ASIGNATURA,
                x=COL_PCT_TOTAL_ACUMULADO,
                orientation='h',
                text=top_5[COL_PCT_TOTAL_ACUMULADO].apply(lambda x: f"{x*100:.1f}%"),
                color=COL_PCT_TOTAL_ACUMULADO,
                color_continuous_scale="Greens",
                labels={COL_PCT_TOTAL_ACUMULADO: "Tasa Aprobados", COL_ASIGNATURA: ""}
            )
            fig_top.update_layout(yaxis=dict(autorange="reversed"), xaxis_tickformat=".0%", height=320, showlegend=False)
            st.plotly_chart(fig_top, use_container_width=True)

        with r_col2:
            st.markdown(f"#### ⚠️ 5 Asignaturas con Menor Tasa de Aprobados ({degree_name})")
            bottom_5 = df_enriched.nsmallest(5, COL_PCT_TOTAL_ACUMULADO)
            fig_bot = px.bar(
                bottom_5,
                y=COL_ASIGNATURA,
                x=COL_PCT_TOTAL_ACUMULADO,
                orientation='h',
                text=bottom_5[COL_PCT_TOTAL_ACUMULADO].apply(lambda x: f"{x*100:.1f}%"),
                color=COL_PCT_TOTAL_ACUMULADO,
                color_continuous_scale="Reds_r",
                labels={COL_PCT_TOTAL_ACUMULADO: "Tasa Aprobados", COL_ASIGNATURA: ""}
            )
            fig_bot.update_layout(yaxis=dict(autorange="reversed"), xaxis_tickformat=".0%", height=320, showlegend=False)
            st.plotly_chart(fig_bot, use_container_width=True)

    with tab_download:
        st.subheader(f"Informe Excel Maquetado para {degree_name}")
        st.markdown(f"""
        El informe Excel descargable incluirá la titulación **"{degree_name}"** en:
        - **Nombre de la Pestaña Excel**: `{degree_name[:31]}`
        - **Cabecera principal (Celda A1)**: `INFORME ESTADÍSTICO DE RESULTADOS - {degree_name.upper()}`
        - **Subtítulo (Celda A2)**: `Titulación: {degree_name}`
        - **Fila final de Totales**: `TOTALES GLOBALES ({degree_name.upper()})`
        """)
        
        st.download_button(
            label=f"📥 DESCARGAR INFORME EN EXCEL DE {degree_name.upper()} (.XLSX)",
            data=excel_io.getvalue(),
            file_name=f"Informe_Notas_{degree_name.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
