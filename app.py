import streamlit as st
import datetime
import os
import json
import calendar as pycal
import pandas as pd
import re
from database import (
    db_init, db_import_excel, get_tasks, add_task, update_task_completion,
    update_task_priority, update_task_details, delete_task, get_last_finalized_date,
    get_undone_days_before, finalize_day, get_reports, get_days_with_pending_tasks,
    get_users, add_user, delete_user, get_overdue_pending_tasks,
    get_tasks_by_user_and_date, get_completed_tasks_count, get_last_report,
    get_last_export_info, save_last_export_info, update_report_suggestions
)
from ai_helper import generate_alternatives, generate_single_alternative
from excel_helper import export_weekly_tasks_to_excel, get_week_dates

# Initialize database and import data
db_init()
excel_file = "Excel de datos pizarra - copia.xlsx"
if os.path.exists(excel_file):
    db_import_excel(excel_file)

# ----------------- SESSION STATE & INITIAL USER -----------------
# Dynamically load users from DB
all_users = get_users()
if not all_users:
    all_users = ["MARY CRUZ", "CPC.SHEYLA", "CPC.HECTOR"]
    for u in all_users:
        add_user(u)

if "current_user" not in st.session_state or st.session_state["current_user"] not in all_users:
    st.session_state["current_user"] = all_users[0]
selected_user = st.session_state["current_user"]

# Query total notification counts
today = datetime.date.today()
today_str = today.strftime("%Y-%m-%d")
today_pending_tasks = [t for t in get_tasks(selected_user, today_str) if t['completed'] == 0]
overdue_pending_tasks = get_overdue_pending_tasks(selected_user, today_str)
total_notifications = len(today_pending_tasks) + len(overdue_pending_tasks)

# Style badge injection if notifications exist
badge_css = ""
if total_notifications > 0:
    badge_css = f"""
    div[data-testid="stPopover"] > button::after {{
        content: "{total_notifications}" !important;
        position: absolute !important;
        top: 2px !important;
        right: 2px !important;
        background-color: #ef4444 !important;
        color: white !important;
        font-size: 11px !important;
        font-weight: bold !important;
        border-radius: 50% !important;
        width: 20px !important;
        height: 20px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        border: 2px solid #1f77b4 !important;
    }}
    """

# ----------------- PAGE CONFIG -----------------
st.set_page_config(
    page_title="CUADROpz",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------- GLOBAL STYLING (DARK/LIGHT MODE OPTIMIZATIONS) -----------------
st.markdown(f"""
    <style>
    /* Base configuration (Light Mode) */
    .stApp {{
        background-color: #FAF8F2 !important;
    }}
    
    .metric-card {{
        background-color: white;
        border: 1px solid #E6E2D8;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    }}
    
    .fixed-header {{
        border-bottom: 2px solid #E6E2D8;
        padding-bottom: 10px;
        margin-bottom: 20px;
    }}
    
    /* Day column containers in Light Mode */
    div[data-testid="column"] {{
        background-color: #F8F5EE;
        border: 1px solid #E6E2D8;
        border-radius: 8px;
        padding: 10px;
    }}
    
    /* Popover FAB Button styling (Always Blue) */
    div[data-testid="stPopover"] {{
        position: fixed !important;
        bottom: 30px !important;
        right: 30px !important;
        z-index: 9999 !important;
    }}
    div[data-testid="stPopover"] button {{
        border-radius: 50% !important;
        width: 60px !important;
        height: 60px !important;
        font-size: 24px !important;
        box-shadow: 0 4px 14px rgba(31, 119, 180, 0.4) !important;
        background-color: #1f77b4 !important;
        color: white !important;
        border: none !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        position: relative !important;
        transition: all 0.2s ease !important;
    }}
    div[data-testid="stPopover"] button:hover {{
        transform: scale(1.08) !important;
        background-color: #175d8f !important;
        box-shadow: 0 6px 20px rgba(31, 119, 180, 0.6) !important;
    }}
    div[data-testid="stPopover"] div[data-testid="stPopoverBody"] {{
        position: absolute !important;
        bottom: 74px !important;
        right: 0px !important;
        width: 340px !important;
        max-height: 480px !important;
        overflow-y: auto !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.15) !important;
        border: 1px solid rgba(128, 128, 128, 0.2) !important;
        background-color: var(--background-color) !important;
    }}
    
    /* Inject badge dynamically */
    {badge_css}
    
    /* ========================================================================= */
    /* Dark Mode Configurations (prefers-color-scheme: dark) */
    /* ========================================================================= */
    @media (prefers-color-scheme: dark) {{
        :root {{
            --background-color: #121212 !important;
            --secondary-background-color: #1e1e1e !important;
            --text-color: #f3f4f6 !important;
        }}
        
        .stApp {{
            background-color: #121212 !important;
            color: #f3f4f6 !important;
        }}
        
        /* Dark Sidebar override */
        section[data-testid="stSidebar"] {{
            background-color: #1e1e1e !important;
            color: #f3f4f6 !important;
        }}
        section[data-testid="stSidebar"] * {{
            color: #f3f4f6 !important;
        }}
        section[data-testid="stSidebar"] hr {{
            border-color: #374151 !important;
        }}
        
        /* Day column containers in Dark Mode (slightly darker than bg) */
        div[data-testid="column"] {{
            background-color: #1e1e1e !important;
            border: 1px solid #2d2d2d !important;
        }}
        
        /* Task card boxes and Metric cards in Dark Mode */
        div[data-testid="stVerticalBlockBorderWrapper"] {{
            background-color: #181818 !important;
            border: 1px solid #2d2d2d !important;
        }}
        
        .metric-card {{
            background-color: #181818 !important;
            border: 1px solid #2d2d2d !important;
            color: #f3f4f6 !important;
        }}
        
        /* High contrast text colors for headings and paragraphs */
        h1, h2, h3, h4, h5, h6, p, span, label, li {{
            color: #f3f4f6 !important;
        }}
        div[data-testid="stMarkdownContainer"] p {{
            color: #f3f4f6 !important;
        }}
    }}
    
    /* Legend color items */
    .legend-box {{
        display: inline-block;
        width: 12px;
        height: 12px;
        border-radius: 3px;
        margin-right: 6px;
    }}
    
    /* Clean table formatting */
    .report-card {{
        border: 1px solid rgba(128, 128, 128, 0.15);
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 12px;
    }}
    </style>
""", unsafe_allow_html=True)

# ----------------- SESSION STATE INIT -----------------
if "nav_selection" not in st.session_state:
    st.session_state["nav_selection"] = "🏠 Inicio"  # Default active page is Inicio on startup

if "selected_date" not in st.session_state:
    st.session_state["selected_date"] = datetime.date.today()

# ----------------- SIDEBAR -----------------
with st.sidebar:
    st.title("💼 CUADROpz")
    st.subheader("Control de Producción")
    st.divider()
    
    selected_user = st.selectbox(
        "Usuario Activo:",
        all_users,
        index=all_users.index(st.session_state["current_user"])
    )
    st.session_state["current_user"] = selected_user
    
    # Expandable section for User Management
    with st.expander("👥 Gestión de Usuarios"):
        st.markdown("**Agregar Usuario:**")
        new_name = st.text_input("Nombre del colaborador:", key="new_user_name_input")
        if st.button("Agregar Nuevo Usuario", use_container_width=True):
            if new_name.strip():
                if add_user(new_name.strip()):
                    st.success(f"Usuario '{new_name}' creado.")
                    st.session_state["current_user"] = new_name.strip()
                    st.rerun()
                else:
                    st.error("El usuario ya existe.")
            else:
                st.error("El nombre no puede estar vacío.")
                
        st.divider()
        
        st.markdown("**Eliminar Usuario:**")
        user_to_del = st.selectbox("Seleccione usuario:", all_users, key="delete_user_selectbox")
        if st.button("Eliminar Usuario", type="primary", use_container_width=True):
            if user_to_del == selected_user:
                st.warning("No puedes eliminarte a ti mismo.")
            else:
                if delete_user(user_to_del):
                    st.success(f"Usuario '{user_to_del}' eliminado.")
                    all_users = get_users()
                    if all_users:
                        st.session_state["current_user"] = all_users[0]
                    st.rerun()
                else:
                    st.error("No se pudo eliminar el usuario.")
                    
    # AI Configuration Expandable Section
    with st.expander("Configuración de IA", icon=":material/settings:", expanded=False):
        st.text_input(
            "Clave API de OpenAI",
            type="password",
            key="openai_api_key",
            placeholder="sk-...",
        )
        st.caption("Opcional: solo necesario para regenerar sugerencias en Informes")
                    
    st.divider()
    
    # Navigation Links
    nav_options = [
        "🏠 Inicio",
        "📋 Pizarra",
        "📊 Informes",
        "📥 Exportar",
        "📅 Calendario"
    ]
    
    # Sync navigation buttons
    for opt in nav_options:
        btn_type = "primary" if st.session_state["nav_selection"] == opt else "secondary"
        if st.button(opt, type=btn_type, use_container_width=True):
            st.session_state["nav_selection"] = opt
            st.rerun()
            
    st.divider()
    
    # Show Today's Finalization Status
    today_reports = get_reports(selected_user, today_str, today_str)
    status_msg = "✅ Día Finalizado" if today_reports else "⏳ Cierre Pendiente"
    st.markdown(f"**Estado de Hoy:**\n`{status_msg}`")

# ----------------- FIXED HEADER -----------------
st.markdown(
    f"""
    <div class="fixed-header">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <span style="font-size: 1.8rem; font-weight: 700; color: #0284c7;">📊 CUADROpz</span>
            <span style="font-size: 1.1rem; color: var(--text-color);">Usuario Activo: <b>{selected_user}</b></span>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# ----------------- DIALOG: FINALIZATION PREVIOUS DAYS -----------------
undone_days = get_undone_days_before(selected_user, today.strftime("%Y-%m-%d"))

if undone_days:
    oldest_undone = undone_days[0]
    st.warning(f"⚠️ **Atención:** No has finalizado el día de trabajo anterior (**{oldest_undone}**). Por favor ciérralo antes de continuar.")
    
    @st.dialog("Finalizar Día Anterior")
    def finalize_previous_dialog(date_str):
        st.markdown(f"### Cerrar Día: **{date_str}**")
        st.write("Las tareas no resueltas de este día se trasladarán al siguiente día hábil disponible de forma automática.")
        
        # Load tasks
        prev_tasks = get_tasks(selected_user, date_str)
        done_tasks = [t for t in prev_tasks if t['completed'] == 1]
        unresolved_tasks = [t for t in prev_tasks if t['completed'] == 0]
        
        st.write(f"- Completadas: **{len(done_tasks)}**")
        st.write(f"- Pendientes: **{len(unresolved_tasks)}**")
        
        if st.button("Finalizar e Ir a Siguiente Día", type="primary", use_container_width=True):
            with st.spinner("Procesando informe..."):
                resolved = [{'description': t['description'], 'time_info': t['time_info']} for t in done_tasks]
                unresolved = [{'description': t['description'], 'time_info': t['time_info'], 'priority': t['priority']} for t in unresolved_tasks]
                
                # Suggestions
                key = st.session_state.get("openai_api_key", "")
                suggs = generate_alternatives(key, unresolved)
                
                # DB Commit
                finalize_day(selected_user, date_str, resolved, unresolved, suggs)
                
                st.success(f"Día {date_str} cerrado con éxito.")
                st.rerun()
                
    if st.button(f"Procesar Cierre del {oldest_undone}", type="primary"):
        finalize_previous_dialog(oldest_undone)

# ----------------- PAGE: DASHBOARD (INICIO) -----------------
if st.session_state["nav_selection"] == "🏠 Inicio":
    st.subheader("🏠 Panel de Inicio / Dashboard")
    st.write("Resumen ejecutivo del día actual y avance de los últimos 7 días hábiles.")
    
    # Calculations for Card 1 (Pizarra)
    today_tasks = get_tasks_by_user_and_date(selected_user, today_str)
    total_t = len(today_tasks)
    comp_t = get_completed_tasks_count(selected_user, today_str)
    pend_t = total_t - comp_t
    progress_pct = int((comp_t / total_t) * 100) if total_t > 0 else 0

    if total_t == 0:
        pizarra_resumen = "Aún no tienes tareas para hoy"
        pizarra_avance = "Completado: 0%"
    else:
        pizarra_resumen = f"Tienes **{pend_t}** tareas pendientes de hoy de un total de **{total_t}**"
        pizarra_avance = f"Completado: **{progress_pct}%**"

    # Calculations for Card 2 (Informes)
    last_report = get_last_report(selected_user)
    if last_report:
        try:
            report_date_dt = datetime.datetime.strptime(last_report['date'], "%Y-%m-%d")
            report_date_str = report_date_dt.strftime("%d/%m/%Y")
        except Exception:
            report_date_str = last_report['date']
        informes_resumen = f"Último informe generado: **{report_date_str}**"
        unresolved_count = len(last_report.get('unresolved_tasks', []))
        informes_pendientes = f"Tareas no resueltas pendientes: **{unresolved_count}**"
    else:
        informes_resumen = "No hay informes generados"
        informes_pendientes = "Tareas no resueltas pendientes: **0**"

    # Calculations for Card 3 (Exportar)
    export_info = get_last_export_info()
    if export_info:
        try:
            export_date_dt = datetime.datetime.strptime(export_info['date'], "%Y-%m-%d")
            export_date_str = export_date_dt.strftime("%d/%m/%Y")
        except Exception:
            export_date_str = export_info['date']
        exportar_resumen = f"Última exportación: **{export_date_str}**"
        exportar_count = export_info.get('tasks_count', 0)
        exportar_tareas = f"Total de tareas exportadas en la última semana: **{exportar_count}**"
    else:
        exportar_resumen = "Última exportación: **Ninguna**"
        exportar_tareas = "Total de tareas exportadas en la última semana: **0**"

    # Row of Cards
    col_c1, col_c2, col_c3 = st.columns(3)
    
    with col_c1:
        with st.container(border=True):
            st.markdown("### 🎯 Pizarra")
            st.write(pizarra_resumen)
            st.write(pizarra_avance)
            st.write("")
            if st.button("Ir a Pizarra", key="go_to_pizarra_btn", width="stretch"):
                st.session_state["nav_selection"] = "📋 Pizarra"
                st.rerun()
                
    with col_c2:
        with st.container(border=True):
            st.markdown("### 📊 Informes")
            st.write(informes_resumen)
            st.write(informes_pendientes)
            st.write("")
            if st.button("Ir a Informes", key="go_to_informes_btn", width="stretch"):
                st.session_state["nav_selection"] = "📊 Informes"
                st.rerun()
                
    with col_c3:
        with st.container(border=True):
            st.markdown("### 📥 Exportar")
            st.write(exportar_resumen)
            st.write(exportar_tareas)
            st.write("")
            if st.button("Ir a Exportar", key="go_to_exportar_btn", width="stretch"):
                st.session_state["nav_selection"] = "📥 Exportar"
                st.rerun()
    
    # 7 Working Days Chart
    st.markdown("### 📊 Avance de los Últimos 7 Días Hábiles")
    
    # Get last 7 working days (excluding Sunday)
    working_days = []
    check_date = today
    while len(working_days) < 7:
        if check_date.weekday() != 6:  # Skip Sunday
            working_days.append(check_date.strftime("%Y-%m-%d"))
        check_date -= datetime.timedelta(days=1)
    working_days.reverse()  # Oldest to newest
    
    chart_data = []
    has_data = False
    for d_str in working_days:
        d_tasks = get_tasks(selected_user, d_str)
        comp = sum(1 for t in d_tasks if t['completed'] == 1)
        pend = sum(1 for t in d_tasks if t['completed'] == 0)
        if len(d_tasks) > 0:
            has_data = True
            
        d_obj = datetime.datetime.strptime(d_str, "%Y-%m-%d").date()
        day_lbl = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"][d_obj.weekday()]
        label = f"{day_lbl} {d_obj.strftime('%d/%m')}"
        
        chart_data.append({
            "Fecha": label,
            "Completadas": comp,
            "Pendientes": pend
        })
        
    if has_data:
        df_chart = pd.DataFrame(chart_data)
        st.bar_chart(
            df_chart,
            x="Fecha",
            y=["Completadas", "Pendientes"],
            color=["#22c55e", "#ef4444"], # Green for completed, Red for pending
            stack=True
        )
    else:
        st.info("No se encontraron registros de tareas para los últimos 7 días hábiles.")
        
    st.divider()
    
    # Pending tasks list at bottom
    st.markdown("### ⏳ Actividades Pendientes del Día")
    today_undone = [t for t in today_tasks if t['completed'] == 0]
    
    if not today_undone:
        st.success("🎉 ¡Excelente trabajo! No tienes actividades pendientes programadas para hoy.")
    else:
        for idx, t in enumerate(today_undone):
            time_str = f"🕒 ({t['time_info']})" if t['time_info'] else ""
            prio = t['priority']
            prio_color = ":red[🔴 Alta]" if prio == "Alta" else (":orange[🟠 Media]" if prio == "Media" else (":green[🟢 Baja]" if prio == "Baja" else "⚪ Normal"))
            # Dark mode friendly orange text color for pending tasks
            st.markdown(f"**{idx+1}.** <span style='color: #ffb74d; font-weight: 600;'>{t['description']}</span> {time_str} — Prioridad: {prio_color}", unsafe_allow_html=True)
            
        st.write("")
        if st.button("✏️ Completar actividades en Pizarra", type="primary"):
            st.session_state["nav_selection"] = "📋 Pizarra"
            st.rerun()


# ----------------- PAGE: PIZARRA (WEEKLY BOARD VIEW) -----------------
elif st.session_state["nav_selection"] == "📋 Pizarra":
    # Snap selected date to Monday
    sel_date = st.session_state["selected_date"]
    monday_diff = sel_date.weekday()
    monday_date = sel_date - datetime.timedelta(days=monday_diff)
    saturday_date = monday_date + datetime.timedelta(days=5)
    
    months_es = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    month_name = months_es[monday_date.month - 1]
    
    # Pizarra top navigation layout
    col_nav, col_mid, col_right = st.columns([4, 5, 4])
    with col_nav:
        col_p, col_h, col_n = st.columns([1, 1, 1])
        with col_p:
            if st.button("◀", key="prev_week_btn", help="Semana Anterior"):
                st.session_state["selected_date"] -= datetime.timedelta(days=7)
                st.rerun()
        with col_h:
            if st.button("Hoy", key="hoy_week_btn", help="Semana Actual"):
                st.session_state["selected_date"] = datetime.date.today()
                st.rerun()
        with col_n:
            if st.button("▶", key="next_week_btn", help="Siguiente Semana"):
                st.session_state["selected_date"] += datetime.timedelta(days=7)
                st.rerun()

    with col_mid:
        st.markdown(f"### **Semana del {monday_date.day} al {saturday_date.day} de {month_name}**")
        st.caption("1 PIZARRA(S) ACTIVA(S)")

    with col_right:
        col_act1, col_act2 = st.columns(2)
        with col_act1:
            if st.button("📥 Exportar semana (Excel)", key="export_week_top_btn", use_container_width=True):
                st.session_state["nav_selection"] = "📥 Exportar"
                st.rerun()
        with col_act2:
            @st.dialog("Agregar Actividad Rápida")
            def quick_add_dialog():
                target_w_day = st.selectbox("Día de la semana:", ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"])
                w_idx_sel = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"].index(target_w_day)
                target_day_str = (monday_date + datetime.timedelta(days=w_idx_sel)).strftime("%Y-%m-%d")
                
                desc = st.text_input("Descripción:")
                time_info = st.text_input("Horario (Opcional):", placeholder="Ej: 9:00 AM A 10:00 AM")
                prio = st.selectbox("Prioridad:", ["Normal", "Baja", "Media", "Alta"])
                
                if st.button("Agregar a la Pizarra", type="primary", use_container_width=True):
                    if desc.strip():
                        add_task(selected_user, target_day_str, desc.strip(), time_info.strip(), prio)
                        st.success("Actividad agregada con éxito.")
                        st.rerun()
                    else:
                        st.error("La descripción es obligatoria.")
                        
            if st.button("+ Nueva pizarra", key="new_board_top_btn", use_container_width=True):
                quick_add_dialog()
                
    st.write("")
    
    # 2. Main Weekly Board User Card
    with st.container(border=True):
        col_u_hdr, col_u_info = st.columns([8, 4])
        with col_u_hdr:
            st.markdown(f"### 🔽 🟡 {selected_user}")
        with col_u_info:
            st.markdown("<p style='text-align: right; margin-top: 10px; color: #94a3b8; font-weight: 500;'>9 objetivos base / día</p>", unsafe_allow_html=True)
            
        st.divider()
        
        # 6 Column day grids
        day_cols = st.columns(6)
        day_names = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"]
        
        for i in range(6):
            date_obj = monday_date + datetime.timedelta(days=i)
            date_str = date_obj.strftime("%Y-%m-%d")
            
            with day_cols[i]:
                col_dh, col_da = st.columns([3, 1])
                with col_dh:
                    st.markdown(f"##### **{day_names[i]}**\n<span style='color: gray;'>{date_obj.strftime('%d/%m')}</span>", unsafe_allow_html=True)
                with col_da:
                    @st.dialog(f"Agregar Actividad ({day_names[i]})")
                    def add_day_task_dialog(target_date_str):
                        add_desc = st.text_input("Descripción:")
                        add_time = st.text_input("Horario (Opcional):", placeholder="Ej: 9:00 AM A 10:00 AM")
                        add_prio = st.selectbox("Prioridad:", ["Normal", "Baja", "Media", "Alta"])
                        
                        if st.button("Agregar a Pizarra", type="primary", use_container_width=True):
                            if add_desc.strip():
                                add_task(selected_user, target_date_str, add_desc.strip(), add_time.strip(), add_prio)
                                st.success("Actividad agendada.")
                                st.rerun()
                            else:
                                st.error("La descripción es obligatoria.")
                                
                    if st.button("➕", key=f"add_btn_day_{i}"):
                        add_day_task_dialog(date_str)
                        
                st.divider()
                
                day_tasks = get_tasks(selected_user, date_str)
                reports_exist = get_reports(selected_user, date_str, date_str)
                day_is_finalized = len(reports_exist) > 0
                
                total_slots = max(9, len(day_tasks))
                
                for slot_idx in range(total_slots):
                    if slot_idx < len(day_tasks):
                        task = day_tasks[slot_idx]
                        t_id = task['id']
                        desc = task['description']
                        time_info = task['time_info']
                        completed = task['completed']
                        prio = task['priority']
                        carried = task['carried_over_from']
                        
                        with st.container(border=True):
                            c_chk, c_txt = st.columns([1, 4])
                            with c_chk:
                                chk_val = st.checkbox(
                                    "Ok",
                                    value=bool(completed),
                                    key=f"col_chk_{t_id}",
                                    disabled=day_is_finalized,
                                    label_visibility="collapsed"
                                )
                                if chk_val != bool(completed) and not day_is_finalized:
                                    update_task_completion(t_id, 1 if chk_val else 0)
                                    st.rerun()
                            with c_txt:
                                # High contrast completed (green-crossed) / pending (orange) color overrides
                                if completed:
                                    st.markdown(f"<span style='color: #4caf50; text-decoration: line-through; font-size: 0.9rem; font-weight: 500;'>{desc}</span>", unsafe_allow_html=True)
                                else:
                                    carry_lbl = " 🔄" if carried else ""
                                    st.markdown(f"<span style='color: #ffb74d; font-size: 0.9rem; font-weight: 600;'>{desc}{carry_lbl}</span>", unsafe_allow_html=True)
                                    
                            c_time, c_act = st.columns([2, 1])
                            with c_time:
                                time_lbl = time_info or "Sin horario"
                                st.caption(f"🕒 {time_lbl}")
                            with c_act:
                                if not day_is_finalized:
                                    col_e, col_d = st.columns(2)
                                    with col_e:
                                        @st.dialog("Editar Tarea")
                                        def edit_task_dialog(task_obj):
                                            edit_desc = st.text_input("Descripción:", value=task_obj['description'])
                                            edit_time = st.text_input("Horario:", value=task_obj['time_info'] or "")
                                            edit_prio = st.selectbox("Prioridad:", ["Normal", "Baja", "Media", "Alta"], index=["Normal", "Baja", "Media", "Alta"].index(task_obj['priority']))
                                            
                                            if st.button("Guardar Cambios", type="primary", use_container_width=True):
                                                if edit_desc.strip():
                                                    update_task_details(task_obj['id'], edit_desc.strip(), edit_time.strip())
                                                    update_task_priority(task_obj['id'], edit_prio)
                                                    st.success("Actividad modificada.")
                                                    st.rerun()
                                                else:
                                                    st.error("La descripción no puede estar vacía.")
                                        if st.button("✏️", key=f"e_b_{t_id}", help="Editar"):
                                            edit_task_dialog(task)
                                    with col_d:
                                        if st.button("🗑️", key=f"d_b_{t_id}", help="Eliminar"):
                                            delete_task(t_id)
                                            st.rerun()
                    else:
                        with st.container(border=True):
                            st.markdown(f"<p style='color: #94a3b8; font-style: italic; font-size: 0.85rem; margin: 0;'>{slot_idx+1}. Objetivo sin definir...</p>", unsafe_allow_html=True)
                            st.write("")
                            
                st.write("")
                
                # Cerrar día button under each column day
                btn_close_disabled = day_is_finalized or not day_tasks
                
                @st.dialog(f"Cerrar {day_names[i]}")
                def close_day_column_dialog(target_date_str, day_name_lbl):
                    st.markdown(f"### Cerrar Día: **{day_name_lbl} {target_date_str}**")
                    st.write("Las tareas pendientes se registrarán en el informe diario de IA y se trasladarán automáticamente a tu agenda de mañana.")
                    
                    day_tasks_to_close = get_tasks(selected_user, target_date_str)
                    resolved_t = [t for t in day_tasks_to_close if t['completed'] == 1]
                    unresolved_t = [t for t in day_tasks_to_close if t['completed'] == 0]
                    
                    st.write(f"- Cumplidas: **{len(resolved_t)}**")
                    st.write(f"- Pendientes: **{len(unresolved_t)}**")
                    
                    if st.button("Confirmar Cierre", type="primary", use_container_width=True):
                        with st.spinner("Guardando informe..."):
                            resolved_data = [{'description': t['description'], 'time_info': t['time_info']} for t in resolved_t]
                            unresolved_data = [{'description': t['description'], 'time_info': t['time_info'], 'priority': t['priority']} for t in unresolved_t]
                            
                            key = st.session_state.get("openai_api_key", "")
                            suggs = generate_alternatives(key, unresolved_data)
                            
                            finalize_day(selected_user, target_date_str, resolved_data, unresolved_data, suggs)
                            
                            st.success("Día cerrado de forma exitosa.")
                            st.session_state["nav_selection"] = "📊 Informes"
                            st.rerun()
                            
                if st.button("Cerrar día →", key=f"close_col_btn_{i}", use_container_width=True, disabled=btn_close_disabled):
                    close_day_column_dialog(date_str, day_names[i])


# ----------------- PAGE: INFORMES -----------------
elif st.session_state["nav_selection"] == "📊 Informes":
    st.subheader("📊 Historial de Informes de IA")
    st.write("Consulte los reportes diarios cerrados y las alternativas de solución recomendadas por la Inteligencia Artificial.")
    
    col_inf1, col_inf2, col_inf3 = st.columns(3)
    with col_inf1:
        f_user = st.selectbox("Filtrar Usuario:", ["TODOS"] + all_users)
    with col_inf2:
        f_start = st.date_input("Fecha Inicio:", today - datetime.timedelta(days=14))
    with col_inf3:
        f_end = st.date_input("Fecha Fin:", today)
        
    selected_f_user = None if f_user == "TODOS" else f_user
    reports = get_reports(selected_f_user, f_start.strftime("%Y-%m-%d"), f_end.strftime("%Y-%m-%d"))
    
    if not reports:
        st.info("No se encontraron informes cerrados en el rango seleccionado.")
        
        check_date_str = f_end.strftime("%Y-%m-%d")
        existing_report = get_reports(selected_user, check_date_str, check_date_str)
        if not existing_report and selected_user in all_users:
            st.markdown(f"**¿Deseas cerrar el día {check_date_str} manualmente para generar un informe?**")
            
            @st.dialog("Generar Informe Manual")
            def generate_manual_dialog(date_str):
                m_tasks = get_tasks(selected_user, date_str)
                if not m_tasks:
                    st.error("No hay actividades programadas en este día para cerrar.")
                    return
                done = [t for t in m_tasks if t['completed'] == 1]
                unresolved_m = [t for t in m_tasks if t['completed'] == 0]
                
                if st.button("Generar Informe del Día", type="primary", use_container_width=True):
                    resolved_data = [{'description': t['description'], 'time_info': t['time_info']} for t in done]
                    unresolved_data = [{'description': t['description'], 'time_info': t['time_info'], 'priority': t['priority']} for t in unresolved_m]
                    key = st.session_state.get("openai_api_key", "")
                    suggs = generate_alternatives(key, unresolved_data)
                    finalize_day(selected_user, date_str, resolved_data, unresolved_data, suggs)
                    st.success("Informe generado.")
                    st.rerun()
                    
            if st.button(f"Cerrar Día {check_date_str} y Generar"):
                generate_manual_dialog(check_date_str)
    else:
        for rep in reports:
            r_user = rep['user_name']
            r_date = rep['date']
            resolved = rep['resolved_tasks']
            unresolved = rep['unresolved_tasks']
            alts = rep['alternatives_of_solution']
            
            # Format report section as "ALTERNATIVAS DE SOLUCION" (without "IA")
            with st.expander(f"📋 Informe: {r_date} | Colaborador: {r_user} ({len(resolved)} cumplidas / {len(unresolved)} pendientes)"):
                col_l1, col_l2 = st.columns(2)
                with col_l1:
                    st.markdown(":green[**Actividades Resueltas**]")
                    if not resolved:
                        st.write("Ninguna.")
                    else:
                        for idx, t in enumerate(resolved):
                            time_str = f"🕒 ({t['time_info']})" if t.get('time_info') else ""
                            st.write(f"✓ {t['description']} {time_str}")
                            
                with col_l2:
                    st.markdown(":red[**Actividades No Resueltas**]")
                    if not unresolved:
                        st.write("Ninguna.")
                    else:
                        for idx, t in enumerate(unresolved):
                            time_str = f"🕒 ({t['time_info']})" if t.get('time_info') else ""
                            st.markdown(f"✗ {t['description']} {time_str}")
                            
                st.divider()
                st.markdown("**💡 ALTERNATIVAS DE SOLUCION:**")
                if not alts or not unresolved:
                    st.write("No se requirieron sugerencias.")
                else:
                    for idx, t in enumerate(unresolved):
                        suggestion = alts[idx] if idx < len(alts) else "Reagendar para mañana."
                        st.markdown(f"* **Actividad:** {t['description']}")
                        
                        col_sug_text, col_sug_btn = st.columns([6, 1])
                        with col_sug_text:
                            st.info(f"💡 Sugerencia: {suggestion}")
                        with col_sug_btn:
                            st.write("") # vertical spacing
                            if st.button("🔄 Regenerar", key=f"regen_{rep['id']}_{idx}", width="stretch"):
                                key = st.session_state.get("openai_api_key", "")
                                if not key or not key.strip():
                                    st.warning("No se puede regenerar sin clave API de OpenAI")
                                else:
                                    with st.spinner("Regenerando sugerencia..."):
                                        new_sug = generate_single_alternative(key, t['description'], suggestion)
                                        if new_sug:
                                            updated_alts = list(alts)
                                            while len(updated_alts) <= idx:
                                                updated_alts.append("Reagendar para mañana.")
                                            updated_alts[idx] = new_sug
                                            update_report_suggestions(rep['id'], updated_alts)
                                            st.success("Sugerencia regenerada con éxito.")
                                            st.rerun()
                                        else:
                                            st.error("No se pudo obtener respuesta de la IA. Verifique su API key.")


# ----------------- PAGE: EXPORTAR -----------------
elif st.session_state["nav_selection"] == "📥 Exportar":
    st.subheader("📥 Exportación Semanal a Excel")
    st.write("Descargue el cuadro de producción semanal en formato original Excel, respetando los colores de prioridad de la matriz.")
    
    export_day = st.date_input("Seleccione un día del periodo semanal:", today)
    monday_diff = export_day.weekday()
    monday_date = export_day - datetime.timedelta(days=monday_diff)
    
    snap_monday_str = monday_date.strftime("%Y-%m-%d")
    saturday_date = monday_date + datetime.timedelta(days=5)
    
    st.info(f"📅 **Rango de semana laboral:** Lunes {monday_date.strftime('%d/%m/%Y')} al Sábado {saturday_date.strftime('%d/%m/%Y')}")
    
    week_days = get_week_dates(snap_monday_str)
    tasks_count = 0
    for u in all_users:
        for d in week_days:
            tasks_count += len(get_tasks(u, d))
            
    st.write(f"Cantidad total de tareas en la base de datos para esta semana: **{tasks_count}**")
    
    filename = f"Cuadro_de_Produccion_{snap_monday_str}.xlsx"
    temp_path = filename
    
    if st.button("📊 Generar Documento de Excel", type="primary", use_container_width=True):
        if tasks_count == 0:
            st.warning("No hay tareas registradas para esta semana en la base de datos, el Excel se exportará en blanco.")
        with st.spinner("Generando archivo formateado..."):
            try:
                export_weekly_tasks_to_excel(snap_monday_str, temp_path)
                save_last_export_info(datetime.date.today().strftime("%Y-%m-%d"), tasks_count)
                with open(temp_path, "rb") as file:
                    st.download_button(
                        label="💾 Descargar Archivo Excel",
                        data=file,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                st.success("Documento generado satisfactoriamente.")
            except Exception as e:
                st.error(f"Error en generación: {e}")


# ----------------- PAGE: CALENDARIO -----------------
elif st.session_state["nav_selection"] == "📅 Calendario":
    st.subheader("📅 Planificador Mensual de Actividades")
    st.write("Visualización a pantalla completa del mes para revisar la distribución de actividades asignadas.")
    
    c_col1, c_col2 = st.columns(2)
    with c_col1:
        cal_year = st.selectbox("Seleccione Año:", [2026, 2027], index=0, key="full_cal_year")
    with c_col2:
        month_names = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        cal_month = st.selectbox("Seleccione Mes:", month_names, index=today.month - 1, key="full_cal_month")
        
    month_idx = month_names.index(cal_month) + 1
    pending_days_map = get_days_with_pending_tasks(selected_user)
    
    cal = pycal.Calendar(firstweekday=0)
    month_days = cal.monthdayscalendar(cal_year, month_idx)
    
    st.markdown(f"#### Distribución Mensual: **{cal_month} {cal_year}**")
    
    day_headers = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    grid_cols = st.columns(7)
    
    for idx, dh in enumerate(day_headers):
        grid_cols[idx].markdown(f"**{dh}**")
        
    for week in month_days:
        grid_cols = st.columns(7)
        for idx, day in enumerate(week):
            if day == 0:
                grid_cols[idx].write("")
            else:
                day_date_str = f"{cal_year}-{month_idx:02d}-{day:02d}"
                day_tasks = get_tasks(selected_user, day_date_str)
                total_t = len(day_tasks)
                pending_t = sum(1 for t in day_tasks if t['completed'] == 0)
                
                with grid_cols[idx].container(border=True):
                    st.markdown(f"### {day}")
                    if total_t > 0:
                        if pending_t > 0:
                            st.markdown(f":red[⏳ {pending_t} pend.]")
                        else:
                            st.markdown(":green[🟢 Todo OK]")
                    else:
                        st.markdown("*Sin tareas*")
                        
    st.divider()
    
    inspect_day = st.slider("Deslice para inspeccionar detalles del día:", 1, pycal.monthrange(cal_year, month_idx)[1], today.day)
    inspect_date_str = f"{cal_year}-{month_idx:02d}-{inspect_day:02d}"
    
    st.markdown(f"##### Actividades del **{inspect_date_str}**:")
    insp_tasks = get_tasks(selected_user, inspect_date_str)
    
    if not insp_tasks:
        st.info("No hay actividades registradas en esta fecha.")
    else:
        for idx, pt in enumerate(insp_tasks):
            completed = pt['completed']
            prio = pt['priority']
            time_info = pt['time_info'] or "Sin horario"
            
            # Use high-contrast status colors in Dark Mode
            if completed:
                st.markdown(f"✅ <span style='color: #4caf50; text-decoration: line-through;'>{pt['description']}</span> *({time_info})* - Prioridad: **{prio}**", unsafe_allow_html=True)
            else:
                st.markdown(f"⏳ <span style='color: #ffb74d; font-weight: 600;'>{pt['description']}</span> *({time_info})* - Prioridad: **{prio}**", unsafe_allow_html=True)


# ----------------- FLOATING ALERTS & NOTIFICATIONS CENTER (FAB) -----------------
# Clock-based local time suggestions
current_hour = datetime.datetime.now().hour

def classify_task_time(time_info):
    if not time_info:
        return "Resto del día"
    
    time_info_clean = str(time_info).upper()
    is_morning = False
    is_afternoon = False
    
    if "AM" in time_info_clean:
        is_morning = True
    elif "PM" in time_info_clean:
        hours = re.findall(r'\b(1|2|3|4|5|13|14|15|16|17)\b', time_info_clean)
        if hours:
            is_afternoon = True
        else:
            is_afternoon = True
    else:
        hours = re.findall(r'\b(8|9|10|11)\b', time_info_clean)
        if hours:
            is_morning = True
        hours_pm = re.findall(r'\b(12|1|2|3|4|5|13|14|15|16|17)\b', time_info_clean)
        if hours_pm:
            is_afternoon = True
            
    if is_morning:
        return "Mañana"
    elif is_afternoon:
        return "Tarde"
    else:
        return "Resto del día"

# Classify tasks
morning_tasks = []
afternoon_tasks = []
other_tasks = []

for t in today_pending_tasks:
    cat = classify_task_time(t['time_info'])
    if cat == "Mañana":
        morning_tasks.append(t)
    elif cat == "Tarde":
        afternoon_tasks.append(t)
    else:
        other_tasks.append(t)

# Suggested advice messages depending on time range
suggestion_msg = ""
if overdue_pending_tasks:
    suggestion_msg = f"Tienes {len(overdue_pending_tasks)} tareas atrasadas de ayer. Finaliza el día anterior para reorganizarlas."
elif current_hour < 12:
    if today_pending_tasks:
        suggestion_msg = f"Buenos días, tienes {len(today_pending_tasks)} tareas para esta mañana. ¡Empieza con la más urgente!"
    else:
        suggestion_msg = "¡Buenos días! No tienes tareas pendientes programadas para hoy."
elif 12 <= current_hour < 18:
    if today_pending_tasks:
        hours_left = 18 - current_hour
        suggestion_msg = f"Faltan {hours_left} horas para terminar el día. Prioriza las tareas más críticas."
    else:
        suggestion_msg = "¡Buenas tardes! Estás al día con tus tareas de hoy."
else:
    suggestion_msg = "¡Tu horario de trabajo principal para hoy ha culminado!"

# FAB label displaying a notification bell (badge handled by CSS)
fab_label = "🔔"

st.markdown("<div id='floating-fab-anchor'></div>", unsafe_allow_html=True)
with st.popover(fab_label):
    st.markdown("### 🔔 Centro de Alertas y Notificaciones")
    st.markdown(f"**Usuario:** `{selected_user}`")
    st.divider()
    
    # 1. Tareas Atrasadas
    st.markdown("**⚠️ Tareas Atrasadas (Días Anteriores):**")
    if not overdue_pending_tasks:
        st.success("🎉 ¡Sin tareas atrasadas!")
    else:
        for pt in overdue_pending_tasks:
            time_lbl = f" ({pt['time_info']})" if pt['time_info'] else ""
            st.markdown(f"- **{pt['description']}** *(del {pt['date']})*{time_lbl}")
            
    st.divider()
    
    # 2. Tareas de Hoy Agrupadas por Horario
    st.markdown("**📅 Tareas de Hoy:**")
    
    # Mañana (8am - 12pm)
    st.markdown("⏰ **Mañana (8am - 12pm):**")
    if not morning_tasks:
        st.caption("Sin tareas programadas para esta mañana.")
    else:
        for pt in morning_tasks:
            time_lbl = f" ({pt['time_info']})" if pt['time_info'] else ""
            st.markdown(f"- {pt['description']}{time_lbl}")
            
    # Tarde (1pm - 5pm)
    st.markdown("🌤️ **Tarde (1pm - 5pm):**")
    if not afternoon_tasks:
        st.caption("Sin tareas programadas para esta tarde.")
    else:
        for pt in afternoon_tasks:
            time_lbl = f" ({pt['time_info']})" if pt['time_info'] else ""
            st.markdown(f"- {pt['description']}{time_lbl}")
            
    # Resto del día
    st.markdown("🌙 **Resto del día o sin horario:**")
    if not other_tasks:
        st.caption("Sin tareas en este rango.")
    else:
        for pt in other_tasks:
            time_lbl = f" ({pt['time_info']})" if pt['time_info'] else ""
            st.markdown(f"- {pt['description']}{time_lbl}")
            
    st.divider()
    
    # 3. Sugerencias automáticas
    st.markdown("**💡 Sugerencia del Sistema:**")
    st.info(suggestion_msg)
