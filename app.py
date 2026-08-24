import streamlit as st
import datetime
import os
import json
import calendar as pycal
import pandas as pd
import re
import plotly.express as px
from database import (
    db_init, db_import_excel, get_tasks, add_task, update_task_completion,
    update_task_priority, update_task_details, delete_task, get_last_finalized_date,
    get_undone_days_before, finalize_day, get_reports, get_days_with_pending_tasks,
    get_users, add_user, delete_user, get_overdue_pending_tasks,
    get_tasks_by_user_and_date, get_completed_tasks_count, get_last_report,
    get_last_export_info, save_last_export_info, update_report_suggestions,
    get_all_users_with_admin_status, toggle_admin_status, delete_user_by_admin,
    add_user_with_role
)
import importlib
import ai_helper
importlib.reload(ai_helper)
from ai_helper import generate_alternatives, generate_single_alternative
from excel_helper import export_weekly_tasks_to_excel, get_week_dates
from email_helper import send_email, load_email_config, save_email_config

import google.generativeai as genai
from google.genai import Client as GenAIClient
from google.genai import types

class CustomGenAIClient(GenAIClient):
    def __init__(self, *args, **kwargs):
        api_key = kwargs.get("api_key") or getattr(genai, "api_key", None)
        if not api_key:
            api_key = st.session_state.get("gemini_api_key", "")
        if not api_key:
            api_key = os.environ.get("GEMINI_API_KEY", "")
        kwargs["api_key"] = api_key
        super().__init__(*args, **kwargs)

genai.Client = CustomGenAIClient

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
    div:has(#alerts-popover-anchor) + div div[data-testid="stPopover"] > button::after {{
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
        border: 2px solid #64748B !important;
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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    /* System variables for Color Palette */
    :root {{
        --primary-color: #1E3A8A;
        --secondary-color: #3B82F6;
        --accent-color: #10B981;
        --warning-color: #F59E0B;
        --bg-color: #f0f4f8;
        --text-color: #1E293B;
        --card-bg: #FFFFFF;
        --card-border: #E2E8F0;
        --column-bg: #F5F0E8;
        --column-border: #E2E8F0;
        --sidebar-bg: #FFFFFF;
        --sidebar-text: #1E293B;
        --sidebar-shadow: 4px 0 20px rgba(0, 0, 0, 0.05);
    }}
    
    /* Dark Theme Overrides */
    @media (prefers-color-scheme: dark) {{
        :root {{
            --bg-color: #0F172A;
            --text-color: #E2E8F0;
            --card-bg: #1E293B;
            --card-border: #334155;
            --column-bg: #1E293B;
            --column-border: #334155;
            --sidebar-bg: #1E293B;
            --sidebar-text: #E2E8F0;
            --sidebar-shadow: none;
        }}
    }}
    
    [data-theme="dark"], [data-theme="dark"] .stApp, .stApp[data-theme="dark"], [data-testid="stAppViewContainer"][data-theme="dark"] {{
        --bg-color: #0F172A;
        --text-color: #E2E8F0;
        --card-bg: #1E293B;
        --card-border: #334155;
        --column-bg: #1E293B;
        --column-border: #334155;
        --sidebar-bg: #1E293B;
        --sidebar-text: #E2E8F0;
        --sidebar-shadow: none;
    }}
    
    .stApp {{
        background-color: var(--bg-color) !important;
        color: var(--text-color) !important;
        line-height: 1.6 !important;
    }}
    
    .stApp {{
        font-family: 'Inter', sans-serif !important;
    }}
    
    /* Main title gradient header */
    .gradient-text {{
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%) !important;
        -webkit-background-clip: text !important;
        -webkit-text-fill-color: transparent !important;
        background-clip: text !important;
        font-weight: 800 !important;
    }}
    
    .fixed-header {{
        border-bottom: 1px solid var(--card-border) !important;
        padding-bottom: 16px !important;
        margin-bottom: 24px !important;
    }}
    
    /* Cards styling as requested */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: var(--card-bg) !important;
        border: 1px solid var(--card-border) !important;
        box-shadow: 0 8px 30px rgba(0,0,0,0.08) !important;
        border-radius: 16px !important;
        padding: 24px !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease !important;
        animation: fadeInUp 0.5s ease-out both !important;
    }}
    
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {{
        transform: translateY(-4px) !important;
        box-shadow: 0 12px 40px rgba(0,0,0,0.12) !important;
        border-color: var(--secondary-color) !important;
    }}
    
    /* Fondo blanco para las tarjetas del dashboard */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: #FFFFFF !important;
    }}
    
    /* Remove card inner margin padding override to look neat */
    div[data-testid="stVerticalBlockBorderWrapper"] > div {{
        padding: 0 !important;
    }}
    
    /* Fondo arena para las columnas de la pizarra */
    div[data-testid="column"] {{
        background-color: #F5F0E8 !important;
        border-radius: 12px !important;
        padding: 16px !important;
        transition: background-color 0.3s ease, border-color 0.3s ease !important;
    }}
    
    /* Si el selector anterior no funciona, prueba con este más agresivo */
    div[data-testid="stVerticalBlock"] > div[data-testid="column"] {{
        background-color: #F5F0E8 !important;
        border-radius: 12px !important;
        padding: 16px !important;
    }}
    
    /* Para modo oscuro, también aplicamos el color arena */
    [data-theme="dark"] div[data-testid="column"] {{
        background-color: #F5F0E8 !important;
        border-color: #F5F0E8 !important;
    }}
    
    /* Button styles */
    button[data-testid="baseButton-primary"], button[kind="primary"] {{
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 12px rgba(30, 58, 138, 0.25) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease, opacity 0.2s ease !important;
    }}
    button[data-testid="baseButton-primary"]:hover, button[kind="primary"]:hover {{
        transform: scale(1.02) !important;
        box-shadow: 0 6px 16px rgba(30, 58, 138, 0.4) !important;
        opacity: 0.95 !important;
    }}
    
    button[data-testid="baseButton-secondary"], button[kind="secondary"] {{
        background: transparent !important;
        color: var(--text-color) !important;
        border: 1px solid var(--card-border) !important;
        border-radius: 12px !important;
        padding: 10px 24px !important;
        font-weight: 500 !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease !important;
    }}
    button[data-testid="baseButton-secondary"]:hover, button[kind="secondary"]:hover {{
        transform: scale(1.02) !important;
        border-color: var(--secondary-color) !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05) !important;
    }}
    
    /* Sidebar customization */
    section[data-testid="stSidebar"] {{
        background-color: var(--sidebar-bg) !important;
        box-shadow: var(--sidebar-shadow) !important;
        border-right: 1px solid var(--card-border) !important;
    }}
    
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] li {{
        color: var(--sidebar-text) !important;
    }}
    
    section[data-testid="stSidebar"] hr {{
        border-color: var(--card-border) !important;
    }}
    
    /* Compactar barra lateral */
    .stSidebar {{
        padding: 16px 12px !important;
    }}
    .stSidebar .stSelectbox,
    .stSidebar .stTextInput,
    .stSidebar .stButton,
    .stSidebar .stExpander,
    .stSidebar .stMarkdown {{
        margin-bottom: 8px !important;
        line-height: 1.6 !important;
    }}
    .stSidebar .stExpander {{
        margin-top: 4px !important;
    }}
    .stSidebar .stExpander .stExpanderHeader {{
        padding: 8px 0px !important;
        font-size: 0.95rem !important;
    }}
    .stSidebar .stRadio > div {{
        gap: 6px !important;
    }}
    .stSidebar .stMetric {{
        margin-bottom: 4px !important;
    }}
    .stSidebar h1, .stSidebar h2, .stSidebar h3, .stSidebar h4 {{
        margin-top: 8px !important;
        margin-bottom: 12px !important;
    }}
    
    /* Botones visuales de navegacion en la barra lateral */
    section[data-testid="stSidebar"] button[data-testid="baseButton-primary"],
    section[data-testid="stSidebar"] button[kind="primary"] {{
        display: block !important;
        width: 100% !important;
        padding: 12px 16px !important;
        margin-bottom: 8px !important;
        border-radius: 12px !important;
        background: #1E3A8A !important;
        color: white !important;
        border: 1px solid #1E3A8A !important;
        font-weight: 500 !important;
        font-size: 1rem !important;
        text-align: left !important;
        box-shadow: 0 4px 12px rgba(30, 58, 138, 0.3) !important;
        transition: all 0.2s ease !important;
        font-family: 'Inter', sans-serif !important;
    }}

    section[data-testid="stSidebar"] button[data-testid="baseButton-primary"]:hover,
    section[data-testid="stSidebar"] button[kind="primary"]:hover {{
        background: #1e40af !important;
        border-color: #1e40af !important;
        transform: translateX(4px) !important;
    }}

    section[data-testid="stSidebar"] button[data-testid="baseButton-secondary"],
    section[data-testid="stSidebar"] button[kind="secondary"] {{
        display: block !important;
        width: 100% !important;
        padding: 12px 16px !important;
        margin-bottom: 8px !important;
        border: 1px solid #e5e7eb !important;
        border-radius: 12px !important;
        background: transparent !important;
        color: #4b5563 !important;
        font-weight: 500 !important;
        font-size: 1rem !important;
        text-align: left !important;
        transition: all 0.2s ease !important;
        font-family: 'Inter', sans-serif !important;
    }}

    section[data-testid="stSidebar"] button[data-testid="baseButton-secondary"]:hover,
    section[data-testid="stSidebar"] button[kind="secondary"]:hover {{
        background: #f3f4f6 !important;
        border-color: #9ca3af !important;
        color: #1e293b !important;
        transform: translateX(4px) !important;
    }}

    /* Modo Oscuro para botones secundarios inactivos */
    [data-theme="dark"] section[data-testid="stSidebar"] button[data-testid="baseButton-secondary"],
    [data-theme="dark"] section[data-testid="stSidebar"] button[kind="secondary"] {{
        border-color: #334155 !important;
        color: #94a3b8 !important;
    }}
    [data-theme="dark"] section[data-testid="stSidebar"] button[data-testid="baseButton-secondary"]:hover,
    [data-theme="dark"] section[data-testid="stSidebar"] button[kind="secondary"]:hover {{
        background: #1e293b !important;
        border-color: #475569 !important;
        color: #f1f5f9 !important;
    }}
    
    /* Rounded active user selectbox */
    div[data-testid="stSelectbox"] > div {{
        border-radius: 12px !important;
    }}
    div[data-testid="stSelectbox"] div[role="combobox"] {{
        border-radius: 12px !important;
        border: 1px solid var(--card-border) !important;
        background-color: var(--card-bg) !important;
    }}
    
    /* Botón flotante del asistente - siempre visible */
    div[data-testid="stPopover"], div[data-testid="stPopover"] > div {{
        position: fixed !important;
        bottom: 30px !important;
        right: 30px !important;
        width: 60px !important;
        height: 60px !important;
        z-index: 999999 !important;
        padding: 0 !important;
        margin: 0 !important;
    }}

    div[data-testid="stPopover"] button, div[data-testid="stPopover"] > button {{
        border-radius: 50% !important;
        width: 60px !important;
        height: 60px !important;
        background-color: #ffffff !important;
        color: #1E3A8A !important;
        border: 2px solid #1E3A8A !important;
        box-shadow: 0 4px 14px rgba(0,0,0,0.15) !important;
        font-size: 28px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        padding: 0 !important;
        transition: all 0.3s ease !important;
    }}

    div[data-testid="stPopover"] button:hover, div[data-testid="stPopover"] > button:hover {{
        transform: scale(1.08) !important;
        box-shadow: 0 6px 20px rgba(30, 58, 138, 0.3) !important;
    }}

    /* Contenido del popover */
    div[data-testid="stPopover"] div[data-testid="stPopoverBody"] {{
        position: absolute !important;
        bottom: 74px !important;
        right: 0px !important;
        width: 380px !important;
        max-height: 500px !important;
        overflow-y: auto !important;
        border-radius: 16px !important;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2) !important;
        border: 1px solid #e5e7eb !important;
        background-color: var(--card-bg, #ffffff) !important;
    }}

    /* Estilo Alerts Popover (🔔) a la izquierda del chatbot */
    div:has(#alerts-popover-anchor) + div div[data-testid="stPopover"],
    div:has(#alerts-popover-anchor) + div div[data-testid="stPopover"] > div {{
        right: 105px !important;
    }}
    div:has(#alerts-popover-anchor) + div div[data-testid="stPopover"] button {{
        border-radius: 50% !important;
        width: 60px !important;
        height: 60px !important;
        font-size: 24px !important;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.2) !important;
        background-color: #64748B !important;
        color: white !important;
        border: none !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        position: relative !important;
        transition: all 0.3s ease !important;
    }}
    div:has(#alerts-popover-anchor) + div div[data-testid="stPopover"] button:hover {{
        transform: scale(1.08) !important;
        background-color: #475569 !important;
        box-shadow: 0 6px 20px rgba(71, 85, 105, 0.4) !important;
    }}
    
    {badge_css}
    
    /* Custom subheadings and secondary labels */
    h2, h3, .subtitle {{
        color: #64748B !important;
        font-weight: 500 !important;
    }}
    [data-theme="dark"] h2, [data-theme="dark"] h3 {{
        color: #94A3B8 !important;
    }}
    
    /* Animación fadeInUp */
    @keyframes fadeInUp {{
        from {{
            opacity: 0;
            transform: translateY(12px);
        }}
        to {{
            opacity: 1;
            transform: translateY(0);
        }}
    }}
    
    /* Legend items */
    .legend-box {{
        display: inline-block;
        width: 12px;
        height: 12px;
        border-radius: 3px;
        margin-right: 6px;
    }}
    
    /* Report Card styling */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.report-card-anchor) {{
        background-color: var(--card-bg) !important;
        border-radius: 12px !important;
        padding: 20px 24px !important;
        margin-bottom: 20px !important;
        border-left: 4px solid #3B82F6 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.06) !important;
        transition: all 0.2s ease !important;
        animation: fadeInUp 0.5s ease-out both !important;
    }}
    
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.report-card-anchor):hover {{
        box-shadow: 0 8px 24px rgba(0,0,0,0.12) !important;
        transform: translateY(-2px) !important;
    }}
    
    /* Metadata formatting inside report cards */
    .report-date {{
        font-weight: 700 !important;
        color: var(--primary-color) !important;
        font-size: 1.15rem !important;
        margin-bottom: 4px !important;
    }}
    
    .report-user {{
        font-weight: 500 !important;
        color: #64748B !important;
        font-size: 1rem !important;
        margin-bottom: 8px !important;
    }}
    
    [data-theme="dark"] .report-user {{
        color: #94A3B8 !important;
    }}
    
    .report-stats {{
        font-size: 1rem !important;
        font-weight: 600 !important;
        margin-top: 8px !important;
        margin-bottom: 8px !important;
    }}
    
    .report-completed {{
        color: #10B981 !important;
    }}
    
    .report-pending {{
        color: #F59E0B !important;
    }}

    /* Anchors for Dashboard Cards Left Border Colors */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.pizarra-card-anchor) {{
        border-left: 4px solid #10b981 !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.informes-card-anchor) {{
        border-left: 4px solid #3b82f6 !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.exportar-card-anchor) {{
        border-left: 4px solid #f59e0b !important;
    }}
    
    /* Metric layout adjustments inside dashboard cards */
    .card-metric-container {{
        margin: 16px 0;
        display: flex;
        align-items: baseline;
        gap: 8px;
    }}
    .card-metric-val {{
        font-size: 2.2rem;
        font-weight: 800;
        line-height: 1;
    }}
    .card-metric-lbl {{
        font-size: 0.95rem;
        color: #64748B;
        font-weight: 500;
    }}
    [data-theme="dark"] .card-metric-lbl {{
        color: #94A3B8;
    }}
    
    /* Metrics coloring for completed vs pending */
    div[data-testid="column"]:has(.metric-pendientes) div[data-testid="stMetricValue"] {{
        color: #f59e0b !important;
    }}
    div[data-testid="column"]:has(.metric-completadas) div[data-testid="stMetricValue"] {{
        color: #10b981 !important;
    }}
    
    /* Global metric font-size increase */
    div[data-testid="stMetricValue"] {{
        font-size: 2.2rem !important;
        font-weight: 800 !important;
    }}

    /* Reset styling for nested columns inside pizarra task containers */
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="column"] {{
        background-color: transparent !important;
        border: none !important;
        padding: 0 !important;
        box-shadow: none !important;
        border-radius: 0 !important;
    }}
    
    /* Adjust spacing for buttons inside task container columns */
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="column"] button {{
        margin-right: 8px !important;
    }}
    
    /* Compact padding of the task cards inside the pizarra board */
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlockBorderWrapper"] {{
        padding: 12px !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
    }}

    /* Estilo de tarjetas personalizadas HTML del Dashboard */
    .custom-card {{
        background-color: #FFFFFF !important;
        border-radius: 16px !important;
        padding: 24px !important;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08) !important;
        margin-bottom: 16px !important;
        border: 1px solid #E2E8F0 !important;
    }}
    .custom-card.pizarra {{
        border-left: 4px solid #10b981 !important;
    }}
    .custom-card.informes {{
        border-left: 4px solid #3b82f6 !important;
    }}
    .custom-card.exportar {{
        border-left: 4px solid #f59e0b !important;
    }}
    
    /* Estilo mejorado para los botones del dashboard */
    div[data-testid="element-container"]:has(.custom-card.pizarra) + div[data-testid="element-container"] button,
    div[data-testid="element-container"]:has(.custom-card.informes) + div[data-testid="element-container"] button,
    div[data-testid="element-container"]:has(.custom-card.exportar) + div[data-testid="element-container"] button {{
        font-weight: 700 !important;
        font-size: 1.1rem !important;
        border: 2px solid #1E3A8A !important;
        padding: 10px 24px !important;
    }}
    </style>
""", unsafe_allow_html=True)

# Load email config and start background scheduler
email_config = load_email_config()

@st.cache_resource
def start_email_scheduler_service():
    import threading
    import time
    import datetime
    
    def scheduler_loop():
        while True:
            cfg = load_email_config()
            if cfg.get("auto_send_enabled", False):
                now = datetime.datetime.now()
                # Check if it is exactly 8:00 AM (runs once per day)
                if now.hour == 8 and now.minute == 0:
                    last_sent = cfg.get("last_sent_date", "")
                    today_s = now.strftime("%Y-%m-%d")
                    if last_sent != today_s:
                        # Fetch pending tasks for all users
                        from database import get_users, get_tasks, get_overdue_pending_tasks
                        users = get_users()
                        
                        smtp_srv = cfg.get("smtp_server", "smtp.gmail.com")
                        smtp_prt = cfg.get("smtp_port", "587")
                        sender_em = cfg.get("sender_email", "")
                        sender_pw = cfg.get("sender_password", "")
                        recipient_em = cfg.get("recipient_email", "")
                        
                        if sender_em and sender_pw and recipient_em:
                            body = f"Hola,\n\nEste es el resumen automatizado de tareas pendientes de las 8:00 AM para hoy ({today_s}):\n\n"
                            has_tasks = False
                            
                            for u in users:
                                today_pending = [t for t in get_tasks(u, today_s) if t['completed'] == 0]
                                overdue_pending = get_overdue_pending_tasks(u, today_s)
                                
                                if today_pending or overdue_pending:
                                    has_tasks = True
                                    body += f"👤 Colaborador: {u}\n"
                                    
                                    if today_pending:
                                        body += "  📅 TAREAS DE HOY:\n"
                                        for idx, t in enumerate(today_pending):
                                            t_lbl = f"({t['time_info']})" if t['time_info'] else "Sin horario"
                                            body += f"    - {t['description']} - Horario: {t_lbl} | Prioridad: {t['priority']}\n"
                                    if overdue_pending:
                                        body += "  ⚠️ ATRASADAS:\n"
                                        for idx, t in enumerate(overdue_pending):
                                            t_lbl = f"({t['time_info']})" if t['time_info'] else "Sin horario"
                                            body += f"    - {t['description']} (del {t['date']}) - {t_lbl}\n"
                                    body += "\n"
                            
                            if has_tasks:
                                from email_helper import send_email
                                send_email(
                                    smtp_srv, smtp_prt, sender_em, sender_pw, 
                                    recipient_em, f"📋 Resumen Diario Automatizado - Tareas Pendientes {today_s}", 
                                    body
                                )
                            
                            # Mark as sent for today
                            cfg["last_sent_date"] = today_s
                            save_email_config(cfg)
            time.sleep(30)
            
    t = threading.Thread(target=scheduler_loop, daemon=True)
    t.start()
    return t
    
start_email_scheduler_service()

# Helper to render structured suggestions
def render_solution_suggestion(suggestion):
    if isinstance(suggestion, dict):
        prio = suggestion.get("prioridad", "Media")
        reasign = suggestion.get("reasignacion_sugerida", "")
        text = suggestion.get("alternativa_solucion", "")
        
        # Color code the priority tag
        prio_color = "red" if prio == "Alta" else ("orange" if prio == "Media" else "green")
        st.markdown(f"**Prioridad recomendada por la IA:** :{prio_color}[{prio}]")
        
        if reasign and reasign.strip() and reasign.strip().upper() != "NONE":
            st.markdown(f"👥 **Reasignación sugerida:** `{reasign}`")
            
        st.info(f"💡 Sugerencia de Solución:\n{text}")
    else:
        # Fallback for old text format
        st.info(f"💡 Sugerencia de Solución:\n{suggestion}")

# ----------------- SESSION STATE INIT -----------------
if "admin_mode" not in st.session_state:
    st.session_state["admin_mode"] = False

if "nav_selection" not in st.session_state:
    st.session_state["nav_selection"] = "🏠 Inicio"  # Default active page is Inicio on startup

if "selected_date" not in st.session_state:
    st.session_state["selected_date"] = datetime.date.today()

# ----------------- SIDEBAR -----------------
with st.sidebar:
    # 1. Logo y encabezado corporativo
    st.markdown(
        """
        <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 6px;">
            <span style="font-size: 2.2rem;">📊</span>
            <span style="font-weight: 800; font-size: 2.2rem; background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">CUADROpz</span>
        </div>
        <div style="font-size: 0.75rem; font-weight: 700; color: #64748B; letter-spacing: 1.5px; text-transform: uppercase; margin-bottom: 20px; padding-left: 2px;">
            Control de Producción
        </div>
        """,
        unsafe_allow_html=True
    )
    st.markdown("---")
    
    selected_user = st.selectbox(
        "Usuario Activo:",
        all_users,
        index=all_users.index(st.session_state.get("current_user", all_users[0])),
        key="sidebar_active_user_selectbox"
    )
    if selected_user != st.session_state.get("current_user"):
        st.session_state["current_user"] = selected_user
        st.session_state["user_name"] = selected_user
        st.rerun()
        
    st.markdown("---")
    
    # 2. NAVEGACIÓN (Primera posición)
    nav_options = [
        "🏠 Inicio",
        "📋 Pizarra",
        "📊 Informes",
        "📥 Exportar",
        "📅 Calendario"
    ]
    if st.session_state.get("admin_mode", False):
        nav_options.append("👑 Admin")
        
    st.markdown("**Navegación:**")
    current_sel = st.session_state.get("nav_selection", "🏠 Inicio")
    for opt in nav_options:
        btn_type = "primary" if current_sel == opt else "secondary"
        if st.button(opt, type=btn_type, use_container_width=True, key=f"nav_btn_{opt}"):
            st.session_state["nav_selection"] = opt
            st.rerun()
            
    # 3. ESTADO DE HOY (Segunda posición)
    st.markdown("---")
    today_tasks = get_tasks(st.session_state.get("current_user", "MARY CRUZ"), today_str)
    
    if not today_tasks:
        st.markdown("📊 **Estado de Hoy:**\n*Sin tareas programadas*")
    else:
        st.markdown("📊 **Estado de Hoy:**")
        completed_count = len([t for t in today_tasks if t['completed'] == 1])
        pending_count = len([t for t in today_tasks if t['completed'] == 0])
        
        col_stat1, col_stat2 = st.columns(2)
        with col_stat1:
            st.markdown('<div class="metric-pendientes"></div>', unsafe_allow_html=True)
            st.metric("⏳ Pendientes", pending_count)
        with col_stat2:
            st.markdown('<div class="metric-completadas"></div>', unsafe_allow_html=True)
            st.metric("✅ Completadas", completed_count)
        
    st.markdown("---")
    
    # Button to send daily summary manually
    if st.button("📧 Enviar Resumen de Hoy", key="send_today_summary_btn_manual", use_container_width=True):
        cfg_srv = email_config.get("smtp_server", "smtp.gmail.com")
        cfg_prt = email_config.get("smtp_port", "587")
        cfg_snd = email_config.get("sender_email", "")
        cfg_pwd = email_config.get("sender_password", "")
        cfg_rcp = email_config.get("recipient_email", "")
        
        if not cfg_snd or not cfg_pwd or not cfg_rcp:
            st.error("Por favor completa las credenciales de correo en el panel de administración.")
        else:
            today_pending = [t for t in get_tasks(st.session_state.get("current_user", "MARY CRUZ"), today_str) if t['completed'] == 0]
            overdue_pending = get_overdue_pending_tasks(st.session_state.get("current_user", "MARY CRUZ"), today_str)
            
            if not today_pending and not overdue_pending:
                body_text = f"Hola,\n\nNo tienes tareas pendientes programadas para hoy ({today_str}).\n\n¡Excelente día!\nCUADROpz"
            else:
                body_text = f"Hola,\n\nEste es tu resumen de tareas pendientes para hoy ({today_str}):\n\n"
                
                if today_pending:
                    body_text += "📅 TAREAS DE HOY:\n"
                    for idx, t in enumerate(today_pending):
                        time_lbl = f"({t['time_info']})" if t['time_info'] else "Sin horario"
                        body_text += f"{idx+1}. {t['description']} - Horario: {time_lbl} | Prioridad: {t['priority']}\n"
                    body_text += "\n"
                    
                if overdue_pending:
                    body_text += "⚠️ TAREAS ATRASADAS (Días Anteriores):\n"
                    for idx, t in enumerate(overdue_pending):
                        time_lbl = f"({t['time_info']})" if t['time_info'] else "Sin horario"
                        body_text += f"- {t['description']} (del {t['date']}) - {time_lbl}\n"
                    body_text += "\n"
                
                body_text += "Puedes acceder a la aplicación para actualizarlas en Streamlit Cloud.\n\nSaludos,\nEquipo CUADROpz"
                
            with st.spinner("Enviando..."):
                success, msg = send_email(
                    cfg_srv, cfg_prt, cfg_snd, cfg_pwd, cfg_rcp, 
                    f"📋 Resumen de tareas pendientes - {today_str}", 
                    body_text
                )
                if success:
                    st.success("¡Resumen enviado exitosamente!")
                else:
                    st.error(msg)
                    
    st.divider()
    
    # Admin Mode Toggle at the bottom of the sidebar
    @st.dialog("Activar Modo Admin")
    def enter_admin_mode_dialog():
        st.write("Por favor, ingrese la contraseña de administrador:")
        password_input = st.text_input("Contraseña", type="password", key="admin_password_dialog_input")
        if st.button("Confirmar", type="primary", use_container_width=True, key="confirm_admin_pwd_btn"):
            admin_pwd = "admin123"
            try:
                if "ADMIN_PASSWORD" in st.secrets:
                    admin_pwd = st.secrets["ADMIN_PASSWORD"]
            except Exception:
                pass
            
            if password_input == admin_pwd:
                st.session_state["admin_mode"] = True
                st.success("Modo Admin activado.")
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

    if not st.session_state.get("admin_mode", False):
        if st.button("👑 Modo Admin", type="secondary", use_container_width=True, key="activate_admin_mode_btn"):
            enter_admin_mode_dialog()
    else:
        st.markdown("<p style='color: #10B981; font-weight: bold; margin-bottom: 5px; text-align: center;'>🟢 Modo Admin Activo</p>", unsafe_allow_html=True)
        if st.button("🔒 Desactivar Modo Admin", type="secondary", use_container_width=True, key="deactivate_admin_mode_btn"):
            st.session_state["admin_mode"] = False
            if st.session_state.get("nav_selection") == "👑 Admin":
                st.session_state["nav_selection"] = "🏠 Inicio"
            st.rerun()

# ----------------- FIXED HEADER -----------------
st.markdown(
    f"""
    <div class="fixed-header">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <span class="gradient-text" style="font-size: 2.2rem; font-weight: 800; letter-spacing: -0.5px;">📊 CUADROpz</span>
            <span style="font-size: 1rem; color: #64748B; font-weight: 500;">Usuario Activo: <b style="color: var(--primary-color);">{selected_user}</b></span>
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
                key = st.session_state.get("gemini_api_key", "")
                suggs = generate_alternatives(key, unresolved, all_users)
                
                # DB Commit
                finalize_day(selected_user, date_str, resolved, unresolved, suggs)
                
                st.success(f"Día {date_str} cerrado con éxito.")
                st.rerun()
                
    if st.button(f"Procesar Cierre del {oldest_undone}", type="primary"):
        finalize_previous_dialog(oldest_undone)

# ----------------- ROUTE PROTECTION -----------------
if st.session_state.get("nav_selection") == "👑 Admin" and not st.session_state.get("admin_mode", False):
    st.session_state["nav_selection"] = "🏠 Inicio"
    st.rerun()

# ----------------- PAGE: DASHBOARD (INICIO) -----------------
if st.session_state["nav_selection"] == "🏠 Inicio":
    st.subheader("🏠 Panel de Inicio / Dashboard")
    st.write("Resumen ejecutivo del día actual y avance de los últimos 7 días hábiles.")
    st.markdown("<p style='color: #64748B; font-size: 1.15rem; font-style: italic; margin-top: -15px; margin-bottom: 25px;'>Tu productividad al día</p>", unsafe_allow_html=True)
    
    # Calculations for Card 1 (Pizarra)
    today_tasks = get_tasks_by_user_and_date(selected_user, today_str)
    total_t = len(today_tasks)
    comp_t = get_completed_tasks_count(selected_user, today_str)
    pend_t = total_t - comp_t
    progress_pct = int((comp_t / total_t) * 100) if total_t > 0 else 0

    # Cálculo de unresolved_count para la tarjeta de Informes
    last_report = get_last_report(selected_user)
    if last_report:
        unresolved_count = len(last_report.get('unresolved_tasks', []))
    else:
        unresolved_count = 0

    if total_t == 0:
        pizarra_resumen = "Aún no tienes tareas para hoy"
        pizarra_avance = "Completado: 0%"
    else:
        pizarra_resumen = f"Tienes **{pend_t}** tareas pendientes de hoy de un total de **{total_t}**"
        pizarra_avance = f"Completado: **{progress_pct}%**"

    # Calculations for Card 2 (Informes)
    if last_report:
        try:
            report_date_dt = datetime.datetime.strptime(last_report['date'], "%Y-%m-%d")
            report_date_str = report_date_dt.strftime("%d/%m/%Y")
        except Exception:
            report_date_str = last_report['date']
        informes_resumen = f"Último informe generado: **{report_date_str}**"
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
        st.markdown(f"""
        <div class="custom-card pizarra">
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 16px;">
                <span style="display: flex; align-items: center; justify-content: center; width: 40px; height: 40px; border-radius: 50%; background-color: #d1fae5; font-size: 1.5rem;">🎯</span>
                <span style="font-weight: 700; font-size: 1.4rem; color: #1E293B;">Pizarra</span>
            </div>
            <div class="card-metric-container">
                <div class="card-metric-val" style="color: #f59e0b;">{pend_t}</div>
                <div class="card-metric-lbl">tareas pendientes hoy</div>
            </div>
            <div class="card-metric-container">
                <div class="card-metric-val" style="color: #10b981;">{progress_pct}%</div>
                <div class="card-metric-lbl">de avance</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Ir a Pizarra", key="go_to_pizarra_btn", use_container_width=True):
            st.session_state["nav_selection"] = "📋 Pizarra"
            st.rerun()
                
    with col_c2:
        st.markdown(f"""
        <div class="custom-card informes">
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 16px;">
                <span style="display: flex; align-items: center; justify-content: center; width: 40px; height: 40px; border-radius: 50%; background-color: #dbeafe; font-size: 1.5rem;">📊</span>
                <span style="font-weight: 700; font-size: 1.4rem; color: #1E293B;">Informes</span>
            </div>
            <div class="card-metric-container">
                <div class="card-metric-val" style="color: #f59e0b;">{unresolved_count}</div>
                <div class="card-metric-lbl">pendientes de resolver</div>
            </div>
            <div style='font-size: 0.95rem; color: #64748B; margin-bottom: 12px; font-weight: 500;'>{informes_resumen}</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Ir a Informes", key="go_to_informes_btn", use_container_width=True):
            st.session_state["nav_selection"] = "📊 Informes"
            st.rerun()
                
    with col_c3:
        st.markdown(f"""
        <div class="custom-card exportar">
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 16px;">
                <span style="display: flex; align-items: center; justify-content: center; width: 40px; height: 40px; border-radius: 50%; background-color: #fef3c7; font-size: 1.5rem;">📥</span>
                <span style="font-weight: 700; font-size: 1.4rem; color: #1E293B;">Exportar</span>
            </div>
            <div class="card-metric-container">
                <div class="card-metric-val" style="color: #3b82f6;">{exportar_count}</div>
                <div class="card-metric-lbl">tareas exportadas</div>
            </div>
            <div style='font-size: 0.95rem; color: #64748B; margin-bottom: 12px; font-weight: 500;'>{exportar_resumen}</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Ir a Exportar", key="go_to_exportar_btn", use_container_width=True):
            st.session_state["nav_selection"] = "📥 Exportar"
            st.rerun()
    
    # ----------------- PLOTLY DASHBOARD CHARTS -----------------
    st.markdown("### 📊 Gráficos Estadísticos del Dashboard")
    
    # 1. Determine dynamic theme (dark vs. light) based on Streamlit context
    plotly_template = "plotly_white"
    try:
        if st.context.theme.type == "dark":
            plotly_template = "plotly_dark"
    except Exception:
        plotly_template = "plotly_white"

    # Fetch data for Weekly Evolution (Last 7 working days)
    working_days = []
    check_date = today
    while len(working_days) < 7:
        if check_date.weekday() != 6:  # Skip Sunday
            working_days.append(check_date.strftime("%Y-%m-%d"))
        check_date -= datetime.timedelta(days=1)
    working_days.reverse()
    
    chart_data = []
    has_data = False
    for d_str in working_days:
        d_tasks = get_tasks(selected_user, d_str)
        comp = sum(1 for t in d_tasks if t['completed'] == 1)
        total = len(d_tasks)
        if total > 0:
            has_data = True
            
        d_obj = datetime.datetime.strptime(d_str, "%Y-%m-%d").date()
        day_lbl = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"][d_obj.weekday()]
        label = f"{day_lbl} {d_obj.strftime('%d/%m')}"
        
        chart_data.append({
            "Fecha": label,
            "Tareas Totales": total,
            "Tareas Completadas": comp
        })
    df_chart = pd.DataFrame(chart_data)

    # Fetch data for Current Week (Monday to Saturday)
    monday_diff = today.weekday()
    monday_date = today - datetime.timedelta(days=monday_diff)
    week_days_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
    
    prod_data = []
    atrasadas_data = []
    for i, day_name in enumerate(week_days_es):
        date_obj = monday_date + datetime.timedelta(days=i)
        date_str = date_obj.strftime("%Y-%m-%d")
        tasks = get_tasks(selected_user, date_str)
        total = len(tasks)
        comp = sum(1 for t in tasks if t['completed'] == 1)
        incomplete = sum(1 for t in tasks if t['completed'] == 0)
        rate = (comp / total * 100) if total > 0 else 0.0
        
        prod_data.append({
            "Día": f"{day_name} ({date_obj.strftime('%d/%m')})",
            "Cumplimiento (%)": round(rate, 1),
            "Total": total,
            "Completadas": comp
        })
        atrasadas_data.append({
            "Día": f"{day_name} ({date_obj.strftime('%d/%m')})",
            "Tareas Incompletas": incomplete
        })
    df_prod = pd.DataFrame(prod_data)
    df_atrasadas = pd.DataFrame(atrasadas_data)

    # Fetch data for Today's Donut Chart
    today_comp = sum(1 for t in today_tasks if t['completed'] == 1)
    today_atrasadas = sum(1 for t in today_tasks if t['completed'] == 0 and t['carried_over_from'] is not None and t['carried_over_from'] != "")
    today_pendientes = len(today_tasks) - today_comp - today_atrasadas

    # Grid Layout: 2 Columns of 2 Charts each
    col_g1, col_g2 = st.columns(2)
    col_g3, col_g4 = st.columns(2)

    # Chart 1: Evolución Semanal (Line chart)
    with col_g1:
        if has_data:
            fig_evol = px.line(
                df_chart, 
                x="Fecha", 
                y=["Tareas Totales", "Tareas Completadas"], 
                markers=True,
                title="📈 Gráfico 1 - Evolución Semanal (Últimos 7 Días Hábiles)",
                color_discrete_map={"Tareas Totales": "#1E3A8A", "Tareas Completadas": "#10B981"},
                labels={"value": "Cantidad de Tareas", "variable": "Métrica"}
            )
            fig_evol.update_layout(
                template=plotly_template,
                hovermode="x unified",
                margin=dict(l=20, r=20, t=50, b=20),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_evol, use_container_width=True)
        else:
            st.info("No se encontraron registros de tareas para la Evolución Semanal.")

    # Chart 2: Estado Actual (Donut chart)
    with col_g2:
        if len(today_tasks) > 0:
            df_donut = pd.DataFrame({
                "Estado": ["Completadas", "Pendientes", "Atrasadas (Arrastradas)"],
                "Cantidad": [today_comp, today_pendientes, today_atrasadas]
            })
            # Filter out 0 counts to make the pie readable
            df_donut = df_donut[df_donut["Cantidad"] > 0]
            
            fig_donut = px.pie(
                df_donut, 
                values="Cantidad", 
                names="Estado", 
                hole=0.4,
                title="🍩 Gráfico 2 - Estado Actual de Hoy",
                color="Estado",
                color_discrete_map={
                    "Completadas": "#10B981",
                    "Pendientes": "#3B82F6",
                    "Atrasadas (Arrastradas)": "#F59E0B"
                }
            )
            fig_donut.update_traces(textinfo='percent+value')
            fig_donut.update_layout(
                template=plotly_template,
                margin=dict(l=20, r=20, t=50, b=20),
                legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5)
            )
            st.plotly_chart(fig_donut, use_container_width=True)
        else:
            st.info("🍩 Tareas de Hoy: Aún no hay tareas programadas para hoy.")

    # Chart 3: Productividad Diaria (Bar Chart, Green)
    with col_g3:
        fig_prod = px.bar(
            df_prod,
            x="Día",
            y="Cumplimiento (%)",
            hover_data=["Total", "Completadas"],
            text="Cumplimiento (%)",
            title="🏆 Gráfico 3 - Productividad Diaria (Semana Actual)",
            color_discrete_sequence=["#10B981"]
        )
        fig_prod.update_traces(textposition='outside', texttemplate='%{text}%')
        fig_prod.update_layout(
            template=plotly_template,
            yaxis_range=[0, 115],
            margin=dict(l=20, r=20, t=50, b=20)
        )
        st.plotly_chart(fig_prod, use_container_width=True)

    # Chart 4: Tareas Atrasadas (Bar Chart, Amber)
    with col_g4:
        fig_atrasadas = px.bar(
            df_atrasadas,
            x="Día",
            y="Tareas Incompletas",
            text="Tareas Incompletas",
            title="⏳ Gráfico 4 - Tareas Atrasadas por Día (Semana Actual)",
            color_discrete_sequence=["#F59E0B"]
        )
        fig_atrasadas.update_traces(textposition='outside')
        fig_atrasadas.update_layout(
            template=plotly_template,
            yaxis=dict(dtick=1),
            margin=dict(l=20, r=20, t=50, b=20)
        )
        st.plotly_chart(fig_atrasadas, use_container_width=True)
        
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
            # Dark mode friendly warning (amber) text color for pending tasks
            st.markdown(f"**{idx+1}.** <span style='color: #F59E0B; font-weight: 600;'>{t['description']}</span> {time_str} — Prioridad: {prio_color}", unsafe_allow_html=True)
            
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
                                # Accent emerald for completed (green-crossed) / Warning amber for pending (orange) color overrides
                                if completed:
                                    st.markdown(f"<span style='color: #10B981; text-decoration: line-through; font-size: 0.9rem; font-weight: 400;'>{desc}</span>", unsafe_allow_html=True)
                                else:
                                    carry_lbl = " 🔄" if carried else ""
                                    st.markdown(f"<span style='color: #F59E0B; font-size: 0.9rem; font-weight: 400;'>{desc}{carry_lbl}</span>", unsafe_allow_html=True)
                                    
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
                            
                            key = st.session_state.get("gemini_api_key", "")
                            suggs = generate_alternatives(key, unresolved_data, all_users)
                            
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
                    key = st.session_state.get("gemini_api_key", "")
                    suggs = generate_alternatives(key, unresolved_data, all_users)
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
            
            with st.container(border=True):
                # Anchor to style the container as a report-card
                st.markdown('<div class="report-card-anchor"></div>', unsafe_allow_html=True)
                
                # Report metadata in separate lines
                st.markdown(f'<div class="report-date">📅 Fecha: {r_date}</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="report-user">👤 Colaborador: {r_user}</div>', unsafe_allow_html=True)
                st.markdown(
                    f'<div class="report-stats">'
                    f'<span class="report-completed">✓ {len(resolved)} cumplidas</span>'
                    f' &nbsp;|&nbsp; '
                    f'<span class="report-pending">✗ {len(unresolved)} pendientes</span>'
                    f'</div>',
                    unsafe_allow_html=True
                )
                
                st.divider()
                
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
                        suggestion = alts[idx] if idx < len(alts) else {"prioridad": "Media", "reasignacion_sugerida": "", "alternativa_solucion": "Reagendar para mañana."}
                        st.markdown(f"* **Actividad:** {t['description']}")
                        
                        col_sug_text, col_sug_btn = st.columns([6, 1])
                        with col_sug_text:
                            render_solution_suggestion(suggestion)
                        with col_sug_btn:
                            st.write("") # vertical spacing
                            if st.button("🔄 Regenerar", key=f"regen_{rep['id']}_{idx}", width="stretch"):
                                key = st.session_state.get("gemini_api_key", "")
                                if not key or not key.strip():
                                    st.warning("No se puede regenerar sin clave API de Gemini")
                                else:
                                    with st.spinner("Regenerando sugerencia..."):
                                        new_sug = generate_single_alternative(key, t['description'], suggestion, all_users)
                                        if new_sug:
                                            updated_alts = list(alts)
                                            while len(updated_alts) <= idx:
                                                updated_alts.append({"prioridad": "Media", "reasignacion_sugerida": "", "alternativa_solucion": "Reagendar para mañana."})
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
            
            # Use high-contrast status colors: emerald for completed, amber for pending
            if completed:
                st.markdown(f"✅ <span style='color: #10B981; text-decoration: line-through;'>{pt['description']}</span> *({time_info})* - Prioridad: **{prio}**", unsafe_allow_html=True)
            else:
                st.markdown(f"⏳ <span style='color: #F59E0B; font-weight: 600;'>{pt['description']}</span> *({time_info})* - Prioridad: **{prio}**", unsafe_allow_html=True)


# ----------------- PAGE: 👑 ADMIN -----------------
elif st.session_state["nav_selection"] == "👑 Admin":
    if not st.session_state.get("admin_mode", False):
        st.session_state["nav_selection"] = "🏠 Inicio"
        st.rerun()

    st.subheader("👑 Panel de Administración Global")
    st.write("Visualiza métricas generales, gestiona colaboradores, audita tareas y configura servicios.")

    # 1. Global statistics
    import sqlite3
    try:
        conn = sqlite3.connect("cuadropz.db")
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users")
        total_users = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM tasks")
        total_tasks = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM reports")
        total_reports = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM tasks WHERE completed = 1")
        completed_tasks = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM tasks WHERE completed = 0")
        pending_tasks = cursor.fetchone()[0]
        conn.close()
    except Exception as e:
        st.error(f"Error al cargar estadísticas: {e}")
        total_users = 0
        total_tasks = 0
        total_reports = 0
        completed_tasks = 0
        pending_tasks = 0

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1:
        with st.container(border=True):
            st.metric("👥 Colaboradores", total_users)
    with col_m2:
        with st.container(border=True):
            st.metric("📋 Tareas Totales", total_tasks)
    with col_m3:
        with st.container(border=True):
            st.metric("📊 Informes Generados", total_reports)
    with col_m4:
        with st.container(border=True):
            st.metric("✅ Tareas Completadas vs ⏳ Pendientes", f"{completed_tasks} / {pending_tasks}")

    st.divider()

    # Create Tabs for Admin Sections to keep it clean and organized
    tab_users, tab_audit, tab_config = st.tabs(["👥 Gestión de Usuarios", "🔍 Auditoría de Tareas", "⚙️ Configuración del Sistema"])

    with tab_users:
        st.markdown("### 👥 Gestión de Usuarios")
        users_data = get_all_users_with_admin_status()
        
        # Header columns
        col_th1, col_th2, col_th3, col_th4 = st.columns([3, 2, 2, 2])
        col_th1.markdown("**Nombre**")
        col_th2.markdown("**Rol**")
        col_th3.markdown("**Acción Rol**")
        col_th4.markdown("**Eliminar**")
        
        # Display users table
        for u_row in users_data:
            u_name = u_row['name']
            is_admin = u_row['is_admin']
            role_lbl = "👑 Administrador" if is_admin else "👤 Colaborador"
            
            col_u1, col_u2, col_u3, col_u4 = st.columns([3, 2, 2, 2])
            with col_u1:
                st.write(f"**{u_name}**")
            with col_u2:
                st.write(role_lbl)
            with col_u3:
                if st.button("Cambiar Rol", key=f"toggle_role_{u_name}", use_container_width=True):
                    toggle_admin_status(u_name)
                    st.success(f"Rol de {u_name} actualizado.")
                    st.rerun()
            with col_u4:
                with st.popover("🗑️ Eliminar", use_container_width=True):
                    st.warning(f"¿Seguro que deseas eliminar a {u_name}? Se borrarán todas sus tareas e informes.")
                    if st.button("Sí, eliminar", key=f"confirm_delete_{u_name}", type="primary", use_container_width=True):
                        if u_name == st.session_state.get("current_user"):
                            st.error("No puedes eliminar al usuario activo en esta sesión.")
                        else:
                            if delete_user_by_admin(u_name):
                                st.success(f"Usuario {u_name} eliminado.")
                                st.rerun()
                            else:
                                st.error("Error al eliminar usuario.")
        
        st.markdown("---")
        st.markdown("#### ➕ Agregar Nuevo Usuario")
        col_add1, col_add2, col_add3 = st.columns([4, 3, 3])
        with col_add1:
            new_u_name = st.text_input("Nombre del nuevo colaborador:", key="admin_add_user_name")
        with col_add2:
            new_u_role = st.selectbox("Rol:", ["Colaborador", "Administrador"], key="admin_add_user_role")
        with col_add3:
            st.write("") # spacing
            st.write("")
            if st.button("Agregar Usuario", type="primary", use_container_width=True, key="admin_add_user_btn"):
                if new_u_name.strip():
                    is_admin_flag = 1 if new_u_role == "Administrador" else 0
                    if add_user_with_role(new_u_name.strip(), is_admin_flag):
                        st.success(f"Usuario {new_u_name.strip()} agregado con éxito.")
                        st.rerun()
                    else:
                        st.error("Error: El usuario ya existe.")
                else:
                    st.error("El nombre no puede estar vacío.")

    with tab_audit:
        st.markdown("### 🔍 Auditoría de Tareas por Usuario")
        list_users = [u['name'] for u in get_all_users_with_admin_status()]
        selected_inspect_user = st.selectbox("Seleccionar colaborador para inspeccionar:", list_users, key="inspect_tasks_user_select")
        
        col_date1, col_date2 = st.columns(2)
        with col_date1:
            inspect_start_date = st.date_input("Fecha Inicio tareas:", today - datetime.timedelta(days=7), key="inspect_start_date")
        with col_date2:
            inspect_end_date = st.date_input("Fecha Fin tareas:", today + datetime.timedelta(days=7), key="inspect_end_date")
            
        try:
            conn = sqlite3.connect("cuadropz.db")
            cursor = conn.cursor()
            cursor.execute("""
                SELECT date, description, time_info, priority, completed, carried_over_from
                FROM tasks
                WHERE user_name = ? AND date >= ? AND date <= ?
                ORDER BY date DESC, order_num ASC
            """, (selected_inspect_user, inspect_start_date.strftime("%Y-%m-%d"), inspect_end_date.strftime("%Y-%m-%d")))
            rows = cursor.fetchall()
            conn.close()
        except Exception as e:
            st.error(f"Error al cargar tareas: {e}")
            rows = []
            
        if not rows:
            st.info(f"No hay tareas registradas para {selected_inspect_user} en el rango seleccionado.")
        else:
            for r in rows:
                t_date, desc, time_info, prio, completed, carried = r
                t_date_dt = datetime.datetime.strptime(t_date, "%Y-%m-%d").date()
                t_date_str = t_date_dt.strftime("%d/%m/%Y")
                
                completed_icon = "✅" if completed else "⏳"
                carried_lbl = " 🔄" if carried else ""
                time_lbl = f"🕒 ({time_info})" if time_info else ""
                prio_color = "red" if prio == "Alta" else ("orange" if prio == "Media" else "green")
                
                col_t1, col_t2 = st.columns([2, 8])
                with col_t1:
                    st.write(f"**{t_date_str}**")
                with col_t2:
                    st.markdown(f"{completed_icon} **{desc}**{carried_lbl} {time_lbl} — Prioridad: :{prio_color}[{prio}]")

    with tab_config:
        st.markdown("### ⚙️ Configuración de IA (Gemini)")
        st.text_input(
            "Clave API de Gemini:",
            type="password",
            key="gemini_api_key",
            placeholder="AIzaSy...",
        )
        st.caption("Necesario para regenerar sugerencias y responder al asistente virtual en el chat.")
        
        st.divider()
        st.markdown("### 📧 Configuración de Correo Electrónico (SMTP)")
        
        smtp_server = st.text_input("Servidor SMTP:", value=email_config.get("smtp_server", "smtp.gmail.com"), key="admin_smtp_server")
        smtp_port = st.text_input("Puerto SMTP:", value=email_config.get("smtp_port", "587"), key="admin_smtp_port")
        sender_email = st.text_input("Correo Emisor:", value=email_config.get("sender_email", ""), key="admin_sender_email", placeholder="tu@gmail.com")
        sender_password = st.text_input("Contraseña de Aplicación:", value=email_config.get("sender_password", ""), type="password", key="admin_sender_password", placeholder="xxxx xxxx")
        recipient_email = st.text_input("Correo Destinatario:", value=email_config.get("recipient_email", ""), key="admin_recipient_email", placeholder="destinatario@gmail.com")
        
        auto_send_enabled = st.checkbox("Activar recordatorio automático (8:00 AM)", value=email_config.get("auto_send_enabled", False), key="admin_auto_send_enabled")
        
        col_email_btn1, col_email_btn2 = st.columns(2)
        with col_email_btn1:
            if st.button("💾 Guardar Configuración", key="admin_save_email_config_btn", use_container_width=True):
                new_cfg = {
                    "smtp_server": smtp_server,
                    "smtp_port": smtp_port,
                    "sender_email": sender_email,
                    "sender_password": sender_password,
                    "recipient_email": recipient_email,
                    "auto_send_enabled": auto_send_enabled,
                    "last_sent_date": email_config.get("last_sent_date", "")
                }
                if save_email_config(new_cfg):
                    st.success("Configuración de correo guardada con éxito.")
                    email_config = new_cfg
                else:
                    st.error("Error al guardar la configuración.")
        with col_email_btn2:
            if st.button("🧪 Probar Conexión", key="admin_test_email_btn", use_container_width=True):
                if not sender_email or not sender_password or not recipient_email:
                    st.warning("Faltan datos de envío para realizar la prueba.")
                else:
                    with st.spinner("Enviando correo de prueba..."):
                        success, msg = send_email(
                            smtp_server, smtp_port, sender_email, sender_password, 
                            recipient_email, "🧪 Correo de prueba - CUADROpz Admin", 
                            "Este es un correo de prueba enviado desde la pestaña de administración de CUADROpz."
                        )
                        if success:
                            st.success("¡Correo de prueba enviado con éxito!")
                        else:
                            st.error(msg)


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

st.markdown("<div id='alerts-popover-anchor'></div>", unsafe_allow_html=True)
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

# ----------------- GLOBAL AI CHAT ASSISTANT (FAB) -----------------
# Initialize chatbot messages history in session state
if "assistant_messages" not in st.session_state:
    st.session_state["assistant_messages"] = [
        {"role": "assistant", "content": "Hola, soy tu asistente de CUADROpz. ¿En qué puedo ayudarte hoy?"}
    ]
    
# Get Gemini API key safely from state or secrets
gemini_key = st.session_state.get("gemini_api_key", "")
if not gemini_key or not gemini_key.strip():
    import os
    gemini_key = os.environ.get("GEMINI_API_KEY", "")
if not gemini_key or not gemini_key.strip():
    try:
        gemini_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        pass
        
# Gather chatbot context data
u_active = st.session_state.get("current_user", "MARY CRUZ")
today_tasks_list = get_tasks(u_active, today_str)
completed_list = [t for t in today_tasks_list if t['completed'] == 1]
pending_list = [t for t in today_tasks_list if t['completed'] == 0]
overdue_count = len(get_overdue_pending_tasks(u_active, today_str))

completed_str = "\n".join([f"- {t['description']} (Horario: {t.get('time_info') or 'No especificado'})" for t in completed_list]) if completed_list else "Ninguna"
pending_str = "\n".join([f"- {t['description']} (Horario: {t.get('time_info') or 'No especificado'})" for t in pending_list]) if pending_list else "Ninguna"

system_prompt = (
    "Eres un asistente inteligente integrado en la aplicación 'CUADROpz' (Control de Producción).\n"
    f"Usuario activo actual: {u_active}\n"
    f"Fecha de hoy: {today_str}\n\n"
    "--- CONTEXTO DE HOY ---\n"
    f"Tareas completadas hoy:\n{completed_str}\n\n"
    f"Tareas pendientes de hoy:\n{pending_str}\n\n"
    f"Tareas atrasadas acumuladas de días anteriores: {overdue_count}\n"
    "------------------------\n\n"
    "Responde las consultas del usuario basándote en este contexto. Sé profesional, conciso y de gran ayuda en la gestión de sus tareas."
)

st.markdown("<div id='chat-popover-anchor'></div>", unsafe_allow_html=True)
with st.popover("", icon=":material/chat:"):
    st.markdown("### 💬 Asistente Virtual")
    st.caption("Resuelve dudas sobre tus tareas de hoy")
    
    # Display history
    for msg in st.session_state["assistant_messages"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            
    # Show suggestions if history has <= 1 message (only the welcome message)
    if len(st.session_state["assistant_messages"]) <= 1:
        st.markdown("**Preguntas sugeridas:**")
        suggestions = [
            "¿Qué tareas tengo pendientes hoy?",
            "¿Cuáles son mis tareas más urgentes?",
            "¿Cuántas tareas he completado esta semana?",
            "¿Qué tareas tengo atrasadas?",
            "¿Cómo puedo priorizar mis tareas?",
            "¿Cuál es mi productividad de esta semana?",
            "¿Cómo uso la pizarra?",
            "¿Cómo genero un informe?",
            "¿Cómo exporto a Excel?"
        ]
        for sugg in suggestions:
            if st.button(sugg, key=f"sugg_btn_{sugg}", use_container_width=True):
                st.session_state["assistant_messages"].append({"role": "user", "content": sugg})
                if not gemini_key or not gemini_key.strip():
                    st.session_state["assistant_messages"].append({
                        "role": "assistant",
                        "content": "Configura la clave API en la barra lateral para usar el asistente."
                    })
                else:
                    with st.spinner("Pensando..."):
                        try:
                            genai.configure(api_key=gemini_key.strip())
                            client = genai.Client(http_options={'api_version': 'v1'})
                            
                            # Build chat history for Gemini
                            history = []
                            for h_msg in st.session_state["assistant_messages"][:-1]:
                                role_map = "user" if h_msg["role"] == "user" else "model"
                                history.append(types.Content(role=role_map, parts=[types.Part(text=h_msg["content"])]))
                                
                            chat = client.chats.create(
                                model='gemini-3.1-flash-lite',
                                history=history,
                                config=types.GenerateContentConfig(
                                    system_instruction=system_prompt,
                                    temperature=0.7
                                )
                            )
                            response = chat.send_message(sugg)
                            assistant_response = response.text
                            
                            st.session_state["assistant_messages"].append({"role": "assistant", "content": assistant_response})
                        except Exception as e:
                            st.session_state["assistant_messages"].append({
                                "role": "assistant",
                                "content": f"Error al conectar con la IA: {str(e)}"
                            })
                st.rerun()
            
    # Chat input
    if prompt := st.chat_input("Escribe tu consulta aquí...", key="global_chat_input"):
        # Display user message in chat message container
        st.session_state["assistant_messages"].append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
            
        # Call Gemini API
        if not gemini_key or not gemini_key.strip():
            with st.chat_message("assistant"):
                st.error("Configura la clave API en la barra lateral para usar el asistente.")
        else:
            with st.spinner("Pensando..."):
                try:
                    genai.configure(api_key=gemini_key.strip())
                    client = genai.Client(http_options={'api_version': 'v1'})
                    
                    # Build chat history for Gemini
                    history = []
                    for h_msg in st.session_state["assistant_messages"][:-1]:
                        role_map = "user" if h_msg["role"] == "user" else "model"
                        history.append(types.Content(role=role_map, parts=[types.Part(text=h_msg["content"])]))
                        
                    chat = client.chats.create(
                        model='gemini-3.1-flash-lite',
                        history=history,
                        config=types.GenerateContentConfig(
                            system_instruction=system_prompt,
                            temperature=0.7
                        )
                    )
                    response = chat.send_message(prompt)
                    assistant_response = response.text
                    
                    st.session_state["assistant_messages"].append({"role": "assistant", "content": assistant_response})
                    
                    # Re-run to update UI with history
                    st.rerun()
                except Exception as e:
                    with st.chat_message("assistant"):
                        st.error(f"Error al conectar con la IA: {str(e)}")
