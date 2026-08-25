from nicegui import ui, app
import datetime
import os
import json
import calendar
import pandas as pd
import plotly.express as px
from database import (
    db_init, get_tasks, add_task, update_task_completion,
    update_task_priority, update_task_details, delete_task,
    get_users, add_user, delete_user, get_last_finalized_date,
    get_undone_days_before, finalize_day, get_reports,
    get_all_users_with_admin_status, toggle_admin_status,
    delete_user_by_admin, add_user_with_role,
    get_tasks_by_user_and_date, get_completed_tasks_count,
    get_last_report, get_last_export_info, update_report_suggestions,
    save_last_export_info, get_overdue_pending_tasks, update_task_execution_status
)
from ai_helper import generate_alternatives, generate_single_alternative
from excel_helper import export_weekly_tasks_to_excel, get_week_dates
from email_helper import load_email_config, save_email_config, send_email

# Initialize database
db_init()

# Load email config and Gemini key at startup
email_config = load_email_config()
if email_config.get("gemini_api_key"):
    os.environ["GEMINI_API_KEY"] = email_config.get("gemini_api_key")

@ui.page('/')
def main_page():
    # Load admin_mode from persistent user storage
    admin_mode = app.storage.user.get('admin_mode', False)
    
    # Client-specific state (scopes variables to this user session/connection)
    state = {
        'current_user': None,
        'selected_date': datetime.date.today(),
        'nav_selection': '🏠 Inicio',
        'admin_mode': admin_mode,
        'informes_filter_user': 'TODOS',
        'informes_filter_start': (datetime.date.today() - datetime.timedelta(days=14)).strftime("%Y-%m-%d"),
        'informes_filter_end': datetime.date.today().strftime("%Y-%m-%d"),
        'export_date': datetime.date.today().strftime("%Y-%m-%d"),
        'calendar_year': datetime.date.today().year,
        'calendar_month': datetime.date.today().month,
        'calendar_selected_day': datetime.date.today().day,
        'assistant_messages': [
            {"role": "assistant", "content": "Hola, soy tu asistente de CUADROpz. ¿En qué puedo ayudarte hoy?"}
        ],
        'chat_visible': False,
        'chat_loading': False,
        'chat_minimized': False
    }
    
    # Load users from database
    all_users = get_users(include_admins=state['admin_mode'])
    if not all_users:
        all_users = get_users(include_admins=state['admin_mode'])
        if not all_users:
            all_users = ["CPC.SHEYLA", "CPC.HECTOR"]
            for u in all_users:
                add_user(u)
            
    state['current_user'] = all_users[0]
    
    # ----------------- UI LAYOUT STRUCTURE -----------------
    
    # Header Top Bar
    with ui.header().classes('bg-blue-900 text-white p-4 flex justify-between items-center shadow-md'):
        with ui.row().classes('items-center gap-3'):
            ui.button(on_click=lambda: left_drawer.toggle(), icon='menu').props('flat round color=white')
            ui.label('📊 CUADROpz').classes('text-xl font-extrabold tracking-tight')
            
        with ui.row().classes('items-center gap-3'):
            ui.label('Usuario Activo:').classes('text-sm text-blue-200 font-medium')
            user_header_label = ui.label(state['current_user']).classes('text-sm font-bold bg-blue-800 px-3 py-1 rounded-full')
            
    # Sidebar Drawer
    with ui.left_drawer().classes('bg-slate-50 border-r p-6') as left_drawer:
        # Brand Logo and title
        with ui.row().classes('items-center gap-3 mb-6'):
            ui.label('📊').classes('text-3xl')
            with ui.column().classes('gap-0'):
                ui.label('CUADROpz').classes('text-2xl font-black text-blue-900 leading-none')
                ui.label('Control de Producción').classes('text-[10px] uppercase font-bold text-slate-400 tracking-wider')
                
        ui.separator().classes('mb-6')
        
        # User selector
        ui.label('Usuario Activo:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider mb-2')
        user_select = ui.select(all_users, value=state['current_user'], on_change=lambda e: change_user(e.value)).classes('w-full mb-6')
        
        ui.separator().classes('mb-6')
        
        # Navigation Options list
        ui.label('Navegación:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider mb-2')
        nav_buttons = {}
        nav_container = ui.column().classes('w-full gap-2')

    # Main content container
    with ui.column().classes('w-full p-6 bg-slate-50 min-h-screen'):
        content_area = ui.column().classes('w-full')

    # ----------------- BOARD LOGIC & INTERACTIONS -----------------
    
    def reload_users_list():
        nonlocal all_users
        all_users = get_users(include_admins=state['admin_mode'])
        if not all_users:
            all_users = ["CPC.SHEYLA", "CPC.HECTOR"]
        user_select.options = all_users
        if state['current_user'] not in all_users:
            state['current_user'] = all_users[0]
            user_header_label.text = all_users[0]
            refresh_layout()
        user_select.value = state['current_user']

    def change_user(user_name):
        state['current_user'] = user_name
        user_header_label.text = user_name
        ui.notify(f"Usuario activo cambiado a: {user_name}")
        refresh_layout()
        
    def select_nav(selection_name):
        state['nav_selection'] = selection_name
        refresh_layout()
        
    def refresh_navigation_styles():
        for key_name, button_widget in nav_buttons.items():
            if state['nav_selection'] == key_name:
                button_widget.classes(replace='w-full text-left justify-start py-2 px-4 rounded-xl font-medium transition-all bg-blue-900 text-white hover:bg-blue-800 shadow-md')
            else:
                button_widget.classes(replace='w-full text-left justify-start py-2 px-4 rounded-xl font-medium transition-all text-slate-800 bg-transparent hover:bg-slate-100')
                
    def prev_week():
        state['selected_date'] -= datetime.timedelta(days=7)
        refresh_pizarra()
        
    def go_to_today():
        state['selected_date'] = datetime.date.today()
        refresh_pizarra()
        
    def next_week():
        state['selected_date'] += datetime.timedelta(days=7)
        refresh_pizarra()
        
    def toggle_task(task_id, completed_value):
        update_task_completion(task_id, 1 if completed_value else 0)
        ui.notify('Estado de actividad actualizado')
        refresh_pizarra()
        
    def save_new_task(dialog, date_str, desc, time_info, priority):
        if not desc or not desc.strip():
            ui.notify('La descripción es requerida', type='warning')
            return
        add_task(state['current_user'], date_str, desc.strip(), time_info.strip() or None, priority)
        dialog.close()
        ui.notify('Actividad agregada con éxito')
        refresh_pizarra()
        
    def save_edited_task(dialog, task_id, desc, time_info, priority):
        if not desc or not desc.strip():
            ui.notify('La descripción es requerida', type='warning')
            return
        update_task_details(task_id, desc.strip(), time_info.strip() or None)
        update_task_priority(task_id, priority)
        dialog.close()
        ui.notify('Actividad modificada con éxito')
        refresh_pizarra()
        
    def confirm_delete_task(dialog, task_id):
        delete_task(task_id)
        dialog.close()
        ui.notify('Actividad eliminada')
        refresh_pizarra()
        
    # ----------------- DIALOG BUILDERS -----------------
    
    def show_add_task_dialog(target_date_str=None):
        if not target_date_str:
            target_date_str = state['selected_date'].strftime("%Y-%m-%d")
            
        with ui.dialog() as dialog, ui.card().classes('w-96 p-6 gap-4'):
            ui.label('Agregar Actividad').classes('text-lg font-bold text-slate-800')
            desc_input = ui.input('Descripción').classes('w-full')
            time_input = ui.input('Horario (Opcional)', placeholder='Ej: 9:00 AM A 10:00 AM').classes('w-full')
            prio_select = ui.select(['Normal', 'Baja', 'Media', 'Alta'], value='Normal', label='Prioridad').classes('w-full')
            
            with ui.row().classes('w-full justify-end gap-2 mt-4'):
                ui.button('Cancelar', on_click=dialog.close).props('flat')
                ui.button('Agregar', on_click=lambda: save_new_task(dialog, target_date_str, desc_input.value, time_input.value, prio_select.value)).classes('bg-blue-600 text-white font-semibold')
        dialog.open()
        
    def show_edit_task_dialog(task):
        with ui.dialog() as dialog, ui.card().classes('w-96 p-6 gap-4'):
            ui.label('Editar Actividad').classes('text-lg font-bold text-slate-800')
            desc_input = ui.input('Descripción', value=task['description']).classes('w-full')
            time_input = ui.input('Horario', value=task['time_info'] or '').classes('w-full')
            prio_select = ui.select(['Normal', 'Baja', 'Media', 'Alta'], value=task['priority'], label='Prioridad').classes('w-full')
            
            with ui.row().classes('w-full justify-end gap-2 mt-4'):
                ui.button('Cancelar', on_click=dialog.close).props('flat')
                ui.button('Guardar', on_click=lambda: save_edited_task(dialog, task['id'], desc_input.value, time_input.value, prio_select.value)).classes('bg-blue-600 text-white font-semibold')
        dialog.open()
        
    def show_delete_dialog(task_id):
        with ui.dialog() as dialog, ui.card().classes('w-80 p-6 gap-4'):
            ui.label('Confirmar Eliminación').classes('text-lg font-bold text-red-600')
            ui.label('¿Estás seguro de que deseas eliminar esta actividad?').classes('text-sm text-slate-600')
            
            with ui.row().classes('w-full justify-end gap-2 mt-4'):
                ui.button('Cancelar', on_click=dialog.close).props('flat')
                ui.button('Eliminar', on_click=lambda: confirm_delete_task(dialog, task_id)).classes('bg-red-600 text-white font-semibold')
        dialog.open()

    def show_finalize_previous_dialog(date_str):
        prev_tasks = get_tasks(state['current_user'], date_str, include_admin_users=state['admin_mode'])
        done_tasks = [t for t in prev_tasks if t['completed'] == 1]
        unresolved_tasks = [t for t in prev_tasks if t['completed'] == 0]
        
        with ui.dialog() as dialog, ui.card().classes('w-96 p-6 gap-4'):
            ui.label(f'Cerrar Día Anterior: {date_str}').classes('text-lg font-bold text-slate-800')
            ui.label('Las tareas no resueltas de este día se trasladarán al siguiente día hábil disponible de forma automática.').classes('text-sm text-slate-600')
            
            with ui.column().classes('bg-slate-50 p-3 rounded-lg border border-slate-100 w-full mb-2'):
                ui.label(f"✓ Actividades cumplidas: {len(done_tasks)}").classes('text-xs text-green-600 font-bold')
                ui.label(f"✗ Actividades pendientes: {len(unresolved_tasks)}").classes('text-xs text-amber-600 font-bold')
                
            async def run_finalize():
                dialog.close()
                loading = ui.notification('Procesando cierre de día e informe de IA...', spinner=True, timeout=0)
                try:
                    resolved_data = [{'description': t['description'], 'time_info': t['time_info'], 'carried_over_from': t.get('carried_over_from')} for t in done_tasks]
                    unresolved_data = [{'description': t['description'], 'time_info': t['time_info'], 'priority': t['priority'], 'carried_over_from': t.get('carried_over_from')} for t in unresolved_tasks]
                    
                    gemini_key = os.environ.get("GEMINI_API_KEY", "")
                    
                    import asyncio
                    from functools import partial
                    loop = asyncio.get_event_loop()
                    suggs = await loop.run_in_executor(None, partial(generate_alternatives, gemini_key, unresolved_data, all_users))
                    
                    await loop.run_in_executor(None, partial(finalize_day, state['current_user'], date_str, resolved_data, unresolved_data, suggs))
                    
                    loading.dismiss()
                    ui.notify(f"Día {date_str} cerrado con éxito.", type='success')
                    refresh_layout()
                except Exception as e:
                    loading.dismiss()
                    ui.notify(f"Error al cerrar el día: {e}", type='negative')
            
            with ui.row().classes('w-full justify-end gap-2 mt-4'):
                ui.button('Cancelar', on_click=dialog.close).props('flat')
                ui.button('Finalizar e Ir a Siguiente Día', on_click=run_finalize).classes('bg-blue-600 text-white font-semibold')
        dialog.open()

    def show_close_day_dialog(date_str, day_name_lbl):
        day_tasks_to_close = get_tasks(state['current_user'], date_str, include_admin_users=state['admin_mode'])
        resolved_t = [t for t in day_tasks_to_close if t['completed'] == 1]
        unresolved_t = [t for t in day_tasks_to_close if t['completed'] == 0]
        
        with ui.dialog() as dialog, ui.card().classes('w-96 p-6 gap-4'):
            ui.label(f'Cerrar Día: {day_name_lbl} {date_str}').classes('text-lg font-bold text-slate-800')
            ui.label('Las tareas pendientes se registrarán en el informe diario de IA y se trasladarán automáticamente a tu agenda de mañana.').classes('text-sm text-slate-600')
            
            with ui.column().classes('bg-slate-50 p-3 rounded-lg border border-slate-100 w-full mb-2'):
                ui.label(f"✓ Actividades cumplidas: {len(resolved_t)}").classes('text-xs text-green-600 font-bold')
                ui.label(f"✗ Actividades pendientes: {len(unresolved_t)}").classes('text-xs text-amber-600 font-bold')
                
            async def run_finalize():
                dialog.close()
                loading = ui.notification('Cerrando día y compilando informe...', spinner=True, timeout=0)
                try:
                    resolved_data = [{'description': t['description'], 'time_info': t['time_info'], 'carried_over_from': t.get('carried_over_from')} for t in resolved_t]
                    unresolved_data = [{'description': t['description'], 'time_info': t['time_info'], 'priority': t['priority'], 'carried_over_from': t.get('carried_over_from')} for t in unresolved_t]
                    
                    gemini_key = os.environ.get("GEMINI_API_KEY", "")
                    
                    import asyncio
                    from functools import partial
                    loop = asyncio.get_event_loop()
                    suggs = await loop.run_in_executor(None, partial(generate_alternatives, gemini_key, unresolved_data, all_users))
                    
                    await loop.run_in_executor(None, partial(finalize_day, state['current_user'], date_str, resolved_data, unresolved_data, suggs))
                    
                    loading.dismiss()
                    ui.notify(f"Día {date_str} cerrado con éxito.", type='success')
                    refresh_layout()
                except Exception as e:
                    loading.dismiss()
                    ui.notify(f"Error al cerrar el día: {e}", type='negative')
                    
            with ui.row().classes('w-full justify-end gap-2 mt-4'):
                ui.button('Cancelar', on_click=dialog.close).props('flat')
                ui.button('Confirmar Cierre', on_click=run_finalize).classes('bg-blue-600 text-white font-semibold')
        dialog.open()

    def export_weekly_excel():
        selected_date = state['selected_date']
        monday_diff = selected_date.weekday()
        monday_date = selected_date - datetime.timedelta(days=monday_diff)
        monday_str = monday_date.strftime("%Y-%m-%d")
        week_days = [(monday_date + datetime.timedelta(days=j)).strftime("%Y-%m-%d") for j in range(6)]
        
        tasks_count = 0
        all_users_list = get_users(include_admins=state['admin_mode'])
        for u in all_users_list:
            for d in week_days:
                tasks_count += len(get_tasks(u, d, include_admin_users=state['admin_mode']))
                
        filename = f"Cuadro_de_Produccion_{monday_str}.xlsx"
        try:
            export_weekly_tasks_to_excel(monday_str, filename, include_admin_users=state['admin_mode'])
            save_last_export_info(datetime.date.today().strftime("%Y-%m-%d"), tasks_count)
            ui.download(filename)
            ui.notify('Excel generado y descargado con éxito.')
        except Exception as e:
            ui.notify(f'Error al exportar: {e}', type='negative')

    # ----------------- REFRESH WRAPPERS -----------------
    
    def refresh_pizarra():
        content_area.clear()
        with content_area:
            selected_date = state['selected_date']
            monday_diff = selected_date.weekday()
            monday_date = selected_date - datetime.timedelta(days=monday_diff)
            saturday_date = monday_date + datetime.timedelta(days=5)
            
            months_es = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
            month_name = months_es[monday_date.month - 1]
            
            undone_days = get_undone_days_before(state['current_user'], datetime.date.today().strftime("%Y-%m-%d"), include_admin_users=state['admin_mode'])
            oldest_undone = undone_days[0] if undone_days else None
            
            # Nav bar header panel
            with ui.row().classes('w-full justify-between items-center bg-white p-4 rounded-xl shadow-sm border border-slate-100 mb-6'):
                # Left side: Date navigation & Static state indicator
                with ui.row().classes('items-center gap-3'):
                    # Week selector navigation
                    ui.button('◀', on_click=prev_week).props('flat dense').classes('text-slate-600 font-bold text-lg')
                    
                    with ui.column().classes('gap-0'):
                        ui.label(f"Semana del {monday_date.day} al {saturday_date.day} de {month_name}").classes('text-base font-bold text-slate-800 leading-none')
                        ui.label("1 PIZARRA(S) ACTIVA(S)").classes('text-[10px] text-green-600 font-black tracking-wider uppercase mt-1')
                        
                    ui.button('▶', on_click=next_week).props('flat dense').classes('text-slate-600 font-bold text-lg')
                    
                    ui.button('Hoy', on_click=go_to_today).props('outline dense color=primary').classes('px-3 ml-2')
                
                # Right side action buttons
                with ui.row().classes('items-center gap-2'):
                    if oldest_undone:
                        ui.button(f"Procesar Cierre del {oldest_undone}", on_click=lambda: show_finalize_previous_dialog(oldest_undone)).classes('bg-amber-600 text-white rounded-lg px-4 py-2 font-bold shadow-sm hover:bg-amber-700 transition-colors')
                        
                    ui.button('Exportar semana (Excel)', on_click=export_weekly_excel).classes('bg-green-700 text-white rounded-lg px-4 py-2 font-semibold shadow-sm hover:bg-green-800 transition-colors')
                    ui.button('➕ Nueva pizarra', on_click=lambda: show_add_task_dialog()).classes('bg-blue-700 text-white rounded-lg px-4 py-2 font-semibold shadow-sm hover:bg-blue-800 transition-colors')
            
            # Responsive 6 column layout
            day_names = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"]
            with ui.grid().classes('w-full grid-cols-1 md:grid-cols-3 lg:grid-cols-6 gap-4'):
                for i in range(6):
                    date_obj = monday_date + datetime.timedelta(days=i)
                    date_str = date_obj.strftime("%Y-%m-%d")
                    
                    # Load tasks and finalize status
                    day_tasks = get_tasks(state['current_user'], date_str, include_admin_users=state['admin_mode'])
                    reports_exist = get_reports(state['current_user'], date_str, date_str, include_admin_users=state['admin_mode'])
                    day_is_finalized = len(reports_exist) > 0
                    
                    # Column box (sand background)
                    with ui.column().classes('bg-[#F5F0E8] p-3 rounded-xl shadow-sm border border-slate-200/60 min-h-[550px] gap-3 flex flex-col'):
                        with ui.row().classes('w-full justify-between items-center border-b border-slate-300 pb-2 mb-1'):
                            with ui.column().classes('gap-0'):
                                ui.label(day_names[i]).classes('font-black text-xs text-slate-800 tracking-wider')
                                ui.label(date_obj.strftime('%d/%m')).classes('text-[10px] text-slate-500 font-bold')
                            
                            if day_is_finalized:
                                ui.label('✅ Día cerrado').classes('text-[9px] text-green-700 font-black px-2 py-0.5 bg-green-50 border border-green-200 rounded')
                            else:
                                ui.button(icon='add', on_click=lambda d=date_str: show_add_task_dialog(d)).props('flat round dense size=sm')
                            
                        for task in day_tasks:
                            t_id = task['id']
                            desc = task['description']
                            time_info = task['time_info']
                            completed = task['completed']
                            prio = task['priority']
                            carried = task['carried_over_from']
                            
                            # Priority color border left coding
                            border_color = 'border-slate-300'
                            if prio == 'Alta':
                                border_color = 'border-red-500'
                            elif prio == 'Media':
                                border_color = 'border-amber-500'
                            elif prio == 'Baja':
                                border_color = 'border-green-500'
                                
                            # Pure white compact task card
                            with ui.card().classes(f'w-full p-3 bg-white border-l-4 {border_color} shadow-sm gap-2 hover:shadow-md transition-shadow rounded-lg'):
                                with ui.row().classes('w-full items-start no-wrap gap-2'):
                                    if day_is_finalized:
                                        checkbox = ui.checkbox(value=bool(completed)).props('dense')
                                        checkbox.disable()
                                    else:
                                        checkbox = ui.checkbox(value=bool(completed), on_change=lambda e, tid=t_id: toggle_task(tid, e.value)).props('dense')
                                    
                                    desc_style = 'text-xs text-slate-700 font-normal leading-tight w-full break-words'
                                    if completed:
                                        desc_style += ' line-through text-green-500'
                                        
                                    carry_lbl = " 🔄" if carried else ""
                                    ui.label(f"{desc}{carry_lbl}").classes(desc_style)
                                    
                                # Footer actions inside the card
                                with ui.row().classes('w-full justify-between items-center mt-1 pt-1 border-t border-slate-50'):
                                    time_lbl = f"🕒 {time_info}" if time_info else "Sin horario"
                                    ui.label(time_lbl).classes('text-[9px] text-slate-400 font-medium')
                                    
                                    with ui.row().classes('gap-1'):
                                        edit_btn = ui.button(icon='edit', on_click=lambda _, t=task: show_edit_task_dialog(t)).props('flat round dense size=xs color=primary')
                                        delete_btn = ui.button(icon='delete', on_click=lambda _, tid=t_id: show_delete_dialog(tid)).props('flat round dense size=xs color=negative')
                                        if day_is_finalized:
                                            edit_btn.disable()
                                            delete_btn.disable()
                        
                        # Spacer to push close day button to the bottom
                        ui.element('div').classes('flex-grow')
                        
                        # Cerrar día button
                        btn_close_disabled = day_is_finalized or not day_tasks
                        if day_is_finalized:
                            ui.label('✅ Día cerrado').classes('text-xs text-green-600 font-bold text-center w-full py-2 bg-green-50 rounded-lg border border-green-200 shadow-sm')
                        else:
                            btn = ui.button('Cerrar día →', on_click=lambda _, d=date_str, n=day_names[i]: show_close_day_dialog(d, n))
                            btn.classes('w-full bg-blue-600 text-white font-bold rounded-lg text-xs py-2 shadow-sm hover:bg-blue-700 transition-colors')
                            if btn_close_disabled:
                                btn.disable()
                                        
    def refresh_dashboard():
        content_area.clear()
        with content_area:
            selected_user = state['current_user']
            today_date = datetime.date.today()
            today_str = today_date.strftime("%Y-%m-%d")
            
            # Header Info
            ui.label('Panel de Inicio / Dashboard').classes('text-2xl font-bold text-slate-800')
            ui.label('Resumen ejecutivo del día actual y avance de los últimos 7 días hábiles.').classes('text-slate-500 italic text-sm mt-1 mb-6')
            
            # Metric Card 1: Pizarra
            today_tasks = get_tasks_by_user_and_date(selected_user, today_str, include_admin_users=state['admin_mode'])
            total_tasks_today = len(today_tasks)
            comp_tasks_today = get_completed_tasks_count(selected_user, today_str, include_admin_users=state['admin_mode'])
            pend_tasks_today = total_tasks_today - comp_tasks_today
            progress_pct = int((comp_tasks_today / total_tasks_today) * 100) if total_tasks_today > 0 else 0
            
            # Metric Card 2: Informes
            last_report = get_last_report(selected_user, include_admin_users=state['admin_mode'])
            if last_report:
                try:
                    report_date_dt = datetime.datetime.strptime(last_report['date'], "%Y-%m-%d")
                    report_date_str = report_date_dt.strftime("%d/%m/%Y")
                except Exception:
                    report_date_str = last_report['date']
                informes_resumen = f"Último informe generado: {report_date_str}"
                unresolved_count = len(last_report.get('unresolved_tasks', []))
                informes_pendientes = f"Tareas no resueltas pendientes: {unresolved_count}"
            else:
                informes_resumen = "No hay informes generados"
                informes_pendientes = "Tareas no resueltas pendientes: 0"
                unresolved_count = 0
                
            # Metric Card 3: Exportar
            export_info = get_last_export_info()
            if export_info:
                try:
                    export_date_dt = datetime.datetime.strptime(export_info['date'], "%Y-%m-%d")
                    export_date_str = export_date_dt.strftime("%d/%m/%Y")
                except Exception:
                    export_date_str = export_info['date']
                exportar_resumen = f"Última exportación: {export_date_str}"
                exportar_count = export_info.get('tasks_count', 0)
                exportar_tareas = f"Total de tareas exportadas en la última semana: {exportar_count}"
            else:
                exportar_resumen = "Última exportación: Ninguna"
                exportar_tareas = "Total de tareas exportadas en la última semana: 0"
                exportar_count = 0

            # Render 3 Metric Cards
            with ui.grid().classes('w-full grid-cols-1 md:grid-cols-3 gap-4 mb-8'):
                # Card 1: Pizarra
                with ui.card().classes('bg-white shadow-lg rounded-xl p-6 border-l-4 border-green-500 gap-3 justify-between hover:shadow-xl transition-shadow'):
                    with ui.column().classes('gap-1'):
                        ui.label('🎯 Pizarra').classes('text-lg font-bold text-slate-800')
                        ui.label(f"Tienes {pend_tasks_today} tareas pendientes de hoy de un total de {total_tasks_today}").classes('text-xs text-slate-500 font-medium')
                        ui.label(f"Completado: {progress_pct}%").classes('text-sm text-green-600 font-bold')
                    ui.button('Ir a Pizarra', on_click=lambda: select_nav('📋 Pizarra')).classes('bg-green-600 text-white rounded-lg font-bold text-xs py-2 w-full mt-2')

                # Card 2: Informes
                with ui.card().classes('bg-white shadow-lg rounded-xl p-6 border-l-4 border-blue-500 gap-3 justify-between hover:shadow-xl transition-shadow'):
                    with ui.column().classes('gap-1'):
                        ui.label('📊 Informes').classes('text-lg font-bold text-slate-800')
                        ui.label(informes_resumen).classes('text-xs text-slate-500 font-medium')
                        ui.label(informes_pendientes).classes('text-sm text-blue-600 font-bold')
                    ui.button('Ir a Informes', on_click=lambda: select_nav('📊 Informes')).classes('bg-blue-600 text-white rounded-lg font-bold text-xs py-2 w-full mt-2')

                # Card 3: Exportar
                with ui.card().classes('bg-white shadow-lg rounded-xl p-6 border-l-4 border-amber-500 gap-3 justify-between hover:shadow-xl transition-shadow'):
                    with ui.column().classes('gap-1'):
                        ui.label('📥 Exportar').classes('text-lg font-bold text-slate-800')
                        ui.label(exportar_resumen).classes('text-xs text-slate-500 font-medium')
                        ui.label(exportar_tareas).classes('text-sm text-amber-600 font-bold')
                    ui.button('Ir a Exportar', on_click=lambda: select_nav('📥 Exportar')).classes('bg-amber-600 text-white rounded-lg font-bold text-xs py-2 w-full mt-2')

            # Render Charts Section
            ui.label('📊 Gráficos Estadísticos del Dashboard').classes('text-xl font-bold text-slate-800 mb-4')
            
            # Fetch data for Weekly Evolution (Last 7 working days)
            working_days = []
            check_date = today_date
            while len(working_days) < 7:
                if check_date.weekday() != 6:  # Skip Sunday
                    working_days.append(check_date.strftime("%Y-%m-%d"))
                check_date -= datetime.timedelta(days=1)
            working_days.reverse()
            
            chart_data = []
            has_data = False
            for d_str in working_days:
                d_tasks = get_tasks(selected_user, d_str, include_admin_users=state['admin_mode'])
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
            monday_diff = today_date.weekday()
            monday_date = today_date - datetime.timedelta(days=monday_diff)
            week_days_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
            
            prod_data = []
            atrasadas_data = []
            for i, day_name in enumerate(week_days_es):
                date_obj = monday_date + datetime.timedelta(days=i)
                date_str = date_obj.strftime("%Y-%m-%d")
                tasks = get_tasks(selected_user, date_str, include_admin_users=state['admin_mode'])
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

            today_comp = sum(1 for t in today_tasks if t['completed'] == 1)
            today_atrasadas = sum(1 for t in today_tasks if t['completed'] == 0 and t['carried_over_from'] is not None and t['carried_over_from'] != "")
            today_pendientes = len(today_tasks) - today_comp - today_atrasadas

            # Grid of 2x2 Plotly Charts
            with ui.grid().classes('w-full grid-cols-1 md:grid-cols-2 gap-6 mb-8'):
                
                # Chart 1: Evolución Semanal
                with ui.card().classes('bg-white shadow-md rounded-xl p-4 w-full h-[400px]'):
                    if has_data:
                        fig_evol = px.line(
                            df_chart, 
                            x="Fecha", 
                            y=["Tareas Totales", "Tareas Completadas"], 
                            markers=True,
                            title="📈 Gráfico 1 - Evolución Semanal (Últimos 7 Días Hábiles)",
                            color_discrete_map={"Tareas Totales": "#1E3A8A", "Tareas Completadas": "#10B981"},
                            labels={"value": "Cantidad", "variable": "Métrica"}
                        )
                        fig_evol.update_layout(
                            hovermode="x unified",
                            margin=dict(l=20, r=20, t=50, b=20),
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                        )
                        ui.plotly(fig_evol).classes('w-full h-full')
                    else:
                        ui.label("Evolución Semanal: No se encontraron registros de tareas.").classes('text-slate-400 text-sm m-auto')

                # Chart 2: Estado Actual (Donut)
                with ui.card().classes('bg-white shadow-md rounded-xl p-4 w-full h-[400px]'):
                    if len(today_tasks) > 0:
                        df_donut = pd.DataFrame({
                            "Estado": ["Completadas", "Pendientes", "Atrasadas (Arrastradas)"],
                            "Cantidad": [today_comp, today_pendientes, today_atrasadas]
                        })
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
                            margin=dict(l=20, r=20, t=50, b=20),
                            legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5)
                        )
                        ui.plotly(fig_donut).classes('w-full h-full')
                    else:
                        ui.label("🍩 Tareas de Hoy: Aún no hay tareas programadas para hoy.").classes('text-slate-400 text-sm m-auto')

                # Chart 3: Productividad Diaria
                with ui.card().classes('bg-white shadow-md rounded-xl p-4 w-full h-[400px]'):
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
                        yaxis_range=[0, 115],
                        margin=dict(l=20, r=20, t=50, b=20)
                    )
                    ui.plotly(fig_prod).classes('w-full h-full')

                # Chart 4: Tareas Atrasadas
                with ui.card().classes('bg-white shadow-md rounded-xl p-4 w-full h-[400px]'):
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
                        yaxis=dict(dtick=1),
                        margin=dict(l=20, r=20, t=50, b=20)
                    )
                    ui.plotly(fig_atrasadas).classes('w-full h-full')

            # Render Incomplete tasks list
            ui.separator().classes('my-6')
            ui.label('⏳ Actividades Pendientes del Día').classes('text-xl font-bold text-slate-800 mb-4')
            today_undone = [t for t in today_tasks if t['completed'] == 0]
            
            if not today_undone:
                with ui.row().classes('w-full items-center gap-3 p-4 bg-green-50 rounded-xl border border-green-200 text-green-700 shadow-sm'):
                    ui.label('🎉').classes('text-2xl')
                    ui.label('¡Excelente trabajo! No tienes actividades pendientes programadas para hoy.').classes('font-bold text-sm')
            else:
                with ui.column().classes('w-full gap-2 mb-6'):
                    for idx, t in enumerate(today_undone):
                        time_str = f"🕒 ({t['time_info']})" if t['time_info'] else ""
                        prio = t['priority']
                        prio_color = "text-red-500" if prio == "Alta" else ("text-amber-500" if prio == "Media" else ("text-green-500" if prio == "Baja" else "text-slate-500"))
                        
                        with ui.row().classes('w-full items-center justify-between p-3 bg-white rounded-lg border border-slate-100 shadow-sm'):
                            with ui.row().classes('items-center gap-2'):
                                ui.label(f"{idx+1}.").classes('font-bold text-slate-400')
                                ui.label(t['description']).classes('text-slate-700 font-semibold text-xs')
                                if time_str:
                                    ui.label(time_str).classes('text-[10px] text-slate-400')
                            ui.label(f"Prioridad: {prio}").classes(f'text-xs font-bold {prio_color}')
                            
                ui.button('✏️ Completar actividades en Pizarra', on_click=lambda: select_nav('📋 Pizarra')).classes('bg-blue-900 text-white font-semibold rounded-lg px-4 py-2 mt-4 shadow-sm hover:bg-blue-800 transition-colors')

    def render_solution_suggestion_nicegui(suggestion):
        if isinstance(suggestion, dict):
            prio = suggestion.get("prioridad", "Media")
            reasign = suggestion.get("reasignacion_sugerida", "")
            text = suggestion.get("alternativa_solucion", "")
            
            prio_color = 'bg-red-100 text-red-700'
            if prio == 'Media':
                prio_color = 'bg-amber-100 text-amber-700'
            elif prio == 'Baja':
                prio_color = 'bg-green-100 text-green-700'
                
            with ui.column().classes('gap-1 w-full bg-slate-50 p-3 rounded-lg border border-slate-100 mt-2'):
                with ui.row().classes('items-center gap-2'):
                    ui.label('Prioridad Recomendada:').classes('text-xs text-slate-500 font-bold')
                    ui.label(prio).classes(f'text-[10px] font-black px-2 py-0.5 rounded-full {prio_color}')
                if reasign and reasign.strip() and reasign.strip().upper() != "NONE":
                    with ui.row().classes('items-center gap-2'):
                        ui.label('👥 Reasignación Sugerida:').classes('text-xs text-slate-500 font-bold')
                        ui.label(reasign).classes('text-xs font-bold text-slate-700 bg-blue-50 px-2 py-0.5 rounded')
                
                ui.label(f"💡 Sugerencia de Solución:\n{text}").classes('text-xs text-slate-600 font-medium whitespace-pre-line mt-1')
        else:
            with ui.column().classes('gap-1 w-full bg-slate-50 p-3 rounded-lg border border-slate-100 mt-2'):
                ui.label(f"💡 Sugerencia de Solución:\n{suggestion}").classes('text-xs text-slate-600 font-medium whitespace-pre-line')

    def update_filter_user(val):
        state['informes_filter_user'] = val
        refresh_informes()
        
    def update_filter_start(val):
        state['informes_filter_start'] = val
        refresh_informes()
        
    def update_filter_end(val):
        state['informes_filter_end'] = val
        refresh_informes()

    async def regen_sugg(report_id, idx, task_desc, previous_sugg, alts_list, user_name=None):
        gemini_key = os.environ.get("GEMINI_API_KEY", "")
        if not gemini_key or not gemini_key.strip():
            ui.notify('No se puede regenerar sin clave API de Gemini', type='warning')
            return
            
        loading = ui.notification('Regenerando sugerencia con IA...', spinner=True, timeout=0)
        try:
            import asyncio
            from functools import partial
            loop = asyncio.get_event_loop()
            
            new_sug = await loop.run_in_executor(None, partial(generate_single_alternative, gemini_key, task_desc, user_name, previous_sugg, all_users))
            
            if new_sug:
                updated_alts = list(alts_list)
                while len(updated_alts) <= idx:
                    updated_alts.append({"prioridad": "Media", "reasignacion_sugerida": "", "alternativa_solucion": "Reagendar para mañana."})
                updated_alts[idx] = new_sug
                
                await loop.run_in_executor(None, partial(update_report_suggestions, report_id, updated_alts))
                
                loading.dismiss()
                ui.notify('Sugerencia de solución regenerada.', type='success')
                refresh_informes()
            else:
                loading.dismiss()
                ui.notify('No se pudo obtener respuesta de la IA. Verifique su API key.', type='negative')
        except Exception as e:
            loading.dismiss()
            ui.notify(f'Error al regenerar: {e}', type='negative')

    def refresh_informes():
        content_area.clear()
        with content_area:
            ui.label('📊 Historial de Informes de IA').classes('text-2xl font-bold text-slate-800')
            ui.label('Consulte los reportes diarios cerrados y las alternativas de solución recomendadas por la Inteligencia Artificial.').classes('text-slate-500 italic text-sm mt-1 mb-6')
            
            # Filters block
            with ui.row().classes('w-full justify-between items-end bg-white p-4 rounded-xl shadow-sm border border-slate-100 mb-6 gap-4'):
                # User filter dropdown
                with ui.column().classes('w-full md:w-1/4 gap-1'):
                    ui.label('Filtrar Colaborador:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider')
                    user_filter_select = ui.select(['TODOS'] + all_users, value=state.get('informes_filter_user', 'TODOS'), on_change=lambda e: update_filter_user(e.value)).classes('w-full')
                
                # Start date filter (browser native date selector via type=date input)
                with ui.column().classes('w-full md:w-1/4 gap-1'):
                    ui.label('Fecha Inicio:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider')
                    start_date_input = ui.input(value=state.get('informes_filter_start', (datetime.date.today() - datetime.timedelta(days=14)).strftime("%Y-%m-%d")), on_change=lambda e: update_filter_start(e.value)).props('type=date').classes('w-full')
                
                # End date filter
                with ui.column().classes('w-full md:w-1/4 gap-1'):
                    ui.label('Fecha Fin:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider')
                    end_date_input = ui.input(value=state.get('informes_filter_end', datetime.date.today().strftime("%Y-%m-%d")), on_change=lambda e: update_filter_end(e.value)).props('type=date').classes('w-full')
            
            # Fetch reports
            filter_user = None if state['informes_filter_user'] == 'TODOS' else state['informes_filter_user']
            start_date_str = state['informes_filter_start']
            end_date_str = state['informes_filter_end']
            
            reports = get_reports(filter_user, start_date_str, end_date_str, include_admin_users=state['admin_mode'])
            
            if not reports:
                with ui.card().classes('w-full p-8 items-center justify-center text-center bg-white shadow-sm border border-slate-100 rounded-xl mt-6'):
                    ui.label('🔍 No se encontraron informes').classes('text-lg font-bold text-slate-700 mb-1')
                    ui.label('No se encontraron informes cerrados en el rango seleccionado.').classes('text-slate-400 text-sm')
                return
                
            # Render reports list
            with ui.column().classes('w-full gap-6'):
                for rep in reports:
                    rep_id = rep['id']
                    rep_date = rep['date']
                    rep_user = rep['user_name']
                    
                    # Decouple resolved / unresolved tasks from dict
                    resolved = rep.get('resolved_tasks', [])
                    unresolved = rep.get('unresolved_tasks', [])
                    alts = rep.get('alternatives_of_solution', [])
                    
                    # Metric summary label
                    metrics_str = f"✓ {len(resolved)} cumplidas | ✗ {len(unresolved)} pendientes"
                    
                    # Report Card
                    with ui.card().classes('w-full p-6 bg-white shadow-md border border-slate-100 rounded-xl gap-4'):
                        # Header
                        with ui.row().classes('w-full justify-between items-center border-b border-slate-100 pb-3 mb-2'):
                            with ui.row().classes('items-center gap-3'):
                                ui.label(f"📅 Fecha: {rep_date}").classes('text-sm font-bold text-slate-800')
                                ui.label(f"👤 Colaborador: {rep_user}").classes('text-sm font-bold text-slate-500')
                            ui.label(metrics_str).classes('text-xs font-black bg-blue-50 text-blue-700 px-3 py-1 rounded-full')
                            
                        # Two column layout (Resolved vs Unresolved)
                        with ui.grid().classes('w-full grid-cols-1 md:grid-cols-2 gap-6'):
                            # Left Column: Resolved
                            with ui.column().classes('gap-3'):
                                ui.label('✅ Actividades Resueltas').classes('text-xs font-bold text-green-700 uppercase tracking-wider border-b pb-1 mb-1')
                                if not resolved:
                                    ui.label('Ninguna actividad completada.').classes('text-xs text-slate-400 italic')
                                else:
                                    for t in resolved:
                                        t_time = f"🕒 {t.get('time_info')}" if t.get('time_info') else "Sin horario"
                                        carried_from = t.get('carried_over_from')
                                        cof_str = ""
                                        if carried_from:
                                            try:
                                                cof_dt = datetime.datetime.strptime(carried_from, "%Y-%m-%d")
                                                cof_str = f" 🔄 (arrastrada del {cof_dt.strftime('%d/%m')})"
                                            except Exception:
                                                cof_str = f" 🔄 (arrastrada del {carried_from})"
                                                
                                        with ui.card().classes('w-full p-3 bg-white border-l-4 border-green-500 shadow-sm rounded-lg hover:shadow-md transition-shadow gap-1'):
                                            ui.label(f"{t['description']}{cof_str}").classes('text-xs text-slate-700 font-semibold leading-tight')
                                            ui.label(t_time).classes('text-[9px] text-slate-400 font-medium')
                                        
                            # Right Column: Unresolved
                            with ui.column().classes('gap-3'):
                                ui.label('❌ Actividades No Resueltas').classes('text-xs font-bold text-red-700 uppercase tracking-wider border-b pb-1 mb-1')
                                if not unresolved:
                                    ui.label('Ninguna actividad pendiente.').classes('text-xs text-slate-400 italic')
                                else:
                                    for idx, t in enumerate(unresolved):
                                        t_time = f"🕒 {t.get('time_info')}" if t.get('time_info') else "Sin horario"
                                        prio = t.get('priority', 'Normal')
                                        carried_from = t.get('carried_over_from')
                                        cof_str = ""
                                        if carried_from:
                                            try:
                                                cof_dt = datetime.datetime.strptime(carried_from, "%Y-%m-%d")
                                                cof_str = f" 🔄 (arrastrada del {cof_dt.strftime('%d/%m')})"
                                            except Exception:
                                                cof_str = f" 🔄 (arrastrada del {carried_from})"
                                                
                                        # Priority color border left coding
                                        border_color = 'border-slate-300'
                                        if prio == 'Alta':
                                            border_color = 'border-red-500'
                                        elif prio == 'Media':
                                            border_color = 'border-amber-500'
                                        elif prio == 'Baja':
                                            border_color = 'border-green-500'
                                            
                                        with ui.card().classes(f'w-full p-4 bg-white border-l-4 {border_color} shadow-sm rounded-xl hover:shadow-md transition-shadow gap-2'):
                                            # Header task info inside card
                                            with ui.column().classes('gap-0.5'):
                                                ui.label(f"{t['description']}{cof_str}").classes('text-xs text-slate-700 font-bold leading-tight')
                                                with ui.row().classes('items-center gap-2 mt-1'):
                                                    ui.label(t_time).classes('text-[9px] text-slate-400 font-medium')
                                                    ui.label(f"Prioridad: {prio}").classes('text-[9px] font-bold text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded')
                                            
                                            # Embed IA Suggestion directly inside unresolved activity card
                                            suggestion = alts[idx] if idx < len(alts) else {"prioridad": "Media", "reasignacion_sugerida": "", "alternativa_solucion": "Reagendar para mañana."}
                                            
                                            with ui.column().classes('w-full gap-1 p-2.5 bg-slate-50 border border-slate-100 rounded-lg mt-1'):
                                                with ui.row().classes('w-full justify-between items-center no-wrap gap-2'):
                                                    ui.label('💡 Alternativa de IA:').classes('text-[9px] font-bold text-slate-400 uppercase tracking-wider')
                                                    
                                                    # Regenerate button inside card
                                                    ui.button(icon='refresh', on_click=lambda _, r_id=rep_id, idx_val=idx, t_desc=t['description'], s_val=suggestion, list_alts=alts, u_name=rep_user: regen_sugg(r_id, idx_val, t_desc, s_val, list_alts, u_name)).props('flat round dense size=xs color=primary')
                                                    
                                                render_solution_suggestion_nicegui(suggestion)
                                                
                                            # Field "EN AVANCE O MOTIVO DE NO EJECUCIÓN"
                                            exec_statuses = rep.get('execution_status', [])
                                            current_reason = exec_statuses[idx] if idx < len(exec_statuses) else ""
                                            
                                            reason_input = ui.textarea(
                                                label="En avance o motivo de no ejecución",
                                                placeholder="Explica por qué no se ejecutó (opcional)",
                                                value=current_reason
                                            ).classes('w-full text-xs mt-2').props('dense outlined rows=2')
                                            
                                            async def save_reason(r_id=rep_id, idx_val=idx, text_area=reason_input, list_reasons=exec_statuses):
                                                updated_reasons = list(list_reasons) if list_reasons else []
                                                while len(updated_reasons) <= idx_val:
                                                    updated_reasons.append("")
                                                updated_reasons[idx_val] = text_area.value.strip() if text_area.value else ""
                                                
                                                import asyncio
                                                from functools import partial
                                                loop = asyncio.get_event_loop()
                                                await loop.run_in_executor(None, partial(update_task_execution_status, r_id, updated_reasons))
                                                ui.notify("Motivo guardado con éxito.", type='success')
                                                refresh_informes()
                                                
                                            ui.button('💾 Guardar Motivo', on_click=save_reason).classes('mt-1 self-end text-xs bg-slate-700 text-white font-semibold rounded px-3 py-1 shadow-sm hover:bg-slate-800 transition-colors')

    def update_export_date(val):
        state['export_date'] = val
        refresh_exportar()
        
    async def run_excel_export(monday_str, count):
        import tempfile
        import os
        import asyncio
        from functools import partial
        
        loading = ui.notification('Generando archivo formateado...', spinner=True, timeout=0)
        try:
            fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            
            loop = asyncio.get_event_loop()
            
            # Export excel in background thread
            await loop.run_in_executor(None, partial(export_weekly_tasks_to_excel, monday_str, temp_path, state['admin_mode']))
            
            # Save export info
            await loop.run_in_executor(None, partial(save_last_export_info, datetime.date.today().strftime("%Y-%m-%d"), count))
            
            # Download file using NiceGUI native ui.download
            filename = f"Cuadro_de_Produccion_{monday_str}.xlsx"
            ui.download(temp_path, filename)
            
            loading.dismiss()
            ui.notify('Documento generado satisfactoriamente.', type='success')
        except Exception as e:
            loading.dismiss()
            ui.notify(f'Error al generar Excel: {e}', type='negative')

    def refresh_exportar():
        content_area.clear()
        with content_area:
            ui.label('📥 Exportación Semanal a Excel').classes('text-2xl font-bold text-slate-800')
            ui.label('Descargue el cuadro de producción semanal en formato original Excel, respetando los colores de prioridad de la matriz.').classes('text-slate-500 italic text-sm mt-1 mb-6')
            
            # Selector container card
            with ui.card().classes('w-full p-6 bg-white shadow-md border border-slate-100 rounded-xl gap-4'):
                ui.label('Seleccione un día del periodo semanal:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider')
                
                # Date input defaulting to selected export_date
                export_date_input = ui.input(value=state.get('export_date', datetime.date.today().strftime("%Y-%m-%d")), on_change=lambda e: update_export_date(e.value)).props('type=date').classes('w-full max-w-md')
                
                # Calculations section
                snap_date_str = state.get('export_date', datetime.date.today().strftime("%Y-%m-%d"))
                try:
                    export_day = datetime.datetime.strptime(snap_date_str, "%Y-%m-%d").date()
                except Exception:
                    export_day = datetime.date.today()
                    
                monday_diff = export_day.weekday()
                monday_date = export_day - datetime.timedelta(days=monday_diff)
                snap_monday_str = monday_date.strftime("%Y-%m-%d")
                saturday_date = monday_date + datetime.timedelta(days=5)
                
                # Rango de semana alert banner
                with ui.row().classes('w-full items-center gap-3 p-4 bg-blue-50 rounded-xl border border-blue-200 text-blue-800 shadow-sm'):
                    ui.label('📅').classes('text-xl')
                    ui.label(f"Rango de semana laboral: Lunes {monday_date.strftime('%d/%m/%Y')} al Sábado {saturday_date.strftime('%d/%m/%Y')}").classes('font-bold text-xs')
                
                # Count tasks in database for selected week
                week_days = get_week_dates(snap_monday_str)
                tasks_count = 0
                for u in all_users:
                    for d in week_days:
                        tasks_count += len(get_tasks(u, d, include_admin_users=state['admin_mode']))
                        
                # Count info text
                with ui.row().classes('items-center gap-2 mt-2'):
                    ui.label('Cantidad total de tareas en la base de datos para esta semana:').classes('text-sm text-slate-600')
                    ui.label(str(tasks_count)).classes('text-sm font-black text-slate-800 bg-slate-100 px-2 py-0.5 rounded')
                
                # Warning if 0 tasks
                if tasks_count == 0:
                    with ui.row().classes('w-full items-center gap-2 p-3 bg-amber-50 rounded-lg border border-amber-200 text-amber-800 text-xs font-semibold'):
                        ui.label('⚠️')
                        ui.label('No hay tareas registradas para esta semana en la base de datos, el Excel se exportará en blanco.')
                
                # Generate button
                ui.button('📊 Generar Documento de Excel', on_click=lambda: run_excel_export(snap_monday_str, tasks_count)).classes('w-full md:w-auto bg-blue-900 text-white font-bold rounded-lg px-6 py-2.5 shadow hover:bg-blue-800 transition-colors mt-4')

    def update_calendar_year(val):
        state['calendar_year'] = int(val)
        max_d = calendar.monthrange(state['calendar_year'], state['calendar_month'])[1]
        if state['calendar_selected_day'] > max_d:
            state['calendar_selected_day'] = max_d
        refresh_calendario()
        
    def update_calendar_month(val):
        state['calendar_month'] = int(val)
        max_d = calendar.monthrange(state['calendar_year'], state['calendar_month'])[1]
        if state['calendar_selected_day'] > max_d:
            state['calendar_selected_day'] = max_d
        refresh_calendario()
        
    def update_calendar_selected_day(val):
        state['calendar_selected_day'] = int(val)
        refresh_calendario()

    def go_to_pizarra_date(day_num):
        selected_date_str = f"{state['calendar_year']:04d}-{state['calendar_month']:02d}-{day_num:02d}"
        state['selected_date'] = datetime.datetime.strptime(selected_date_str, "%Y-%m-%d").date()
        state['nav_selection'] = '📋 Pizarra'
        refresh_layout()

    def refresh_calendario():
        content_area.clear()
        with content_area:
            ui.label('📅 Planificador Mensual de Actividades').classes('text-2xl font-bold text-slate-800')
            ui.label('Haga clic en cualquier día para ver sus tareas en la pizarra.').classes('text-slate-500 italic text-sm mt-1 mb-6')
            
            year = state['calendar_year']
            month = state['calendar_month']
            selected_day = state['calendar_selected_day']
            selected_user = state['current_user']
            
            # Selector of Month and Year in a row
            with ui.row().classes('w-full justify-between items-center bg-white p-4 rounded-xl shadow-sm border border-slate-100 mb-6 gap-4'):
                with ui.row().classes('items-center gap-4'):
                    # Year Selector
                    with ui.column().classes('gap-1'):
                        ui.label('Año:').classes('text-[10px] font-bold text-slate-400 uppercase tracking-wider')
                        years_opts = [2026, 2027, 2028, 2029, 2030]
                        ui.select(years_opts, value=year, on_change=lambda e: update_calendar_year(e.value)).classes('w-32')
                    
                    # Month Selector
                    with ui.column().classes('gap-1'):
                        ui.label('Mes:').classes('text-[10px] font-bold text-slate-400 uppercase tracking-wider')
                        months_dict = {
                            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                            5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                            9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
                        }
                        ui.select(months_dict, value=month, on_change=lambda e: update_calendar_month(e.value)).classes('w-44')
                
                # Legend description
                with ui.row().classes('items-center gap-3 text-xs font-semibold text-slate-500 bg-slate-50 px-3 py-2 rounded-lg border border-slate-100'):
                    with ui.row().classes('items-center gap-1'):
                        ui.element('div').classes('w-2.5 h-2.5 bg-green-500 rounded-full')
                        ui.label('Completadas/Sin tareas')
                    with ui.row().classes('items-center gap-1'):
                        ui.element('div').classes('w-2.5 h-2.5 bg-amber-500 rounded-full')
                        ui.label('Pendientes')
                    with ui.row().classes('items-center gap-1'):
                        ui.element('div').classes('w-2.5 h-2.5 bg-red-500 rounded-full')
                        ui.label('Atrasadas')

            # Render Calendar Grid
            week_days_headers = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
            
            # Container of grid
            with ui.card().classes('w-full p-6 bg-white shadow-md border border-slate-100 rounded-xl gap-4'):
                # Week headers row
                with ui.row().classes('grid grid-cols-7 gap-2 w-full text-center mb-2'):
                    for header in week_days_headers:
                        ui.label(header).classes('text-xs font-bold text-slate-400 uppercase tracking-wider')
                
                # Calendar month days matrix calculation
                cal = calendar.Calendar(firstweekday=0)
                month_weeks = cal.monthdayscalendar(year, month)
                
                # Loop through weeks to render rows
                for week in month_weeks:
                    with ui.row().classes('grid grid-cols-7 gap-2 w-full'):
                        for day in week:
                            if day == 0:
                                ui.element('div').classes('p-2 min-h-[70px] bg-slate-50/50 rounded-lg border border-transparent')
                            else:
                                day_str = f"{year:04d}-{month:02d}-{day:02d}"
                                day_tasks = get_tasks(selected_user, day_str, include_admin_users=state['admin_mode'])
                                total_tasks = len(day_tasks)
                                
                                bg_class = 'bg-green-50 border-green-100'
                                dot_class = 'bg-green-500'
                                
                                if total_tasks > 0:
                                    incomplete_tasks = [t for t in day_tasks if t['completed'] == 0]
                                    if len(incomplete_tasks) > 0:
                                        has_overdue = any(t.get('carried_over_from') for t in incomplete_tasks)
                                        if has_overdue:
                                            bg_class = 'bg-red-50 border-red-100'
                                            dot_class = 'bg-red-500'
                                        else:
                                            bg_class = 'bg-amber-50 border-amber-100'
                                            dot_class = 'bg-amber-500'
                                else:
                                    bg_class = 'bg-slate-50/70 border-slate-100 text-slate-400'
                                    dot_class = 'bg-slate-300'
                                
                                is_selected = (day == selected_day)
                                select_border = 'ring-2 ring-blue-500 ring-offset-1' if is_selected else 'border border-slate-200/60'
                                
                                with ui.card().classes(f'p-2 text-center min-h-[75px] justify-between cursor-pointer shadow-sm hover:shadow-md hover:scale-[1.02] hover:bg-slate-100/50 transition-all duration-200 rounded-xl {bg_class} {select_border}') as day_cell:
                                    day_cell.on('click', lambda _, d=day: go_to_pizarra_date(d))
                                    
                                    ui.label(str(day)).classes('text-sm font-bold text-slate-700 w-full text-left leading-none')
                                    
                                    with ui.row().classes('w-full justify-between items-center no-wrap mt-2'):
                                        if total_tasks > 0:
                                            ui.label(f"{total_tasks} act.").classes('text-[9px] text-slate-500 font-bold')
                                            ui.element('div').classes(f'w-2 h-2 rounded-full {dot_class}')
                                        else:
                                            ui.label('').classes('text-[9px]')
                                            ui.element('div').classes(f'w-2 h-2 rounded-full {dot_class}')
            
            # Selection slider for day details
            max_days_in_month = calendar.monthrange(year, month)[1]
            if selected_day > max_days_in_month:
                selected_day = max_days_in_month
                
            ui.separator().classes('my-6')
            
            with ui.column().classes('w-full gap-2 mb-2'):
                ui.label(f"Seleccione el día para ver actividades detalladas (1 al {max_days_in_month}):").classes('text-xs font-bold text-slate-500 uppercase tracking-wider')
                ui.slider(min=1, max=max_days_in_month, value=selected_day, on_change=lambda e: update_calendar_selected_day(e.value)).classes('w-full mb-2')
            
            # Display detailed list of tasks
            selected_date_str = f"{year:04d}-{month:02d}-{selected_day:02d}"
            
            ui.label(f"📅 Actividades del {selected_date_str}").classes('text-lg font-bold text-slate-800 mb-4')
            
            selected_day_tasks = get_tasks(selected_user, selected_date_str, include_admin_users=state['admin_mode'])
            
            if not selected_day_tasks:
                with ui.row().classes('w-full items-center gap-3 p-4 bg-green-50 rounded-xl border border-green-200 text-green-700 shadow-sm'):
                    ui.label('🎉').classes('text-2xl')
                    ui.label('No hay actividades registradas o programadas para este día.').classes('font-bold text-sm')
            else:
                with ui.column().classes('w-full gap-3 mb-6'):
                    for t in selected_day_tasks:
                        t_time = f"🕒 {t.get('time_info')}" if t.get('time_info') else "Sin horario"
                        prio = t.get('priority', 'Normal')
                        completed = t['completed']
                        carried = t['carried_over_from']
                        
                        border_color = 'border-slate-300'
                        if prio == 'Alta':
                            border_color = 'border-red-500'
                        elif prio == 'Media':
                            border_color = 'border-amber-500'
                        elif prio == 'Baja':
                            border_color = 'border-green-500'
                            
                        with ui.card().classes(f'w-full p-4 bg-white border-l-4 {border_color} shadow-sm rounded-xl hover:shadow transition-shadow gap-2'):
                            with ui.row().classes('w-full justify-between items-center no-wrap'):
                                with ui.row().classes('items-center gap-2 w-full'):
                                    if completed:
                                        ui.label('✓').classes('text-green-600 font-black text-sm bg-green-50 border border-green-200 rounded px-1.5 py-0.5')
                                        desc_style = 'text-xs text-slate-500 font-semibold line-through break-words'
                                    else:
                                        ui.label('✗').classes('text-amber-600 font-black text-sm bg-amber-50 border border-amber-200 rounded px-1.5 py-0.5')
                                        desc_style = 'text-xs text-slate-800 font-semibold break-words'
                                        
                                    carry_lbl = " 🔄 (arrastrada)" if carried else ""
                                    ui.label(f"{t['description']}{carry_lbl}").classes(desc_style)
                                
                                with ui.row().classes('items-center gap-2 no-wrap shrink-0'):
                                    ui.label(t_time).classes('text-[10px] text-slate-400 font-medium')
                                    ui.label(prio).classes('text-[9px] font-black text-slate-500 bg-slate-100 px-2 py-0.5 rounded')

    # Chatbot assistant functions
    def toggle_chat_card():
        state['chat_visible'] = not state['chat_visible']
        if state['chat_visible']:
            refresh_chat_messages()

    def toggle_chat_minimize():
        state['chat_minimized'] = not state['chat_minimized']
        if state['chat_minimized']:
            chat_card.classes(replace='fixed bottom-24 right-6 w-[385px] h-[50px] z-[9999] bg-white shadow-2xl rounded-2xl border border-slate-100 flex flex-col p-0 overflow-hidden transition-all duration-200')
            minimize_btn.props('icon=keyboard_arrow_up')
        else:
            chat_card.classes(replace='fixed bottom-24 right-6 w-[385px] h-[520px] z-[9999] bg-white shadow-2xl rounded-2xl border border-slate-100 flex flex-col p-0 overflow-hidden transition-all duration-200')
            minimize_btn.props('icon=keyboard_arrow_down')

    def refresh_chat_messages():
        with messages_area:
            messages_area.clear()
            for msg in state['assistant_messages']:
                align_class = 'justify-end' if msg['role'] == 'user' else 'justify-start'
                bubble_color = 'bg-blue-100 text-slate-800 rounded-tr-none border border-blue-200 shadow-sm' if msg['role'] == 'user' else 'bg-white text-slate-800 rounded-tl-none border border-slate-100 shadow-sm'
                
                with ui.row().classes(f'w-full {align_class} no-wrap gap-2'):
                    if msg['role'] == 'assistant':
                        ui.label('🤖').classes('text-sm mt-1')
                    with ui.card().classes(f'p-2.5 rounded-xl {bubble_color} max-w-[80%] gap-1'):
                        if msg.get('is_thinking'):
                            with ui.row().classes('items-center gap-2 no-wrap'):
                                ui.spinner(size='xs', color='primary')
                                ui.label(msg['content']).classes('text-xs italic text-slate-500')
                        else:
                            ui.markdown(msg['content']).classes('text-xs leading-normal')
                    if msg['role'] == 'user':
                        ui.label('👤').classes('text-sm mt-1')
                        
            # Show suggestions if empty / only welcome msg
            if len(state['assistant_messages']) <= 1:
                ui.label('Preguntas sugeridas:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider mt-4 mb-1 w-full')
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
                    with ui.card().classes('w-full p-2 bg-slate-50 hover:bg-slate-100/80 border border-slate-200/50 rounded-lg cursor-pointer transition-colors shadow-sm') as sugg_btn:
                        sugg_btn.on('click', lambda _, s=sugg: send_chat_message(s))
                        ui.label(s).classes('text-xs font-medium text-slate-700 w-full text-left')
            
            # Scroll to bottom
            ui.run_javascript("setTimeout(() => { const el = document.getElementById('chat_messages'); if(el) { el.scrollTop = el.scrollHeight; } }, 50)")

    async def send_chat_message(text_val):
        if not text_val or not text_val.strip():
            return
            
        # 1. Immediate UI Updates
        chat_input.disable()
        chat_send_btn.disable()
        chat_input.value = '' # Clear input field
        
        # Append user message
        state['assistant_messages'].append({"role": "user", "content": text_val.strip()})
        # Append temporary thinking message
        state['assistant_messages'].append({"role": "assistant", "content": "Pensando...", "is_thinking": True})
        refresh_chat_messages()
        
        try:
            # Get Gemini key
            gemini_key = os.environ.get("GEMINI_API_KEY", "")
            
            if not gemini_key or not gemini_key.strip():
                if state['assistant_messages'] and state['assistant_messages'][-1].get("is_thinking"):
                    state['assistant_messages'].pop()
                state['assistant_messages'].append({
                    "role": "assistant",
                    "content": "⚠️ Configura la clave API de Gemini en las variables de entorno para usar el asistente."
                })
                return
                
            # Gather chatbot context data
            u_active = state['current_user'] or "MARY CRUZ"
            today_s = datetime.date.today().strftime("%Y-%m-%d")
            
            import asyncio
            from functools import partial
            loop = asyncio.get_event_loop()
            
            today_tasks_list = await loop.run_in_executor(None, partial(get_tasks, u_active, today_s, state['admin_mode']))
            completed_list = [t for t in today_tasks_list if t['completed'] == 1]
            pending_list = [t for t in today_tasks_list if t['completed'] == 0]
            overdue_tasks = await loop.run_in_executor(None, partial(get_overdue_pending_tasks, u_active, today_s, state['admin_mode']))
            overdue_count = len(overdue_tasks)
            
            completed_str = "\n".join([f"- {t['description']} (Horario: {t.get('time_info') or 'No especificado'})" for t in completed_list]) if completed_list else "Ninguna"
            pending_str = "\n".join([f"- {t['description']} (Horario: {t.get('time_info') or 'No especificado'})" for t in pending_list]) if pending_list else "Ninguna"
            
            system_prompt = (
                "Eres un asistente inteligente integrado en la aplicación 'CUADROpz' (Control de Producción).\n"
                f"Usuario activo actual: {u_active}\n"
                f"Fecha de hoy: {today_s}\n\n"
                "--- CONTEXTO DE HOY ---\n"
                f"Tareas completadas hoy:\n{completed_str}\n\n"
                f"Tareas pendientes de hoy:\n{pending_str}\n\n"
                f"Tareas atrasadas acumuladas de días anteriores: {overdue_count}\n"
                "------------------------\n\n"
                "Responde las consultas del usuario basándote en este contexto. Sé profesional, conciso y de gran ayuda en la gestión de sus tareas."
            )
            
            # Limit history to the last 5 messages (excluding the thinking message and the user prompt we just appended)
            recent_history = state['assistant_messages'][:-2]
            if len(recent_history) > 5:
                recent_history = recent_history[-5:]
                
            history = []
            for h_msg in recent_history:
                role_map = "user" if h_msg["role"] == "user" else "model"
                history.append({"role": role_map, "parts": [{"text": h_msg["content"]}]})
                
            from google import genai
            from google.genai import types
            
            import queue
            import threading
            
            formatted_history = []
            for h in history:
                formatted_history.append(
                    types.Content(
                        role=h["role"],
                        parts=[types.Part(text=h["parts"][0]["text"])]
                    )
                )
                
            q = queue.Queue()
            
            def run_gemini_streaming():
                try:
                    client = genai.Client(http_options={'api_version': 'v1'})
                    try:
                        chat = client.chats.create(
                            model='gemini-3.6-flash',
                            history=formatted_history,
                            config=types.GenerateContentConfig(
                                system_instruction=system_prompt,
                                temperature=0.5
                            )
                        )
                    except Exception as e_flash:
                        print("Error creating chat with gemini-3.6-flash, trying fallback to gemini-2.0-flash-lite:", e_flash)
                        chat = client.chats.create(
                            model='gemini-2.0-flash-lite',
                            history=formatted_history,
                            config=types.GenerateContentConfig(
                                system_instruction=system_prompt,
                                temperature=0.5
                            )
                        )
                    
                    for chunk in chat.send_message_stream(text_val.strip()):
                        if chunk.text:
                            q.put(chunk.text)
                except Exception as e:
                    q.put(e)
                finally:
                    q.put(None)
                    
            # Start streaming thread
            threading.Thread(target=run_gemini_streaming, daemon=True).start()
            
            # Read from queue in async thread
            async def consume_queue():
                first_chunk = True
                while True:
                    chunk_val = await loop.run_in_executor(None, q.get)
                    if chunk_val is None:
                        break
                    if isinstance(chunk_val, Exception):
                        raise chunk_val
                        
                    if first_chunk:
                        # Clear thinking state and set is_thinking to False
                        if state['assistant_messages'] and state['assistant_messages'][-1].get("is_thinking"):
                            state['assistant_messages'][-1]['is_thinking'] = False
                            state['assistant_messages'][-1]['content'] = ""
                        first_chunk = False
                        
                    state['assistant_messages'][-1]['content'] += chunk_val
                    refresh_chat_messages()
                    
            try:
                # 30 second timeout for the entire generation stream
                await asyncio.wait_for(consume_queue(), timeout=30.0)
            except asyncio.TimeoutError:
                if state['assistant_messages'] and state['assistant_messages'][-1].get("is_thinking"):
                    state['assistant_messages'].pop()
                state['assistant_messages'].append({
                    "role": "assistant",
                    "content": "⏳ La respuesta está tomando más tiempo de lo esperado. Intenta de nuevo o reformula tu pregunta."
                })
            except Exception as e_inner:
                if state['assistant_messages'] and state['assistant_messages'][-1].get("is_thinking"):
                    state['assistant_messages'].pop()
                state['assistant_messages'].append({
                    "role": "assistant",
                    "content": f"❌ Error al conectar con la IA. Verifica tu conexión o clave API: {str(e_inner)}"
                })
                
        except Exception as e:
            if state['assistant_messages'] and state['assistant_messages'][-1].get("is_thinking"):
                state['assistant_messages'].pop()
            state['assistant_messages'].append({
                "role": "assistant",
                "content": f"❌ Error general en el chat: {str(e)}"
            })
            
        finally:
            chat_input.enable()
            chat_send_btn.enable()
            chat_input.focus()
            refresh_chat_messages()

    # Chat Float Button and Card container (Global, always visible)
    with ui.button(icon='chat', on_click=lambda: toggle_chat_card()).classes('fixed bottom-6 right-6 z-[9999] rounded-full w-14 h-14 bg-white text-[#1E3A8A] border-2 border-[#1E3A8A] shadow-2xl hover:scale-110 transition-all duration-200').props('flat'):
        pass
        
    chat_card = ui.card().classes('fixed bottom-24 right-6 w-[385px] h-[520px] z-[9999] bg-white shadow-2xl rounded-2xl border border-slate-100 flex flex-col p-0 overflow-hidden transition-all duration-200')
    chat_card.bind_visibility_from(state, 'chat_visible')
    
    with chat_card:
        # Header banner
        with ui.row().classes('w-full bg-[#1E3A8A] text-white p-3 flex justify-between items-center no-wrap'):
            with ui.row().classes('items-center gap-2'):
                ui.label('💬').classes('text-lg')
                with ui.column().classes('gap-0'):
                    ui.label('Asistente Virtual').classes('text-xs font-bold leading-none')
                    ui.label('CUADROpz AI Chat').classes('text-[9px] text-blue-200 font-medium leading-none')
            with ui.row().classes('items-center gap-1'):
                minimize_btn = ui.button(icon='keyboard_arrow_down', on_click=toggle_chat_minimize).props('flat round dense size=sm color=white')
                ui.button(icon='close', on_click=lambda: toggle_chat_card()).props('flat round dense size=sm color=white')
            
        # Messages Container (scrollable)
        messages_area = ui.column().classes('w-full flex-grow overflow-y-auto p-4 gap-3 bg-slate-50').props('id="chat_messages"')
        
        # Loader banner "pensando..."
        with ui.row().classes('w-full px-4 py-2 bg-slate-100 border-t border-slate-200/50 justify-start items-center gap-2') as loader_area:
            ui.spinner(size='xs', color='primary')
            ui.label('Asistente está pensando...').classes('text-[10px] text-slate-500 italic')
        loader_area.set_visibility(False)
            
        # Input bar
        with ui.row().classes('w-full p-3 border-t border-slate-100 bg-white items-center gap-2 no-wrap'):
            chat_input = ui.input(placeholder='Escribe tu consulta aquí...').classes('flex-grow text-xs').props('dense outlined rounded')
            chat_send_btn = ui.button(icon='send').props('flat round dense color=primary')
            
            chat_input.on('keydown.enter', lambda: send_chat_message(chat_input.value))
            chat_send_btn.on('click', lambda: send_chat_message(chat_input.value))

    # Admin Mode dialogs, functions, and views
    admin_content = None

    # Auth Dialog
    auth_dialog = ui.dialog()
    with auth_dialog:
        with ui.card().classes('p-6 bg-white shadow-2xl rounded-2xl w-[320px] gap-4'):
            ui.label('🔐 Acceso de Administrador').classes('text-lg font-bold text-slate-800 w-full text-center')
            ui.label('Ingrese la contraseña de administrador para desbloquear los controles globales.').classes('text-xs text-slate-400 w-full text-center mb-2')
            
            pwd_input = ui.input('Contraseña:', password=True, password_toggle_button=True).classes('w-full').props('outlined autofocus')
            pwd_input.on('keydown.enter', lambda: validate_admin_login(pwd_input.value))
            
            with ui.row().classes('w-full justify-end gap-2'):
                ui.button('Cancelar', on_click=auth_dialog.close).props('flat')
                ui.button('Validar', on_click=lambda: validate_admin_login(pwd_input.value)).classes('bg-blue-900 text-white font-bold rounded-lg px-4')
                
    def show_admin_login_dialog():
        pwd_input.value = ''
        auth_dialog.open()
        
    def validate_admin_login(password_value):
        ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin1995")
        print(f"[DEBUG ADMIN] Solicitado ingreso. Password cargado en env: '{os.environ.get('ADMIN_PASSWORD')}' (usando fallback: 'admin1995')")
        if password_value == ADMIN_PASSWORD:
            state['admin_mode'] = True
            app.storage.user['admin_mode'] = True
            auth_dialog.close()
            ui.notify('✅ Modo Admin activado correctamente.', type='success')
            reload_users_list()
            refresh_sidebar_navigation()
            refresh_layout()
        else:
            ui.notify('❌ Contraseña incorrecta. Intenta nuevamente.', type='negative')
            
    # Confirm Deactivation Dialog
    confirm_deactivate_dialog = ui.dialog()
    with confirm_deactivate_dialog:
        with ui.card().classes('p-6 bg-white shadow-2xl rounded-2xl w-[320px] gap-4'):
            ui.label('Confirmar Salida').classes('text-base font-bold text-slate-800')
            ui.label('¿Está seguro de que desea salir del Modo Administrador?').classes('text-xs text-slate-500')
            with ui.row().classes('w-full justify-end gap-2'):
                ui.button('Cancelar', on_click=confirm_deactivate_dialog.close).props('flat')
                ui.button('Desactivar', on_click=lambda: confirm_deactivate_mode()).classes('bg-red-600 text-white font-bold rounded-lg px-4')
                
    def deactivate_admin_mode():
        confirm_deactivate_dialog.open()
        
    def confirm_deactivate_mode():
        confirm_deactivate_dialog.close()
        state['admin_mode'] = False
        app.storage.user['admin_mode'] = False
        if state['nav_selection'] == '👑 Admin':
            state['nav_selection'] = '🏠 Inicio'
        ui.notify('Modo Administrador desactivado.', type='info')
        reload_users_list()
        refresh_sidebar_navigation()
        refresh_layout()

    # User deletion confirm dialog
    delete_user_confirm_dialog = ui.dialog()
    with delete_user_confirm_dialog:
        with ui.card().classes('p-6 bg-white shadow-2xl rounded-2xl w-[320px] gap-4'):
            ui.label('⚠️ Confirmar Eliminación').classes('text-base font-bold text-slate-800')
            ui.label('¿Está seguro de que desea eliminar a este colaborador? Se borrarán todas sus tareas e informes asociados permanentemente.').classes('text-xs text-slate-500')
            with ui.row().classes('w-full justify-end gap-2'):
                ui.button('Cancelar', on_click=delete_user_confirm_dialog.close).props('flat')
                ui.button('Eliminar', on_click=lambda: execute_delete_user()).classes('bg-red-600 text-white font-bold rounded-lg px-4')

    def confirm_delete_user(target_username):
        state['delete_target_user'] = target_username
        delete_user_confirm_dialog.open()
        
    def execute_delete_user():
        target_username = state.get('delete_target_user')
        if target_username:
            delete_user_by_admin(target_username)
            ui.notify(f'Usuario {target_username} eliminado correctamente.', type='success')
            nonlocal all_users
            all_users = get_users(include_admins=state['admin_mode'])
            user_select.options = all_users
            
            if state['current_user'] == target_username:
                if all_users:
                    change_user(all_users[0])
                else:
                    state['current_user'] = None
                    
            delete_user_confirm_dialog.close()
            refresh_admin_dashboard()
            refresh_sidebar_navigation()
            refresh_layout()

    def add_admin_user(name_val, is_adm):
        if not name_val or not name_val.strip():
            ui.notify('El nombre del colaborador no puede estar vacío.', type='warning')
            return
        if add_user_with_role(name_val.strip().upper(), is_adm):
            ui.notify('Colaborador agregado con éxito.', type='success')
            nonlocal all_users
            all_users = get_users(include_admins=state['admin_mode'])
            user_select.options = all_users
            refresh_admin_dashboard()
            refresh_sidebar_navigation()
        else:
            ui.notify('Error al agregar colaborador. El nombre ya existe.', type='negative')
            
    def run_toggle_admin(target_username):
        toggle_admin_status(target_username)
        ui.notify(f'Rol del usuario {target_username} actualizado.', type='info')
        refresh_admin_dashboard()
        refresh_sidebar_navigation()

    def save_central_settings(gemini_api_key, smtp_server, smtp_port, sender_email, sender_password, recipient_email, auto_send):
        cfg = {
            "gemini_api_key": gemini_api_key.strip(),
            "smtp_server": smtp_server.strip(),
            "smtp_port": smtp_port.strip(),
            "sender_email": sender_email.strip(),
            "sender_password": sender_password.strip(),
            "recipient_email": recipient_email.strip(),
            "auto_send_enabled": auto_send,
            "last_sent_date": load_email_config().get("last_sent_date", "")
        }
        if save_email_config(cfg):
            os.environ["GEMINI_API_KEY"] = gemini_api_key.strip()
            ui.notify('Configuración guardada con éxito.', type='success')
            refresh_admin_dashboard()
            refresh_sidebar_navigation()
        else:
            ui.notify('Error al guardar la configuración.', type='negative')

    @ui.refreshable
    def refresh_admin_dashboard():
        if not admin_content:
            return
        admin_content.clear()
        with admin_content:
            all_db_users = get_all_users_with_admin_status()
            total_users = len(all_db_users)
            
            from database import db_connect
            conn = db_connect()
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM tasks")
            total_tasks = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM reports")
            total_reports = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM tasks WHERE completed = 1")
            completed_tasks = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM tasks WHERE completed = 0")
            pending_tasks = cursor.fetchone()[0]
            conn.close()
            
            # --- SECTION 1: GLOBAL STATS ---
            ui.label('📊 Estadísticas Globales').classes('text-lg font-bold text-slate-800 mb-2')
            with ui.grid().classes('w-full grid-cols-1 md:grid-cols-4 gap-4 mb-6'):
                with ui.card().classes('p-4 bg-white border-l-4 border-blue-500 shadow-sm rounded-xl gap-1'):
                    ui.label('Usuarios Registrados').classes('text-[10px] font-bold text-slate-400 uppercase tracking-wider')
                    ui.label(str(total_users)).classes('text-2xl font-black text-slate-800')
                with ui.card().classes('p-4 bg-white border-l-4 border-indigo-500 shadow-sm rounded-xl gap-1'):
                    ui.label('Tareas Totales').classes('text-[10px] font-bold text-slate-400 uppercase tracking-wider')
                    ui.label(str(total_tasks)).classes('text-2xl font-black text-slate-800')
                with ui.card().classes('p-4 bg-white border-l-4 border-green-500 shadow-sm rounded-xl gap-1'):
                    ui.label('Informes Cerrados').classes('text-[10px] font-bold text-slate-400 uppercase tracking-wider')
                    ui.label(str(total_reports)).classes('text-2xl font-black text-slate-800')
                ratio = int(completed_tasks / total_tasks * 100) if total_tasks > 0 else 100
                with ui.card().classes('p-4 bg-white border-l-4 border-amber-500 shadow-sm rounded-xl gap-1'):
                    ui.label('Cumplimiento General').classes('text-[10px] font-bold text-slate-400 uppercase tracking-wider')
                    ui.label(f"{ratio}%").classes('text-2xl font-black text-slate-800')
                    
            if total_tasks > 0:
                with ui.card().classes('w-full p-4 bg-white shadow-sm border border-slate-100 rounded-xl mb-6 items-center justify-center'):
                    df_global = pd.DataFrame({
                        "Estado": ["Cumplidas", "Pendientes"],
                        "Cantidad": [completed_tasks, pending_tasks]
                    })
                    fig_global = px.pie(df_global, names="Estado", values="Cantidad", hole=0.5, color="Estado",
                                       color_discrete_map={"Cumplidas": "#10B981", "Pendientes": "#F59E0B"})
                    fig_global.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=220, showlegend=True)
                    ui.plotly(fig_global).classes('w-full max-w-md')
            
            # --- SECTION 2: USER MANAGEMENT ---
            ui.separator().classes('my-4')
            ui.label('👥 Gestión de Colaboradores').classes('text-lg font-bold text-slate-800 mb-2')
            with ui.card().classes('w-full p-4 bg-white shadow-sm border border-slate-100 rounded-xl mb-4 gap-3'):
                ui.label('Agregar Nuevo Colaborador:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider')
                with ui.row().classes('w-full items-center gap-4 flex-wrap no-wrap md:no-wrap'):
                    name_input = ui.input(placeholder='Nombre del colaborador...').classes('flex-grow text-xs').props('outlined dense')
                    admin_checkbox = ui.checkbox('¿Es Administrador?').props('dense')
                    ui.button('Añadir Usuario', on_click=lambda: add_admin_user(name_input.value, admin_checkbox.value)).classes('bg-blue-900 text-white font-semibold rounded-lg px-4 py-2 text-xs shadow-sm')
            
            with ui.card().classes('w-full p-4 bg-white shadow-sm border border-slate-100 rounded-xl gap-2'):
                ui.label('Lista de Usuarios y Roles:').classes('text-xs font-bold text-slate-400 uppercase tracking-wider border-b pb-1 mb-2')
                for u in all_db_users:
                    u_name = u['name']
                    is_admin = u['is_admin']
                    role_lbl = "Administrador" if is_admin else "Colaborador Normal"
                    role_color = 'bg-blue-50 text-blue-800' if is_admin else 'bg-slate-100 text-slate-600'
                    
                    with ui.row().classes('w-full justify-between items-center p-3 border border-slate-100 rounded-lg hover:bg-slate-50/50 transition-colors gap-2'):
                        with ui.row().classes('items-center gap-3'):
                            ui.label('👤').classes('text-lg')
                            with ui.column().classes('gap-0'):
                                ui.label(u_name).classes('text-xs font-bold text-slate-800')
                                ui.label(role_lbl).classes(f'text-[9px] font-black px-1.5 py-0.5 rounded {role_color}')
                        
                        with ui.row().classes('items-center gap-2'):
                            ui.button('Cambiar Rol', on_click=lambda _, unm=u_name: run_toggle_admin(unm)).props('flat dense size=sm color=primary').classes('text-xs')
                            del_btn = ui.button('Eliminar', on_click=lambda _, unm=u_name: confirm_delete_user(unm)).props('flat dense size=sm color=negative').classes('text-xs')
                            if u_name == state['current_user']:
                                del_btn.disable()

            # --- SECTION 3: SYSTEM CENTRAL CONFIG ---
            ui.separator().classes('my-4')
            ui.label('⚙️ Configuración del Sistema').classes('text-lg font-bold text-slate-800 mb-2')
            
            cfg = load_email_config()
            with ui.card().classes('w-full p-6 bg-white shadow-md border border-slate-100 rounded-xl gap-4'):
                ui.label('⚙️ Configuración de IA (Gemini)').classes('text-xs font-bold text-slate-400 uppercase tracking-wider border-b pb-1')
                with ui.column().classes('w-full gap-1'):
                    ui.label('Clave API de Gemini:').classes('text-xs font-bold text-slate-500')
                    sys_gemini_input = ui.input(value=cfg.get('gemini_api_key', ''), password=True, password_toggle_button=True).classes('w-full max-w-lg text-xs').props('outlined dense')
                    
                ui.label('📧 Configuración de Correo Electrónico (SMTP)').classes('text-xs font-bold text-slate-400 uppercase tracking-wider border-b pb-1 mt-4')
                with ui.grid().classes('w-full grid-cols-1 md:grid-cols-2 gap-4'):
                    with ui.column().classes('gap-1'):
                        ui.label('Servidor SMTP:').classes('text-xs font-bold text-slate-500')
                        sys_smtp_input = ui.input(value=cfg.get('smtp_server', 'smtp.gmail.com')).classes('w-full text-xs').props('outlined dense')
                    with ui.column().classes('gap-1'):
                        ui.label('Puerto SMTP:').classes('text-xs font-bold text-slate-500')
                        sys_port_input = ui.input(value=cfg.get('smtp_port', '587')).classes('w-full text-xs').props('outlined dense')
                    with ui.column().classes('gap-1'):
                        ui.label('Correo Emisor:').classes('text-xs font-bold text-slate-500')
                        sys_sender_input = ui.input(value=cfg.get('sender_email', '')).classes('w-full text-xs').props('outlined dense')
                    with ui.column().classes('gap-1'):
                        ui.label('Contraseña de Aplicación:').classes('text-xs font-bold text-slate-500')
                        sys_pass_input = ui.input(value=cfg.get('sender_password', ''), password=True, password_toggle_button=True).classes('w-full text-xs').props('outlined dense')
                    with ui.column().classes('gap-1'):
                        ui.label('Correo Destinatario:').classes('text-xs font-bold text-slate-500')
                        sys_recipient_input = ui.input(value=cfg.get('recipient_email', '')).classes('w-full text-xs').props('outlined dense')
                    with ui.column().classes('gap-1 justify-center'):
                        sys_auto_send = ui.checkbox('Activar recordatorio automático (8:00 AM)', value=cfg.get('auto_send_enabled', False)).props('dense')
                
                ui.button('💾 Guardar Configuración Centralizada', on_click=lambda: save_central_settings(
                    sys_gemini_input.value, sys_smtp_input.value, sys_port_input.value, sys_sender_input.value, sys_pass_input.value, sys_recipient_input.value, sys_auto_send.value
                )).classes('w-full md:w-auto bg-blue-900 text-white font-bold rounded-lg px-6 py-2.5 shadow hover:bg-blue-800 transition-colors mt-4')

    def refresh_admin():
        content_area.clear()
        with content_area:
            ui.label('👑 Panel de Administración').classes('text-2xl font-bold text-slate-800')
            ui.label('Administre colaboradores, analice estadísticas de uso y modifique credenciales del sistema.').classes('text-slate-500 italic text-sm mt-1 mb-6')
            
            nonlocal admin_content
            admin_content = ui.column().classes('w-full gap-4')
            refresh_admin_dashboard()

    @ui.refreshable
    def refresh_sidebar_navigation():
        nav_container.clear()
        with nav_container:
            nav_items = [
                ("🏠 Inicio", "🏠 Inicio"),
                ("📋 Pizarra", "📋 Pizarra"),
                ("📊 Informes", "📊 Informes"),
                ("📥 Exportar", "📥 Exportar"),
                ("📅 Calendario", "📅 Calendario")
            ]
            if state['admin_mode']:
                nav_items.append(("👑 Admin", "👑 Admin"))
                
            for display_name, target_selection in nav_items:
                btn = ui.button(display_name, on_click=lambda ts=target_selection: select_nav(ts)).props('flat').classes('w-full text-left justify-start py-2 px-4 rounded-xl font-medium transition-all')
                nav_buttons[target_selection] = btn
            
            refresh_navigation_styles()
            


            ui.separator().classes('my-4')
            if state['admin_mode']:
                ui.label('🟢 Modo Admin Activo').classes('text-xs font-bold text-green-600 text-center w-full bg-green-50 py-1.5 rounded-lg border border-green-200 shadow-sm mb-2')
                ui.button('🔒 Desactivar Modo', on_click=deactivate_admin_mode).classes('w-full bg-slate-600 text-white font-bold rounded-xl py-2 shadow hover:bg-slate-700 transition-colors text-xs')
            else:
                ui.button('👑 Modo Admin', on_click=show_admin_login_dialog).classes('w-full bg-blue-900 text-white font-bold rounded-xl py-2 shadow hover:bg-blue-800 transition-colors text-xs')

    def refresh_layout():
        refresh_navigation_styles()
        if state['nav_selection'] == '🏠 Inicio':
            refresh_dashboard()
        elif state['nav_selection'] == '📋 Pizarra':
            refresh_pizarra()
        elif state['nav_selection'] == '📊 Informes':
            refresh_informes()
        elif state['nav_selection'] == '📥 Exportar':
            refresh_exportar()
        elif state['nav_selection'] == '📅 Calendario':
            refresh_calendario()
        elif state['nav_selection'] == '👑 Admin':
            if not state['admin_mode']:
                state['nav_selection'] = '🏠 Inicio'
                refresh_layout()
            else:
                refresh_admin()
        else:
            content_area.clear()
            with content_area:
                with ui.card().classes('w-full p-12 items-center justify-center text-center bg-white shadow-sm border border-slate-100 rounded-2xl mt-12 gap-4'):
                    ui.label('🚀 Próximamente').classes('text-3xl font-black text-blue-900')
                    ui.label(f'La pestaña "{state["nav_selection"]}" será migrada en la Fase 2 del desarrollo.').classes('text-slate-500 text-sm')
                    ui.button('Volver a la Pizarra', on_click=lambda: select_nav('📋 Pizarra')).classes('bg-blue-900 text-white font-semibold rounded-lg px-4 py-2 shadow-sm')
                    
    # Initial page render load
    refresh_sidebar_navigation()
    refresh_layout()

# NiceGUI run configurations
if __name__ in {'__main__', '__mp_main__'}:
    ui.run(port=8080, title="CUADROpz", reload=True, storage_secret="cuadropz_secret_key_2026")
