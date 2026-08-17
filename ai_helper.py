import os
import re
import json
from openai import OpenAI

def generate_alternatives(api_key, unresolved_tasks, team_members=None):
    """
    Generates alternative solutions, priorities, and reassignments for a list of unresolved tasks.
    If api_key is not provided or empty, or if the OpenAI call fails, 
    it falls back to a smart, rule-based suggestion engine.
    
    unresolved_tasks: List of dicts, each with 'description', 'time_info', etc.
    team_members: List of names of available team members.
    Returns: List of dicts, where each dict has 'prioridad', 'reasignacion_sugerida', 'alternativa_solucion'.
    """
    if not unresolved_tasks:
        return []
        
    members_str = ", ".join(team_members) if team_members else "MARY CRUZ, CPC.SHEYLA, CPC.HECTOR"
    suggestions = []
    
    # Check if we can use OpenAI
    use_openai = False
    client = None
    if api_key and str(api_key).strip():
        try:
            client = OpenAI(api_key=api_key.strip())
            use_openai = True
        except Exception as e:
            print("Error initializing OpenAI client:", e)
            
    if use_openai and client:
        try:
            tasks_str = "\n".join([f"{i+1}. Tarea: {task['description']} (Horario: {task.get('time_info') or 'No especificado'})" 
                                   for i, task in enumerate(unresolved_tasks)])
            
            prompt = (
                "Eres un experto senior en gestión de proyectos y administración de empresas.\n"
                "A continuación tienes una lista de tareas corporativas que no se pudieron completar hoy.\n"
                "Para cada tarea, debes sugerir una solución, una prioridad y determinar si es conveniente reasignarla a otro miembro del equipo.\n"
                f"Los miembros disponibles del equipo son: [{members_str}].\n\n"
                "Devuelve la respuesta en formato JSON que consista estrictamente en una lista de objetos, donde cada objeto represente una tarea en orden y contenga exactamente estas claves:\n"
                "1. \"prioridad\": Prioridad recomendada (debe ser uno de estos valores exactos: 'Alta', 'Media', 'Baja').\n"
                "2. \"reasignacion_sugerida\": Nombre de otro miembro del equipo al que se sugiere reasignar si corresponde (elige el más calificado de la lista de miembros, o déjalo vacío \"\" si no es necesario reasignar o si el usuario actual ya es el adecuado).\n"
                "3. \"alternativa_solucion\": 2 o 3 pasos concretos y accionables para resolver la tarea pendiente (máximo 3-4 líneas).\n\n"
                "Ejemplo de formato de respuesta:\n"
                "[\n"
                "  {\n"
                "    \"prioridad\": \"Alta\",\n"
                "    \"reasignacion_sugerida\": \"CPC.HECTOR\",\n"
                "    \"alternativa_solucion\": \"Para resolver la conciliación pendiente: 1) Clasificar los comprobantes de cobro pendientes. 2) Comparar los saldos auxiliares contra el extracto bancario corporativo.\"\n"
                "  }\n"
                "]\n\n"
                "No agregues ninguna introducción, ni explicaciones adicionales, ni bloques de código fuera del JSON. Responde únicamente con el JSON válido.\n\n"
                f"Lista de tareas no resueltas:\n{tasks_str}"
            )
            
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Eres un consultor empresarial senior que responde únicamente en formato JSON (arreglo de objetos)."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=1500
            )
            
            content = response.choices[0].message.content.strip()
            if content.startswith("```json"):
                content = content[7:]
            if content.endswith("```"):
                content = content[:-3]
            content = content.strip()
            
            parsed = json.loads(content)
            if isinstance(parsed, list) and len(parsed) == len(unresolved_tasks):
                cleaned = []
                for item in parsed:
                    if isinstance(item, dict):
                        cleaned.append({
                            "prioridad": item.get("prioridad", "Media"),
                            "reasignacion_sugerida": item.get("reasignacion_sugerida", ""),
                            "alternativa_solucion": item.get("alternativa_solucion", "Reagendar para mañana.")
                        })
                    else:
                        cleaned.append({
                            "prioridad": "Media",
                            "reasignacion_sugerida": "",
                            "alternativa_solucion": str(item)
                        })
                return cleaned
            else:
                print("OpenAI response size mismatch or invalid structure, falling back to rule-based engine.")
        except Exception as e:
            print("OpenAI API execution error:", e, "- falling back to rule-based engine.")
            
    # Fallback Rule-Based Engine
    for task in unresolved_tasks:
        desc = task['description'].upper()
        
        prio = "Media"
        reasign = ""
        sol = ""
        
        if any(x in desc for x in ["COBRANZA", "COBRAR", "PAGO", "FACTURA", "EMITIR", "DEVOLUCION"]):
            prio = "Alta"
            reasign = "CPC.SHEYLA"
            sol = "Para resolver esta cobranza, recomiendo: 1) Enviar un correo formal con copia a gerencia para escalar el caso. 2) Programar una llamada con el cliente para negociar un plan de pagos."
        elif any(x in desc for x in ["BALANCE", "CONTABILIDAD", "BALANCE", "CONTABLE", "SUNAT", "GASTOS", "CIERRE"]):
            prio = "Alta"
            reasign = "CPC.HECTOR"
            sol = "Para resolver este balance contable, recomiendo: 1) Clasificar los comprobantes pendientes de registro. 2) Conciliar las cuentas bancarias con el libro auxiliar."
        elif any(x in desc for x in ["REUNION", "COORDINAR", "PLANIFICAR", "ACUERDOS"]):
            prio = "Media"
            sol = "Para organizar esta reunión, recomiendo: 1) Redactar una agenda clara con los puntos clave. 2) Coordinar la disponibilidad horaria por correo electrónico."
        elif any(x in desc for x in ["CONTRATO", "LEGAL", "FIRMA", "CARTA", "NOTARIAL"]):
            prio = "Alta"
            reasign = "CPC.HECTOR"
            sol = "Para avanzar este contrato, recomiendo: 1) Redactar los términos principales usando una plantilla legal estandarizada. 2) Agendar una revisión express con el asesor jurídico."
        elif any(x in desc for x in ["INVENTARIO", "KARDEX", "STOCK"]):
            prio = "Media"
            sol = "Para conciliar el Kardex, recomiendo: 1) Realizar un conteo físico rápido de los productos críticos. 2) Actualizar las existencias en el sistema ERP."
        elif any(x in desc for x in ["MARKETING", "PUBLICIDAD", "REDES"]):
            prio = "Media"
            reasign = "MARY CRUZ"
            sol = "Para la campaña publicitaria, recomiendo: 1) Calendarizar las publicaciones semanales en la plataforma de redes. 2) Analizar las métricas del periodo anterior."
        else:
            sol = "Para completar esta actividad, recomiendo: 1) Dividir la tarea en 2 sub-actividades manejables. 2) Reservar un bloque enfocado de 45 minutos sin distracciones."
            
        suggestions.append({
            "prioridad": prio,
            "reasignacion_sugerida": reasign,
            "alternativa_solucion": sol
        })
            
    return suggestions

def generate_single_alternative(api_key, task_desc, previous_suggestion=None, team_members=None):
    """
    Generates a single alternative solution for a task returning a JSON object.
    """
    members_str = ", ".join(team_members) if team_members else "MARY CRUZ, CPC.SHEYLA, CPC.HECTOR"
    
    if not api_key or not str(api_key).strip():
        return {
            "prioridad": "Media",
            "reasignacion_sugerida": "",
            "alternativa_solucion": "Para completar esta actividad, recomiendo: 1) Dividir la tarea en 2 sub-actividades. 2) Dedicar 45 minutos de trabajo concentrado."
        }
        
    try:
        client = OpenAI(api_key=api_key.strip())
        
        prompt = (
            "Eres un experto senior en gestión de proyectos y administración de empresas.\n"
            f"A continuación tienes una tarea que no se pudo completar hoy:\n"
            f"Tarea: {task_desc}\n\n"
        )
        if previous_suggestion:
            prev_text = previous_suggestion.get("alternativa_solucion", str(previous_suggestion)) if isinstance(previous_suggestion, dict) else str(previous_suggestion)
            prompt += (
                f"La sugerencia de solución anterior fue:\n\"{prev_text}\"\n\n"
                "IMPORTANTE: Genera una propuesta alternativa que tenga un enfoque o perspectiva diferente a la anterior.\n"
            )
            
        prompt += (
            f"Los miembros disponibles del equipo son: [{members_str}].\n\n"
            "Responde únicamente en formato JSON con un objeto que tenga exactamente las siguientes claves:\n"
            "1. \"prioridad\": 'Alta', 'Media' o 'Baja'.\n"
            "2. \"reasignacion_sugerida\": Nombre de otro miembro del equipo si corresponde, o vacío \"\".\n"
            "3. \"alternativa_solucion\": La propuesta de solución detallada (máximo 3-4 líneas).\n\n"
            "No agregues bloques de código, introducciones ni explicaciones adicionales fuera del JSON."
        )
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "Eres un consultor empresarial senior que responde únicamente en formato JSON (un solo objeto)."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=500
        )
        
        content = response.choices[0].message.content.strip()
        if content.startswith("```json"):
            content = content[7:]
        if content.endswith("```"):
            content = content[:-3]
        content = content.strip()
        
        item = json.loads(content)
        if isinstance(item, dict):
            return {
                "prioridad": item.get("prioridad", "Media"),
                "reasignacion_sugerida": item.get("reasignacion_sugerida", ""),
                "alternativa_solucion": item.get("alternativa_solucion", "Reagendar para mañana.")
            }
    except Exception as e:
        print("Error generating single alternative:", e)
        
    return {
        "prioridad": "Media",
        "reasignacion_sugerida": "",
        "alternativa_solucion": "Para completar esta actividad, recomiendo redactar una minuta y revaluar el plan de entrega."
    }
