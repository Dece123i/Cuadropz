import os
import re
from openai import OpenAI

def generate_alternatives(api_key, unresolved_tasks):
    """
    Generates alternative solutions for a list of unresolved tasks.
    If api_key is not provided or empty, or if the OpenAI call fails, 
    it falls back to a smart, rule-based suggestion engine.
    
    unresolved_tasks: List of dicts, each with 'description', 'time_info', etc.
    Returns: List of string suggestions corresponding to each task in the list.
    """
    if not unresolved_tasks:
        return []
        
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
            # We will batch the requests or make them one by one. 
            # Doing it in a single prompt is much faster and saves API costs/tokens.
            tasks_str = "\n".join([f"{i+1}. Tarea: {task['description']} (Horario: {task.get('time_info') or 'No especificado'})" 
                                   for i, task in enumerate(unresolved_tasks)])
            
            prompt = (
                "Eres un experto senior en gestión de proyectos, finanzas y administración de empresas.\n"
                "A continuación tienes una lista de tareas de la empresa que no pudieron cumplirse hoy.\n"
                "Para cada tarea, propón una alternativa de solución detallada y en tono de experto (máximo 3-4 líneas) que cumpla estrictamente con:\n"
                "1. Ofrecer 2 o 3 pasos concretos y accionables para resolver la tarea pendiente.\n"
                "2. Incluir una breve recomendación de prioridad justificada (ej. 'Prioridad alta, ya que...').\n"
                "Devuelve el resultado en formato JSON estructurado como un arreglo de strings, donde cada elemento corresponda en orden a la sugerencia para cada tarea de la lista.\n"
                "Ejemplo de formato de respuesta:\n"
                '[\n  "Para resolver esta cobranza, recomiendo: 1) Enviar un correo formal con copia a gerencia para escalar el caso. 2) Programar una llamada con el cliente para negociar un plan de pagos. Prioridad alta, ya que el cobro pendiente afecta la liquidez del mes."\n]\n\n'
                "No agregues explicaciones adicionales, ni introducciones, ni bloques de código fuera del JSON. Responde únicamente con el JSON válido.\n\n"
                f"Lista de tareas no resueltas:\n{tasks_str}"
            )
            
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Eres un consultor empresarial senior que responde únicamente en formato JSON (arreglo de strings)."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=1000
            )
            
            content = response.choices[0].message.content.strip()
            # Clean possible markdown block formatting
            if content.startswith("```json"):
                content = content[7:]
            if content.endswith("```"):
                content = content[:-3]
            content = content.strip()
            
            import json
            parsed = json.loads(content)
            if isinstance(parsed, list) and len(parsed) == len(unresolved_tasks):
                return [str(item) for item in parsed]
            else:
                print("OpenAI response size mismatch or invalid structure, falling back to rule-based engine.")
        except Exception as e:
            print("OpenAI API execution error:", e, "- falling back to rule-based engine.")
            
    # Fallback Rule-Based Engine
    for task in unresolved_tasks:
        desc = task['description'].upper()
        
        if any(x in desc for x in ["COBRANZA", "COBRAR", "PAGO", "FACTURA", "EMITIR", "DEVOLUCION"]):
            suggestions.append("Para resolver esta cobranza, recomiendo: 1) Enviar un correo formal con copia a gerencia para escalar el caso. 2) Programar una llamada con el cliente para negociar un plan de pagos. Prioridad alta, ya que el cobro pendiente afecta directamente la liquidez del mes.")
        elif any(x in desc for x in ["BALANCE", "CONTABILIDAD", "BALANCE", "CONTABLE", "SUNAT", "GASTOS", "CIERRE"]):
            suggestions.append("Para resolver este balance contable, recomiendo: 1) Clasificar los comprobantes pendientes de registro. 2) Conciliar las cuentas bancarias con el libro auxiliar. Prioridad alta, dado que previene multas tributarias ante la SUNAT.")
        elif any(x in desc for x in ["REUNION", "COORDINAR", "PLANIFICAR", "ACUERDOS"]):
            suggestions.append("Para organizar esta reunión, recomiendo: 1) Redactar una agenda clara con los puntos clave. 2) Coordinar la disponibilidad horaria por correo electrónico. Prioridad media, necesaria para alinear los entregables del equipo de trabajo.")
        elif any(x in desc for x in ["CONTRATO", "LEGAL", "FIRMA", "CARTA", "NOTARIAL"]):
            suggestions.append("Para avanzar este contrato, recomiendo: 1) Redactar los términos principales usando una plantilla legal estandarizada. 2) Agendar una revisión express con el asesor jurídico. Prioridad alta, ya que mitiga riesgos contractuales importantes.")
        elif any(x in desc for x in ["INVENTARIO", "KARDEX", "STOCK"]):
            suggestions.append("Para conciliar el Kardex, recomiendo: 1) Realizar un conteo físico rápido de los productos críticos. 2) Actualizar las existencias en el sistema ERP. Prioridad media, importante para asegurar el abastecimiento del área.")
        elif any(x in desc for x in ["MARKETING", "PUBLICIDAD", "REDES"]):
            suggestions.append("Para la campaña publicitaria, recomiendo: 1) Calendarizar las publicaciones semanales en la plataforma de redes. 2) Analizar las métricas del periodo anterior. Prioridad media, ayuda a sostener el posicionamiento de marca en canales digitales.")
        else:
            suggestions.append("Para completar esta actividad, recomiendo: 1) Dividir la tarea en 2 sub-actividades manejables. 2) Reservar un bloque enfocado de 45 minutos sin distracciones. Prioridad media, para retomar el ritmo de entrega semanal.")
            
    return suggestions

def generate_single_alternative(api_key, task_desc, previous_suggestion=None):
    """
    Generates a single alternative solution for a task, requesting a different approach 
    or perspective than the previous suggestion.
    """
    if not api_key or not str(api_key).strip():
        return None
        
    try:
        client = OpenAI(api_key=api_key.strip())
        
        prompt = (
            "Eres un experto senior en gestión de proyectos, finanzas y administración de empresas.\n"
            f"A continuación tienes una tarea que no pudo cumplirse hoy:\n"
            f"Tarea: {task_desc}\n\n"
        )
        if previous_suggestion:
            prompt += (
                f"La sugerencia anterior fue:\n\"{previous_suggestion}\"\n\n"
                "IMPORTANTE: Genera una alternativa de solución que tenga un ENFOQUE o PERSPECTIVA DIFERENTE a la sugerencia anterior. "
                "Por ejemplo, si la primera sugerencia fue enviar correo y llamar, la segunda podría ser escalar a gerencia y revisar contrato. "
                "Ofrece 2 o 3 pasos concretos y accionables diferentes.\n"
            )
        else:
            prompt += "Ofrece 2 o 3 pasos concretos y accionables para resolver la tarea pendiente.\n"
            
        prompt += (
            "La respuesta debe ser una sugerencia detallada y en tono de experto (máximo 3-4 líneas), que incluya una breve recomendación de prioridad justificada.\n"
            "Devuelve ÚNICAMENTE la sugerencia como texto plano, sin formato JSON, sin bloques de código, sin introducciones ni explicaciones adicionales."
        )
        
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "Eres un consultor empresarial senior que responde únicamente con el texto directo de la sugerencia."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=300
        )
        
        suggestion = response.choices[0].message.content.strip()
        return suggestion
    except Exception as e:
        print("Error generating single alternative:", e)
        return None

