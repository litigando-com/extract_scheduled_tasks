import csv
import sys

try:
    import win32com.client
except ImportError:
    print("Error: Necesitas instalar pywin32.")
    print("Ejecuta: pip install pywin32")
    sys.exit(1)


def get_root_tasks():
    """
    Obtiene las tareas del folder raíz (/) del Task Scheduler
    usando la API COM de Windows directamente.
    """
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()

    root_folder = scheduler.GetFolder("\\")
    tasks = root_folder.GetTasks(0)  # 0 = incluir tareas ocultas también

    task_list = []

    for task in tasks:
        nombre = task.Name
        estado = normalize_status(task.State)
        triggers = get_triggers(task)
        acciones = get_actions(task)

        task_list.append({
            "NombreTarea": nombre,
            "Estado": estado,
            "Desencadenador": triggers,
            "Accion": acciones,
        })

    return task_list


def normalize_status(state_code):
    """
    Normaliza el código de estado numérico del Task Scheduler.
    1 = Deshabilitado
    2 = En cola
    3 = Listo (Habilitado)
    4 = En ejecución
    """
    status_map = {
        0: "Desconocido",
        1: "Deshabilitado",
        2: "Habilitado",   # En cola
        3: "Habilitado",   # Listo
        4: "Habilitado",   # En ejecución
    }
    return status_map.get(state_code, "Desconocido")


def get_triggers(task):
    """Obtiene los desencadenadores de una tarea."""
    try:
        definition = task.Definition
        triggers = definition.Triggers
        trigger_list = []

        # Tipos de trigger del Task Scheduler
        trigger_types = {
            0: "Evento",
            1: "Hora específica",
            2: "Diario",
            3: "Semanal",
            4: "Mensual",
            5: "Mensual (día de semana)",
            6: "Al estar inactivo",
            7: "Al registrar tarea",
            8: "Al iniciar el sistema",
            9: "Al iniciar sesión",
            11: "Evento personalizado",
        }

        for i in range(1, triggers.Count + 1):
            trigger = triggers.Item(i)
            tipo = trigger_types.get(trigger.Type, f"Tipo {trigger.Type}")

            # Intentar obtener más detalle
            detail = ""
            try:
                if trigger.StartBoundary:
                    detail = f" ({trigger.StartBoundary})"
            except Exception:
                pass

            trigger_list.append(f"{tipo}{detail}")

        return " | ".join(trigger_list) if trigger_list else "Sin desencadenador"

    except Exception:
        return "No disponible"


def get_actions(task):
    """
    Obtiene las acciones de una tarea.
    Formato: "Iniciar un programa: ruta/programa argumentos"
    """
    try:
        definition = task.Definition
        actions = definition.Actions
        action_list = []

        # Tipos de acción del Task Scheduler
        action_types = {
            0: "Iniciar un programa",
            5: "Ejecutar handler COM",
            6: "Enviar correo",
            7: "Mostrar mensaje",
        }

        for i in range(1, actions.Count + 1):
            action = actions.Item(i)
            tipo = action_types.get(action.Type, f"Tipo {action.Type}")

            if action.Type == 0:  # Exec action
                programa = action.Path or ""
                argumentos = action.Arguments or ""

                if argumentos:
                    action_list.append(f"{tipo}: {programa} {argumentos}")
                else:
                    action_list.append(f"{tipo}: {programa}")
            else:
                action_list.append(tipo)

        return " | ".join(action_list) if action_list else "Sin acción"

    except Exception:
        return "No disponible"


def export_to_csv(tasks, output_file="tareas_programadas.csv"):
    """Exporta las tareas a un archivo CSV con delimitador |."""
    fieldnames = ["NombreTarea", "Estado", "Desencadenador", "Accion"]

    with open(output_file, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter="|")
        writer.writeheader()
        for task in tasks:
            writer.writerow(task)

    print(f"Tareas en raíz (/): {len(tasks)}")
    print(f"Exportadas a '{output_file}'")


def main():
    print("Obteniendo tareas programadas de Windows (solo folder /)...\n")

    tasks = get_root_tasks()

    if not tasks:
        print("No se encontraron tareas en el folder raíz.")
        return

    # Preview en consola
    print(f"{'Nombre':<40} {'Estado':<15} {'Acción'}")
    print("-" * 100)
    for t in tasks:
        print(f"{t['NombreTarea']:<40} {t['Estado']:<15} {t['Accion']}")

    print()
    export_to_csv(tasks)
    print("¡Listo!")


if __name__ == "__main__":
    main()
