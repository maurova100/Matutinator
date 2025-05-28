import tkinter as tk
from tkinter import messagebox
import datetime
import os
import sys
import webbrowser

class RoutineAssistant:
    def __init__(self, root):
        self.root = root
        self.root.title("Asistente de Rutina Matutina")
        self.root.geometry("600x450")
        
        # Configuración inicial
        self.setup_symbols()
        self.setup_tasks()
        self.setup_autostart()
        
        # Variables de estado
        self.completed = [False] * len(self.tasks)
        self.current_task = 0
        
        # Interfaz
        self.create_widgets()
        self.show_task()
    
    def setup_symbols(self):
        """Configura símbolos compatibles"""
        try:
            "✓✗".encode(sys.stdout.encoding)
            self.check_symbol = "✓"
            self.cross_symbol = "✗"
        except (UnicodeEncodeError, AttributeError):
            self.check_symbol = "[X]"
            self.cross_symbol = "[ ]"
    
    def setup_tasks(self):
        """Define las tareas y sus acciones asociadas"""
        self.tasks = [
            "Chequear saldo YPF en Ruta y cargas observadas",
            "Chequear cronograma de tareas CRER",
            "Pagar Facturas y FR CUT del día",
            "Revisar BandejaCO e imprmimir extracto",
            "Revisar Outlook e imprimir extracto",
            "Revisar WhatsApp",
            "Actualizar cuardeno y organizar el día"
        ]
        
        self.task_actions = {
            "Revisar Outlook e imprimir extracto": {
                "type": "program",
                "path": "outlook"  # Cambia esto según tu configuración
            },
            "Chequear saldo YPF en Ruta y cargas observadas": {
                "type": "url",
                "path": "https://procesos.inta.gob.ar/portalprocesos#"  # Tu URL real aquí
            },
            "Chequear cronograma de tareas CRER": {
                "type": "url",
                "path": "https://docs.google.com/spreadsheets/d/1K9NmslL0RhEpz-td7BPfWnFMGPIoH3hxtRbTEIaW3hQ/edit?gid=915351991#gid=915351991"  # Tu URL real aquí
            },
            "Revisar BandejaCO e imprmimir extracto": {
                "type": "url",
                "path": "https://euc.gde.gob.ar/ccoo-web/"  # Tu URL real aquí
            },
            "Pagar Facturas y FR CUT del día": {
                "type": "program",
                "path": "C:\Program Files (x86)\EsigaIntaDesktop\EsigaIntaDesktop.exe"  # esiga
            },
            "Revisar WhatsApp": {
                "type": "url",
                "path": "https://web.whatsapp.com/"  # esiga
            }            
        }
    
    def setup_autostart(self):
        """Configura el inicio automático"""
        try:
            startup_folder = os.path.join(
                os.getenv('APPDATA'),
                'Microsoft', 'Windows', 'Start Menu',
                'Programs', 'Startup'
            )
            bat_path = os.path.join(startup_folder, 'rutina_matutina.bat')
            
            if not os.path.exists(bat_path):
                with open(bat_path, 'w') as f:
                    script_path = os.path.abspath(__file__)
                    f.write(f'start "" /min pythonw.exe "{script_path}"')
        except Exception as e:
            print("Advertencia: No se pudo configurar el inicio automático:", e)
    
    def create_widgets(self):
        """Crea la interfaz de usuario"""
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Cabecera
        tk.Label(
            main_frame,
            text="Asistente de Rutina Matutina",
            font=('Arial', 16, 'bold'),
            fg='#2E7D32'
        ).pack(pady=10)
        
        # Contador de tareas
        self.counter_label = tk.Label(
            main_frame,
            text="",
            font=('Arial', 10)
        )
        self.counter_label.pack()
        
        # Tarea actual
        self.task_label = tk.Label(
            main_frame,
            text="",
            font=('Arial', 12),
            wraplength=500,
            justify=tk.LEFT
        )
        self.task_label.pack(pady=20)
        
        # Botón de acción (para programas/URLs)
        self.action_btn = tk.Button(
            main_frame,
            text="Abrir recurso relacionado",
            command=self.open_resource,
            bg='#FF9800',
            fg='white',
            font=('Arial', 10),
            state=tk.DISABLED
        )
        self.action_btn.pack(pady=5)
        
        # Botón de completado
        self.complete_btn = tk.Button(
            main_frame,
            text="Marcar como completado",
            command=self.complete_task,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 10)
        )
        self.complete_btn.pack(pady=10)
        
        # Botón de reporte
        self.report_btn = tk.Button(
            main_frame,
            text="Generar Reporte",
            command=self.generate_report,
            bg='#2196F3',
            fg='white',
            font=('Arial', 10)
        )
        self.report_btn.pack(pady=10)
    
    def show_task(self):
        """Muestra la tarea actual"""
        if self.current_task < len(self.tasks):
            task = self.tasks[self.current_task]
            self.counter_label.config(
                text=f"Tarea {self.current_task + 1} de {len(self.tasks)}")
            self.task_label.config(text=task)
            
            # Habilitar botón de acción si hay una acción definida
            if task in self.task_actions:
                self.action_btn.config(state=tk.NORMAL)
            else:
                self.action_btn.config(state=tk.DISABLED)
        else:
            self.task_label.config(text="¡Todas las tareas completadas!")
            self.complete_btn.config(state=tk.DISABLED)
            self.action_btn.config(state=tk.DISABLED)
            self.counter_label.config(text="")
    
    def complete_task(self):
        """Marca la tarea como completada"""
        self.completed[self.current_task] = True
        self.current_task += 1
        self.show_task()
    
    def open_resource(self):
        """Abre el programa o URL asociada a la tarea actual"""
        task = self.tasks[self.current_task]
        if task in self.task_actions:
            action = self.task_actions[task]
            try:
                if action["type"] == "program":
                    os.startfile(action["path"])
                elif action["type"] == "url":
                    webbrowser.open(action["path"])
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el recurso:\n{e}")
    
    def generate_report(self):
        """Genera un reporte en formato Markdown"""
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        time_now = datetime.datetime.now().strftime("%H:%M")
        
        report = f"# Reporte de tareas - {today}\n\n"
        report += f"**Generado a las:** {time_now}\n\n"
        report += "## Resumen\n\n"
        
        completed_count = sum(self.completed)
        progress = int((completed_count / len(self.tasks)) * 100)
        report += f"- **Tareas completadas:** {completed_count}/{len(self.tasks)}\n"
        report += f"- **Progreso:** {progress}%\n\n"
        
        report += "## Detalle de tareas\n\n"
        for i, task in enumerate(self.tasks):
            status = self.check_symbol if self.completed[i] else self.cross_symbol
            report += f"- {status} {task}\n"
        
        filename = f"Reporte_Tareas_{today.replace('-', '')}.md"
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(report)
            
            if messagebox.askyesno(
                "Reporte Generado",
                f"Reporte guardado en:\n{os.path.abspath(filename)}\n\n¿Abrir archivo?"):
                os.startfile(filename)
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el reporte:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RoutineAssistant(root)
    root.mainloop()