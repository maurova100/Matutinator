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
        self.setup_holidays()
        self.setup_tasks()
        self.setup_autostart()
        
        # Variables de estado
        self.completed = [False] * len(self.filtered_tasks)
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
    
    def setup_holidays(self):
        """Configura días feriados (personaliza esta lista)"""
        today = datetime.datetime.now()
        year = today.year
        self.holidays = [
            datetime.date(year, 1, 1),    # Año Nuevo
            datetime.date(year, 3, 24),   # Día Nacional de la Memoria
            datetime.date(year, 4, 2),    # Día del Veterano
            datetime.date(year, 5, 1),    # Día del Trabajador
            datetime.date(year, 5, 25),   # Día de la Revolución de Mayo
            datetime.date(year, 6, 20),   # Día de la Bandera
            datetime.date(year, 7, 9),    # Día de la Independencia
            datetime.date(year, 12, 8),   # Inmaculada Concepción
            datetime.date(year, 12, 25)   # Navidad
        ]
    
    def is_working_day(self, date):
        """Determina si es día laboral (no fin de semana ni feriado)"""
        if date.weekday() >= 5:  # Sábado (5) o Domingo (6)
            return False
        if date.date() in self.holidays:
            return False
        return True
    
    def next_working_day(self, date):
        """Encuentra el próximo día laboral"""
        next_day = date + datetime.timedelta(days=1)
        while not self.is_working_day(next_day):
            next_day += datetime.timedelta(days=1)
        return next_day
    
    def setup_tasks(self):
        """Define las tareas y sus acciones asociadas"""
        today = datetime.datetime.now()
        day = today.day
        is_quincenal_day = day in (1, 16) and self.is_working_day(today)
        is_monthly_day = day == 1 and self.is_working_day(today)
        
        # Todas las tareas posibles (según tus especificaciones)
        self.all_tasks = [
            # Tareas diarias
            {"task": "Chequear saldo YPF en Ruta y cargas observadas", "frequency": "daily"},
            {"task": "Pagar Facturas y FR CUT del día", "frequency": "daily"},
            {"task": "Revisar BandejaCO e imprimir extracto", "frequency": "daily"},
            {"task": "Revisar Outlook e imprimir extracto", "frequency": "daily"},
            {"task": "Revisar WhatsApp", "frequency": "daily"},
            {"task": "Actualizar cuaderno y organizar el día", "frequency": "daily"},
            
            # Tareas quincenales (solo días 1 y 16 laborales)
            {"task": "Enviar retenciones no CUT IVA-SUSS-Ganancias si corresponde", "frequency": "quincenal"},
            {"task": "Recuperar retenciones no CUT ATER si corresponde", "frequency": "quincenal"},
            {"task": "Chequear facturas de telefonía fija", "frequency": "quincenal"},
            {"task": "Organizar temario y reunión de VT y RI si corresponde", "frequency": "quincenal"},
            {"task": "Chequear cronograma de tareas CRER", "frequency": "quincenal"},
            
            # Tareas mensuales (solo día 1 laboral)
            {"task": "Gestionar pago de servicio de limpieza", "frequency": "monthly"},
            {"task": "Gestionar Pago de alquiler de dispensers", "frequency": "monthly"},
            {"task": "Gestionar EE de descuentos de cuota alimentaria", "frequency": "monthly"},
            {"task": "Iniciar NOTA de conciliación bancaria", "frequency": "monthly"},
            {"task": "Iniciar EE de gastos bancarios", "frequency": "monthly"},
            {"task": "Iniciar EE de pago de retenciones ATER", "frequency": "monthly"}
        ]
        
        # Filtrar tareas según la fecha
        self.filtered_tasks = [
            t["task"] for t in self.all_tasks 
            if (t["frequency"] == "daily") or 
               (t["frequency"] == "quincenal" and is_quincenal_day) or
               (t["frequency"] == "monthly" and is_monthly_day)
        ]
        
        # Mensaje especial si es feriado o fin de semana
        self.special_message = ""
        if not self.is_working_day(today):
            next_workday = self.next_working_day(today)
            self.special_message = (
                f"\n\nAVISO: Hoy es {'fin de semana' if today.weekday() >= 5 else 'feriado'}.\n"
                f"Las tareas quincenales/mensuales se mostrarán el próximo día laboral ({next_workday.strftime('%d/%m')})."
            )
        
        # Acciones asociadas a tareas (según tus especificaciones)
        self.task_actions = {
            "Chequear saldo YPF en Ruta y cargas observadas": {
                "type": "url",
                "path": "https://procesos.inta.gob.ar/portalprocesos#"
            },
            "Pagar Facturas y FR CUT del día": {
                "type": "program",
                "path": r"C:\Program Files (x86)\EsigaIntaDesktop\EsigaIntaDesktop.exe"
            },
            "Revisar BandejaCO e imprimir extracto": {
                "type": "url",
                "path": "https://euc.gde.gob.ar/ccoo-web/"
            },
            "Revisar Outlook e imprimir extracto": {
                "type": "program",
                "path": "outlook"
            },
            "Revisar WhatsApp": {
                "type": "url",
                "path": "https://web.whatsapp.com/"
            },            
            "Organizar temario y reunión de VT y RI si corresponde": {
                "type": "url",
                "path": "https://docs.google.com/forms/d/12oKrn41VZgBKBjSs_Mdgei6PfIWwGTncGPak1YnFv20/edit?pli=1#response=ACYDBNgoPnbtExu642AvnBuRSFWRvnX2LjtGh9E2tc1M8LwTwc-uW5IgP1ubeQVfwwBlq-o"
            },
            "Chequear cronograma de tareas CRER": {
                "type": "url",
                "path": "https://docs.google.com/spreadsheets/d/1K9NmslL0RhEpz-td7BPfWnFMGPIoH3hxtRbTEIaW3hQ/edit?gid=915351991#gid=915351991"
            },
            "Iniciar NOTA de conciliación bancaria": {
                "type": "url",
                "path": "https://bee3.redlink.com.ar/bna3/bee/auth/login"
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
        
        # Mostrar fecha actual
        today = datetime.datetime.now().strftime("%A, %d/%m/%Y")
        tk.Label(
            main_frame,
            text=f"Fecha: {today}",
            font=('Arial', 10, 'bold')
        ).pack()
        
        # Mostrar tipo de día
        day_type = "Día laboral" if self.is_working_day(datetime.datetime.now()) else "Fin de semana/Feriado"
        day_color = '#2E7D32' if self.is_working_day(datetime.datetime.now()) else '#D32F2F'
        tk.Label(
            main_frame,
            text=f"Tipo de día: {day_type}",
            font=('Arial', 10),
            fg=day_color
        ).pack()
        
        # Mensaje especial
        if self.special_message:
            tk.Label(
                main_frame,
                text=self.special_message,
                font=('Arial', 9),
                fg='#D32F2F',
                justify=tk.LEFT
            ).pack(pady=5)
        
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
        self.task_label.pack(pady=15)
        
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
        
        # Botón de reporte simplificado
        self.report_btn = tk.Button(
            main_frame,
            text="Guardar Reporte",
            command=self.generate_report,
            bg='#2196F3',
            fg='white',
            font=('Arial', 10)
        )
        self.report_btn.pack(pady=10)
    
    def show_task(self):
        """Muestra la tarea actual"""
        if self.current_task < len(self.filtered_tasks):
            task = self.filtered_tasks[self.current_task]
            self.counter_label.config(
                text=f"Tarea {self.current_task + 1} de {len(self.filtered_tasks)}")
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
        task = self.filtered_tasks[self.current_task]
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
        """Genera un reporte simple en formato TXT"""
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        filename = f"Reporte_Tareas_{today.replace('-', '')}.txt"
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(f"Reporte de tareas - {today}\n")
                f.write("="*30 + "\n\n")
                
                for i, task in enumerate(self.filtered_tasks):
                    status = "COMPLETADA" if self.completed[i] else "PENDIENTE"
                    f.write(f"{i+1}. {task} - {status}\n")
                
                f.write("\n")
                completed_count = sum(self.completed)
                total_tasks = len(self.filtered_tasks)
                progress = int((completed_count / total_tasks) * 100) if total_tasks > 0 else 100
                f.write(f"Progreso: {progress}% ({completed_count}/{total_tasks} tareas)\n")
                
                if self.special_message:
                    f.write("\n" + self.special_message.replace('\n', ' ') + "\n")
            
            messagebox.showinfo(
                "Reporte Guardado",
                f"Se ha guardado el reporte en:\n{os.path.abspath(filename)}")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el reporte:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RoutineAssistant(root)
    root.mainloop()