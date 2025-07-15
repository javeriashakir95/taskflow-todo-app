
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import json
import os
import re
import platform
import subprocess
from openpyxl import Workbook
from datetime import datetime

# --- File Constants ---
TASKS_FILE = "tasks.json"
EXCEL_FILE = "tasks.xlsx"

# --- UI Constants ---
HEADER_HEIGHT = 70
SIDEBAR_WIDTH = 300

# --- UI Colors (Soft Pastel Professional Theme) ---
COLOR_GRADIENT_LIGHT = "#fdf2f8"       
COLOR_GRADIENT_MEDIUM = "#cce4f6"      
COLOR_SIDEBAR_GRADIENT_END = "#b2c7e6" 
COLOR_SIDEBAR_BTN_BG = "#b2c7e6"       
COLOR_BACKGROUND_LIGHT = "#fefce8"     
COLOR_TEXT_DARK = "#334155"            

class TaskFlowApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TaskFlow - To-Do List App")
        self.root.attributes("-fullscreen", True)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.tasks = []
        self.filtered_tasks = []  # For search results
        self.load_tasks()

        self.setup_ui()

    def setup_ui(self):
        header = tk.Canvas(self.root, height=HEADER_HEIGHT, highlightthickness=0)
        header.pack(fill='x')
        header.create_rectangle(0, 0, self.root.winfo_screenwidth(), HEADER_HEIGHT, fill=COLOR_GRADIENT_LIGHT, outline="")
        tk.Label(self.root, text="📋 TaskFlow", font=('Segoe UI', 24, 'bold'), bg=COLOR_GRADIENT_LIGHT, fg=COLOR_TEXT_DARK).place(x=20, y=10)

        main_app_frame = tk.Frame(self.root, bg=COLOR_BACKGROUND_LIGHT)
        main_app_frame.pack(fill='both', expand=True)

        sidebar = tk.Canvas(main_app_frame, width=SIDEBAR_WIDTH, highlightthickness=0, bg=COLOR_SIDEBAR_GRADIENT_END)
        sidebar.pack(side='left', fill='y')

        nav_frame = tk.Frame(sidebar, bg=COLOR_SIDEBAR_BTN_BG)
        sidebar.create_window((0, 0), window=nav_frame, anchor='nw', width=SIDEBAR_WIDTH)

        tk.Label(nav_frame, text="📌 Navigation", bg=COLOR_SIDEBAR_BTN_BG, font=("Helvetica", 18, "bold")).pack(pady=20)

        for text, command in [("Home", self.show_home), ("Calendar", self.show_calendar), ("Open Excel", self.open_excel_file), ("Exit", self.on_closing)]:
            ttk.Button(nav_frame, text=f"🏠 {text}" if text == "Home" else f"📅 {text}" if text == "Calendar" else f"📂 {text}" if text == "Open Excel" else f"🚪 {text}", command=command).pack(pady=10, fill="x", padx=30)

        self.main_frame = tk.Frame(main_app_frame, bg="white", padx=40, pady=20)
        self.main_frame.pack(side="right", expand=True, fill="both")

        self.show_home()

    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def show_home(self):
        self.clear_main_frame()
        tk.Label(self.main_frame, text="📝 Your Tasks", font=("Helvetica", 20, "bold"), bg="white").pack(pady=20)

        search_frame = tk.Frame(self.main_frame, bg="white")
        search_frame.pack(pady=10)
        tk.Label(search_frame, text="🔍 Search:", font=("Helvetica", 12), bg="white").pack(side="left", padx=(0, 5))
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.pack(side="left")
        ttk.Button(search_frame, text="Go", command=self.search_tasks).pack(side="left", padx=5)
        ttk.Button(search_frame, text="Clear", command=self.clear_search).pack(side="left")

        ttk.Button(self.main_frame, text="🔃 Sort by Date", command=self.sort_tasks).pack(pady=5)

        if not self.filtered_tasks:
            tk.Label(self.main_frame, text="No tasks yet! Schedule one using the Calendar button.", bg="white", font=("Helvetica", 12)).pack(pady=50)
            return

        canvas = tk.Canvas(self.main_frame, bg="white")
        scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="white")

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.task_vars = []
        for idx, task in enumerate(self.filtered_tasks):
            frame = tk.Frame(scroll_frame, bg="white")
            frame.pack(pady=10, anchor="w", padx=10, fill="x")

            var = tk.BooleanVar(value=task.get("done", False))
            self.task_vars.append(var)

            style = ("Helvetica", 13, "overstrike") if var.get() else ("Helvetica", 13)
            color = "gray" if var.get() else "black"

            tk.Checkbutton(frame, variable=var, bg="white", command=lambda i=idx: self.toggle_task(i)).pack(side="left")
            tk.Label(frame, text=f"{task['title']} (Due: {task['date']} {task.get('time', '')})", font=style, fg=color, bg="white").pack(side="left")
            tk.Button(frame, text="✏️ Edit", command=lambda i=idx: self.edit_task(i), bg="#6CA6B9", fg=COLOR_TEXT_DARK, font=("Helvetica", 11, "bold")).pack(side="right", padx=5)
            tk.Button(frame, text="🗑️ Delete", command=lambda i=idx: self.delete_task(i), bg="#EF8181", fg=COLOR_TEXT_DARK, font=("Helvetica", 11, "bold")).pack(side="right", padx=5)

    def search_tasks(self):
        keyword = self.search_entry.get().strip().lower()
        self.filtered_tasks = [task for task in self.tasks if keyword in task["title"].lower()]
        self.show_home()

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.filtered_tasks = self.tasks.copy()
        self.show_home()

    def show_calendar(self, edit_index=None):
        self.clear_main_frame()
        editing = edit_index is not None
        title = "✏️ Edit Task" if editing else "📅 Schedule a Task"
        tk.Label(self.main_frame, text=title, font=("Helvetica", 20, "bold"), bg="white").pack(pady=20)

        form = tk.Frame(self.main_frame, bg="white")
        form.pack(pady=30)

        tk.Label(form, text="Task Title:", bg="white", font=("Helvetica", 13)).grid(row=0, column=0, padx=15, pady=10, sticky="e")
        self.title_entry = ttk.Entry(form, width=35)
        self.title_entry.grid(row=0, column=1, padx=15, pady=10)

        tk.Label(form, text="Due Date:", bg="white", font=("Helvetica", 13)).grid(row=1, column=0, padx=15, pady=10, sticky="e")
        self.date_entry = DateEntry(form, width=33, background='darkblue', foreground='white', borderwidth=2)
        self.date_entry.grid(row=1, column=1, padx=15, pady=10)

        tk.Label(form, text="Due Time (HH:MM):", bg="white", font=("Helvetica", 13)).grid(row=2, column=0, padx=15, pady=10, sticky="e")
        self.time_entry = ttk.Entry(form, width=35)
        self.time_entry.grid(row=2, column=1, padx=15, pady=10)

        if editing:
            task = self.tasks[edit_index]
            self.title_entry.insert(0, task["title"])
            self.date_entry.set_date(datetime.strptime(task["date"], "%Y-%m-%d"))
            self.time_entry.insert(0, task.get("time", ""))

        tk.Button(self.main_frame, text="💾 Save Task", command=lambda: self.save_task(edit_index), bg="#A3D9A5", fg=COLOR_TEXT_DARK, font=("Helvetica", 13, "bold"), width=20).pack(pady=20)

    def toggle_task(self, index):
        self.filtered_tasks[index]["done"] = self.task_vars[index].get()
        for task in self.tasks:
            if task["title"] == self.filtered_tasks[index]["title"]:
                task["done"] = self.filtered_tasks[index]["done"]
                break
        self.save_tasks()
        self.update_excel_file()
        self.show_home()

    def delete_task(self, index):
        task_to_delete = self.filtered_tasks[index]
        if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this task?"):
            self.tasks.remove(task_to_delete)
            self.filtered_tasks.remove(task_to_delete)
            self.save_tasks()
            self.update_excel_file()
            self.show_home()

    def edit_task(self, index):
        actual_index = self.tasks.index(self.filtered_tasks[index])
        self.show_calendar(edit_index=actual_index)

    def save_task(self, edit_index=None):
        title = self.title_entry.get().strip()
        date = self.date_entry.get_date().strftime("%Y-%m-%d")
        time = self.time_entry.get().strip()

        if not title:
            messagebox.showwarning("Input Error", "Task title cannot be empty.")
            return
        if time and not re.match(r"^(?:[01]\d|2[0-3]):[0-5]\d$", time):
            messagebox.showwarning("Input Error", "Invalid time format. Use HH:MM (24-hour format).")
            return

        task = {"title": title, "date": date, "time": time, "done": False}

        if edit_index is not None:
            self.tasks[edit_index] = task
        else:
            self.tasks.append(task)

        self.save_tasks()
        self.update_excel_file()
        messagebox.showinfo("Success", "Task saved!")
        self.filtered_tasks = self.tasks.copy()
        self.show_home()

    def sort_tasks(self):
        try:
            self.tasks.sort(key=lambda x: datetime.strptime(f"{x['date']} {x.get('time', '00:00')}", "%Y-%m-%d %H:%M"))
            self.filtered_tasks = self.tasks.copy()
            self.save_tasks()
            self.show_home()
        except Exception as e:
            messagebox.showerror("Sorting Error", f"Could not sort tasks: {e}")

    def save_tasks(self):
        with open(TASKS_FILE, "w") as f:
            json.dump(self.tasks, f, indent=4)

    def load_tasks(self):
        if os.path.exists(TASKS_FILE):
            try:
                with open(TASKS_FILE, "r") as f:
                    self.tasks = json.load(f)
                    self.filtered_tasks = self.tasks.copy()
            except json.JSONDecodeError:
                messagebox.showwarning("Load Error", "JSON file is corrupted. Starting with empty task list.")
                self.tasks = []
                self.filtered_tasks = []

    def update_excel_file(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Tasks"
        ws.append(["Title", "Date", "Time", "Done"])
        for task in self.tasks:
            ws.append([task["title"], task["date"], task.get("time", ""), "Yes" if task.get("done") else "No"])
        wb.save(EXCEL_FILE)

    def open_excel_file(self):
        self.update_excel_file()
        if os.path.exists(EXCEL_FILE):
            try:
                if platform.system() == "Windows":
                    os.startfile(EXCEL_FILE)
                elif platform.system() == "Darwin":
                    subprocess.call(["open", EXCEL_FILE])
                else:
                    subprocess.call(["xdg-open", EXCEL_FILE])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open Excel file: {e}")
        else:
            messagebox.showinfo("Not Found", "No Excel file found yet. Save a task first.")

    def on_closing(self):
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = TaskFlowApp(root)
    root.mainloop()
