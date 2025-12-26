import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
import subprocess
import platform
import threading
import time
import re
import csv
from datetime import datetime, timedelta
import sys
import os

# Try to import multiping
try:
    from multiping import multi_ping
    HAS_MULTIPING = True
except ImportError:
    HAS_MULTIPING = False

# Import openpyxl for Excel support
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# --- VALIDATION UTILS ---

def clean_address(address):
    """Очистка и проверка адреса хоста."""
    address = address.strip()
    if not address:
        raise ValueError("Адрес не может быть пустым.")

    ipv4_pattern = r"^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    ipv6_pattern = r"^(?:[0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}$"
    domain_pattern = r"^[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$"

    if not (re.match(ipv4_pattern, address) or re.match(ipv6_pattern, address) or re.match(domain_pattern, address)):
        raise ValueError(f"Некорректный адрес: {address}. Должен быть IPv4, IPv6 или доменным именем.")
    
    return address

def clean_group_name(group_name):
    """Очистка имени группы или подгруппы."""
    group_name = group_name.strip()
    if not group_name:
        raise ValueError("Имя группы не может быть пустым.")
    if not re.match(r"^[a-zA-Z0-9_\sа-яА-ЯәңғұүқөһіӘҢҒҰҮҚӨҺІ\-\.\(\)]+$", group_name):
        raise ValueError("Некорректное имя группы. Используйте буквы, цифры и пробелы.")
    return group_name

# --- DATABASE MANAGER ---

class DatabaseManager:
    def __init__(self, db_name="monitoring_gui.db"):
        self.conn = sqlite3.connect(db_name, check_same_thread=False)
        self.cursor = self.conn.cursor()
        self.init_db()

    def init_db(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS groups (
                group_name TEXT PRIMARY KEY
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS ping_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                group_name TEXT,
                subgroup TEXT,
                address TEXT,
                status TEXT,
                latency REAL
            )
        """)
        self.conn.commit()

    def create_group(self, group_name):
        group_name = clean_group_name(group_name)
        self.cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS hosts_{group_name} (
                address TEXT PRIMARY KEY,
                description TEXT,
                subgroup TEXT,
                ping_interval INTEGER DEFAULT 600,
                last_ping_time TEXT,
                offline_since TEXT
            )
        """)
        self.cursor.execute("INSERT OR IGNORE INTO groups (group_name) VALUES (?)", (group_name,))
        self.conn.commit()
        return group_name

    def delete_group(self, group_name):
        group_name = clean_group_name(group_name)
        self.cursor.execute(f"DROP TABLE IF EXISTS hosts_{group_name}")
        self.cursor.execute("DELETE FROM groups WHERE group_name = ?", (group_name,))
        self.cursor.execute("DELETE FROM ping_results WHERE group_name = ?", (group_name,))
        self.conn.commit()

    def get_groups(self):
        self.cursor.execute("SELECT group_name FROM groups")
        return [row[0] for row in self.cursor.fetchall()]

    def add_host(self, group_name, address, description, subgroup=None, ping_interval=600):
        group_name = clean_group_name(group_name)
        address = clean_address(address)
        if subgroup:
            subgroup = clean_group_name(subgroup)
        self.cursor.execute(f"""
            INSERT OR REPLACE INTO hosts_{group_name} 
            (address, description, subgroup, ping_interval, last_ping_time, offline_since) 
            VALUES (?, ?, ?, ?, NULL, NULL)
        """, (address, description, subgroup, ping_interval))
        self.conn.commit()

    def remove_host(self, group_name, address):
        group_name = clean_group_name(group_name)
        address = clean_address(address)
        self.cursor.execute(f"DELETE FROM hosts_{group_name} WHERE address = ?", (address,))
        self.conn.commit()

    def get_hosts(self, group_name):
        try:
            group_name = clean_group_name(group_name)
            self.cursor.execute(f"""
                SELECT address, description, subgroup, ping_interval, last_ping_time, offline_since 
                FROM hosts_{group_name}
            """)
            return self.cursor.fetchall()
        except sqlite3.OperationalError:
            return []

    def update_ping_time(self, group_name, address, status):
        """Обновить last_ping_time и offline_since"""
        group_name = clean_group_name(group_name)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if status == "Online":
            self.cursor.execute(f"""
                UPDATE hosts_{group_name} 
                SET last_ping_time = ?, offline_since = NULL 
                WHERE address = ?
            """, (now, address))
        else:  # Offline
            # Если уже offline, не менять offline_since. Если был online, установить сейчас
            self.cursor.execute(f"""
                SELECT offline_since FROM hosts_{group_name} WHERE address = ?
            """, (address,))
            row = self.cursor.fetchone()
            if row and row[0] is None:
                # Был online, переходит offline
                self.cursor.execute(f"""
                    UPDATE hosts_{group_name} 
                    SET last_ping_time = ?, offline_since = ? 
                    WHERE address = ?
                """, (now, now, address))
            else:
                # Уже offline
                self.cursor.execute(f"""
                    UPDATE hosts_{group_name} 
                    SET last_ping_time = ? 
                    WHERE address = ?
                """, (now, address))
        self.conn.commit()
            
    def get_all_hosts_stats(self):
        """Возвращает статистику по всем группам с их статусами."""
        groups = self.get_groups()
        stats = []
        
        for group in groups:
            hosts = self.get_hosts(group)
            total = len(hosts)
            
            online_count = 0
            offline_count = 0
            critical_offline = False  # есть offline >= 1 часа
            warning_offline = False   # есть offline < 1 часа
            
            for host in hosts:
                address = host[0]
                offline_since = host[5]  # последний элемент в кортеже хоста
                
                self.cursor.execute("""
                    SELECT status FROM ping_results 
                    WHERE group_name = ? AND address = ? 
                    ORDER BY id DESC LIMIT 1
                """, (group, address))
                row = self.cursor.fetchone()
                if row:
                    if row[0] == "Online":
                        online_count += 1
                    elif row[0] == "Offline":
                        offline_count += 1
                        # Проверить, критично ли offline
                        if offline_since:
                            try:
                                offline_time = datetime.strptime(offline_since, "%Y-%m-%d %H:%M:%S")
                                now = datetime.now()
                                diff = (now - offline_time).total_seconds()
                                if diff >= 3600:  # >= 1 часа
                                    critical_offline = True
                                else:
                                    warning_offline = True
                            except:
                                critical_offline = True
                        else:
                            warning_offline = True
            
            # Определить статус группы
            if critical_offline:
                group_status = "critical"  # красный
            elif warning_offline or (offline_count > 0 and online_count > 0):
                group_status = "warning"   # жёлтый
            elif online_count == total or (total == 0):
                group_status = "healthy"   # зелёный
            else:
                group_status = "unknown"
            
            stats.append({
                "group": group,
                "total": total,
                "online": online_count,
                "offline": offline_count,
                "unknown": total - online_count - offline_count,
                "status": group_status
            })
        return stats
    
    def get_overall_status(self):
        """Получить общий статус всей системы."""
        stats = self.get_all_hosts_stats()
        
        critical_count = sum(1 for s in stats if s["status"] == "critical")
        warning_count = sum(1 for s in stats if s["status"] == "warning")
        
        if critical_count > 0:
            return "critical"  # красный
        elif warning_count > 0:
            return "warning"   # жёлтый
        else:
            return "healthy"   # зелёный

    def log_result(self, group_name, subgroup, address, status, latency=None):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.cursor.execute(
            "INSERT INTO ping_results (timestamp, group_name, subgroup, address, status, latency) VALUES (?, ?, ?, ?, ?, ?)",
            (timestamp, group_name, subgroup, address, status, latency)
        )
        self.update_ping_time(group_name, address, status)
        self.conn.commit()

    def close(self):
        self.conn.close()

# --- PING ENGINE ---

def ping_host_subprocess(address, timeout=1.0):
    """Fallback: Пингует хост, используя системную утилиту ping."""
    param = '-n' if platform.system().lower() == 'windows' else '-c'
    timeout_param = '-w' if platform.system().lower() == 'windows' else '-W'
    
    timeout_val = str(int(timeout * 1000)) if platform.system().lower() == 'windows' else str(int(timeout))
    if platform.system().lower() != 'windows':
         timeout_val = str(max(1, int(timeout)))

    command = ['ping', param, '1', timeout_param, timeout_val, address]

    try:
        startupinfo = None
        if platform.system().lower() == 'windows':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
        start_time = time.time()
        result = subprocess.run(
            command, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE,
            startupinfo=startupinfo,
            text=True
        )
        end_time = time.time()
        
        duration = end_time - start_time
        
        if result.returncode == 0:
            return True, duration
        else:
            return False, None
    except Exception:
        return False, None

# --- GUI APPLICATION ---

class PingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ping Monitor")
        self.geometry("1200x650")
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure(bg="#f0f0f0")

        self.db = DatabaseManager()
        self.monitoring = False
        self.monitor_thread = None
        self.monitor_interval = 2.0 
        self.current_group = None
        self.host_status_cache = {}  # Кэш статусов для определения цвета индикатора

        self.create_widgets()
        self.load_groups()
        self.update_overall_status()
        
        if not HAS_OPENPYXL:
            messagebox.showwarning("Внимание", "openpyxl не установлен. Экспорт/импорт могут не работать.")
        if not HAS_MULTIPING:
            messagebox.showinfo("Info", "Библиотека multiping не найдена. Используется системный ping (медленнее).")

    def create_widgets(self):
        # Create Tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Tab 1: Monitor
        self.tab_monitor = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_monitor, text="Мониторинг")
        
        # Tab 2: Statistics
        self.tab_stats = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stats, text="Статистика")
        
        # --- TAB 1 CONTENT ---
        self.create_monitor_widgets(self.tab_monitor)
        
        # --- TAB 2 CONTENT ---
        self.create_stats_widgets(self.tab_stats)
        
        # --- Overall Status Bar ---
        overall_frame = ttk.Frame(self)
        overall_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=3)
        
        ttk.Label(overall_frame, text="Общий статус:").pack(side=tk.LEFT, padx=5)
        self.overall_status_label = ttk.Label(overall_frame, text="●", font=("Arial", 14))
        self.overall_status_label.pack(side=tk.LEFT, padx=3)
        
        ttk.Separator(overall_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # --- Status Bar (Shared) ---
        self.status_var = tk.StringVar(value="Готов")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def create_monitor_widgets(self, parent):
        # --- Toolbar ---
        toolbar = ttk.Frame(parent)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        ttk.Label(toolbar, text="Группа:").pack(side=tk.LEFT, padx=5)
        self.group_combo = ttk.Combobox(toolbar, state="readonly", width=20)
        self.group_combo.pack(side=tk.LEFT, padx=5)
        self.group_combo.bind("<<ComboboxSelected>>", self.on_group_select)

        ttk.Button(toolbar, text="Создать", command=self.create_group_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Удалить", command=self.delete_group).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(toolbar, text="Добавить хост", command=self.add_host_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Удалить хост", command=self.remove_host).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        ttk.Button(toolbar, text="Экспорт (Excel)", command=self.export_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Импорт (Excel)", command=self.import_data).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        self.btn_start = ttk.Button(toolbar, text="Старт", command=self.toggle_monitoring)
        self.btn_start.pack(side=tk.LEFT, padx=5)
        
        # --- Treeview with custom columns ---
        columns = ("status_indicator", "address", "description", "subgroup", "ping_interval", "latency", "last_check")
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="extended", height=20)
        
        self.tree.heading("status_indicator", text="")
        self.tree.heading("address", text="Адрес")
        self.tree.heading("description", text="Описание")
        self.tree.heading("subgroup", text="Подгруппа")
        self.tree.heading("ping_interval", text="Интервал (мин)")
        self.tree.heading("latency", text="Задержка (сек)")
        self.tree.heading("last_check", text="Последняя проверка")

        self.tree.column("status_indicator", width=30)
        self.tree.column("address", width=140)
        self.tree.column("description", width=180)
        self.tree.column("subgroup", width=90)
        self.tree.column("ping_interval", width=80)
        self.tree.column("latency", width=80)
        self.tree.column("last_check", width=130)

        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure tags for status colors
        self.tree.tag_configure("status_green", foreground="#2ecc71")
        self.tree.tag_configure("status_yellow", foreground="#f39c12")
        self.tree.tag_configure("status_red", foreground="#e74c3c")
        self.tree.tag_configure("status_gray", foreground="#95a5a6")

    def create_stats_widgets(self, parent):
        toolbar = ttk.Frame(parent)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="Обновить статистику", command=self.refresh_stats).pack(side=tk.LEFT, padx=5)
        
        columns = ("status", "group", "total", "online", "offline", "unknown")
        self.stats_tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        
        self.stats_tree.heading("status", text="")
        self.stats_tree.heading("group", text="Группа")
        self.stats_tree.heading("total", text="Всего хостов")
        self.stats_tree.heading("online", text="Доступно")
        self.stats_tree.heading("offline", text="Недоступно")
        self.stats_tree.heading("unknown", text="Неизвестно")
        
        self.stats_tree.column("status", width=25)
        self.stats_tree.column("group", width=150)
        self.stats_tree.column("total", width=100)
        self.stats_tree.column("online", width=80)
        self.stats_tree.column("offline", width=80)
        self.stats_tree.column("unknown", width=80)
        
        # Configure tags for status colors
        self.stats_tree.tag_configure("status_healthy", foreground="#2ecc71")
        self.stats_tree.tag_configure("status_warning", foreground="#f39c12")
        self.stats_tree.tag_configure("status_critical", foreground="#e74c3c")
        
        self.stats_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

    def on_tab_change(self, event):
        selected_tab = self.notebook.select()
        tab_text = self.notebook.tab(selected_tab, "text")
        if tab_text == "Статистика":
            self.refresh_stats()

    def refresh_stats(self):
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)
            
        stats = self.db.get_all_hosts_stats()
        for s in stats:
            tag = f"status_{s['status']}"
            self.stats_tree.insert("", "end", values=(
                "●", s["group"], s["total"], s["online"], s["offline"], s["unknown"]
            ), tags=(tag,))
        
        self.update_overall_status()
    
    def update_overall_status(self):
        """Обновить отображение общего статуса."""
        overall = self.db.get_overall_status()
        color_map = {
            "healthy": "#2ecc71",   # зелёный
            "warning": "#f39c12",   # жёлтый
            "critical": "#e74c3c"   # красный
        }
        
        self.overall_status_label.config(foreground=color_map.get(overall, "#95a5a6"))
        
        status_text_map = {
            "healthy": "Все системы в норме",
            "warning": "Есть проблемы (< 1ч)",
            "critical": "Критичные проблемы (>= 1ч)"
        }
        
        if self.monitoring:
            # Не перезаписываем статус если идёт мониторинг
            return

    # --- Group Actions ---
    
    def load_groups(self):
        groups = self.db.get_groups()
        self.group_combo['values'] = groups
        if groups:
            if self.current_group and self.current_group in groups:
                self.group_combo.set(self.current_group)
            else:
                self.group_combo.current(0)
                self.on_group_select(None)
        else:
            self.group_combo.set('')
            self.clear_table()

    def on_group_select(self, event):
        self.current_group = self.group_combo.get()
        self.refresh_table()

    def create_group_dialog(self):
        name = tk.simpledialog.askstring("Новая группа", "Введите имя группы:")
        if name:
            try:
                self.db.create_group(name)
                self.load_groups()
                self.group_combo.set(name)
                self.on_group_select(None)
            except ValueError as e:
                messagebox.showerror("Ошибка", str(e))

    def delete_group(self):
        if not self.current_group:
            return
        if messagebox.askyesno("Подтверждение", f"Удалить группу '{self.current_group}'?"):
            self.db.delete_group(self.current_group)
            self.current_group = None
            self.load_groups()

    # --- Host Actions ---

    def add_host_dialog(self):
        if not self.current_group:
            messagebox.showwarning("Внимание", "Сначала создайте или выберите группу.")
            return
            
        dialog = tk.Toplevel(self)
        dialog.title("Добавить хост")
        dialog.geometry("350x320")
        
        ttk.Label(dialog, text="Адрес (IP/Домен):").pack(pady=5)
        addr_entry = ttk.Entry(dialog, width=30)
        addr_entry.pack(pady=5)
        
        ttk.Label(dialog, text="Описание:").pack(pady=5)
        desc_entry = ttk.Entry(dialog, width=30)
        desc_entry.pack(pady=5)
        
        ttk.Label(dialog, text="Подгруппа (необязательно):").pack(pady=5)
        sub_entry = ttk.Entry(dialog, width=30)
        sub_entry.pack(pady=5)

        ttk.Label(dialog, text="Интервал пинга (минуты, стандарт 10):").pack(pady=5)
        interval_var = tk.StringVar(value="10")
        interval_entry = ttk.Entry(dialog, width=30, textvariable=interval_var)
        interval_entry.pack(pady=5)

        def save():
            try:
                interval_minutes = int(interval_var.get())
                interval_seconds = interval_minutes * 60
                self.db.add_host(
                    self.current_group,
                    addr_entry.get(),
                    desc_entry.get(),
                    sub_entry.get() or None,
                    interval_seconds
                )
                self.refresh_table()
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("Ошибка", str(e))
            except sqlite3.Error as e:
                messagebox.showerror("Ошибка БД", str(e))

        ttk.Button(dialog, text="Сохранить", command=save).pack(pady=10)

    def remove_host(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        
        for item_id in selected_item:
            item = self.tree.item(item_id)
            address = item['values'][1]  # address is at index 1 (after status_indicator)
            self.db.remove_host(self.current_group, address)
        
        self.refresh_table()

    def export_data(self):
        if not self.current_group:
            return
        
        if not HAS_OPENPYXL:
            messagebox.showerror("Ошибка", "openpyxl не установлен. Невозможно экспортировать.")
            return
        
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        
        if filename:
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "hosts"
                
                ws.append(["address", "description", "subgroup", "ping_interval_min"])
                hosts = self.db.get_hosts(self.current_group)
                for h in hosts:
                    # h[3] is ping_interval in seconds, convert to minutes
                    interval_min = h[3] // 60 if h[3] else 10
                    ws.append([h[0], h[1], h[2] or "", interval_min])
                
                # Auto-adjust column widths
                for col in ws.columns:
                    max_length = 0
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    ws.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width
                
                wb.save(filename)
                messagebox.showinfo("Успех", "Данные экспортированы в Excel.")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))

    def import_data(self):
        if not self.current_group:
            messagebox.showwarning("Внимание", "Выберите группу.")
            return
        
        if not HAS_OPENPYXL:
            messagebox.showerror("Ошибка", "openpyxl не установлен. Невозможно импортировать.")
            return
        
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        
        if filename:
            try:
                wb = load_workbook(filename)
                ws = wb.active
                
                rows = list(ws.iter_rows(values_only=True))
                if not rows:
                    return
                
                for row in rows[1:]:
                    if row and len(row) >= 2:
                        addr = row[0]
                        desc = row[1]
                        sub = row[2] if len(row) > 2 else None
                        interval_min = int(row[3]) if len(row) > 3 and row[3] else 10
                        interval_sec = interval_min * 60
                        if addr:
                            self.db.add_host(self.current_group, addr, desc, sub, interval_sec)
                
                self.refresh_table()
                messagebox.showinfo("Успех", "Данные импортированы из Excel.")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))

    def get_status_color(self, address):
        """Определить цвет индикатора на основе статуса и времени."""
        if address not in self.host_status_cache:
            return "gray"  # unknown
        
        status, offline_since = self.host_status_cache[address]
        
        if status == "Online":
            return "green"
        elif status == "Offline":
            if offline_since:
                try:
                    offline_time = datetime.strptime(offline_since, "%Y-%m-%d %H:%M:%S")
                    now = datetime.now()
                    diff = (now - offline_time).total_seconds()
                    if diff < 3600:  # < 1 часа
                        return "yellow"
                    else:
                        return "red"
                except:
                    return "red"
            else:
                return "red"
        else:
            return "gray"

    def refresh_table(self):
        self.clear_table()
        if not self.current_group:
            return
        
        hosts = self.db.get_hosts(self.current_group)
        for host in hosts:
            address, desc, subgroup, interval_sec, last_ping_time, offline_since = host
            
            # Get last status from cache or DB
            self.db.cursor.execute("""
                SELECT status FROM ping_results 
                WHERE group_name = ? AND address = ? 
                ORDER BY id DESC LIMIT 1
            """, (self.current_group, address))
            row = self.db.cursor.fetchone()
            status = row[0] if row else "Unknown"
            
            # Store in cache
            self.host_status_cache[address] = (status, offline_since)
            
            # Convert interval to minutes for display
            interval_min = interval_sec // 60 if interval_sec else 10
            
            # Get color and set tag
            color_name = self.get_status_color(address)
            tag = f"status_{color_name}"
            
            # Insert with placeholder for status indicator
            self.tree.insert("", "end", values=("●", address, desc, subgroup or "", interval_min, "-", ""), tags=(tag,))

    def clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    # --- Monitoring ---

    def toggle_monitoring(self):
        if self.monitoring:
            self.monitoring = False
            self.btn_start.configure(text="Старт")
            self.status_var.set("Мониторинг остановлен")
        else:
            if not self.current_group:
                messagebox.showwarning("Внимание", "Выберите группу.")
                return
            self.monitoring = True
            self.btn_start.configure(text="Стоп")
            self.status_var.set("Мониторинг запущен...")
            self.monitor_thread = threading.Thread(target=self.monitoring_loop, daemon=True)
            self.monitor_thread.start()

    def monitoring_loop(self):
        while self.monitoring:
            if not self.current_group:
                break
            
            hosts = self.db.get_hosts(self.current_group) 
            if not hosts:
                time.sleep(1)
                continue

            addresses = [h[0] for h in hosts]

            if HAS_MULTIPING:
                try:
                    responses, no_responses = multi_ping(addresses, timeout=1, retry=1, ignore_lookup_errors=True)
                    
                    timestamp = datetime.now().strftime("%H:%M:%S")

                    for addr, latency in responses.items():
                        if not self.monitoring: break
                        self.db.log_result(self.current_group, None, addr, "Online", latency)
                        self.update_row(addr, "Online", f"{latency:.3f}", timestamp)

                    for addr in no_responses:
                        if not self.monitoring: break
                        self.db.log_result(self.current_group, None, addr, "Offline", None)
                        self.update_row(addr, "Offline", "-", timestamp)

                except Exception as e:
                    print(f"Multiping error: {e}")
            else:
                import concurrent.futures
                with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                    future_to_addr = {executor.submit(ping_host_subprocess, h[0]): h for h in hosts}
                    for future in concurrent.futures.as_completed(future_to_addr):
                        if not self.monitoring: break
                        host_data = future_to_addr[future]
                        address = host_data[0]
                        try:
                            success, latency = future.result()
                            status = "Online" if success else "Offline"
                            latency_str = f"{latency:.3f}" if latency else "-"
                            timestamp = datetime.now().strftime("%H:%M:%S")
                            self.db.log_result(self.current_group, None, address, status, latency)
                            self.update_row(address, status, latency_str, timestamp)
                        except Exception:
                            pass

            if self.monitoring:
                try:
                    selected = self.notebook.select()
                    if self.notebook.tab(selected, "text") == "Статистика":
                         self.refresh_stats()
                    # Периодически обновляем общий статус
                    self.update_overall_status()
                except:
                    pass
                time.sleep(self.monitor_interval)

    def update_row(self, address, status, latency, timestamp):
        try:
            for item_id in self.tree.get_children():
                vals = self.tree.item(item_id)['values']
                if vals and str(vals[1]) == str(address):  # address is at index 1
                    # Update cache with new status
                    offline_since = self.host_status_cache.get(address, ("Unknown", None))[1]
                    if status == "Offline" and offline_since is None:
                        offline_since = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    elif status == "Online":
                        offline_since = None
                    
                    self.host_status_cache[address] = (status, offline_since)
                    
                    # Get color based on new status
                    color = self.get_status_color(address)
                    tag = f"status_{color}"
                    
                    # Create colored indicator
                    indicator = "●"
                    
                    new_vals = (indicator, vals[1], vals[2], vals[3], vals[4], latency, timestamp)
                    self.tree.item(item_id, values=new_vals, tags=(tag,))
                    break
        except Exception:
            pass

if __name__ == "__main__":
    app = PingApp()
    app.mainloop()
