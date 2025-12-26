import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import subprocess
import platform
import threading
import time
import re
import queue
import concurrent.futures
from datetime import datetime

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

DEFAULT_PING_INTERVAL = 600
PING_BATCH_SIZE = 50
MAX_PING_WORKERS = 20
STATS_REFRESH_INTERVAL = 5


def clean_address(address):
    """Очистка и проверка адреса хоста."""
    address = address.strip()
    if not address:
        raise ValueError("Адрес не может быть пустым.")

    ipv4_pattern = r"^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    ipv6_pattern = r"^(?:[0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}$"
    domain_pattern = r"^[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$"

    if not (
        re.match(ipv4_pattern, address)
        or re.match(ipv6_pattern, address)
        or re.match(domain_pattern, address)
    ):
        raise ValueError(
            f"Некорректный адрес: {address}. Должен быть IPv4, IPv6 или доменным именем."
        )

    return address


def clean_group_name(group_name):
    """Очистка имени группы или подгруппы."""
    group_name = group_name.strip()
    if not group_name:
        raise ValueError("Имя группы не может быть пустым.")
    if not re.match(r"^[a-zA-Z0-9_\sа-яА-ЯәңғұүқөһіӘҢҒҰҮҚӨҺІ\-\.\(\)]+$", group_name):
        raise ValueError("Некорректное имя группы. Используйте буквы, цифры и пробелы.")
    return group_name


class DatabaseManager:
    def __init__(self, db_name="monitoring_gui.db"):
        self.conn = sqlite3.connect(db_name, check_same_thread=False)
        self.cursor = self.conn.cursor()
        self.lock = threading.RLock()
        self.init_db()
        self.migrate_legacy_tables()

    def init_db(self):
        with self.lock:
            self.cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS groups (
                    group_name TEXT PRIMARY KEY
                )
            """
            )
            self.cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS hosts (
                    group_name TEXT NOT NULL,
                    address TEXT NOT NULL,
                    description TEXT,
                    subgroup TEXT,
                    ping_interval INTEGER DEFAULT 600,
                    last_ping_time TEXT,
                    offline_since TEXT,
                    PRIMARY KEY (group_name, address)
                )
            """
            )
            self.cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS ping_results (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT,
                    group_name TEXT,
                    subgroup TEXT,
                    address TEXT,
                    status TEXT,
                    latency REAL
                )
            """
            )
            self.cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_hosts_group
                ON hosts (group_name)
            """
            )
            self.cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_ping_results_group_address
                ON ping_results (group_name, address)
            """
            )
            self.conn.commit()

    def migrate_legacy_tables(self):
        with self.lock:
            self.cursor.execute(
                """
                SELECT name FROM sqlite_master
                WHERE type='table' AND name LIKE 'hosts_%'
            """
            )
            legacy_tables = [row[0] for row in self.cursor.fetchall()]
            for table_name in legacy_tables:
                group_name = table_name[len("hosts_") :]
                if not group_name:
                    continue
                self.cursor.execute(
                    "INSERT OR IGNORE INTO groups (group_name) VALUES (?)",
                    (group_name,),
                )
                self.cursor.execute(
                    f"SELECT address, description, subgroup, ping_interval, last_ping_time, offline_since FROM {table_name}"
                )
                rows = self.cursor.fetchall()
                if rows:
                    self.cursor.executemany(
                        """
                        INSERT OR REPLACE INTO hosts
                        (group_name, address, description, subgroup, ping_interval, last_ping_time, offline_since)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                        [
                            (
                                group_name,
                                row[0],
                                row[1],
                                row[2],
                                row[3],
                                row[4],
                                row[5],
                            )
                            for row in rows
                        ],
                    )
                self.cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
            self.conn.commit()

    def create_group(self, group_name):
        group_name = clean_group_name(group_name)
        with self.lock:
            self.cursor.execute(
                "INSERT OR IGNORE INTO groups (group_name) VALUES (?)", (group_name,)
            )
            self.conn.commit()
        return group_name

    def delete_group(self, group_name):
        group_name = clean_group_name(group_name)
        with self.lock:
            self.cursor.execute("DELETE FROM groups WHERE group_name = ?", (group_name,))
            self.cursor.execute("DELETE FROM hosts WHERE group_name = ?", (group_name,))
            self.cursor.execute(
                "DELETE FROM ping_results WHERE group_name = ?", (group_name,)
            )
            self.conn.commit()

    def get_groups(self):
        with self.lock:
            self.cursor.execute("SELECT group_name FROM groups")
            return [row[0] for row in self.cursor.fetchall()]

    def add_host(self, group_name, address, description, subgroup=None, ping_interval=DEFAULT_PING_INTERVAL):
        group_name = clean_group_name(group_name)
        address = clean_address(address)
        if subgroup:
            subgroup = clean_group_name(subgroup)
        with self.lock:
            self.cursor.execute(
                """
                INSERT OR REPLACE INTO hosts
                (group_name, address, description, subgroup, ping_interval, last_ping_time, offline_since)
                VALUES (?, ?, ?, ?, ?, NULL, NULL)
            """,
                (group_name, address, description, subgroup, ping_interval),
            )
            self.conn.commit()

    def update_host(self, group_name, address, description, subgroup=None, ping_interval=DEFAULT_PING_INTERVAL):
        group_name = clean_group_name(group_name)
        address = clean_address(address)
        if subgroup:
            subgroup = clean_group_name(subgroup)
        with self.lock:
            self.cursor.execute(
                """
                UPDATE hosts
                SET description = ?, subgroup = ?, ping_interval = ?
                WHERE group_name = ? AND address = ?
            """,
                (description, subgroup, ping_interval, group_name, address),
            )
            self.conn.commit()

    def remove_host(self, group_name, address):
        group_name = clean_group_name(group_name)
        address = clean_address(address)
        with self.lock:
            self.cursor.execute(
                "DELETE FROM hosts WHERE group_name = ? AND address = ?",
                (group_name, address),
            )
            self.conn.commit()

    def get_hosts(self, group_name):
        group_name = clean_group_name(group_name)
        with self.lock:
            self.cursor.execute(
                """
                SELECT address, description, subgroup, ping_interval, last_ping_time, offline_since
                FROM hosts
                WHERE group_name = ?
            """,
                (group_name,),
            )
            return self.cursor.fetchall()

    def get_last_status(self, group_name, address):
        group_name = clean_group_name(group_name)
        address = clean_address(address)
        with self.lock:
            self.cursor.execute(
                """
                SELECT status FROM ping_results
                WHERE group_name = ? AND address = ?
                ORDER BY id DESC LIMIT 1
            """,
                (group_name, address),
            )
            row = self.cursor.fetchone()
            return row[0] if row else "Unknown"

    def get_recent_results(self, group_name, address, limit=50):
        group_name = clean_group_name(group_name)
        address = clean_address(address)
        with self.lock:
            self.cursor.execute(
                """
                SELECT timestamp, status, latency
                FROM ping_results
                WHERE group_name = ? AND address = ?
                ORDER BY id DESC
                LIMIT ?
            """,
                (group_name, address, limit),
            )
            return list(reversed(self.cursor.fetchall()))

    def update_ping_time(self, group_name, address, status, commit=True):
        group_name = clean_group_name(group_name)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.lock:
            if status == "Online":
                self.cursor.execute(
                    """
                    UPDATE hosts
                    SET last_ping_time = ?, offline_since = NULL
                    WHERE group_name = ? AND address = ?
                """,
                    (now, group_name, address),
                )
            else:
                self.cursor.execute(
                    """
                    SELECT offline_since FROM hosts
                    WHERE group_name = ? AND address = ?
                """,
                    (group_name, address),
                )
                row = self.cursor.fetchone()
                if row and row[0] is None:
                    self.cursor.execute(
                        """
                        UPDATE hosts
                        SET last_ping_time = ?, offline_since = ?
                        WHERE group_name = ? AND address = ?
                    """,
                        (now, now, group_name, address),
                    )
                else:
                    self.cursor.execute(
                        """
                        UPDATE hosts
                        SET last_ping_time = ?
                        WHERE group_name = ? AND address = ?
                    """,
                        (now, group_name, address),
                    )
            if commit:
                self.conn.commit()

    def log_results_batch(self, entries):
        if not entries:
            return
        with self.lock:
            self.cursor.executemany(
                """
                INSERT INTO ping_results
                (timestamp, group_name, subgroup, address, status, latency)
                VALUES (?, ?, ?, ?, ?, ?)
            """,
                [
                    (
                        entry["timestamp"],
                        entry["group_name"],
                        entry["subgroup"],
                        entry["address"],
                        entry["status"],
                        entry["latency"],
                    )
                    for entry in entries
                ],
            )
            for entry in entries:
                self.update_ping_time(
                    entry["group_name"], entry["address"], entry["status"], commit=False
                )
            self.conn.commit()

    def log_result(self, group_name, subgroup, address, status, latency=None):
        entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "group_name": group_name,
            "subgroup": subgroup,
            "address": address,
            "status": status,
            "latency": latency,
        }
        self.log_results_batch([entry])

    def get_all_hosts_stats(self):
        """Возвращает статистику по всем группам с их статусами."""
        with self.lock:
            self.cursor.execute(
                """
                SELECT h.group_name, h.address, h.offline_since, pr.status
                FROM hosts h
                LEFT JOIN (
                    SELECT pr.group_name, pr.address, pr.status
                    FROM ping_results pr
                    INNER JOIN (
                        SELECT group_name, address, MAX(id) AS max_id
                        FROM ping_results
                        GROUP BY group_name, address
                    ) latest
                    ON pr.id = latest.max_id
                ) pr
                ON pr.group_name = h.group_name AND pr.address = h.address
            """
            )
            rows = self.cursor.fetchall()
            self.cursor.execute("SELECT group_name FROM groups")
            all_groups = [row[0] for row in self.cursor.fetchall()]

        stats_by_group = {}
        for group, address, offline_since, status in rows:
            group_stats = stats_by_group.setdefault(
                group,
                {
                    "group": group,
                    "total": 0,
                    "online": 0,
                    "offline": 0,
                    "unknown": 0,
                    "status": "unknown",
                },
            )
            group_stats["total"] += 1

            if status == "Online":
                group_stats["online"] += 1
            elif status == "Offline":
                group_stats["offline"] += 1
            else:
                group_stats["unknown"] += 1

            if status == "Offline":
                if offline_since:
                    try:
                        offline_time = datetime.strptime(
                            offline_since, "%Y-%m-%d %H:%M:%S"
                        )
                        if (datetime.now() - offline_time).total_seconds() >= 3600:
                            group_stats["status"] = "critical"
                        elif group_stats["status"] != "critical":
                            group_stats["status"] = "warning"
                    except ValueError:
                        group_stats["status"] = "critical"
                elif group_stats["status"] != "critical":
                    group_stats["status"] = "warning"

        for group_stats in stats_by_group.values():
            if group_stats["status"] == "unknown":
                if group_stats["offline"] > 0:
                    group_stats["status"] = "warning"
                elif group_stats["total"] == 0:
                    group_stats["status"] = "healthy"
                else:
                    group_stats["status"] = "healthy"
        for group in all_groups:
            stats_by_group.setdefault(
                group,
                {
                    "group": group,
                    "total": 0,
                    "online": 0,
                    "offline": 0,
                    "unknown": 0,
                    "status": "healthy",
                },
            )
        return list(stats_by_group.values())

    def get_overall_status(self):
        stats = self.get_all_hosts_stats()
        critical_count = sum(1 for s in stats if s["status"] == "critical")
        warning_count = sum(1 for s in stats if s["status"] == "warning")

        if critical_count > 0:
            return "critical"
        if warning_count > 0:
            return "warning"
        return "healthy"

    def close(self):
        with self.lock:
            self.conn.close()


def ping_host_subprocess(address, timeout=1.0):
    """Пингует хост, используя системную утилиту ping."""
    param = "-n" if platform.system().lower() == "windows" else "-c"
    timeout_param = "-w" if platform.system().lower() == "windows" else "-W"

    timeout_val = str(int(timeout * 1000)) if platform.system().lower() == "windows" else str(
        int(timeout)
    )
    if platform.system().lower() != "windows":
        timeout_val = str(max(1, int(timeout)))

    command = ["ping", param, "1", timeout_param, timeout_val, address]

    try:
        startupinfo = None
        if platform.system().lower() == "windows":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        start_time = time.time()
        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            startupinfo=startupinfo,
            text=True,
        )
        end_time = time.time()

        duration = end_time - start_time

        if result.returncode == 0:
            return True, duration
        return False, None
    except Exception:
        return False, None


class PingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ping Monitor")
        self.geometry("1200x650")

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.configure(bg="#f0f0f0")
        self.apply_theme()

        self.db = DatabaseManager()
        self.monitoring = False
        self.monitor_thread = None
        self.monitor_interval = 1.0
        self.current_group = None
        self.host_status_cache = {}
        self.ui_queue = queue.Queue()
        self.last_stats_refresh = 0
        self.selected_host = None

        self.create_widgets()
        self.load_groups()
        self.update_overall_status()
        self.after(200, self.process_ui_queue)

        if not HAS_OPENPYXL:
            messagebox.showwarning(
                "Внимание",
                "openpyxl не установлен. Импорт/экспорт Excel будет недоступен.",
            )

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tab_monitor = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_monitor, text="Мониторинг")

        self.tab_stats = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stats, text="Статистика")

        self.create_monitor_widgets(self.tab_monitor)
        self.create_stats_widgets(self.tab_stats)

        overall_frame = ttk.Frame(self)
        overall_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=3)

        ttk.Label(overall_frame, text="Общий статус:").pack(side=tk.LEFT, padx=5)
        self.overall_status_label = ttk.Label(overall_frame, text="●", font=("Arial", 14))
        self.overall_status_label.pack(side=tk.LEFT, padx=3)

        ttk.Separator(overall_frame, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=5
        )

        self.status_var = tk.StringVar(value="Готов")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def apply_theme(self):
        self.style.configure("TFrame", background="#f7f9fc")
        self.style.configure("TLabel", background="#f7f9fc", foreground="#1f2937")
        self.style.configure("TButton", padding=6)
        self.style.configure("TNotebook", background="#f7f9fc")
        self.style.configure("TNotebook.Tab", padding=[12, 6])
        self.style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#1f2937",
            rowheight=26,
        )
        self.style.configure(
            "Treeview.Heading",
            background="#e5e7eb",
            foreground="#111827",
            font=("Segoe UI", 10, "bold"),
        )
        self.style.map("Treeview", background=[("selected", "#dbeafe")])
        self.style.configure("TSeparator", background="#e5e7eb")

    def create_monitor_widgets(self, parent):
        toolbar = ttk.Frame(parent)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        ttk.Label(toolbar, text="Группа:").pack(side=tk.LEFT, padx=5)
        self.group_combo = ttk.Combobox(toolbar, state="readonly", width=20)
        self.group_combo.pack(side=tk.LEFT, padx=5)
        self.group_combo.bind("<<ComboboxSelected>>", self.on_group_select)

        ttk.Button(toolbar, text="Создать", command=self.create_group_dialog).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(toolbar, text="Удалить", command=self.delete_group).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(toolbar, text="Добавить хост", command=self.add_host_dialog).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(toolbar, text="Редактировать хост", command=self.edit_host_dialog).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(toolbar, text="Удалить хост", command=self.remove_host).pack(
            side=tk.LEFT, padx=2
        )

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(toolbar, text="Экспорт (Excel)", command=self.export_data).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(toolbar, text="Импорт (Excel)", command=self.import_data).pack(
            side=tk.LEFT, padx=2
        )

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Label(toolbar, text="Поиск:").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind("<KeyRelease>", lambda event: self.refresh_table())
        ttk.Button(toolbar, text="Очистить", command=self.clear_search).pack(
            side=tk.LEFT, padx=2
        )

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        self.btn_start = ttk.Button(toolbar, text="Старт", command=self.toggle_monitoring)
        self.btn_start.pack(side=tk.LEFT, padx=5)

        content = ttk.Frame(parent)
        content.pack(fill=tk.BOTH, expand=True)

        paned = ttk.Panedwindow(content, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        table_frame = ttk.Frame(paned)
        detail_frame = ttk.Frame(paned)
        paned.add(table_frame, weight=3)
        paned.add(detail_frame, weight=1)

        columns = (
            "status_indicator",
            "address",
            "description",
            "subgroup",
            "ping_interval",
            "latency",
            "last_check",
        )
        self.tree = ttk.Treeview(
            table_frame, columns=columns, show="headings", selectmode="extended", height=20
        )

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

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("status_green", foreground="#2ecc71")
        self.tree.tag_configure("status_yellow", foreground="#f39c12")
        self.tree.tag_configure("status_red", foreground="#e74c3c")
        self.tree.tag_configure("status_gray", foreground="#95a5a6")

        self.tree.bind("<<TreeviewSelect>>", self.on_host_select)

        ttk.Label(detail_frame, text="Детали хоста", font=("Segoe UI", 12, "bold")).pack(
            anchor="w", padx=10, pady=(10, 6)
        )

        info_frame = ttk.Frame(detail_frame)
        info_frame.pack(fill=tk.X, padx=10)

        self.detail_vars = {
            "address": tk.StringVar(value="—"),
            "description": tk.StringVar(value="—"),
            "subgroup": tk.StringVar(value="—"),
            "interval": tk.StringVar(value="—"),
            "status": tk.StringVar(value="—"),
            "latency": tk.StringVar(value="—"),
            "last_check": tk.StringVar(value="—"),
        }

        for label, key in [
            ("Адрес:", "address"),
            ("Описание:", "description"),
            ("Подгруппа:", "subgroup"),
            ("Интервал:", "interval"),
            ("Статус:", "status"),
            ("Задержка:", "latency"),
            ("Последняя проверка:", "last_check"),
        ]:
            row = ttk.Frame(info_frame)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=label, width=16, anchor="w").pack(side=tk.LEFT)
            ttk.Label(row, textvariable=self.detail_vars[key], anchor="w").pack(
                side=tk.LEFT, fill=tk.X, expand=True
            )

        ttk.Label(detail_frame, text="График задержек", font=("Segoe UI", 11, "bold")).pack(
            anchor="w", padx=10, pady=(16, 6)
        )
        self.chart_canvas = tk.Canvas(
            detail_frame, height=200, background="#ffffff", highlightthickness=1
        )
        self.chart_canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.chart_canvas.bind("<Configure>", lambda event: self.update_chart())

    def create_stats_widgets(self, parent):
        toolbar = ttk.Frame(parent)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        ttk.Button(toolbar, text="Обновить статистику", command=self.refresh_stats).pack(
            side=tk.LEFT, padx=5
        )

        columns = ("status", "group", "total", "online", "offline", "unknown")
        self.stats_tree = ttk.Treeview(
            parent, columns=columns, show="headings", selectmode="browse"
        )

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
            self.stats_tree.insert(
                "",
                "end",
                values=("●", s["group"], s["total"], s["online"], s["offline"], s["unknown"]),
                tags=(tag,),
            )

        self.update_overall_status()

    def update_overall_status(self):
        overall = self.db.get_overall_status()
        color_map = {
            "healthy": "#2ecc71",
            "warning": "#f39c12",
            "critical": "#e74c3c",
        }
        self.overall_status_label.config(foreground=color_map.get(overall, "#95a5a6"))

    def clear_search(self):
        self.search_var.set("")
        self.refresh_table()

    def load_groups(self):
        groups = self.db.get_groups()
        self.group_combo["values"] = groups
        if groups:
            if self.current_group and self.current_group in groups:
                self.group_combo.set(self.current_group)
            else:
                self.group_combo.current(0)
                self.on_group_select(None)
        else:
            self.group_combo.set("")
            self.clear_table()

    def on_group_select(self, event):
        self.current_group = self.group_combo.get()
        self.selected_host = None
        self.reset_detail_panel()
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

    def add_host_dialog(self):
        if not self.current_group:
            messagebox.showwarning("Внимание", "Сначала создайте или выберите группу.")
            return

        dialog = tk.Toplevel(self)
        dialog.title("Добавить хост")
        dialog.geometry("350x340")

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
                if interval_minutes <= 0:
                    raise ValueError("Интервал должен быть больше 0 минут.")
                interval_seconds = interval_minutes * 60
                self.db.add_host(
                    self.current_group,
                    addr_entry.get(),
                    desc_entry.get(),
                    sub_entry.get() or None,
                    interval_seconds,
                )
                self.refresh_table()
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("Ошибка", str(e))
            except sqlite3.Error as e:
                messagebox.showerror("Ошибка БД", str(e))

        ttk.Button(dialog, text="Сохранить", command=save).pack(pady=10)

    def edit_host_dialog(self):
        if not self.current_group:
            messagebox.showwarning("Внимание", "Выберите группу.")
            return

        selected_item = self.tree.selection()
        if len(selected_item) != 1:
            messagebox.showwarning("Внимание", "Выберите один хост для редактирования.")
            return

        item = self.tree.item(selected_item[0])
        address = item["values"][1]
        hosts = self.db.get_hosts(self.current_group)
        host_map = {h[0]: h for h in hosts}
        host_data = host_map.get(address)
        if not host_data:
            return

        dialog = tk.Toplevel(self)
        dialog.title("Редактировать хост")
        dialog.geometry("350x340")

        ttk.Label(dialog, text="Адрес (IP/Домен):").pack(pady=5)
        addr_entry = ttk.Entry(dialog, width=30)
        addr_entry.insert(0, host_data[0])
        addr_entry.configure(state="disabled")
        addr_entry.pack(pady=5)

        ttk.Label(dialog, text="Описание:").pack(pady=5)
        desc_entry = ttk.Entry(dialog, width=30)
        desc_entry.insert(0, host_data[1] or "")
        desc_entry.pack(pady=5)

        ttk.Label(dialog, text="Подгруппа (необязательно):").pack(pady=5)
        sub_entry = ttk.Entry(dialog, width=30)
        sub_entry.insert(0, host_data[2] or "")
        sub_entry.pack(pady=5)

        ttk.Label(dialog, text="Интервал пинга (минуты):").pack(pady=5)
        interval_minutes = host_data[3] // 60 if host_data[3] else 10
        interval_var = tk.StringVar(value=str(interval_minutes))
        interval_entry = ttk.Entry(dialog, width=30, textvariable=interval_var)
        interval_entry.pack(pady=5)

        def save():
            try:
                interval_minutes = int(interval_var.get())
                if interval_minutes <= 0:
                    raise ValueError("Интервал должен быть больше 0 минут.")
                interval_seconds = interval_minutes * 60
                self.db.update_host(
                    self.current_group,
                    host_data[0],
                    desc_entry.get(),
                    sub_entry.get() or None,
                    interval_seconds,
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
            address = item["values"][1]
            self.db.remove_host(self.current_group, address)

        self.refresh_table()

    def export_data(self):
        if not self.current_group:
            return

        if not HAS_OPENPYXL:
            messagebox.showerror("Ошибка", "openpyxl не установлен. Невозможно экспортировать.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
        )

        if filename:
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "hosts"

                ws.append(["address", "description", "subgroup", "ping_interval_min"])
                hosts = self.db.get_hosts(self.current_group)
                for h in hosts:
                    interval_min = h[3] // 60 if h[3] else 10
                    ws.append([h[0], h[1], h[2] or "", interval_min])

                for col in ws.columns:
                    max_length = 0
                    for cell in col:
                        cell_value = str(cell.value) if cell.value is not None else ""
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
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
        if address not in self.host_status_cache:
            return "gray"

        status, offline_since = self.host_status_cache[address]

        if status == "Online":
            return "green"
        if status == "Offline":
            if offline_since:
                try:
                    offline_time = datetime.strptime(offline_since, "%Y-%m-%d %H:%M:%S")
                    diff = (datetime.now() - offline_time).total_seconds()
                    if diff < 3600:
                        return "yellow"
                    return "red"
                except ValueError:
                    return "red"
            return "red"
        return "gray"

    def refresh_table(self):
        self.clear_table()
        if not self.current_group:
            return

        search_text = self.search_var.get().strip().lower()

        hosts = self.db.get_hosts(self.current_group)
        for host in hosts:
            address, desc, subgroup, interval_sec, last_ping_time, offline_since = host

            if search_text:
                haystack = " ".join(
                    [str(address), str(desc or ""), str(subgroup or "")]
                ).lower()
                if search_text not in haystack:
                    continue

            status = self.db.get_last_status(self.current_group, address)

            self.host_status_cache[address] = (status, offline_since)

            interval_min = interval_sec // 60 if interval_sec else 10
            color_name = self.get_status_color(address)
            tag = f"status_{color_name}"

            last_check_display = last_ping_time or "-"

            self.tree.insert(
                "",
                "end",
                values=("●", address, desc, subgroup or "", interval_min, "-", last_check_display),
                tags=(tag,),
            )

    def clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def reset_detail_panel(self):
        for key in self.detail_vars:
            self.detail_vars[key].set("—")
        self.chart_canvas.delete("all")

    def on_host_select(self, event):
        selection = self.tree.selection()
        if not selection:
            return
        item = self.tree.item(selection[0])
        address = item["values"][1]
        self.selected_host = (self.current_group, address)
        self.update_detail_panel(address)
        self.update_chart()

    def update_detail_panel(self, address):
        if not self.current_group or not address:
            return
        hosts = self.db.get_hosts(self.current_group)
        host_map = {h[0]: h for h in hosts}
        host_data = host_map.get(address)
        if not host_data:
            return

        address, desc, subgroup, interval_sec, last_ping_time, _ = host_data
        status = self.db.get_last_status(self.current_group, address)
        latest_results = self.db.get_recent_results(self.current_group, address, limit=1)
        latest_latency = latest_results[-1][2] if latest_results else None

        self.detail_vars["address"].set(address)
        self.detail_vars["description"].set(desc or "—")
        self.detail_vars["subgroup"].set(subgroup or "—")
        interval_min = interval_sec // 60 if interval_sec else 10
        self.detail_vars["interval"].set(f"{interval_min} мин")
        self.detail_vars["status"].set(status)
        self.detail_vars["latency"].set(f"{latest_latency:.3f} сек" if latest_latency else "—")
        self.detail_vars["last_check"].set(last_ping_time or "—")

    def update_chart(self):
        self.chart_canvas.delete("all")
        if not self.selected_host:
            return
        group, address = self.selected_host
        data = self.db.get_recent_results(group, address, limit=50)
        if not data:
            return

        latencies = [row[2] for row in data if row[2] is not None]
        if not latencies:
            return

        width = self.chart_canvas.winfo_width() or 300
        height = self.chart_canvas.winfo_height() or 200
        padding = 20
        min_latency = min(latencies)
        max_latency = max(latencies)
        span = max(max_latency - min_latency, 0.001)

        points = []
        for idx, latency in enumerate(latencies):
            x = padding + idx * (width - 2 * padding) / max(len(latencies) - 1, 1)
            y = height - padding - (latency - min_latency) / span * (height - 2 * padding)
            points.extend([x, y])

        self.chart_canvas.create_line(
            padding, height - padding, width - padding, height - padding, fill="#e5e7eb"
        )
        self.chart_canvas.create_line(
            padding, padding, padding, height - padding, fill="#e5e7eb"
        )
        self.chart_canvas.create_line(points, fill="#2563eb", width=2, smooth=True)

    def toggle_monitoring(self):
        if self.monitoring:
            self.monitoring = False
            self.btn_start.configure(text="Старт")
            self.status_var.set("Мониторинг остановлен")
        else:
            self.monitoring = True
            self.btn_start.configure(text="Стоп")
            self.status_var.set("Мониторинг запущен...")
            self.monitor_thread = threading.Thread(target=self.monitoring_loop, daemon=True)
            self.monitor_thread.start()

    def is_due(self, last_ping_time, interval_seconds, now):
        if not last_ping_time:
            return True
        try:
            last_time = datetime.strptime(last_ping_time, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return True
        return (now - last_time).total_seconds() >= interval_seconds

    def chunk_hosts(self, hosts, chunk_size):
        for i in range(0, len(hosts), chunk_size):
            yield hosts[i : i + chunk_size]

    def ping_hosts_batch(self, hosts_batch):
        entries = []
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_PING_WORKERS) as executor:
            future_to_host = {
                executor.submit(ping_host_subprocess, host[1]): host for host in hosts_batch
            }
            for future in concurrent.futures.as_completed(future_to_host):
                host = future_to_host[future]
                group_name, address, subgroup = host
                try:
                    success, latency = future.result()
                    status = "Online" if success else "Offline"
                    entries.append(
                        {
                            "timestamp": timestamp,
                            "group_name": group_name,
                            "subgroup": subgroup,
                            "address": address,
                            "status": status,
                            "latency": latency,
                        }
                    )
                except Exception:
                    entries.append(
                        {
                            "timestamp": timestamp,
                            "group_name": group_name,
                            "subgroup": subgroup,
                            "address": address,
                            "status": "Offline",
                            "latency": None,
                        }
                    )
        return entries

    def monitoring_loop(self):
        while self.monitoring:
            groups = self.db.get_groups()
            if not groups:
                time.sleep(self.monitor_interval)
                continue

            now = datetime.now()
            due_hosts = []
            for group in groups:
                hosts = self.db.get_hosts(group)
                for host in hosts:
                    address, desc, subgroup, interval_sec, last_ping_time, _ = host
                    interval_sec = interval_sec or DEFAULT_PING_INTERVAL
                    if self.is_due(last_ping_time, interval_sec, now):
                        due_hosts.append((group, address, subgroup))

            if due_hosts:
                for batch in self.chunk_hosts(due_hosts, PING_BATCH_SIZE):
                    if not self.monitoring:
                        break
                    results = self.ping_hosts_batch(batch)
                    self.db.log_results_batch(results)
                    for result in results:
                        self.ui_queue.put({
                            "type": "row",
                            "group": result["group_name"],
                            "address": result["address"],
                            "status": result["status"],
                            "latency": result["latency"],
                            "timestamp": result["timestamp"],
                        })

            now_ts = time.time()
            if now_ts - self.last_stats_refresh >= STATS_REFRESH_INTERVAL:
                self.ui_queue.put({"type": "stats"})
                self.ui_queue.put({"type": "overall"})
                self.last_stats_refresh = now_ts

            time.sleep(self.monitor_interval)

    def update_row(self, group, address, status, latency, timestamp):
        if self.current_group != group:
            return
        for item_id in self.tree.get_children():
            vals = self.tree.item(item_id)["values"]
            if vals and str(vals[1]) == str(address):
                offline_since = self.host_status_cache.get(address, ("Unknown", None))[1]
                if status == "Offline" and offline_since is None:
                    offline_since = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                elif status == "Online":
                    offline_since = None

                self.host_status_cache[address] = (status, offline_since)

                color = self.get_status_color(address)
                tag = f"status_{color}"

                indicator = "●"
                latency_str = f"{latency:.3f}" if latency else "-"
                new_vals = (
                    indicator,
                    vals[1],
                    vals[2],
                    vals[3],
                    vals[4],
                    latency_str,
                    timestamp,
                )
                self.tree.item(item_id, values=new_vals, tags=(tag,))
                break

        if self.selected_host == (group, address):
            self.update_detail_panel(address)
            self.update_chart()

    def process_ui_queue(self):
        try:
            while True:
                event = self.ui_queue.get_nowait()
                if event["type"] == "row":
                    self.update_row(
                        event["group"],
                        event["address"],
                        event["status"],
                        event["latency"],
                        event["timestamp"],
                    )
                elif event["type"] == "stats":
                    selected = self.notebook.select()
                    if self.notebook.tab(selected, "text") == "Статистика":
                        self.refresh_stats()
                elif event["type"] == "overall":
                    self.update_overall_status()
        except queue.Empty:
            pass
        finally:
            self.after(200, self.process_ui_queue)


if __name__ == "__main__":
    app = PingApp()
    app.mainloop()
