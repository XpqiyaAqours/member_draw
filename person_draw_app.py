import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import hashlib
import datetime
import random
import os
import ctypes
import smtplib
from email.message import EmailMessage

try:
    import openpyxl
except ImportError:
    openpyxl = None


# 全局字体配置
BASE_FONT = ("Microsoft YaHei", 11)
TITLE_FONT = ("Microsoft YaHei", 18, "bold")
SUBTITLE_FONT = ("Microsoft YaHei", 13, "bold")


# -------------------- DPI 适配 & 数据库路径 --------------------


def set_dpi_awareness():
    """
    在 Windows 下启用高 DPI 感知，避免高分屏上界面被系统缩放变小/模糊。
    优先使用 Per-Monitor V2，其次 Per-Monitor，再其次 Process DPI Aware。
    """
    if os.name != "nt":
        return
    try:
        # Windows 10+，DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        awareness_context = ctypes.c_void_p(-4)
        ctypes.windll.user32.SetProcessDpiAwarenessContext(awareness_context)
    except Exception:
        try:
            # Windows 8.1 API
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
        except Exception:
            try:
                # 最老的方式
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def get_db_path():
    """
    将数据库放到 Windows 用户的 AppData\\Roaming\\PersonDrawApp 下。
    例如：C:\\Users\\用户名\\AppData\\Roaming\\PersonDrawApp\\person_draw.db
    """
    appdata = os.getenv("APPDATA")
    if not appdata:
        appdata = os.path.expanduser("~")
    app_dir = os.path.join(appdata, "PersonDrawApp")
    os.makedirs(app_dir, exist_ok=True)
    return os.path.join(app_dir, "person_draw.db")


DB_FILE = get_db_path()


def hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()


# -------------------- 数据库层 --------------------


class Database:
    def __init__(self, db_path=DB_FILE):
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self.init_schema()

    def init_schema(self):
        c = self.conn.cursor()

        # 用户表
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('admin','user')),
                email TEXT
            )
        """
        )

        # 兼容老版本库，尝试补充 email 字段
        try:
            c.execute("ALTER TABLE users ADD COLUMN email TEXT")
        except sqlite3.OperationalError:
            pass

        # 若无用户则创建一个默认管理员
        c.execute("SELECT COUNT(*) AS cnt FROM users")
        if c.fetchone()["cnt"] == 0:
            c.execute(
                "INSERT INTO users (username,password_hash,role,email) VALUES (?,?,?,?)",
                ("admin", hash_password("admin123"), "admin", ""),
            )

        # 专家名库表
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS people (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                phone TEXT NOT NULL
            )
        """
        )

        # 抽签会话索引表（每次完整抽 3 人为一次记录）
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                created_at TEXT NOT NULL
            )
        """
        )

        # 抽签详细日志表
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS draw_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id INTEGER NOT NULL,
                person_id INTEGER NOT NULL,
                order_no INTEGER NOT NULL,
                present INTEGER NOT NULL,
                absent_reason TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(session_id) REFERENCES sessions(id),
                FOREIGN KEY(person_id) REFERENCES people(id)
            )
        """
        )

        self.conn.commit()

    # ---------- 用户相关 ----------

    def register_user(self, username, password, role="user", email=""):
        c = self.conn.cursor()
        try:
            c.execute(
                "INSERT INTO users (username,password_hash,role,email) VALUES (?,?,?,?)",
                (username, hash_password(password), role, email),
            )
            self.conn.commit()
            return True, "注册成功"
        except sqlite3.IntegrityError:
            return False, "用户名已存在"

    def authenticate(self, username, password):
        c = self.conn.cursor()
        c.execute("SELECT * FROM users WHERE username=?", (username,))
        row = c.fetchone()
        if not row:
            return None
        if row["password_hash"] == hash_password(password):
            return row
        return None

    def get_all_users(self):
        c = self.conn.cursor()
        c.execute("SELECT id, username, role, email FROM users ORDER BY id ASC")
        return c.fetchall()

    def update_user(self, user_id, username, role, new_password=None, email=""):
        c = self.conn.cursor()
        # 检查是否存在
        c.execute("SELECT * FROM users WHERE id=?", (user_id,))
        row = c.fetchone()
        if not row:
            raise RuntimeError("用户不存在")

        # 若修改为普通用户，需保证至少还有一个管理员
        if row["role"] == "admin" and role == "user":
            c.execute("SELECT COUNT(*) AS cnt FROM users WHERE role='admin'")
            if c.fetchone()["cnt"] <= 1:
                raise RuntimeError("至少需要保留一个管理员账号")

        if new_password:
            c.execute(
                """
                UPDATE users
                SET username=?, role=?, password_hash=?, email=?
                WHERE id=?
                """,
                (username, role, hash_password(new_password), email, user_id),
            )
        else:
            c.execute(
                """
                UPDATE users
                SET username=?, role=?, email=?
                WHERE id=?
                """,
                (username, role, email, user_id),
            )
        self.conn.commit()

    def delete_user(self, user_id, current_user_id=None):
        c = self.conn.cursor()
        c.execute("SELECT * FROM users WHERE id=?", (user_id,))
        row = c.fetchone()
        if not row:
            raise RuntimeError("用户不存在")

        # 不允许删除当前登录账号，避免状态混乱
        if current_user_id is not None and user_id == current_user_id:
            raise RuntimeError("不能删除当前登录的账号")

        # 若删除管理员，需保证至少还有一个管理员
        if row["role"] == "admin":
            c.execute("SELECT COUNT(*) AS cnt FROM users WHERE role='admin'")
            if c.fetchone()["cnt"] <= 1:
                raise RuntimeError("至少需要保留一个管理员账号")

        c.execute("DELETE FROM users WHERE id=?", (user_id,))
        self.conn.commit()

    def export_users_to_excel(self, path):
        if openpyxl is None:
            raise RuntimeError("缺少 openpyxl 库，无法导出 Excel")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "users"
        ws.append(["用户名", "密码(明文)", "角色(admin/user)", "电子邮箱"])
        c = self.conn.cursor()
        c.execute("SELECT username, role, email FROM users ORDER BY id ASC")
        # 密码无法还原成明文，这里导出为空，方便批量编辑再导入
        for r in c.fetchall():
            ws.append([r["username"], "", r["role"], r["email"] or ""])
        wb.save(path)

    def import_users_from_excel(self, path):
        if openpyxl is None:
            raise RuntimeError("缺少 openpyxl 库，无法导入 Excel")
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            if not row or not row[0]:
                continue
            username = str(row[0]).strip()
            password = str(row[1]).strip() if len(row) > 1 and row[1] else "123456"
            role = (
                str(row[2]).strip()
                if len(row) > 2 and row[2] in ("admin", "user")
                else "user"
            )
            email = str(row[3]).strip() if len(row) > 3 and row[3] else ""
            if not username:
                continue
            # 如果已存在则更新角色并可选重置密码，否则新建
            c = self.conn.cursor()
            c.execute("SELECT id FROM users WHERE username=?", (username,))
            exist = c.fetchone()
            if exist:
                uid = exist["id"]
                # 更新：用户名不变（按 Excel）、角色变更，如提供密码则重置
                if row[1]:
                    self.update_user(uid, username, role, new_password=password, email=email)
                else:
                    self.update_user(uid, username, role, email=email)
            else:
                self.register_user(username, password, role=role, email=email)

    # ---------- 人员表 ----------

    def get_all_people(self):
        c = self.conn.cursor()
        c.execute("SELECT * FROM people ORDER BY id ASC")
        return c.fetchall()

    def add_person(self, name, phone):
        c = self.conn.cursor()
        c.execute(
            "INSERT INTO people (name,phone) VALUES (?,?)",
            (name, phone),
        )
        self.conn.commit()

    def delete_person(self, person_id):
        c = self.conn.cursor()
        c.execute("DELETE FROM people WHERE id=?", (person_id,))
        self.conn.commit()

    def update_person(self, person_id, name, phone):
        c = self.conn.cursor()
        c.execute(
            "UPDATE people SET name=?, phone=? WHERE id=?",
            (name, phone, person_id),
        )
        self.conn.commit()

    # ---------- 抽签 / 日志 ----------

    def create_session(self, title):
        c = self.conn.cursor()
        c.execute(
            "INSERT INTO sessions (title,created_at) VALUES (?,?)",
            (
                title,
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        self.conn.commit()
        return c.lastrowid

    def add_draw_log(self, session_id, person_id, order_no, present, absent_reason):
        c = self.conn.cursor()
        c.execute(
            """
            INSERT INTO draw_logs
            (session_id,person_id,order_no,present,absent_reason,created_at)
            VALUES (?,?,?,?,?,?)
        """,
            (
                session_id,
                person_id,
                order_no,
                1 if present else 0,
                absent_reason,
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        self.conn.commit()

    def get_sessions(self):
        c = self.conn.cursor()
        c.execute("SELECT * FROM sessions ORDER BY id DESC")
        return c.fetchall()

    def get_session_logs(self, session_id):
        """获取某论政项目的所有抽签/补签日志，按记录时间倒序（最新在上）"""
        c = self.conn.cursor()
        c.execute(
            """
            SELECT dl.*, p.name, p.phone
            FROM draw_logs dl
            JOIN people p ON p.id = dl.person_id
            WHERE dl.session_id=?
            ORDER BY dl.created_at DESC, dl.id DESC
        """,
            (session_id,),
        )
        return c.fetchall()

    def get_completed_sessions(self):
        """返回已完成抽签（有 3 人到场）的论政项目"""
        sessions = self.get_sessions()
        result = []
        for s in sessions:
            logs = self.get_session_logs(s["id"])
            present_count = sum(1 for r in logs if r["present"] == 1)
            if present_count == 3:
                result.append(s)
        return result

    def get_present_logs(self, session_id):
        """返回某论政项目中所有到场的抽签记录（用于补抽时选择标记不到场）"""
        c = self.conn.cursor()
        c.execute(
            """
            SELECT dl.*, p.name, p.phone
            FROM draw_logs dl
            JOIN people p ON p.id = dl.person_id
            WHERE dl.session_id=? AND dl.present=1
            ORDER BY dl.order_no ASC, dl.id ASC
        """,
            (session_id,),
        )
        return c.fetchall()

    def add_supplement_absent_log(self, session_id, person_id):
        """补签时新增一条“后续不到场”记录，不修改原记录"""
        order_no = self.get_max_order_no(session_id) + 1
        self.add_draw_log(
            session_id,
            person_id,
            order_no,
            present=False,
            absent_reason="专家后续不到场，进行再次抽选。",
        )

    def get_max_order_no(self, session_id):
        """获取某论政项目中当前最大的 order_no，用于补抽时续写"""
        c = self.conn.cursor()
        c.execute(
            "SELECT COALESCE(MAX(order_no), 0) AS mx FROM draw_logs WHERE session_id=?",
            (session_id,),
        )
        return c.fetchone()["mx"]

    # Excel IO - 人员
    def export_people_to_excel(self, path):
        if openpyxl is None:
            raise RuntimeError("缺少 openpyxl 库，无法导出 Excel")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "people"
        ws.append(["姓名", "手机号"])
        for row in self.get_all_people():
            ws.append([row["name"], row["phone"]])
        wb.save(path)

    def import_people_from_excel(self, path):
        if openpyxl is None:
            raise RuntimeError("缺少 openpyxl 库，无法导入 Excel")
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            if not row or not row[0]:
                continue
            name = str(row[0]).strip()
            phone = str(row[1]).strip() if len(row) > 1 else ""
            if name:
                self.add_person(name, phone)

    def get_mail_config(self):
        """获取邮件配置（若表不存在则初始化）"""
        c = self.conn.cursor()
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS mail_config (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                smtp_server TEXT,
                smtp_port INTEGER,
                use_ssl INTEGER,
                use_tls INTEGER,
                username TEXT,
                password TEXT,
                from_addr TEXT
            )
        """
        )
        c.execute("INSERT OR IGNORE INTO mail_config (id) VALUES (1)")
        c.execute("SELECT * FROM mail_config WHERE id=1")
        row = c.fetchone()
        return {
            "smtp_server": row["smtp_server"],
            "smtp_port": row["smtp_port"] or 0,
            "use_ssl": bool(row["use_ssl"]),
            "use_tls": bool(row["use_tls"]),
            "username": row["username"],
            "password": row["password"],
            "from_addr": row["from_addr"],
        }

    def save_mail_config(self, cfg: dict):
        c = self.conn.cursor()
        c.execute(
            """
            UPDATE mail_config
            SET smtp_server=?, smtp_port=?, use_ssl=?, use_tls=?, username=?, password=?, from_addr=?
            WHERE id=1
        """,
            (
                cfg.get("smtp_server") or "",
                int(cfg.get("smtp_port") or 0),
                1 if cfg.get("use_ssl") else 0,
                1 if cfg.get("use_tls") else 0,
                cfg.get("username") or "",
                cfg.get("password") or "",
                cfg.get("from_addr") or "",
            ),
        )
        self.conn.commit()

    def get_admin_emails(self):
        c = self.conn.cursor()
        c.execute(
            "SELECT email FROM users WHERE role='admin' AND email IS NOT NULL AND email<>''"
        )
        return [r["email"] for r in c.fetchall()]

    # Excel IO - 日志
    def export_logs_to_excel(self, path):
        if openpyxl is None:
            raise RuntimeError("缺少 openpyxl 库，无法导出 Excel")
        wb = openpyxl.Workbook()

        # Sheet 1: sessions
        ws1 = wb.active
        ws1.title = "sessions"
        ws1.append(["ID", "标题", "创建时间"])
        for s in self.get_sessions():
            ws1.append([s["id"], s["title"], s["created_at"]])

        # Sheet 2: draw_logs
        ws2 = wb.create_sheet("draw_logs")
        ws2.append(
            [
                "ID",
                "SessionID",
                "抽签顺序",
                "人员ID",
                "姓名",
                "手机号",
                "到场情况",
                "备注",
                "记录时间",
            ]
        )
        c = self.conn.cursor()
        c.execute(
            """
            SELECT dl.*, p.name, p.phone
            FROM draw_logs dl
            JOIN people p ON p.id = dl.person_id
            ORDER BY dl.session_id DESC, dl.created_at DESC, dl.id DESC
        """
        )

        def present_display(row):
            if row["present"] == 1:
                return "到场"
            if row["absent_reason"] == "专家后续不到场，进行再次抽选。":
                return "后续不到场"
            return "不到场"

        for r in c.fetchall():
            ws2.append(
                [
                    r["id"],
                    r["session_id"],
                    r["order_no"],
                    r["person_id"],
                    r["name"],
                    r["phone"],
                    present_display(r),
                    r["absent_reason"] or "",
                    r["created_at"],
                ]
            )

        wb.save(path)


# -------------------- 各个界面 Frame --------------------


class LoginFrame(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.build_ui()

    def build_ui(self):
        container = ttk.Frame(self, padding=80)
        container.place(relx=0.5, rely=0.5, anchor="center")

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        title = ttk.Label(
            container,
            text="人员管理与抽签系统",
            font=TITLE_FONT,
        )
        title.grid(row=0, column=0, columnspan=2, pady=(0, 50))

        ttk.Label(container, text="用户名:").grid(
            row=1, column=0, sticky="e", padx=20, pady=12
        )
        self.username_var = tk.StringVar()
        ttk.Entry(container, textvariable=self.username_var, width=28).grid(
            row=1, column=1, sticky="w", padx=20, pady=12
        )

        ttk.Label(container, text="密码:").grid(
            row=2, column=0, sticky="e", padx=20, pady=12
        )
        self.password_var = tk.StringVar()
        ttk.Entry(container, textvariable=self.password_var, show="*", width=28).grid(
            row=2, column=1, sticky="w", padx=20, pady=12
        )

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=30)

        ttk.Button(btn_frame, text="登录", command=self.login, width=14).grid(
            row=0, column=0, padx=15
        )

        ttk.Label(
            container,
            text="默认管理员账号: admin / admin123",
            foreground="gray",
        ).grid(row=4, column=0, columnspan=2, pady=(20, 0))

    def login(self):
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        if not username or not password:
            messagebox.showwarning("提示", "请输入用户名和密码")
            return
        user = self.app.db.authenticate(username, password)
        if user:
            self.app.current_user = user
            self.app.show_main_frame()
        else:
            messagebox.showerror("错误", "用户名或密码错误")


class RegisterFrame(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.build_ui()

    def build_ui(self):
        container = ttk.Frame(self, padding=80)
        container.place(relx=0.5, rely=0.5, anchor="center")

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        ttk.Label(
            container, text="注册", font=TITLE_FONT
        ).grid(row=0, column=0, columnspan=2, pady=(0, 50))

        ttk.Label(container, text="用户名:").grid(
            row=1, column=0, sticky="e", padx=20, pady=12
        )
        self.username_var = tk.StringVar()
        ttk.Entry(container, textvariable=self.username_var, width=28).grid(
            row=1, column=1, sticky="w", padx=20, pady=12
        )

        ttk.Label(container, text="密码:").grid(
            row=2, column=0, sticky="e", padx=20, pady=12
        )
        self.password_var = tk.StringVar()
        ttk.Entry(container, textvariable=self.password_var, show="*", width=28).grid(
            row=2, column=1, sticky="w", padx=20, pady=12
        )

        ttk.Label(container, text="确认密码:").grid(
            row=3, column=0, sticky="e", padx=20, pady=12
        )
        self.password2_var = tk.StringVar()
        ttk.Entry(container, textvariable=self.password2_var, show="*", width=28).grid(
            row=3, column=1, sticky="w", padx=20, pady=12
        )

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=30)

        ttk.Button(btn_frame, text="注册", command=self.register, width=14).grid(
            row=0, column=0, padx=15
        )
        ttk.Button(btn_frame, text="返回登录", command=self.back_login, width=14).grid(
            row=0, column=1, padx=15
        )

    def register(self):
        u = self.username_var.get().strip()
        p1 = self.password_var.get().strip()
        p2 = self.password2_var.get().strip()
        if not u or not p1 or not p2:
            messagebox.showwarning("提示", "请完整填写信息")
            return
        if p1 != p2:
            messagebox.showwarning("提示", "两次密码不一致")
            return
        ok, msg = self.app.db.register_user(u, p1, role="user")
        if ok:
            messagebox.showinfo("成功", msg)
            self.back_login()
        else:
            messagebox.showerror("失败", msg)

    def back_login(self):
        self.app.show_login_frame()


class PeopleFrame(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.build_ui()
        self.refresh()

    def build_ui(self):
        outer = ttk.Frame(self, padding=30)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 15))

        ttk.Label(top, text="专家名库", font=SUBTITLE_FONT).pack(
            side="left", padx=8
        )

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)

        columns = ("id", "name", "phone")
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=18,
        )
        self.tree.heading("id", text="ID")
        self.tree.heading("name", text="姓名")
        self.tree.heading("phone", text="手机号")
        self.tree.column("id", width=70, anchor="center", stretch=False)
        self.tree.column("name", width=200, anchor="center", stretch=True)
        self.tree.column("phone", width=260, anchor="center", stretch=True)
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        # 下方中部按钮栏
        btn_bar = ttk.Frame(outer)
        btn_bar.pack(pady=15)

        self.btn_add = ttk.Button(btn_bar, text="新增", command=self.add_person, width=9)
        self.btn_edit = ttk.Button(btn_bar, text="编辑", command=self.edit_person, width=9)
        self.btn_del = ttk.Button(btn_bar, text="删除", command=self.delete_person, width=9)
        self.btn_import = ttk.Button(
            btn_bar, text="Excel 导入", command=self.import_excel, width=11
        )
        self.btn_export = ttk.Button(
            btn_bar, text="Excel 导出", command=self.export_excel, width=11
        )

        for b in (
            self.btn_add,
            self.btn_edit,
            self.btn_del,
            self.btn_import,
            self.btn_export,
        ):
            b.pack(side="left", padx=8)

    def set_admin_mode(self, is_admin):
        state = "normal" if is_admin else "disabled"
        for b in (self.btn_add, self.btn_edit, self.btn_del, self.btn_import):
            b["state"] = state
        self.btn_export["state"] = "normal"

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for row in self.app.db.get_all_people():
            self.tree.insert(
                "",
                "end",
                values=(row["id"], row["name"], row["phone"]),
            )

    def _get_selected_id(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return int(self.tree.item(sel[0], "values")[0])

    def add_person(self):
        self._edit_dialog()

    def edit_person(self):
        pid = self._get_selected_id()
        if not pid:
            messagebox.showwarning("提示", "请先选择一条记录")
            return
        people = {p["id"]: p for p in self.app.db.get_all_people()}
        if pid not in people:
            return
        p = people[pid]
        self._edit_dialog(pid, p["name"], p["phone"])

    def _edit_dialog(self, person_id=None, name="", phone=""):
        dlg = tk.Toplevel(self)
        dlg.title("人员信息")
        dlg.grab_set()
        dlg.resizable(False, False)

        container = ttk.Frame(dlg, padding=25)
        container.grid(row=0, column=0, sticky="nsew")
        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)

        ttk.Label(container, text="姓名:").grid(
            row=0, column=0, padx=8, pady=10, sticky="e"
        )
        name_var = tk.StringVar(value=name)
        ttk.Entry(container, textvariable=name_var, width=28).grid(
            row=0, column=1, padx=8, pady=10, sticky="w"
        )

        ttk.Label(container, text="手机号:").grid(
            row=1, column=0, padx=8, pady=10, sticky="e"
        )
        phone_var = tk.StringVar(value=phone)
        ttk.Entry(container, textvariable=phone_var, width=28).grid(
            row=1, column=1, padx=8, pady=10, sticky="w"
        )

        def on_ok():
            n = name_var.get().strip()
            ph = phone_var.get().strip()
            if not n:
                messagebox.showwarning("提示", "姓名不能为空")
                return
            if person_id is None:
                self.app.db.add_person(n, ph)
            else:
                self.app.db.update_person(person_id, n, ph)
            self.refresh()
            dlg.destroy()

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=(18, 0))
        ttk.Button(btn_frame, text="确定", command=on_ok, width=11).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(btn_frame, text="取消", command=dlg.destroy, width=11).grid(
            row=0, column=1, padx=10
        )

    def delete_person(self):
        pid = self._get_selected_id()
        if not pid:
            messagebox.showwarning("提示", "请先选择一条记录")
            return
        if messagebox.askyesno("确认", "确定要删除该人员吗？"):
            self.app.db.delete_person(pid)
            self.refresh()

    def import_excel(self):
        if openpyxl is None:
            messagebox.showerror("错误", "请先安装 openpyxl 库")
            return
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm")],
        )
        if not path:
            return
        try:
            self.app.db.import_people_from_excel(path)
            self.refresh()
            messagebox.showinfo("成功", "导入完成")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")

    def export_excel(self):
        if openpyxl is None:
            messagebox.showerror("错误", "请先安装 openpyxl 库")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            title="导出专家名库",
            initialfile="people.xlsx",
        )
        if not path:
            return
        try:
            self.app.db.export_people_to_excel(path)
            messagebox.showinfo("成功", "导出完成")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")


class LogsFrame(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.build_ui()
        self.refresh_sessions()

    def build_ui(self):
        outer = ttk.Frame(self, padding=30)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 15))

        ttk.Label(top, text="抽签日志", font=SUBTITLE_FONT).pack(
            side="left", padx=8
        )

        btn_bar = ttk.Frame(top)
        btn_bar.pack(side="right")
        ttk.Button(btn_bar, text="刷新", command=self.refresh_sessions, width=9).pack(
            side="right", padx=5
        )
        ttk.Button(
            btn_bar, text="发送记录至邮箱", command=self.send_logs_email, width=14
        ).pack(side="right", padx=5)
        ttk.Button(
            btn_bar, text="导出全部日志 Excel", command=self.export_logs, width=18
        ).pack(side="right", padx=5)

        main = ttk.Panedwindow(outer, orient="horizontal")
        main.pack(fill="both", expand=True)

        # 左侧：session 列表
        left = ttk.Frame(main, padding=(0, 0, 8, 0))
        columns = ("id", "title", "created_at")
        self.sessions_tree = ttk.Treeview(
            left, columns=columns, show="headings", height=18
        )
        self.sessions_tree.heading("id", text="ID")
        self.sessions_tree.heading("title", text="论政项目名称")
        self.sessions_tree.heading("created_at", text="创建时间")
        self.sessions_tree.column("id", width=70, anchor="center", stretch=False)
        self.sessions_tree.column("title", width=260, stretch=True)
        self.sessions_tree.column("created_at", width=230, anchor="center", stretch=True)
        self.sessions_tree.pack(side="left", fill="both", expand=True)

        vsb1 = ttk.Scrollbar(left, orient="vertical", command=self.sessions_tree.yview)
        self.sessions_tree.configure(yscrollcommand=vsb1.set)
        vsb1.pack(side="right", fill="y")

        self.sessions_tree.bind("<<TreeviewSelect>>", self.on_session_select)

        # 右侧：详细日志（隐藏顺序，第一列为记录时间，按最新在上排序）
        right = ttk.Frame(main, padding=(8, 0, 0, 0))
        columns2 = (
            "created_at",
            "name",
            "phone",
            "present",
            "reason",
        )
        self.logs_tree = ttk.Treeview(
            right, columns=columns2, show="headings", height=18
        )
        self.logs_tree.heading("created_at", text="记录时间")
        self.logs_tree.heading("name", text="姓名")
        self.logs_tree.heading("phone", text="手机号")
        self.logs_tree.heading("present", text="到场情况")
        self.logs_tree.heading("reason", text="备注")

        self.logs_tree.column("created_at", width=180, anchor="center", stretch=False)
        self.logs_tree.column("name", width=120, anchor="center", stretch=False)
        self.logs_tree.column("phone", width=160, anchor="center", stretch=False)
        self.logs_tree.column("present", width=100, anchor="center", stretch=False)
        self.logs_tree.column("reason", width=340, stretch=True)
        self.logs_tree.pack(side="left", fill="both", expand=True)

        vsb2 = ttk.Scrollbar(right, orient="vertical", command=self.logs_tree.yview)
        self.logs_tree.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side="right", fill="y")

        main.add(left, weight=1)
        main.add(right, weight=2)

    def refresh_sessions(self):
        for i in self.sessions_tree.get_children():
            self.sessions_tree.delete(i)
        for s in self.app.db.get_sessions():
            self.sessions_tree.insert(
                "",
                "end",
                values=(s["id"], s["title"], s["created_at"]),
            )
        for i in self.logs_tree.get_children():
            self.logs_tree.delete(i)

    def _get_present_display(self, r):
        """到场情况显示：到场 / 后续不到场 / 不到场"""
        if r["present"] == 1:
            return "到场"
        if r["absent_reason"] == "专家后续不到场，进行再次抽选。":
            return "后续不到场"
        return "不到场"

    def on_session_select(self, event=None):
        sel = self.sessions_tree.selection()
        if not sel:
            return
        sid = int(self.sessions_tree.item(sel[0], "values")[0])
        logs = self.app.db.get_session_logs(sid)
        for i in self.logs_tree.get_children():
            self.logs_tree.delete(i)
        for r in logs:
            present_text = self._get_present_display(r)
            self.logs_tree.insert(
                "",
                "end",
                values=(
                    r["created_at"],
                    r["name"],
                    r["phone"],
                    present_text,
                    r["absent_reason"] or "",
                ),
            )

    def export_logs(self):
        if openpyxl is None:
            messagebox.showerror("错误", "请先安装 openpyxl 库")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            title="导出抽签日志",
            initialfile="logs.xlsx",
        )
        if not path:
            return
        try:
            self.app.db.export_logs_to_excel(path)
            messagebox.showinfo("成功", "导出完成")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    def send_logs_email(self):
        """从日志界面发送一个或多个论政项目的日志到所有管理员邮箱"""
        sel = self.sessions_tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择至少一个论政项目")
            return
        session_ids = [int(self.sessions_tree.item(i, "values")[0]) for i in sel]
        self.app.send_sessions_email(self, session_ids)


class DrawFrame(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.current_session_id = None
        self.present_count = 0
        self.current_person = None
        self.order_no = 0
        self.people_cache = []
        self.view_mode = "choice"
        self.supplement_session_id = None
        self.supplement_vacant_count = 0
        self.supplement_order_base = 0
        self.build_ui()

    def build_ui(self):
        self.outer = ttk.Frame(self, padding=30)
        self.outer.pack(fill="both", expand=True)

        # ---------- 功能选择界面（居中垂直两选项） ----------
        self.choice_frame = ttk.Frame(self.outer)
        choice_inner = ttk.Frame(self.choice_frame, padding=60)
        choice_inner.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(
            choice_inner,
            text="请选择抽签功能",
            font=SUBTITLE_FONT,
        ).pack(pady=(0, 40))

        ttk.Button(
            choice_inner,
            text="新论政项目抽签",
            command=self._show_new_draw,
            width=24,
        ).pack(pady=20)

        ttk.Button(
            choice_inner,
            text="过往论政项目补抽",
            command=self._show_supplement_select,
            width=24,
        ).pack(pady=20)

        # ---------- 新论政项目抽签界面 ----------
        self.new_draw_frame = ttk.Frame(self.outer)
        self._build_new_draw_ui()

        # ---------- 过往论政项目补抽界面 ----------
        self.supplement_frame = ttk.Frame(self.outer)
        self._build_supplement_ui()

        self._show_view("choice")

    def _show_view(self, mode):
        self.view_mode = mode
        for f in (self.choice_frame, self.new_draw_frame, self.supplement_frame):
            f.pack_forget()
        if mode == "choice":
            self.choice_frame.pack(fill="both", expand=True)
        elif mode == "new":
            self.new_draw_frame.pack(fill="both", expand=True)
        else:
            self.supplement_frame.pack(fill="both", expand=True)

    def _show_choice(self):
        self._show_view("choice")

    def _show_new_draw(self):
        self._show_view("new")
        self.title_var.set("")
        self.status_var.set("尚未开始")
        self.present_var.set("0 / 3")
        self.name_var.set("")
        self.phone_var.set("")
        self.set_buttons_state(started=False)
        self.current_session_id = None

    def _show_supplement_select(self):
        self._show_view("supplement")
        self._supplement_step = 1
        self.supplement_session_id = None
        self._refresh_supplement_session_list()
        self._show_supplement_step1()

    def _build_new_draw_ui(self):
        outer = self.new_draw_frame
        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 15))

        ttk.Label(top, text="新论政项目抽签", font=SUBTITLE_FONT).pack(
            side="left", padx=8
        )

        ttk.Button(top, text="返回", command=self._show_choice, width=10).pack(
            side="right", padx=5
        )
        ttk.Button(top, text="开始新抽签", command=self.start_new_session, width=14).pack(
            side="right", padx=8
        )

        middle = ttk.Frame(outer)
        middle.pack(pady=15)

        middle.columnconfigure(0, weight=1)
        middle.columnconfigure(1, weight=1)

        ttk.Label(middle, text="论政项目名称:").grid(
            row=0, column=0, sticky="e", padx=20, pady=10
        )
        self.title_var = tk.StringVar()
        ttk.Entry(middle, textvariable=self.title_var, width=40).grid(
            row=0, column=1, sticky="w", padx=20, pady=10
        )

        ttk.Label(middle, text="当前状态:").grid(
            row=1, column=0, sticky="e", padx=20, pady=10
        )
        self.status_var = tk.StringVar(value="尚未开始")
        ttk.Label(middle, textvariable=self.status_var).grid(
            row=1, column=1, sticky="w", padx=20, pady=10
        )

        ttk.Label(middle, text="已到场人数:").grid(
            row=2, column=0, sticky="e", padx=20, pady=10
        )
        self.present_var = tk.StringVar(value="0 / 3")
        ttk.Label(middle, textvariable=self.present_var).grid(
            row=2, column=1, sticky="w", padx=20, pady=10
        )

        ttk.Separator(outer, orient="horizontal").pack(fill="x", pady=8)

        info = ttk.Frame(outer)
        info.pack(pady=20)

        info.columnconfigure(0, weight=1)
        info.columnconfigure(1, weight=1)

        ttk.Label(info, text="抽中人员姓名:").grid(
            row=0, column=0, sticky="e", padx=20, pady=10
        )
        self.name_var = tk.StringVar()
        ttk.Label(info, textvariable=self.name_var, font=("Microsoft YaHei", 15)).grid(
            row=0, column=1, sticky="w", padx=20, pady=10
        )

        ttk.Label(info, text="手机号:").grid(
            row=1, column=0, sticky="e", padx=20, pady=10
        )
        self.phone_var = tk.StringVar()
        ttk.Label(info, textvariable=self.phone_var).grid(
            row=1, column=1, sticky="w", padx=20, pady=10
        )

        btns = ttk.Frame(outer)
        btns.pack(pady=20)

        self.btn_draw = ttk.Button(btns, text="抽签", command=self.draw_one, width=12)
        self.btn_present = ttk.Button(
            btns, text="到场", command=lambda: self.mark_result(True), width=12
        )
        self.btn_absent = ttk.Button(
            btns, text="不到场", command=lambda: self.mark_result(False), width=12
        )

        self.btn_draw.grid(row=0, column=0, padx=15)
        self.btn_present.grid(row=0, column=1, padx=15)
        self.btn_absent.grid(row=0, column=2, padx=15)

        ttk.Label(
            outer,
            text="说明：一次抽签流程中，将依次确认 3 名到场人员。不到场需填写理由并可继续抽下一位。",
            foreground="gray",
        ).pack(pady=20)

        self.set_buttons_state(started=False)

    def _build_supplement_ui(self):
        outer = self.supplement_frame

        self.supp_step1 = ttk.Frame(outer)
        ttk.Label(self.supp_step1, text="选择过往论政项目", font=SUBTITLE_FONT).pack(
            anchor="w", padx=5, pady=(0, 10)
        )
        supp_table_frame = ttk.Frame(self.supp_step1)
        supp_table_frame.pack(fill="both", expand=True)

        cols = ("id", "title", "created_at")
        self.supp_sessions_tree = ttk.Treeview(
            supp_table_frame, columns=cols, show="headings", height=12
        )
        self.supp_sessions_tree.heading("id", text="ID")
        self.supp_sessions_tree.heading("title", text="论政项目名称")
        self.supp_sessions_tree.heading("created_at", text="创建时间")
        self.supp_sessions_tree.column("id", width=70)
        self.supp_sessions_tree.column("title", width=300)
        self.supp_sessions_tree.column("created_at", width=200)
        self.supp_sessions_tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(supp_table_frame, orient="vertical", command=self.supp_sessions_tree.yview)
        self.supp_sessions_tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        supp_btn1 = ttk.Frame(self.supp_step1)
        supp_btn1.pack(pady=10)
        ttk.Button(supp_btn1, text="选择并继续", command=self._supplement_on_session_selected, width=14).pack(side="right", padx=5)
        ttk.Button(supp_btn1, text="返回", command=self._show_choice, width=10).pack(side="right")

        self.supp_step2 = ttk.Frame(outer)
        ttk.Label(self.supp_step2, text="勾选需改为不到场的到场人员", font=SUBTITLE_FONT).pack(
            anchor="w", padx=5, pady=(0, 10)
        )
        self.supp_check_vars = {}
        self.supp_check_frame = ttk.Frame(self.supp_step2)
        self.supp_check_frame.pack(fill="both", expand=True, padx=10, pady=10)

        supp_btn2 = ttk.Frame(self.supp_step2)
        supp_btn2.pack(pady=10)
        ttk.Button(supp_btn2, text="确认标记（自动新增后续不到场记录）", command=self._supplement_confirm_absent, width=26).pack(side="right", padx=5)
        ttk.Button(supp_btn2, text="上一步", command=self._show_supplement_step1, width=10).pack(side="right")

        self.supp_step3 = ttk.Frame(outer)
        ttk.Label(self.supp_step3, text="补抽到场人员", font=SUBTITLE_FONT).pack(
            anchor="w", padx=5, pady=(0, 10)
        )
        supp_info = ttk.Frame(self.supp_step3)
        supp_info.pack(pady=10)
        supp_info.columnconfigure(0, weight=1)
        supp_info.columnconfigure(1, weight=1)
        ttk.Label(supp_info, text="当前状态:").grid(row=0, column=0, sticky="e", padx=10, pady=5)
        self.supp_status_var = tk.StringVar()
        ttk.Label(supp_info, textvariable=self.supp_status_var).grid(row=0, column=1, sticky="w", padx=10, pady=5)
        ttk.Label(supp_info, text="已补抽到场人数:").grid(row=1, column=0, sticky="e", padx=10, pady=5)
        self.supp_present_var = tk.StringVar()
        ttk.Label(supp_info, textvariable=self.supp_present_var).grid(row=1, column=1, sticky="w", padx=10, pady=5)

        ttk.Separator(self.supp_step3, orient="horizontal").pack(fill="x", pady=5)

        supp_draw_info = ttk.Frame(self.supp_step3)
        supp_draw_info.pack(pady=10)
        supp_draw_info.columnconfigure(0, weight=1)
        supp_draw_info.columnconfigure(1, weight=1)
        ttk.Label(supp_draw_info, text="抽中人员姓名:").grid(row=0, column=0, sticky="e", padx=10, pady=5)
        self.supp_name_var = tk.StringVar()
        ttk.Label(supp_draw_info, textvariable=self.supp_name_var, font=("Microsoft YaHei", 14)).grid(row=0, column=1, sticky="w", padx=10, pady=5)
        ttk.Label(supp_draw_info, text="手机号:").grid(row=1, column=0, sticky="e", padx=10, pady=5)
        self.supp_phone_var = tk.StringVar()
        ttk.Label(supp_draw_info, textvariable=self.supp_phone_var).grid(row=1, column=1, sticky="w", padx=10, pady=5)

        supp_btns = ttk.Frame(self.supp_step3)
        supp_btns.pack(pady=10)
        self.supp_btn_draw = ttk.Button(supp_btns, text="抽签", command=self._supplement_draw_one, width=12)
        self.supp_btn_present = ttk.Button(supp_btns, text="到场", command=lambda: self._supplement_mark_result(True), width=12)
        self.supp_btn_absent = ttk.Button(supp_btns, text="不到场", command=lambda: self._supplement_mark_result(False), width=12)
        self.supp_btn_draw.grid(row=0, column=0, padx=10)
        self.supp_btn_present.grid(row=0, column=1, padx=10)
        self.supp_btn_absent.grid(row=0, column=2, padx=10)

        supp_btn3 = ttk.Frame(self.supp_step3)
        supp_btn3.pack(pady=15)
        ttk.Button(supp_btn3, text="完成补抽", command=self._supplement_finish, width=12).pack(side="right", padx=5)

    def _refresh_supplement_session_list(self):
        for i in self.supp_sessions_tree.get_children():
            self.supp_sessions_tree.delete(i)
        for s in self.app.db.get_completed_sessions():
            self.supp_sessions_tree.insert(
                "", "end",
                values=(s["id"], s["title"], s["created_at"]),
            )

    def _show_supplement_step1(self):
        self.supp_step2.pack_forget()
        self.supp_step3.pack_forget()
        self.supp_step1.pack(fill="both", expand=True)

    def _show_supplement_step2(self):
        self.supp_step1.pack_forget()
        self.supp_step3.pack_forget()
        self.supp_step2.pack(fill="both", expand=True)
        for w in self.supp_check_frame.winfo_children():
            w.destroy()
        self.supp_check_vars.clear()
        self.supp_present_logs_by_id = {}
        logs = self.app.db.get_present_logs(self.supplement_session_id)
        for r in logs:
            var = tk.BooleanVar(value=False)
            self.supp_check_vars[r["id"]] = var
            self.supp_present_logs_by_id[r["id"]] = r
            cb = ttk.Checkbutton(
                self.supp_check_frame,
                text=f"{r['name']}（{r['phone']}）",
                variable=var,
            )
            cb.pack(anchor="w", padx=5, pady=3)

    def _show_supplement_step3(self):
        self.supp_step1.pack_forget()
        self.supp_step2.pack_forget()
        self.supp_step3.pack(fill="both", expand=True)
        self.supplement_vacant_count = len([v for v in self.supp_check_vars.values() if v.get()])
        self.supplement_order_base = self.app.db.get_max_order_no(self.supplement_session_id)
        self.supp_present_count = 0
        self.supp_current_person = None
        self.supp_order_no = self.supplement_order_base
        people = self.app.db.get_all_people()
        self.supp_people_cache = list(people)
        self.supp_status_var.set(f"论政项目 ID {self.supplement_session_id} 补抽，需补 {self.supplement_vacant_count} 人")
        self.supp_present_var.set(f"0 / {self.supplement_vacant_count}")
        self.supp_name_var.set("")
        self.supp_phone_var.set("")
        self.supp_btn_draw["state"] = "normal"
        self.supp_btn_present["state"] = "disabled"
        self.supp_btn_absent["state"] = "disabled"

    def _supplement_on_session_selected(self):
        sel = self.supp_sessions_tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择一个论政项目")
            return
        self.supplement_session_id = int(self.supp_sessions_tree.item(sel[0], "values")[0])
        self._show_supplement_step2()

    def _supplement_confirm_absent(self):
        """补签时标记不到场：新增记录（到场情况=后续不到场，备注=专家后续不到场，进行再次抽选。），不修改原记录"""
        selected = [log_id for log_id, var in self.supp_check_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("提示", "请至少勾选一名需改为不到场的到场人员")
            return
        # 按原抽签顺序依次新增“后续不到场”记录
        selected_sorted = sorted(
            selected,
            key=lambda lid: self.supp_present_logs_by_id[lid]["order_no"] if lid in self.supp_present_logs_by_id else 0,
        )
        for log_id in selected_sorted:
            log_info = self.supp_present_logs_by_id.get(log_id)
            if log_info:
                self.app.db.add_supplement_absent_log(
                    self.supplement_session_id,
                    log_info["person_id"],
                )
        self._show_supplement_step3()

    def _ask_reason_dialog(self, prompt="请输入不到场理由"):
        dlg = tk.Toplevel(self)
        dlg.title(prompt)
        dlg.grab_set()
        dlg.resizable(False, False)
        container = ttk.Frame(dlg, padding=25)
        container.grid(row=0, column=0, sticky="nsew")
        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)
        ttk.Label(container, text="备注:").grid(row=0, column=0, padx=8, pady=8, sticky="nw")
        text = tk.Text(container, width=45, height=5, font=BASE_FONT)
        text.grid(row=0, column=1, padx=8, pady=8)
        res = {"value": None}

        def on_ok():
            res["value"] = text.get("1.0", "end").strip()
            dlg.destroy()

        def on_cancel():
            res["value"] = None
            dlg.destroy()

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(15, 0))
        ttk.Button(btn_frame, text="确定", command=on_ok, width=11).grid(row=0, column=0, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=11).grid(row=0, column=1, padx=10)
        dlg.wait_window()
        return res["value"]

    def _supplement_draw_one(self):
        if self.supp_present_count >= self.supplement_vacant_count:
            messagebox.showinfo("提示", "本次补抽已完成")
            return
        if not self.supp_people_cache:
            messagebox.showwarning("提示", "专家名库为空，无法抽签")
            return
        self.supp_current_person = random.choice(self.supp_people_cache)
        self.supp_order_no += 1
        self.supp_name_var.set(self.supp_current_person["name"])
        self.supp_phone_var.set(self.supp_current_person["phone"])
        self.supp_btn_present["state"] = "normal"
        self.supp_btn_absent["state"] = "normal"

    def _supplement_mark_result(self, present: bool):
        if self.supp_current_person is None:
            messagebox.showwarning("提示", "请先抽签")
            return
        reason = ""
        if not present:
            reason = self._ask_reason_dialog()
            if reason is None:
                return
        self.app.db.add_draw_log(
            self.supplement_session_id,
            self.supp_current_person["id"],
            self.supp_order_no,
            present,
            reason,
        )
        if present:
            self.supp_present_count += 1
            self.supp_present_var.set(f"{self.supp_present_count} / {self.supplement_vacant_count}")
        self.supp_current_person = None
        self.supp_name_var.set("")
        self.supp_phone_var.set("")
        self.supp_btn_present["state"] = "disabled"
        self.supp_btn_absent["state"] = "disabled"
        if self.supp_present_count >= self.supplement_vacant_count:
            self.supp_btn_draw["state"] = "disabled"
            self.supp_status_var.set(f"论政项目 ID {self.supplement_session_id} 补抽已完成")
            messagebox.showinfo("完成", "本次补抽流程已完成")
            # 补抽完成后发送邮件
            self.app.send_sessions_email(self, [self.supplement_session_id])

    def _supplement_finish(self):
        if self.supp_present_count < self.supplement_vacant_count:
            if not messagebox.askyesno("提示", "尚未补足全部名额，确定要结束吗？"):
                return
        self._show_choice()

    def set_buttons_state(self, started):
        if not started:
            self.btn_draw["state"] = "disabled"
            self.btn_present["state"] = "disabled"
            self.btn_absent["state"] = "disabled"
        else:
            self.btn_draw["state"] = "normal"
            self.btn_present["state"] = "disabled"
            self.btn_absent["state"] = "disabled"

    def start_new_session(self):
        title = self.title_var.get().strip()
        if not title:
            if not messagebox.askyesno("提示", "未填写论政项目名称，是否继续？"):
                return
        people = self.app.db.get_all_people()
        if not people:
            messagebox.showwarning("提示", "当前无专家名库，请先在专家名库中添加人员")
            return

        sid = self.app.db.create_session(title or "未命名论政项目")
        self.current_session_id = sid
        self.present_count = 0
        self.order_no = 0
        self.current_person = None
        self.people_cache = list(people)
        self.status_var.set(f"正在进行会话 ID: {sid}")
        self.present_var.set("0 / 3")
        self.name_var.set("")
        self.phone_var.set("")
        self.set_buttons_state(started=True)
        messagebox.showinfo("提示", "已开始新的抽签会话，可点击“抽签”")

    def draw_one(self):
        if self.current_session_id is None:
            messagebox.showwarning("提示", "请先点击“开始新抽签”")
            return
        if self.present_count >= 3:
            messagebox.showinfo("提示", "本次抽签已完成 3 名到场人员")
            return
        if not self.people_cache:
            messagebox.showwarning("提示", "专家名库为空，无法抽签")
            return

        self.current_person = random.choice(self.people_cache)
        self.order_no += 1
        self.name_var.set(self.current_person["name"])
        self.phone_var.set(self.current_person["phone"])
        self.btn_present["state"] = "normal"
        self.btn_absent["state"] = "normal"

    def mark_result(self, present: bool):
        if self.current_person is None:
            messagebox.showwarning("提示", "请先抽签")
            return
        reason = ""
        if not present:
            reason = self._ask_reason()
            if reason is None:
                return

        self.app.db.add_draw_log(
            self.current_session_id,
            self.current_person["id"],
            self.order_no,
            present,
            reason,
        )

        if present:
            self.present_count += 1
            self.present_var.set(f"{self.present_count} / 3")

        self.current_person = None
        self.name_var.set("")
        self.phone_var.set("")
        self.btn_present["state"] = "disabled"
        self.btn_absent["state"] = "disabled"

        if self.present_count >= 3:
            self.btn_draw["state"] = "disabled"
            self.status_var.set(
                f"会话 ID {self.current_session_id} 已完成 (3 人到场)"
            )
            messagebox.showinfo("完成", "本次抽签流程已完成 3 名到场人员")
            # 完成后发送邮件
            self.app.send_sessions_email(self, [self.current_session_id])
        else:
            self.status_var.set(
                f"会话 ID {self.current_session_id} 进行中，已到场 {self.present_count} 人"
            )

    def _ask_reason(self):
        dlg = tk.Toplevel(self)
        dlg.title("请输入不到场理由")
        dlg.grab_set()
        dlg.resizable(False, False)

        container = ttk.Frame(dlg, padding=25)
        container.grid(row=0, column=0, sticky="nsew")
        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)

        ttk.Label(container, text="不到场理由:").grid(
            row=0, column=0, padx=8, pady=8, sticky="nw"
        )
        text = tk.Text(container, width=45, height=6, font=BASE_FONT)
        text.grid(row=0, column=1, padx=8, pady=8)

        res = {"value": None}

        def on_ok():
            res["value"] = text.get("1.0", "end").strip()
            dlg.destroy()

        def on_cancel():
            res["value"] = None
            dlg.destroy()

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(15, 0))
        ttk.Button(btn_frame, text="确定", command=on_ok, width=11).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=11).grid(
            row=0, column=1, padx=10
        )

        dlg.wait_window()
        return res["value"]


class UsersFrame(ttk.Frame):
    """账号管理：仅管理员使用"""

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.build_ui()
        self.refresh()

    def build_ui(self):
        outer = ttk.Frame(self, padding=30)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 15))

        ttk.Label(top, text="账号管理（仅管理员）", font=SUBTITLE_FONT).pack(
            side="left", padx=8
        )

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)

        columns = ("id", "username", "email", "role")
        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=18,
        )
        self.tree.heading("id", text="ID")
        self.tree.heading("username", text="用户名")
        self.tree.heading("email", text="电子邮箱")
        self.tree.heading("role", text="角色")

        self.tree.column("id", width=70, anchor="center", stretch=False)
        self.tree.column("username", width=180, anchor="w", stretch=True)
        self.tree.column("email", width=220, anchor="w", stretch=True)
        self.tree.column("role", width=120, anchor="center", stretch=False)
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        # 下方中部按钮栏
        btn_bar = ttk.Frame(outer)
        btn_bar.pack(pady=15)

        self.btn_add = ttk.Button(btn_bar, text="新增账号", command=self.add_user, width=10)
        self.btn_edit = ttk.Button(btn_bar, text="编辑账号", command=self.edit_user, width=10)
        self.btn_del = ttk.Button(btn_bar, text="删除账号", command=self.delete_user, width=10)
        self.btn_set_admin = ttk.Button(
            btn_bar, text="设为管理员", command=lambda: self.change_role("admin"), width=11
        )
        self.btn_set_user = ttk.Button(
            btn_bar, text="设为普通用户", command=lambda: self.change_role("user"), width=12
        )
        self.btn_import = ttk.Button(
            btn_bar, text="Excel 导入", command=self.import_excel, width=11
        )
        self.btn_export = ttk.Button(
            btn_bar, text="Excel 导出", command=self.export_excel, width=11
        )

        for b in (
            self.btn_add,
            self.btn_edit,
            self.btn_del,
            self.btn_set_admin,
            self.btn_set_user,
            self.btn_import,
            self.btn_export,
        ):
            b.pack(side="left", padx=6)

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for row in self.app.db.get_all_users():
            role_text = "管理员" if row["role"] == "admin" else "普通用户"
            self.tree.insert(
                "",
                "end",
                values=(row["id"], row["username"], row["email"] or "", role_text),
            )

    def _get_selected_id_role(self):
        sel = self.tree.selection()
        if not sel:
            return None, None
        values = self.tree.item(sel[0], "values")
        uid = int(values[0])
        role_text = values[3]
        role = "admin" if role_text == "管理员" else "user"
        return uid, role

    def add_user(self):
        self._edit_dialog()

    def edit_user(self):
        uid, role = self._get_selected_id_role()
        if not uid:
            messagebox.showwarning("提示", "请先选择一个账号")
            return
        # 从 db 获取该用户
        users = {u["id"]: u for u in self.app.db.get_all_users()}
        if uid not in users:
            return
        u = users[uid]
        self._edit_dialog(uid, u["username"], u["role"])

    def _edit_dialog(self, user_id=None, username="", role="user"):
        dlg = tk.Toplevel(self)
        dlg.title("账号信息")
        dlg.grab_set()
        dlg.resizable(False, False)

        container = ttk.Frame(dlg, padding=25)
        container.grid(row=0, column=0, sticky="nsew")
        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)

        ttk.Label(container, text="用户名:").grid(
            row=0, column=0, padx=8, pady=10, sticky="e"
        )
        username_var = tk.StringVar(value=username)
        ttk.Entry(container, textvariable=username_var, width=28).grid(
            row=0, column=1, padx=8, pady=10, sticky="w"
        )

        ttk.Label(container, text="密码:").grid(
            row=1, column=0, padx=8, pady=10, sticky="e"
        )
        pwd_var = tk.StringVar()
        ttk.Entry(container, textvariable=pwd_var, show="*", width=28).grid(
            row=1, column=1, padx=8, pady=10, sticky="w"
        )
        ttk.Label(
            container, text="（留空则不修改密码）", foreground="gray"
        ).grid(row=2, column=1, padx=8, pady=(0, 10), sticky="w")

        ttk.Label(container, text="电子邮箱:").grid(
            row=3, column=0, padx=8, pady=10, sticky="e"
        )
        email_value = username if username and "@" in username else ""
        email_var = tk.StringVar(value=email_value)
        ttk.Entry(container, textvariable=email_var, width=28).grid(
            row=3, column=1, padx=8, pady=10, sticky="w"
        )

        ttk.Label(container, text="角色:").grid(
            row=4, column=0, padx=8, pady=10, sticky="e"
        )
        role_var = tk.StringVar(value="管理员" if role == "admin" else "普通用户")
        cb = ttk.Combobox(
            container,
            textvariable=role_var,
            values=["管理员", "普通用户"],
            state="readonly",
            width=25,
        )
        cb.grid(row=4, column=1, padx=8, pady=10, sticky="w")

        def on_ok():
            name = username_var.get().strip()
            pwd = pwd_var.get().strip()
            email = email_var.get().strip()
            role_choice = "admin" if role_var.get() == "管理员" else "user"
            if not name:
                messagebox.showwarning("提示", "用户名不能为空")
                return
            try:
                if user_id is None:
                    # 新增账号时密码必填
                    if not pwd:
                        messagebox.showwarning("提示", "新增账号时密码不能为空")
                        return
                    ok, msg = self.app.db.register_user(name, pwd, role=role_choice, email=email)
                    if not ok:
                        messagebox.showerror("失败", msg)
                        return
                else:
                    if pwd:
                        self.app.db.update_user(user_id, name, role_choice, pwd, email=email)
                    else:
                        self.app.db.update_user(user_id, name, role_choice, email=email)
                self.refresh()
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("错误", str(e))

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(18, 0))
        ttk.Button(btn_frame, text="确定", command=on_ok, width=11).grid(
            row=0, column=0, padx=10
        )
        ttk.Button(btn_frame, text="取消", command=dlg.destroy, width=11).grid(
            row=0, column=1, padx=10
        )

    def delete_user(self):
        uid, role = self._get_selected_id_role()
        if not uid:
            messagebox.showwarning("提示", "请先选择一个账号")
            return
        if not messagebox.askyesno("确认", "确定要删除该账号吗？"):
            return
        try:
            self.app.db.delete_user(uid, current_user_id=self.app.current_user["id"])
            self.refresh()
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def change_role(self, new_role):
        uid, _old_role = self._get_selected_id_role()
        if not uid:
            messagebox.showwarning("提示", "请先选择一个账号")
            return
        users = {u["id"]: u for u in self.app.db.get_all_users()}
        if uid not in users:
            return
        u = users[uid]
        try:
            self.app.db.update_user(uid, u["username"], new_role)
            self.refresh()
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def import_excel(self):
        if openpyxl is None:
            messagebox.showerror("错误", "请先安装 openpyxl 库")
            return
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm")],
        )
        if not path:
            return
        try:
            self.app.db.import_users_from_excel(path)
            self.refresh()
            messagebox.showinfo(
                "成功",
                "导入完成。\n表头示例：用户名 | 密码(明文) | 角色(admin/user)，空密码则不重置。",
            )
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")

    def export_excel(self):
        if openpyxl is None:
            messagebox.showerror("错误", "请先安装 openpyxl 库")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            title="导出账号列表",
            initialfile="users.xlsx",
        )
        if not path:
            return
        try:
            self.app.db.export_users_to_excel(path)
            messagebox.showinfo(
                "成功",
                "导出完成。\n注意：密码列为空，如需重置密码，可在 Excel 中填写后再导入。",
            )
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")


class MailConfigFrame(ttk.Frame):
    """邮件设置（仅管理员）"""

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.build_ui()
        self.load_config()

    def build_ui(self):
        outer = ttk.Frame(self, padding=30)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="邮件设置（仅管理员）", font=SUBTITLE_FONT).pack(
            anchor="w", padx=5, pady=(0, 15)
        )

        form = ttk.Frame(outer)
        form.pack(pady=10, padx=10, anchor="center")

        for i in range(2):
            form.columnconfigure(i, weight=1)

        ttk.Label(form, text="SMTP 服务器:").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        self.smtp_server_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.smtp_server_var, width=32).grid(
            row=0, column=1, sticky="w", padx=8, pady=6
        )

        ttk.Label(form, text="端口:").grid(row=1, column=0, sticky="e", padx=8, pady=6)
        self.smtp_port_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.smtp_port_var, width=10).grid(
            row=1, column=1, sticky="w", padx=8, pady=6
        )

        ttk.Label(form, text="发件人邮箱:").grid(row=2, column=0, sticky="e", padx=8, pady=6)
        self.from_addr_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.from_addr_var, width=32).grid(
            row=2, column=1, sticky="w", padx=8, pady=6
        )

        ttk.Label(form, text="登录用户名:").grid(row=3, column=0, sticky="e", padx=8, pady=6)
        self.username_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.username_var, width=32).grid(
            row=3, column=1, sticky="w", padx=8, pady=6
        )

        ttk.Label(form, text="登录密码:").grid(row=4, column=0, sticky="e", padx=8, pady=6)
        self.password_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.password_var, show="*", width=32).grid(
            row=4, column=1, sticky="w", padx=8, pady=6
        )

        self.use_ssl_var = tk.BooleanVar(value=True)
        self.use_tls_var = tk.BooleanVar(value=False)

        ssl_frame = ttk.Frame(outer)
        ssl_frame.pack(pady=5)
        ttk.Checkbutton(
            ssl_frame, text="使用 SSL", variable=self.use_ssl_var
        ).pack(side="left", padx=10)
        ttk.Checkbutton(
            ssl_frame, text="使用 STARTTLS", variable=self.use_tls_var
        ).pack(side="left", padx=10)

        btn_frame = ttk.Frame(outer)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="保存配置", command=self.save_config, width=12).pack(
            side="left", padx=10
        )

    def load_config(self):
        cfg = self.app.db.get_mail_config()
        self.smtp_server_var.set(cfg.get("smtp_server") or "")
        self.smtp_port_var.set(str(cfg.get("smtp_port") or ""))
        self.from_addr_var.set(cfg.get("from_addr") or "")
        self.username_var.set(cfg.get("username") or "")
        self.password_var.set(cfg.get("password") or "")
        self.use_ssl_var.set(cfg.get("use_ssl"))
        self.use_tls_var.set(cfg.get("use_tls"))

    def save_config(self):
        try:
            port = int(self.smtp_port_var.get() or 0)
        except ValueError:
            messagebox.showerror("错误", "端口必须是数字")
            return
        cfg = {
            "smtp_server": self.smtp_server_var.get().strip(),
            "smtp_port": port,
            "from_addr": self.from_addr_var.get().strip(),
            "username": self.username_var.get().strip(),
            "password": self.password_var.get(),
            "use_ssl": self.use_ssl_var.get(),
            "use_tls": self.use_tls_var.get(),
        }
        self.app.db.save_mail_config(cfg)
        messagebox.showinfo("成功", "邮件配置已保存")


class MainFrame(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.account_tab_added = False
        self.mail_tab_added = False
        self.build_ui()

    def build_ui(self):
        outer = ttk.Frame(self, padding=30)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x", pady=(0, 15))

        self.user_label = ttk.Label(top, text="")
        self.user_label.pack(side="left", padx=8)

        ttk.Button(top, text="退出登录", command=self.logout, width=11).pack(
            side="right", padx=8
        )

        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        self.people_frame = PeopleFrame(self.notebook, self.app)
        self.draw_frame = DrawFrame(self.notebook, self.app)
        self.logs_frame = LogsFrame(self.notebook, self.app)
        self.account_frame = UsersFrame(self.notebook, self.app)
        self.mail_frame = MailConfigFrame(self.notebook, self.app)

        self.notebook.add(self.people_frame, text="专家名库")
        self.notebook.add(self.draw_frame, text="抽签")
        self.notebook.add(self.logs_frame, text="日志")
        # 账号管理 / 邮件设置 Tab 在管理员登录后动态添加

    def refresh_user_info(self):
        u = self.app.current_user
        if u:
            text = f"当前用户：{u['username']}（{'管理员' if u['role']=='admin' else '普通用户'}）"
            self.user_label.configure(text=text)
            self.people_frame.set_admin_mode(u["role"] == "admin")

            # 管理员才显示账号管理 & 邮件设置 Tab
            if u["role"] == "admin":
                if not self.account_tab_added:
                    self.notebook.add(self.account_frame, text="账号管理")
                    self.account_tab_added = True
                if not self.mail_tab_added:
                    self.notebook.add(self.mail_frame, text="邮件设置")
                    self.mail_tab_added = True
            else:
                if self.account_tab_added:
                    self.notebook.forget(self.account_frame)
                    self.account_tab_added = False
                if self.mail_tab_added:
                    self.notebook.forget(self.mail_frame)
                    self.mail_tab_added = False
        else:
            self.user_label.configure(text="")
            if self.account_tab_added:
                self.notebook.forget(self.account_frame)
                self.account_tab_added = False
            if self.mail_tab_added:
                self.notebook.forget(self.mail_frame)
                self.mail_tab_added = False

    def logout(self):
        if messagebox.askyesno("确认", "确定要退出登录吗？"):
            self.app.current_user = None
            self.app.show_login_frame()


# -------------------- 应用主类 --------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # 初始窗口大小，允许缩放
        self.title("人员管理与抽签系统")
        self.geometry("1400x900")
        self.minsize(1000, 700)
        self.resizable(True, True)

        # 全局字体
        self.option_add("*Font", BASE_FONT)

        # Windows 下使用较为现代的 ttk 主题
        self.style = ttk.Style()
        try:
            if os.name == "nt":
                self.style.theme_use("vista")
        except Exception:
            pass

        # 放大 Notebook Tab 样式（选项卡标签）
        self.style.configure(
            "TNotebook.Tab",
            font=("Microsoft YaHei", 12, "bold"),
            padding=(20, 8),
        )

        # 全局按钮默认内边距稍大
        self.style.configure("TButton", padding=(12, 6))

        # Treeview 行高与字体
        self.style.configure(
            "Treeview",
            rowheight=34,
            font=BASE_FONT,
        )
        self.style.configure(
            "Treeview.Heading",
            font=("Microsoft YaHei", 11, "bold"),
        )

        self.db = Database()
        self.current_user = None

        self.container = ttk.Frame(self, padding=20)
        self.container.pack(fill="both", expand=True)

        self.login_frame = LoginFrame(self.container, self)
        self.register_frame = RegisterFrame(self.container, self)
        self.main_frame = MainFrame(self.container, self)

        self.show_login_frame()

        # 在 Tk 初始化完成后，根据实际 DPI 设置 tk 的 scaling
        self.after(0, self._adjust_scaling_for_dpi)

    def send_sessions_email(self, parent, session_ids):
        """将一个或多个论政项目的日志发送到所有管理员邮箱"""
        cfg = self.db.get_mail_config()
        if not cfg["smtp_server"] or not cfg["from_addr"]:
            messagebox.showerror(parent, "错误", "请先在“邮件设置”中配置 SMTP 服务器和发件人邮箱")
            return False
        admin_emails = self.db.get_admin_emails()
        if not admin_emails:
            messagebox.showerror(parent, "错误", "当前没有配置管理员邮箱，无法发送邮件")
            return False

        # 组装邮件内容
        lines = []
        for sid in session_ids:
            sessions = [s for s in self.db.get_sessions() if s["id"] == sid]
            if not sessions:
                continue
            s = sessions[0]
            lines.append(f"论政项目 ID: {s['id']}")
            lines.append(f"论政项目名称: {s['title'] or ''}")
            lines.append(f"创建时间: {s['created_at']}")
            lines.append("-" * 40)
            logs = self.db.get_session_logs(sid)
            for r in logs:
                if r["present"] == 1:
                    state = "到场"
                elif r["absent_reason"] == "专家后续不到场，进行再次抽选。":
                    state = "后续不到场"
                else:
                    state = "不到场"
                lines.append(
                    f"[{r['created_at']}] 姓名:{r['name']} 手机:{r['phone']} 到场情况:{state} 备注:{r['absent_reason'] or ''}"
                )
            lines.append("=" * 60)
            lines.append("")

        if not lines:
            messagebox.showwarning(parent, "提示", "未找到可发送的日志记录")
            return False

        body = "\n".join(lines)
        subject = "论政项目抽签结果"

        while True:
            try:
                messagebox.showinfo(parent, "提示", "正在发送结果至管理员邮箱，请稍候...")

                if cfg["use_ssl"]:
                    server = smtplib.SMTP_SSL(cfg["smtp_server"], cfg["smtp_port"] or 465, timeout=10)
                else:
                    server = smtplib.SMTP(cfg["smtp_server"], cfg["smtp_port"] or 25, timeout=10)

                try:
                    if cfg["use_tls"]:
                        server.starttls()
                    if cfg["username"]:
                        server.login(cfg["username"], cfg["password"] or "")

                    msg = EmailMessage()
                    msg["Subject"] = subject
                    msg["From"] = cfg["from_addr"]
                    msg["To"] = ", ".join(admin_emails)
                    msg.set_content(body)

                    server.send_message(msg)
                finally:
                    server.quit()

                messagebox.showinfo(parent, "成功", "抽签结果邮件发送成功")
                return True
            except Exception as e:
                retry = messagebox.askretrycancel(
                    "发送失败",
                    f"发送失败：{e}\n请检查网络和邮件配置。\n是否重试发送？",
                    parent=parent,
                )
                if not retry:
                    return False

    def _adjust_scaling_for_dpi(self):
        try:
            pixels_per_inch = self.winfo_fpixels("1i")
            scaling = pixels_per_inch / 72.0
            scaling = max(1.0, min(scaling, 2.0))
            self.tk.call("tk", "scaling", scaling)
        except Exception:
            pass

    def _show_frame(self, frame):
        for w in self.container.winfo_children():
            w.pack_forget()
        frame.pack(fill="both", expand=True)

    def show_login_frame(self):
        self._show_frame(self.login_frame)

    def show_register_frame(self):
        self._show_frame(self.register_frame)

    def show_main_frame(self):
        self.main_frame.refresh_user_info()
        self._show_frame(self.main_frame)


# -------------------- 入口 --------------------

if __name__ == "__main__":
    set_dpi_awareness()
    app = App()
    app.mainloop()