# -*- coding: utf-8 -*-
"""
POS Sector - Dupliciranje kategorija
------------------------------------
Single-file tkinter alat za dupliciranje cijele kategorije
(subkategorije + artikli + ArticleGoods) unutar POS Sector baze.

- Odaberi izvornu kategoriju iz dropdowna
- Dupliciraj u NOVU kategoriju (upišeš ime) ili u POSTOJEĆU
- Sve ide unutar transakcije -> ako bilo što pukne, ROLLBACK

Pokretanje:  py pos_category_duplicator.py
Compile:     py -m PyInstaller --onefile --noconsole --icon=logo.ico pos_category_duplicator.py
"""

import threading
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import pyodbc
except ImportError:
    pyodbc = None

# ----------------------------------------------------------------------
# Tema (isti stil kao ostali POS Sector alati)
# ----------------------------------------------------------------------
PRIMARY      = "#1565c0"
PRIMARY_DARK = "#0d47a1"
ACTION       = "#f57c00"
ACTION_DARK  = "#e65100"
BG           = "#ffffff"
PANEL        = "#f5f7fa"
TEXT         = "#1a1a1a"
MUTED        = "#5f6b7a"
OK_GREEN     = "#2e7d32"
ERR_RED      = "#c62828"

DRIVER = "ODBC Driver 18 for SQL Server"


def conn_str(server, database):
    return (
        f"DRIVER={{{DRIVER}}};SERVER={server};DATABASE={database};"
        f"Trusted_Connection=yes;TrustServerCertificate=yes;"
    )


# ======================================================================
#  DB helperi
# ======================================================================
def list_databases(server):
    """Vrati listu korisničkih baza na serveru (bez sistemskih)."""
    cn = pyodbc.connect(conn_str(server, "master"), autocommit=True, timeout=10)
    try:
        cur = cn.cursor()
        cur.execute(
            "SELECT name FROM sys.databases "
            "WHERE database_id > 4 AND state = 0 ORDER BY name"
        )
        return [r[0] for r in cur.fetchall()]
    finally:
        cn.close()


def soft_delete_col(conn, table):
    """Detektiraj kako se zove soft-delete kolona (Deleted / IsDeleted)."""
    cur = conn.cursor()
    cur.execute(
        "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME = ? AND COLUMN_NAME IN ('Deleted','IsDeleted')",
        table,
    )
    cols = [r[0] for r in cur.fetchall()]
    if "Deleted" in cols:
        return "Deleted"
    if "IsDeleted" in cols:
        return "IsDeleted"
    return None


def load_categories(conn):
    """Vrati [(id_str, name), ...] aktivnih kategorija."""
    col = soft_delete_col(conn, "Categories")
    where = f"WHERE {col} = 0" if col else ""
    cur = conn.cursor()
    cur.execute(
        f"SELECT CAST(Id AS NVARCHAR(50)) AS Id, Name "
        f"FROM Categories {where} ORDER BY Name"
    )
    return [(r.Id, r.Name) for r in cur.fetchall()]


def category_counts(conn, cat_id, skip_deleted):
    """Koliko subkategorija i artikala bi se kopiralo."""
    cur = conn.cursor()
    sub_f = "AND Deleted = 0" if skip_deleted else ""
    cur.execute(
        f"SELECT COUNT(*) FROM SubCategories "
        f"WHERE Category_Id = ? {sub_f}",
        cat_id,
    )
    subs = cur.fetchone()[0]

    art_f = "AND a.Deleted = 0 AND s.Deleted = 0" if skip_deleted else ""
    cur.execute(
        f"SELECT COUNT(*) FROM Articles a "
        f"JOIN SubCategories s ON a.SubCategory_Id = s.Id "
        f"WHERE s.Category_Id = ? {art_f}",
        cat_id,
    )
    arts = cur.fetchone()[0]
    return subs, arts


def create_new_category(conn, source_id, new_name):
    """
    Kreiraj novu kategoriju kopiranjem reda izvorne kategorije
    (sve kolone osim Id/Name/Order ostaju iste). Vrati novi Id (string).
    Dinamički čita kolone -> radi bez obzira na verziju sheme.
    """
    cur = conn.cursor()
    cur.execute(
        "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME = 'Categories' "
        "AND COLUMNPROPERTY(OBJECT_ID('dbo.Categories'), COLUMN_NAME, 'IsComputed') = 0 "
        "ORDER BY ORDINAL_POSITION"
    )
    cols = [r[0] for r in cur.fetchall()]
    if not cols:
        raise RuntimeError("Ne mogu pročitati kolone tablice Categories.")

    collist = ", ".join(f"[{c}]" for c in cols)
    parts = []
    for c in cols:
        lc = c.lower()
        if lc == "id":
            parts.append("@NewId")
        elif lc == "name":
            parts.append("@NewName")
        elif lc == "order":
            parts.append("(SELECT ISNULL(MAX([Order]), 0) + 1 FROM Categories)")
        else:
            parts.append(f"[{c}]")
    selectlist = ", ".join(parts)

    sql = f"""
SET NOCOUNT ON;
DECLARE @NewId UNIQUEIDENTIFIER = NEWID();
DECLARE @NewName NVARCHAR(255) = ?;
DECLARE @SourceId UNIQUEIDENTIFIER = ?;
INSERT INTO Categories ({collist})
SELECT {selectlist} FROM Categories WHERE Id = @SourceId;
SELECT CAST(@NewId AS NVARCHAR(50)) AS NewId;
"""
    cur.execute(sql, new_name, source_id)
    while cur.description is None:
        if not cur.nextset():
            break
    return cur.fetchone()[0]


def duplicate_into_category(conn, old_cat_id, new_cat_id, skip_deleted):
    """
    Tvoj provjereni query: dupliciraj subkategorije, artikle i ArticleGoods
    iz old -> new kategorije. Vrati (subs, arts, goods) brojeve.
    """
    sub_filter = "AND s0.Deleted = 0" if skip_deleted else ""
    art_filter = "AND a.Deleted = 0" if skip_deleted else ""

    sql = f"""
SET NOCOUNT ON;
DECLARE @OldCategoryId UNIQUEIDENTIFIER = ?;
DECLARE @NewCategoryId UNIQUEIDENTIFIER = ?;

-- 1) Subkategorije -> nove GUID-ove
DECLARE @SubMap TABLE (OldId UNIQUEIDENTIFIER, NewId UNIQUEIDENTIFIER);
INSERT INTO @SubMap (OldId, NewId)
SELECT s0.Id, NEWID()
FROM SubCategories s0
WHERE s0.Category_Id = @OldCategoryId {sub_filter};

INSERT INTO SubCategories
    (Id, [Order], Printer, Name, Deleted, Storage_Id, Category_Id, Tag, ExtraPrinter1, ExtraPrinter2)
SELECT
    m.NewId, s.[Order], s.Printer, s.Name, s.Deleted, s.Storage_Id,
    @NewCategoryId, s.Tag, s.ExtraPrinter1, s.ExtraPrinter2
FROM SubCategories s
INNER JOIN @SubMap m ON s.Id = m.OldId;

-- 2) Mapiranje stari artikal -> novi artikal
DECLARE @ArticleMap TABLE (OldId UNIQUEIDENTIFIER, NewId UNIQUEIDENTIFIER, NewSubId UNIQUEIDENTIFIER);
INSERT INTO @ArticleMap (OldId, NewId, NewSubId)
SELECT a.Id, NEWID(), sm.NewId
FROM Articles a
INNER JOIN @SubMap sm ON a.SubCategory_Id = sm.OldId
WHERE 1 = 1 {art_filter};

-- 3) Artikli
INSERT INTO Articles
    (Id, Deleted, Image, Name, Tag, ArticleNumber, [Order], Price, SubCategory_Id, BarCode, Code, ReturnFee, FreeModifiers)
SELECT
    am.NewId, a.Deleted, a.Image, a.Name, a.Tag, a.ArticleNumber, a.[Order],
    a.Price, am.NewSubId, a.BarCode, a.Code, a.ReturnFee, a.FreeModifiers
FROM Articles a
INNER JOIN @ArticleMap am ON a.Id = am.OldId;

-- 4) ArticleGoods
DECLARE @AgCount INT;
INSERT INTO ArticleGoods (Id, Quantity, ValidFrom, ValidUntil, Article_Id, Good_Id)
SELECT NEWID(), ag.Quantity, ag.ValidFrom, ag.ValidUntil, am.NewId, ag.Good_Id
FROM ArticleGoods ag
INNER JOIN @ArticleMap am ON ag.Article_Id = am.OldId;
SET @AgCount = @@ROWCOUNT;

SELECT
    (SELECT COUNT(*) FROM @SubMap)     AS SubCount,
    (SELECT COUNT(*) FROM @ArticleMap) AS ArtCount,
    @AgCount                           AS GoodsCount;
"""
    cur = conn.cursor()
    cur.execute(sql, old_cat_id, new_cat_id)
    while cur.description is None:
        if not cur.nextset():
            break
    row = cur.fetchone()
    return row.SubCount, row.ArtCount, row.GoodsCount


# ======================================================================
#  GUI
# ======================================================================
class App:
    def __init__(self, root):
        self.root = root
        self.conn = None            # konekcija na odabranu bazu
        self.server = ""
        self.source_cats = []       # [(id, name)]
        self.target_cats = []       # [(id, name)]

        root.title("POS Sector — Dupliciranje kategorija")
        root.geometry("640x720")
        root.configure(bg=BG)
        root.minsize(560, 640)

        self._build_styles()
        self._build_ui()

        if pyodbc is None:
            self.log("⚠  pyodbc nije instaliran.  Pokreni:  py -m pip install pyodbc", ERR_RED)
            self.btn_connect.config(state="disabled")

    # ------------------------------------------------------------------
    def _build_styles(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TCombobox", fieldbackground="white", background="white")
        st.configure("Hdr.TLabel", background=BG, foreground=PRIMARY,
                     font=("Segoe UI", 10, "bold"))
        st.configure("Muted.TLabel", background=BG, foreground=MUTED,
                     font=("Segoe UI", 9))

    def _build_ui(self):
        # --- Header ---
        head = tk.Frame(self.root, bg=PRIMARY, height=56)
        head.pack(fill="x")
        head.pack_propagate(False)
        tk.Label(head, text="POS Sector — Dupliciranje kategorija",
                 bg=PRIMARY, fg="white",
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=18)

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=14)

        # --- 1) Konekcija ---
        self._section(body, "1 · Konekcija")
        crow = tk.Frame(body, bg=BG)
        crow.pack(fill="x", pady=(2, 4))
        tk.Label(crow, text="Server:", bg=BG, fg=TEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        self.var_server = tk.StringVar(value="localhost\\SQLEXPRESS")
        tk.Entry(crow, textvariable=self.var_server, width=26,
                 relief="solid", bd=1).pack(side="left", padx=6)
        self.btn_connect = tk.Button(crow, text="Spoji", bg=PRIMARY, fg="white",
                                     activebackground=PRIMARY_DARK, relief="flat",
                                     font=("Segoe UI", 9, "bold"), padx=14,
                                     command=self.on_connect)
        self.btn_connect.pack(side="left", padx=4)

        drow = tk.Frame(body, bg=BG)
        drow.pack(fill="x", pady=(0, 4))
        tk.Label(drow, text="Baza:", bg=BG, fg=TEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        self.cmb_db = ttk.Combobox(drow, state="disabled", width=34)
        self.cmb_db.pack(side="left", padx=6)
        self.cmb_db.bind("<<ComboboxSelected>>", lambda e: self.on_db_selected())

        self.lbl_status = tk.Label(body, text="Nisi spojen.", bg=BG, fg=MUTED,
                                   font=("Segoe UI", 9, "italic"), anchor="w")
        self.lbl_status.pack(fill="x", pady=(2, 10))

        # --- 2) Izvorna kategorija ---
        self._section(body, "2 · Izvorna kategorija")
        self.cmb_source = ttk.Combobox(body, state="disabled")
        self.cmb_source.pack(fill="x", pady=(2, 2))
        self.cmb_source.bind("<<ComboboxSelected>>", lambda e: self.on_source_selected())
        self.lbl_counts = tk.Label(body, text="—", bg=BG, fg=MUTED,
                                   font=("Segoe UI", 9), anchor="w")
        self.lbl_counts.pack(fill="x", pady=(0, 10))

        # --- 3) Odredište ---
        self._section(body, "3 · Odredište")
        self.var_target = tk.StringVar(value="new")

        r1 = tk.Frame(body, bg=BG)
        r1.pack(fill="x")
        tk.Radiobutton(r1, text="Nova kategorija:", variable=self.var_target,
                       value="new", bg=BG, fg=TEXT, font=("Segoe UI", 9),
                       activebackground=BG, command=self._toggle_target).pack(side="left")
        self.var_newname = tk.StringVar()
        self.ent_newname = tk.Entry(r1, textvariable=self.var_newname, width=30,
                                    relief="solid", bd=1)
        self.ent_newname.pack(side="left", padx=6)

        r2 = tk.Frame(body, bg=BG)
        r2.pack(fill="x", pady=(4, 0))
        tk.Radiobutton(r2, text="Postojeća:", variable=self.var_target,
                       value="existing", bg=BG, fg=TEXT, font=("Segoe UI", 9),
                       activebackground=BG, command=self._toggle_target).pack(side="left")
        self.cmb_target = ttk.Combobox(r2, state="disabled", width=32)
        self.cmb_target.pack(side="left", padx=6)

        # --- Opcije ---
        self.var_skip = tk.BooleanVar(value=True)
        tk.Checkbutton(body, text="Preskoči obrisane (Deleted = 1)",
                       variable=self.var_skip, bg=BG, fg=TEXT,
                       activebackground=BG, font=("Segoe UI", 9),
                       command=self.on_source_selected).pack(anchor="w", pady=(10, 8))

        # --- Akcija ---
        self.btn_run = tk.Button(body, text="DUPLICIRAJ KATEGORIJU",
                                 bg=ACTION, fg="white", activebackground=ACTION_DARK,
                                 relief="flat", font=("Segoe UI", 11, "bold"),
                                 pady=10, state="disabled", command=self.on_duplicate)
        self.btn_run.pack(fill="x", pady=(2, 10))

        # --- Log ---
        self._section(body, "Log")
        logfr = tk.Frame(body, bg=BG)
        logfr.pack(fill="both", expand=True)
        self.txt = tk.Text(logfr, height=8, bg=PANEL, fg=TEXT, relief="solid", bd=1,
                           font=("Consolas", 9), wrap="word", state="disabled")
        self.txt.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(logfr, command=self.txt.yview)
        sb.pack(side="right", fill="y")
        self.txt.config(yscrollcommand=sb.set)

    def _section(self, parent, title):
        ttk.Label(parent, text=title, style="Hdr.TLabel").pack(anchor="w", pady=(2, 0))
        tk.Frame(parent, bg="#e0e4ea", height=1).pack(fill="x", pady=(2, 4))

    # ------------------------------------------------------------------
    #  Helperi
    # ------------------------------------------------------------------
    def log(self, msg, color=None):
        self.root.after(0, self._append_log, msg, color)

    def _append_log(self, msg, color):
        self.txt.config(state="normal")
        if color:
            tag = f"c{color}"
            self.txt.tag_config(tag, foreground=color)
            self.txt.insert("end", msg + "\n", tag)
        else:
            self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.config(state="disabled")

    def set_status(self, text, color=MUTED):
        self.root.after(0, lambda: self.lbl_status.config(text=text, fg=color))

    def set_busy(self, busy):
        state = "disabled" if busy else "normal"
        self.btn_connect.config(state=state)
        self.btn_run.config(state="disabled" if busy else self.btn_run["state"])

    def _toggle_target(self):
        if self.var_target.get() == "new":
            self.ent_newname.config(state="normal")
            self.cmb_target.config(state="disabled")
        else:
            self.ent_newname.config(state="disabled")
            self.cmb_target.config(state="readonly" if self.target_cats else "disabled")

    # ------------------------------------------------------------------
    #  Akcije
    # ------------------------------------------------------------------
    def on_connect(self):
        self.server = self.var_server.get().strip()
        if not self.server:
            messagebox.showwarning("Server", "Upiši server.")
            return
        self.set_status("Spajam se…")
        threading.Thread(target=self._do_connect, daemon=True).start()

    def _do_connect(self):
        try:
            dbs = list_databases(self.server)
            def done():
                self.cmb_db.config(values=dbs, state="readonly")
                self.set_status(f"Spojen na {self.server} — {len(dbs)} baza.", OK_GREEN)
                self.log(f"✓ Spojen na {self.server}")
            self.root.after(0, done)
        except Exception as e:
            self.set_status("Greška kod spajanja.", ERR_RED)
            self.log(f"✗ Greška: {e}", ERR_RED)

    def on_db_selected(self):
        db = self.cmb_db.get()
        if not db:
            return
        self.set_status(f"Učitavam bazu {db}…")
        threading.Thread(target=self._do_load_db, args=(db,), daemon=True).start()

    def _do_load_db(self, db):
        try:
            if self.conn:
                try:
                    self.conn.close()
                except Exception:
                    pass
            self.conn = pyodbc.connect(conn_str(self.server, db),
                                       autocommit=True, timeout=15)
            cats = load_categories(self.conn)

            def done():
                self.source_cats = cats
                self.target_cats = cats
                names = [n for _, n in cats]
                self.cmb_source.config(values=names, state="readonly")
                self.cmb_target.config(values=names)
                self.cmb_source.set("")
                self.cmb_target.set("")
                self.lbl_counts.config(text="—")
                self.btn_run.config(state="normal")
                self._toggle_target()
                self.set_status(f"Baza {db} — {len(cats)} kategorija.", OK_GREEN)
                self.log(f"✓ Učitano {len(cats)} kategorija iz baze '{db}'")
            self.root.after(0, done)
        except Exception as e:
            self.set_status("Greška kod učitavanja baze.", ERR_RED)
            self.log(f"✗ Greška: {e}", ERR_RED)

    def on_source_selected(self):
        idx = self.cmb_source.current()
        if idx < 0 or not self.conn:
            return
        cat_id, name = self.source_cats[idx]
        threading.Thread(target=self._do_counts, args=(cat_id, name),
                         daemon=True).start()

    def _do_counts(self, cat_id, name):
        try:
            subs, arts = category_counts(self.conn, cat_id, self.var_skip.get())
            self.root.after(0, lambda: self.lbl_counts.config(
                text=f"„{name}“  →  {subs} subkategorija, {arts} artikala"))
        except Exception as e:
            self.log(f"✗ Greška kod brojanja: {e}", ERR_RED)

    def on_duplicate(self):
        idx = self.cmb_source.current()
        if idx < 0:
            messagebox.showwarning("Izvor", "Odaberi izvornu kategoriju.")
            return
        src_id, src_name = self.source_cats[idx]

        mode = self.var_target.get()
        if mode == "new":
            new_name = self.var_newname.get().strip()
            if not new_name:
                messagebox.showwarning("Ime", "Upiši ime nove kategorije.")
                return
            target_desc = f"NOVA kategorija „{new_name}“"
            tgt_id = None
        else:
            tidx = self.cmb_target.current()
            if tidx < 0:
                messagebox.showwarning("Odredište", "Odaberi postojeću kategoriju.")
                return
            tgt_id, tgt_name = self.target_cats[tidx]
            if tgt_id == src_id:
                messagebox.showwarning("Odredište",
                                       "Izvor i odredište ne smiju biti isti.")
                return
            target_desc = f"postojeću „{tgt_name}“"
            new_name = None

        if not messagebox.askyesno(
            "Potvrda",
            f"Dupliciram „{src_name}“  →  {target_desc}\n\n"
            f"Sve ide unutar transakcije (rollback ako pukne). Nastaviti?"
        ):
            return

        self.set_busy(True)
        self.btn_run.config(state="disabled")
        threading.Thread(target=self._do_duplicate,
                         args=(src_id, src_name, tgt_id, new_name),
                         daemon=True).start()

    def _do_duplicate(self, src_id, src_name, tgt_id, new_name):
        conn = self.conn
        skip = self.var_skip.get()
        try:
            conn.autocommit = False
            if tgt_id is None:
                self.log(f"→ Kreiram novu kategoriju „{new_name}“…")
                tgt_id = create_new_category(conn, src_id, new_name)
                self.log(f"  ✓ Kreirana (Id={tgt_id})")

            self.log(f"→ Dupliciram sadržaj iz „{src_name}“…")
            subs, arts, goods = duplicate_into_category(conn, src_id, tgt_id, skip)
            conn.commit()
            self.log(f"✓ GOTOVO — {subs} subkat., {arts} artikala, "
                     f"{goods} ArticleGoods.  COMMIT.", OK_GREEN)
            self.set_status("Dupliciranje uspješno.", OK_GREEN)
            self.root.after(0, self._refresh_after)
        except Exception as e:
            try:
                conn.rollback()
            except Exception:
                pass
            self.log(f"✗ GREŠKA — sve poništeno (ROLLBACK): {e}", ERR_RED)
            self.set_status("Greška — rollback.", ERR_RED)
        finally:
            try:
                conn.autocommit = True
            except Exception:
                pass
            self.root.after(0, lambda: self.set_busy(False))
            self.root.after(0, lambda: self.btn_run.config(state="normal"))

    def _refresh_after(self):
        """Osvježi liste kategorija nakon uspješnog dupliciranja."""
        try:
            cats = load_categories(self.conn)
            self.source_cats = cats
            self.target_cats = cats
            names = [n for _, n in cats]
            self.cmb_source.config(values=names)
            self.cmb_target.config(values=names)
            self.var_newname.set("")
        except Exception as e:
            self.log(f"(ne mogu osvježiti liste: {e})", MUTED)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
