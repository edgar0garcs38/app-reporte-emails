import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, ttk
import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ── Tema global ───────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

REQUIRED_COLUMNS = [
    "id correo", "fecha recibido", "fecha respondido",
    "estado", "tiempo minutos"
]

# Paleta
C_BG      = "#0f1117"
C_SURFACE = "#1a1d27"
C_CARD    = "#22263a"
C_ACCENT  = "#4f8ef7"
C_ACCENT2 = "#f7614f"
C_GREEN   = "#4ade80"
C_YELLOW  = "#fbbf24"
C_TEXT    = "#e8eaf6"
C_MUTED   = "#6b7280"

mpl.rcParams.update({
    "figure.facecolor": C_BG,
    "axes.facecolor":   C_SURFACE,
    "axes.edgecolor":   C_MUTED,
    "axes.labelcolor":  C_TEXT,
    "text.color":       C_TEXT,
    "xtick.color":      C_MUTED,
    "ytick.color":      C_MUTED,
    "grid.color":       "#2a2d3a",
    "grid.linestyle":   "--",
    "font.family":      "sans-serif",
})


# ── Toast notification ────────────────────────────────────────────────────────
class Toast(ctk.CTkToplevel):
    """Notificacion flotante que aparece y desaparece sola."""

    TYPES = {
        "success": {"emoji": "✅", "color": C_GREEN,   "bg": "#0d2b1a"},
        "error":   {"emoji": "❌", "color": C_ACCENT2, "bg": "#2b100d"},
        "warning": {"emoji": "⚠️",  "color": C_YELLOW,  "bg": "#2b200d"},
        "info":    {"emoji": "ℹ️",  "color": C_ACCENT,  "bg": "#0d1a2b"},
    }

    def __init__(self, master, message: str, kind: str = "success", duration: int = 3000):
        super().__init__(master)
        cfg = self.TYPES.get(kind, self.TYPES["info"])

        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(fg_color=cfg["bg"])

        border = ctk.CTkFrame(self, fg_color=cfg["color"], width=4, corner_radius=0)
        border.pack(side="left", fill="y")

        inner = ctk.CTkFrame(self, fg_color=cfg["bg"], corner_radius=0)
        inner.pack(side="left", fill="both", expand=True, padx=14, pady=12)

        ctk.CTkLabel(
            inner,
            text=f"{cfg['emoji']}  {message}",
            font=("Segoe UI Semibold", 12),
            text_color=cfg["color"],
            wraplength=280,
            justify="left"
        ).pack(anchor="w")

        self.update_idletasks()

        pw = master.winfo_x() + master.winfo_width()
        ph = master.winfo_y() + master.winfo_height()
        w, h = 320, 64
        x = pw - w - 24
        y = ph - h - 24
        self.geometry(f"{w}x{h}+{x}+{y}")

        self._alpha = 0.0
        self._fade_in()
        self.after(duration, self._fade_out)

    def _fade_in(self):
        if self._alpha < 1.0:
            self._alpha = min(self._alpha + 0.08, 1.0)
            self.attributes("-alpha", self._alpha)
            self.after(16, self._fade_in)

    def _fade_out(self):
        if self._alpha > 0.0:
            self._alpha = max(self._alpha - 0.06, 0.0)
            self.attributes("-alpha", self._alpha)
            self.after(16, self._fade_out)
        else:
            self.destroy()


# ── KPI Card ──────────────────────────────────────────────────────────────────
class KPICard(ctk.CTkFrame):
    def __init__(self, master, emoji: str, title: str,
                 value: str = "—", color: str = C_ACCENT, **kw):
        super().__init__(master, fg_color=C_CARD, corner_radius=16, **kw)
        self.configure(border_width=1, border_color="#2e3250")

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=16, pady=(14, 0))

        ctk.CTkLabel(top, text=emoji, font=("Segoe UI", 20)).pack(side="left")
        ctk.CTkLabel(top, text=title, font=("Segoe UI", 11),
                     text_color=C_MUTED).pack(side="left", padx=8)

        self.val_label = ctk.CTkLabel(
            self, text=value,
            font=("Segoe UI Semibold", 30, "bold"),
            text_color=color
        )
        self.val_label.pack(padx=16, pady=(4, 16), anchor="w")

    def set(self, value: str):
        self.val_label.configure(text=value)


# ── App principal ─────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.df = None

        self.title("Email KPI Dashboard")
        self.geometry("1240x800")
        self.configure(fg_color=C_BG)
        self.minsize(960, 620)

        self._build_ui()

    def _build_ui(self):
        # ── SIDEBAR ──
        sidebar = ctk.CTkFrame(self, width=230, fg_color=C_SURFACE, corner_radius=0)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        logo_frame = ctk.CTkFrame(sidebar, fg_color=C_CARD, corner_radius=16)
        logo_frame.pack(padx=20, pady=(30, 16), fill="x")
        ctk.CTkLabel(logo_frame, text="📧", font=("Segoe UI", 36)).pack(pady=(12, 2))
        ctk.CTkLabel(logo_frame, text="Email KPI",
                     font=("Segoe UI Bold", 16, "bold"),
                     text_color=C_TEXT).pack()
        ctk.CTkLabel(logo_frame, text="Dashboard v2.0",
                     font=("Segoe UI", 10),
                     text_color=C_MUTED).pack(pady=(0, 12))

        ctk.CTkButton(
            sidebar, text="📂   Cargar archivo",
            font=("Segoe UI Semibold", 13), height=44,
            fg_color=C_ACCENT, hover_color="#3a74e0",
            corner_radius=12, command=self.load_file
        ).pack(padx=20, pady=(4, 6), fill="x")

        ctk.CTkButton(
            sidebar, text="💾   Exportar Excel",
            font=("Segoe UI Semibold", 13), height=44,
            fg_color=C_CARD, hover_color="#2a2f4a",
            border_width=1, border_color="#3a4060",
            corner_radius=12, command=self.export_excel
        ).pack(padx=20, pady=6, fill="x")

        ctk.CTkFrame(sidebar, height=1, fg_color="#2a2d3a").pack(
            fill="x", padx=20, pady=16)

        ctk.CTkLabel(sidebar, text="ARCHIVO ACTIVO",
                     font=("Segoe UI", 9, "bold"),
                     text_color=C_MUTED).pack(padx=20, anchor="w")

        self.file_label = ctk.CTkLabel(
            sidebar, text="Ninguno cargado",
            font=("Segoe UI", 11), text_color=C_MUTED,
            wraplength=190, justify="left"
        )
        self.file_label.pack(padx=20, pady=(4, 0), anchor="w")

        self.rows_label = ctk.CTkLabel(
            sidebar, text="",
            font=("Segoe UI", 10), text_color=C_MUTED
        )
        self.rows_label.pack(padx=20, pady=(2, 0), anchor="w")

        self.progress = ctk.CTkProgressBar(sidebar, mode="indeterminate",
                                           progress_color=C_ACCENT,
                                           fg_color=C_CARD, height=4)
        self.progress.pack(padx=20, pady=(16, 0), fill="x")
        self.progress.set(0)

        self.progress_label = ctk.CTkLabel(
            sidebar, text="", font=("Segoe UI", 10), text_color=C_ACCENT)
        self.progress_label.pack(padx=20, anchor="w")

        # ── AREA PRINCIPAL ──
        main = ctk.CTkFrame(self, fg_color=C_BG, corner_radius=0)
        main.pack(side="left", fill="both", expand=True)

        hdr = ctk.CTkFrame(main, fg_color="transparent")
        hdr.pack(fill="x", padx=28, pady=(26, 4))
        ctk.CTkLabel(hdr, text="Resumen de correos 📬",
                     font=("Segoe UI Bold", 22, "bold"),
                     text_color=C_TEXT).pack(side="left")

        self.status_dot = ctk.CTkLabel(hdr, text="● Sin datos",
                                       font=("Segoe UI", 11),
                                       text_color=C_MUTED)
        self.status_dot.pack(side="right", padx=4)

        cards_row = ctk.CTkFrame(main, fg_color="transparent")
        cards_row.pack(fill="x", padx=28, pady=10)
        cards_row.columnconfigure((0, 1, 2, 3), weight=1, uniform="col")

        self.card_total  = KPICard(cards_row, "📨", "Total",         color=C_TEXT)
        self.card_resp   = KPICard(cards_row, "✅", "Respondidos",   color=C_GREEN)
        self.card_noresp = KPICard(cards_row, "🔴", "Sin respuesta", color=C_ACCENT2)
        self.card_avg    = KPICard(cards_row, "⏱️", "Promedio",      color=C_ACCENT)

        for i, c in enumerate([self.card_total, self.card_resp,
                                self.card_noresp, self.card_avg]):
            c.grid(row=0, column=i, padx=6, sticky="ew")

        self.tabs = ctk.CTkTabview(
            main, fg_color=C_SURFACE,
            segmented_button_fg_color=C_CARD,
            segmented_button_selected_color=C_ACCENT,
            segmented_button_selected_hover_color="#3a74e0",
            border_width=0
        )
        self.tabs.pack(fill="both", expand=True, padx=28, pady=(6, 20))
        self.tabs.add("📋  Tabla")
        self.tabs.add("📊  Graficos")

        self._build_table_tab()
        self._build_chart_tab()

        # Barra de estado inferior
        statusbar = ctk.CTkFrame(main, fg_color=C_SURFACE, height=28, corner_radius=0)
        statusbar.pack(fill="x", side="bottom")
        statusbar.pack_propagate(False)
        self.statusbar_label = ctk.CTkLabel(
            statusbar,
            text="👋  Bienvenido — carga un archivo CSV o Excel para comenzar",
            font=("Segoe UI", 10), text_color=C_MUTED
        )
        self.statusbar_label.pack(side="left", padx=16)

    # ── Tab Tabla ─────────────────────────────────────────────────────────────
    def _build_table_tab(self):
        tab = self.tabs.tab("📋  Tabla")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.Treeview",
                        background=C_SURFACE, foreground=C_TEXT,
                        rowheight=32, fieldbackground=C_SURFACE,
                        borderwidth=0, font=("Segoe UI", 11))
        style.configure("Dark.Treeview.Heading",
                        background=C_CARD, foreground=C_ACCENT,
                        font=("Segoe UI Semibold", 11, "bold"), relief="flat")
        style.map("Dark.Treeview",
                  background=[("selected", "#2e3a6e")],
                  foreground=[("selected", C_TEXT)])

        self.table_placeholder = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=12)
        self.table_placeholder.pack(fill="both", expand=True)
        ctk.CTkLabel(
            self.table_placeholder,
            text="📂  Carga un archivo para ver los datos aqui",
            font=("Segoe UI", 13), text_color=C_MUTED
        ).pack(expand=True)

        self.table_frame = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=12)

        self.tree = ttk.Treeview(self.table_frame, style="Dark.Treeview", show="headings")
        vsb = ttk.Scrollbar(self.table_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

    # ── Tab Graficos ──────────────────────────────────────────────────────────
    def _build_chart_tab(self):
        tab = self.tabs.tab("📊  Graficos")

        self.chart_placeholder = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=12)
        self.chart_placeholder.pack(fill="both", expand=True)
        ctk.CTkLabel(
            self.chart_placeholder,
            text="📊  Los graficos apareceran aqui despues de cargar datos",
            font=("Segoe UI", 13), text_color=C_MUTED
        ).pack(expand=True)

        self.chart_frame = ctk.CTkFrame(tab, fg_color=C_SURFACE, corner_radius=12)

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _set_status(self, text: str, color: str = C_MUTED):
        self.statusbar_label.configure(text=text, text_color=color)

    def _set_dot(self, text: str, color: str = C_MUTED):
        self.status_dot.configure(text=text, text_color=color)

    def _start_loading(self, msg: str = "Procesando..."):
        self.progress_label.configure(text=msg)
        self.progress.start()

    def _stop_loading(self):
        self.progress.stop()
        self.progress.set(0)
        self.progress_label.configure(text="")

    def _toast(self, message: str, kind: str = "success"):
        Toast(self, message, kind)

    # ── Cargar archivo ────────────────────────────────────────────────────────
    def load_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV / Excel", "*.csv *.xlsx")])
        if not path:
            return

        self._start_loading("⏳ Cargando archivo...")
        self._set_status("📂  Leyendo archivo, por favor espera...", C_ACCENT)
        self.update()

        try:
            df = pd.read_csv(path) if path.endswith(".csv") else pd.read_excel(path)

            if list(df.columns) != REQUIRED_COLUMNS:
                self._stop_loading()
                self._toast(
                    f"Columnas incorrectas.\nEsperadas: {', '.join(REQUIRED_COLUMNS)}",
                    "error"
                )
                self._set_status("❌  Error — columnas incorrectas en el archivo", C_ACCENT2)
                return

            df["fecha recibido"]   = pd.to_datetime(df["fecha recibido"],   errors="coerce")
            df["fecha respondido"] = pd.to_datetime(df["fecha respondido"], errors="coerce")
            self.df = df

            fname = os.path.basename(path)
            self.file_label.configure(text=f"✅  {fname}", text_color=C_GREEN)
            self.rows_label.configure(
                text=f"📝  {len(df)} registros cargados", text_color=C_MUTED)

            self._update_kpis()
            self._show_table()
            self._show_charts()
            self._stop_loading()

            responded = len(df[df["estado"] == "Respondido"])
            rate = (responded / len(df) * 100) if len(df) > 0 else 0

            if rate >= 80:
                self._toast(f"Archivo cargado! Tasa de respuesta: {rate:.0f}% — Excelente!", "success")
            elif rate >= 50:
                self._toast(f"Archivo cargado. Tasa de respuesta: {rate:.0f}%", "info")
            else:
                self._toast(f"Archivo cargado. Tasa de respuesta baja: {rate:.0f}%", "warning")

            self._set_dot(f"● {len(df)} correos", C_GREEN)
            self._set_status(
                f"✅  Cargado: {fname}  —  {len(df)} registros  —  {rate:.0f}% respondidos",
                C_GREEN
            )

        except Exception as e:
            self._stop_loading()
            self._toast(f"Error inesperado: {str(e)}", "error")
            self._set_status(f"❌  Error: {str(e)}", C_ACCENT2)

    # ── KPIs ──────────────────────────────────────────────────────────────────
    def _update_kpis(self):
        df = self.df
        total     = len(df)
        responded = len(df[df["estado"] == "Respondido"])
        no_resp   = len(df[df["estado"] == "Sin Respuesta"])
        avg_time  = df["tiempo minutos"].mean()

        self.card_total.set(str(total))
        self.card_resp.set(str(responded))
        self.card_noresp.set(str(no_resp))
        self.card_avg.set(f"{avg_time:.1f} min")

    # ── Tabla ─────────────────────────────────────────────────────────────────
    def _show_table(self):
        self.table_placeholder.pack_forget()
        self.table_frame.pack(fill="both", expand=True, padx=4, pady=4)

        self.tree.delete(*self.tree.get_children())
        cols = list(self.df.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col.title())
            self.tree.column(col, width=150, anchor="center")

        for i, (_, row) in enumerate(self.df.iterrows()):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=list(row), tags=(tag,))

        self.tree.tag_configure("even", background=C_SURFACE)
        self.tree.tag_configure("odd",  background="#1e2130")

    # ── Graficos ──────────────────────────────────────────────────────────────
    def _show_charts(self):
        self.chart_placeholder.pack_forget()
        for w in self.chart_frame.winfo_children():
            w.destroy()
        self.chart_frame.pack(fill="both", expand=True, padx=4, pady=4)

        df = self.df
        responded = len(df[df["estado"] == "Respondido"])
        no_resp   = len(df[df["estado"] == "Sin Respuesta"])

        fig, axes = plt.subplots(1, 2, figsize=(9, 3.8))
        fig.patch.set_facecolor(C_BG)
        fig.subplots_adjust(wspace=0.35, left=0.06, right=0.97, top=0.88, bottom=0.12)

        ax1 = axes[0]
        wedges, texts, autotexts = ax1.pie(
            [responded, no_resp],
            labels=["Respondido", "Sin Respuesta"],
            colors=[C_ACCENT, C_ACCENT2],
            autopct="%1.1f%%", startangle=140,
            wedgeprops=dict(width=0.55, edgecolor=C_BG, linewidth=2),
            pctdistance=0.75
        )
        for t in texts:     t.set_color(C_MUTED);  t.set_fontsize(10)
        for a in autotexts: a.set_color(C_TEXT);   a.set_fontsize(10)
        ax1.set_title("Distribucion de estado", color=C_TEXT,
                      fontsize=12, pad=10, fontweight="bold")

        ax2 = axes[1]
        bars = ax2.bar(
            ["Respondido", "Sin Respuesta"],
            [responded, no_resp],
            color=[C_ACCENT, C_ACCENT2],
            width=0.5, edgecolor=C_BG, linewidth=1.5
        )
        ax2.set_title("Conteo por estado", color=C_TEXT,
                      fontsize=12, pad=10, fontweight="bold")
        ax2.yaxis.grid(True, alpha=0.3)
        ax2.set_axisbelow(True)
        for bar in bars:
            h = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width() / 2, h + 0.3,
                     str(int(h)), ha="center", va="bottom",
                     color=C_TEXT, fontsize=12, fontweight="bold")

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)

    # ── Exportar ──────────────────────────────────────────────────────────────
    def export_excel(self):
        if self.df is None:
            self._toast("Primero carga un archivo de datos", "warning")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="reporte_kpi.xlsx"
        )
        if not path:
            return

        try:
            self._start_loading("💾 Exportando...")
            self._set_status("💾  Guardando archivo Excel...", C_ACCENT)
            self.update()

            self.df.to_excel(path, index=False)
            self._stop_loading()

            fname = os.path.basename(path)
            self._toast(f"Exportado con exito! {fname}", "success")
            self._set_status(f"💾  Guardado correctamente → {fname}", C_GREEN)

        except Exception as e:
            self._stop_loading()
            self._toast(f"Error al exportar: {str(e)}", "error")


if __name__ == "__main__":
    app = App()
    app.mainloop()
