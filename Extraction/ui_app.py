"""
ui_app.py
─────────
Financial Statement Pipeline — Customtkinter UI (compact layout).
DuckDB / Parquet logging handled entirely by pipeline.py subprocess.
UI only displays stdout and parses run_id from pipeline output.
"""

import customtkinter as ctk
import subprocess
import threading
from pathlib import Path
import os
import sys
import re
import polars as pl

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ════════════════════════════════════════════════════════════════════════
# DYNAMIC PATHS
# ════════════════════════════════════════════════════════════════════════
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
REPORTS       = os.path.join(BASE_DIR, "03_Validated_Report")
XLSX          = os.path.join(BASE_DIR, "Master", "standard_coa_master.xlsx")
PROMPTS       = os.path.join(BASE_DIR, "prompts")
OUTPUT        = os.path.join(BASE_DIR, "04_Validated_output")
MANUAL_OUTPUT = os.path.join(BASE_DIR, "05_Manual_Validation_Required")
COA_MAPPING   = os.path.join(BASE_DIR, "COA_Mapping")
LOG_DIR       = os.path.join(BASE_DIR, "logs")

# ════════════════════════════════════════════════════════════════════════
# MODEL OPTIONS PER PROVIDER  (normalization models)
# ════════════════════════════════════════════════════════════════════════
OPENAI_MODELS = [
    "gpt-5.5", "gpt-4o-mini", "gpt-4o",
    "gpt-4.1", "gpt-4.1-mini", "gpt-5.4-mini",
]

GEMINI_MODELS = [
    "gemini-3.5-flash", "gemini-3.1-flash-lite-preview",
    "gemini-2.5-flash", "gemini-2.5-pro",
]

CLAUDE_MODELS = [
    "Claude Opus 4.5",
    "Claude Sonnet 4.6",
    "Claude Sonnet 4.5",
    "Claude Opus 4.8",
]

PROVIDER_MODELS = {
    "openai": OPENAI_MODELS,
    "gemini": GEMINI_MODELS,
    "claude": CLAUDE_MODELS,
}

PROVIDER_ENV = {
    "openai": "OPENAI_API_KEY",
    "gemini": "GEMINI_API_KEY",
    "claude": "ANTHROPIC_API_KEY",
}

# ════════════════════════════════════════════════════════════════════════
# PAGE IDENTIFICATION MODELS  (used by --id-model flag)
# Grouped visually: Gemini | Claude | OpenAI
# ════════════════════════════════════════════════════════════════════════
ID_MODELS = [
    # ── Gemini ──────────────────────────────────────────────────────
    "gemini-3.5-flash",
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    
    # ── Claude ──────────────────────────────────────────────────────
    "claude-haiku-4-5-20251001",
    "claude-sonnet-4-6",
    "claude-sonnet-4-5",
    "claude-opus-4-5",
    "claude-opus-4-8",
    # ── OpenAI ──────────────────────────────────────────────────────
    "gpt-4o",
    "gpt-4o-mini",
    "gpt-4.1",
    "gpt-4.1-mini",
    "gpt-5.4-mini",
    "gpt-5.5",
]

# Which env-var does each ID model need?
_ID_MODEL_ENV = {
    "gemini": "GEMINI_API_KEY",
    "claude": "ANTHROPIC_API_KEY",
    "gpt":    "OPENAI_API_KEY",
}

def _env_for_id_model(model: str) -> str:
    """Return the env-var name required for the given id-model string."""
    m = model.lower()
    if m.startswith("gemini"):
        return _ID_MODEL_ENV["gemini"]
    if m.startswith("claude"):
        return _ID_MODEL_ENV["claude"]
    if m.startswith("gpt") or m.startswith("o1") or m.startswith("o3"):
        return _ID_MODEL_ENV["gpt"]
    return ""


TIER_DESCRIPTIONS = {
    "openai": {
        "free": "Plain synchronous calls, full cost per call.",
        "paid": "Sync + auto prompt caching (~50% off cached prefix). "
                "Tick Batch for async 50% off everything.",
    },
    "gemini": {
        "free": "Plain calls only — no caching (avoids 429 on free keys).",
        "paid": "Sync + explicit context caching (~75% off cached tokens). "
                "Tick Batch for async 50% off everything.",
    },
    "claude": {
        "free": "Not available. Claude API is paid only.",
        "paid": "Sync + prompt caching. Tick Batch for async ~50% off.",
    },
}

BATCH_ON_HINT     = "✓ Batch ON — async job, ~50% cost, results in ≤24 h."
BATCH_OFF_HINT    = "Sync mode — immediate results per PDF."
LLM_ID_ONLY_HINT  = "🔍 ID-Only mode: LLM finds pages → Python slices PDFs → NO normalization."
LLM_ID_FULL_HINT  = "🔍 LLM finds pages → Python slices → then runs full normalization."




def _load_env_from_registry():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
            for env_var in ["OPENAI_API_KEY", "GEMINI_API_KEY", "ANTHROPIC_API_KEY"]:
                if not os.environ.get(env_var):
                    try:
                        value, _ = winreg.QueryValueEx(key, env_var)
                        os.environ[env_var] = value
                    except FileNotFoundError:
                        pass
    except Exception:
        pass


_load_env_from_registry()


# ════════════════════════════════════════════════════════════════════════
# LOG VIEWER WINDOW
# ════════════════════════════════════════════════════════════════════════

class LogViewerWindow(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Pipeline Log History")
        self.geometry("1100x600")
        self.minsize(800, 400)
        self.resizable(True, True)
        self._build_ui()
        self._query()   # auto-load on open

    def _build_ui(self):
        # ── Toolbar ──────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=12, pady=(10, 4))

        ctk.CTkLabel(
            toolbar, text="Filter (Run ID or keyword):",
            font=("Arial", 11), anchor="w",
        ).pack(side="left", padx=(0, 6))

        self.filter_var = ctk.StringVar()
        self.filter_entry = ctk.CTkEntry(
            toolbar, textvariable=self.filter_var, width=200,
        )
        self.filter_entry.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            toolbar, text="🔍 Search", width=90,
            command=self._query,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            toolbar, text="🔄 Refresh", width=90,
            command=self._query,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            toolbar, text="🗑 Clear Log File", width=120,
            fg_color="#7F1D1D", hover_color="#991B1B",
            command=self._clear_logs,
        ).pack(side="left", padx=(0, 8))

        self.status_label = ctk.CTkLabel(
            toolbar, text="", font=("Arial", 10, "italic"),
            text_color="gray60", anchor="w",
        )
        self.status_label.pack(side="left", padx=(12, 0))

        # ── Log text box ─────────────────────────────────────────────
        self.text = ctk.CTkTextbox(
            self, font=("Consolas", 10), wrap="none",
        )
        self.text.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    def _query(self):
        self.text.delete("1.0", "end")
        parquet_path = os.path.join(LOG_DIR, "pipeline_logs.parquet")

        if not os.path.isfile(parquet_path):
            self.text.insert("end", "No log file found yet.\nRun the pipeline first.\n")
            self.status_label.configure(text="No log file found.")
            return

        try:
            df = pl.read_parquet(parquet_path)

            run_filter = self.filter_var.get().strip()
            if run_filter:
                df = df.filter(
                    pl.col("Remarks").str.contains(run_filter, literal=False)
                    | pl.col("ProcessingId").str.contains(run_filter, literal=False)
                )

            total = len(df)
            shown = df.tail(500)

            for row in shown.iter_rows(named=True):
                time_str  = str(row.get("Time", ""))[:19]
                log_id    = row.get("ProcessingLogId", "")
                pid       = row.get("ProcessingId", "") or ""
                stage     = row.get("Stage", "")
                remarks   = row.get("Remarks", "")
                self.text.insert(
                    "end",
                    f"{time_str}  [#{log_id:<5}]  [{pid:<45}]  [{stage}]  {remarks}\n"
                )

            self.text.see("end")
            filter_note = f' (filtered: "{run_filter}")' if run_filter else ""
            self.status_label.configure(
                text=f"{total} row(s) total{filter_note} — showing last {min(500, total)}",
                text_color="gray60",
            )

        except Exception as ex:
            self.text.insert("end", f"Read error: {ex}\n")
            self.status_label.configure(text=f"Error: {ex}", text_color="#F87171")

    def _clear_logs(self):
        parquet_path = os.path.join(LOG_DIR, "pipeline_logs.parquet")
        try:
            # Write an empty parquet with the correct schema
            import polars as pl
            from log_writer import _empty_df
            _empty_df().write_parquet(parquet_path)
            self.text.delete("1.0", "end")
            self.text.insert("end", "Log file cleared.\n")
            self.status_label.configure(text="Log cleared.", text_color="#3FB950")
        except Exception as ex:
            self.status_label.configure(text=f"Clear failed: {ex}", text_color="#F87171")

# ════════════════════════════════════════════════════════════════════════
# MAIN APP
# ════════════════════════════════════════════════════════════════════════

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Financial Statement Pipeline")
        self.geometry("780x780")          # slightly taller for extra ID-model row
        self.minsize(720, 680)
        self.resizable(True, True)
        self.proc = None
        self._build_ui()
        self._on_provider_change()

    # ════════════════════════════════════════════════════════════════════
    # UI BUILD
    # ════════════════════════════════════════════════════════════════════
    def _build_ui(self):
        top_panel = ctk.CTkFrame(self, fg_color="transparent")
        top_panel.pack(fill="x", padx=14, pady=(8, 0))

        title_row = ctk.CTkFrame(top_panel, fg_color="transparent")
        title_row.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(
            title_row,
            text="📊  Financial Statement Pipeline  —  Government CAFR / ACFR  →  JSON / CSV / Excel",
            font=("Arial", 13, "bold"),
        ).pack(side="left")

        # ── ROW 1: Provider + Model ───────────────────────────────────
        row1 = ctk.CTkFrame(top_panel, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 2))

        ctk.CTkLabel(row1, text="Provider:", width=70, anchor="w",
                     font=("Arial", 11, "bold")).pack(side="left")

        self.provider_var = ctk.StringVar(value="gemini")
        for prov in ["openai", "gemini", "claude"]:
            ctk.CTkRadioButton(
                row1, text=prov.capitalize(),
                variable=self.provider_var, value=prov,
                command=self._on_provider_change,
            ).pack(side="left", padx=(0, 10))

        ctk.CTkLabel(row1, text="   Model:", width=60, anchor="w",
                     font=("Arial", 11, "bold")).pack(side="left")

        self.model_var = ctk.StringVar(value=GEMINI_MODELS[0])
        self.model_menu = ctk.CTkOptionMenu(
            row1, values=GEMINI_MODELS, variable=self.model_var, width=220,
        )
        self.model_menu.pack(side="left", padx=4)

        # ── ROW 2: Tier ───────────────────────────────────────────────
        row2 = ctk.CTkFrame(top_panel, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 2))

        ctk.CTkLabel(row2, text="API Tier:", width=70, anchor="w",
                     font=("Arial", 11, "bold")).pack(side="left")

        self.tier_var = ctk.StringVar(value="free")
        self.free_radio = ctk.CTkRadioButton(
            row2, text="Free", variable=self.tier_var, value="free",
            command=self._on_tier_change,
        )
        self.free_radio.pack(side="left", padx=(0, 12))

        self.paid_radio = ctk.CTkRadioButton(
            row2, text="Paid (caching / cheaper)",
            variable=self.tier_var, value="paid",
            command=self._on_tier_change,
        )
        self.paid_radio.pack(side="left")

        self.tier_hint_label = ctk.CTkLabel(
            row2,
            text=TIER_DESCRIPTIONS["gemini"]["free"],
            text_color="gray60",
            font=("Arial", 10, "italic"),
            wraplength=340, justify="left",
        )
        self.tier_hint_label.pack(side="left", padx=(16, 0))

        # ── ROW 3: Batch + LLM Page ID checkbox ──────────────────────
        row3 = ctk.CTkFrame(top_panel, fg_color="transparent")
        row3.pack(fill="x", pady=(0, 2))

        self.batch_var = ctk.BooleanVar(value=False)
        self.batch_checkbox = ctk.CTkCheckBox(
            row3,
            text="⚡ Batch Processing  (Paid only — 50% cost, async ≤24 h)",
            variable=self.batch_var, onvalue=True, offvalue=False,
            command=self._on_batch_toggle, state="disabled",
            font=("Arial", 11),
        )
        self.batch_checkbox.pack(side="left", padx=(0, 20))

        self.llm_id_var = ctk.BooleanVar(value=True)
        self.llm_id_checkbox = ctk.CTkCheckBox(
            row3,
            text="🔍 LLM Page Identification",
            variable=self.llm_id_var, onvalue=True, offvalue=False,
            command=self._on_llm_id_toggle,
            font=("Arial", 11),
        )
        self.llm_id_checkbox.pack(side="left")

        # ── ROW 3b: ID-Model selector  ← NEW ─────────────────────────
        self.id_model_frame = ctk.CTkFrame(top_panel, fg_color="transparent")
        self.id_model_frame.pack(fill="x", pady=(0, 2))

        ctk.CTkLabel(
            self.id_model_frame,
            text="  ID Model:",
            width=80, anchor="w",
            font=("Arial", 11, "bold"),
        ).pack(side="left", padx=(20, 0))

        self.id_model_var = ctk.StringVar(value="gemini-3.5-flash")
        self.id_model_menu = ctk.CTkOptionMenu(
            self.id_model_frame,
            values=ID_MODELS,
            variable=self.id_model_var,
            width=200,
            command=self._on_id_model_change,
        )
        self.id_model_menu.pack(side="left", padx=4)

        self.id_model_hint = ctk.CTkLabel(
            self.id_model_frame,
            text="",
            font=("Arial", 10, "italic"),
            text_color="gray60",
            anchor="w",
        )
        self.id_model_hint.pack(side="left", padx=(8, 0))

        # Initialise hint text
        self._on_id_model_change(self.id_model_var.get())

        # ── ROW 4: Hints ─────────────────────────────────────────────
        row4 = ctk.CTkFrame(top_panel, fg_color="transparent")
        row4.pack(fill="x", pady=(0, 2))

        self.batch_hint_label = ctk.CTkLabel(
            row4,
            text="Select 'Paid' to enable Batch Processing.",
            text_color="gray55",
            font=("Arial", 10, "italic"),
            width=360, anchor="w",
        )
        self.batch_hint_label.pack(side="left", padx=(24, 0))

        self.llm_id_hint_label = ctk.CTkLabel(
            row4,
            text=LLM_ID_FULL_HINT,
            text_color="gray70",
            font=("Arial", 10, "italic"),
            anchor="w",
        )
        self.llm_id_hint_label.pack(side="left", padx=(4, 0))

        # ── ROW 5: ID-Only sub-option ─────────────────────────────────
        self.llm_id_only_frame = ctk.CTkFrame(top_panel, fg_color="transparent")
        self.llm_id_only_frame.pack(fill="x", pady=(0, 2))

        self.llm_id_only_var = ctk.BooleanVar(value=False)
        self.llm_id_only_checkbox = ctk.CTkCheckBox(
            self.llm_id_only_frame,
            text="📄 Page Extraction Only  (identify + slice pages, skip normalization)",
            variable=self.llm_id_only_var, onvalue=True, offvalue=False,
            command=self._on_llm_id_only_toggle,
            font=("Arial", 10),
            fg_color="#7C3AED", hover_color="#6D28D9",
        )
        self.llm_id_only_checkbox.pack(side="left", padx=(24, 0))

        self.llm_id_only_hint = ctk.CTkLabel(
            self.llm_id_only_frame,
            text="",
            text_color="gray55",
            font=("Arial", 10, "italic"),
        )
        self.llm_id_only_hint.pack(side="left", padx=(8, 0))

        # ── API key status ────────────────────────────────────────────
        self.key_status_label = ctk.CTkLabel(
            top_panel, text="", font=("Arial", 10), text_color="gray70", anchor="w",
        )
        self.key_status_label.pack(fill="x", pady=(0, 2))

        # ── Run / Stop / View Logs ────────────────────────────────────
        btn_frame = ctk.CTkFrame(top_panel, fg_color="transparent")
        btn_frame.pack(pady=(2, 4))

        self.run_button = ctk.CTkButton(
            btn_frame, text="▶  Run Pipeline",
            command=self.run_pipeline,
            font=("Arial", 12, "bold"),
            width=160, height=32,
            fg_color="#2563EB", hover_color="#1D4ED8",
        )
        self.run_button.pack(side="left", padx=6)

        self.stop_button = ctk.CTkButton(
            btn_frame, text="⏹  Stop",
            command=self.stop_pipeline,
            font=("Arial", 12),
            width=100, height=32,
            fg_color="#DC2626", hover_color="#B91C1C",
            state="disabled",
        )
        self.stop_button.pack(side="left", padx=6)

        self.log_view_button = ctk.CTkButton(
            btn_frame, text="📊  View Log History",
            command=self._open_log_viewer,
            font=("Arial", 12),
            width=160, height=32,
            fg_color="#065F46", hover_color="#047857",
        )
        self.log_view_button.pack(side="left", padx=6)

        # ── Log status bar ────────────────────────────────────────────
        self.log_status_label = ctk.CTkLabel(
            top_panel,
            text="📝 Logs: not started",
            font=("Arial", 10, "italic"),
            text_color="gray55",
            anchor="w",
        )
        self.log_status_label.pack(fill="x", padx=4, pady=(0, 2))

        # ── Log box ───────────────────────────────────────────────────
        log_label_row = ctk.CTkFrame(self, fg_color="transparent")
        log_label_row.pack(fill="x", padx=14, pady=(2, 2))
        ctk.CTkLabel(
            log_label_row, text="Pipeline Log:",
            font=("Arial", 11, "bold"),
        ).pack(side="left")

        self.log_box = ctk.CTkTextbox(
            self,
            font=("Consolas", 10),
            wrap="word",
        )
        self.log_box.pack(padx=14, pady=(0, 10), fill="both", expand=True)

    # ════════════════════════════════════════════════════════════════════
    # EVENT HANDLERS
    # ════════════════════════════════════════════════════════════════════
    def _on_provider_change(self):
        provider = self.provider_var.get()
        models = PROVIDER_MODELS.get(provider, [])
        if models:
            self.model_var.set(models[0])
            self.model_menu.configure(values=models)

        if provider == "claude":
            self.tier_var.set("paid")
            self.free_radio.configure(state="disabled")
            self.paid_radio.configure(state="normal")
        else:
            self.free_radio.configure(state="normal")
            self.paid_radio.configure(state="normal")

        self._on_tier_change()
        self._update_key_status(provider)

    def _on_tier_change(self, *_):
        prov = self.provider_var.get()
        tier = self.tier_var.get()
        self.tier_hint_label.configure(text=TIER_DESCRIPTIONS.get(prov, {}).get(tier, ""))

        if tier == "paid":
            self.batch_checkbox.configure(state="normal")
            if self.batch_var.get():
                self.batch_hint_label.configure(text=BATCH_ON_HINT, text_color="#3FB950")
            else:
                self.batch_hint_label.configure(text=BATCH_OFF_HINT, text_color="gray70")
        else:
            self.batch_var.set(False)
            self.batch_checkbox.configure(state="disabled")
            self.batch_hint_label.configure(
                text="Select 'Paid' to enable Batch Processing.",
                text_color="gray55",
            )

    def _on_batch_toggle(self, *_):
        if self.batch_var.get():
            self.batch_hint_label.configure(text=BATCH_ON_HINT, text_color="#3FB950")
        else:
            self.batch_hint_label.configure(text=BATCH_OFF_HINT, text_color="gray70")

    def _on_llm_id_toggle(self, *_):
        if self.llm_id_var.get():
            self.id_model_frame.pack(fill="x", pady=(0, 2))
            self.llm_id_only_frame.pack(fill="x", pady=(0, 2))
            self.llm_id_hint_label.configure(text=LLM_ID_FULL_HINT, text_color="gray70")
        else:
            self.llm_id_only_var.set(False)
            self.id_model_frame.pack_forget()
            self.llm_id_only_frame.pack_forget()
            self.llm_id_hint_label.configure(
                text="Python keyword-based page detection (no LLM call).",
                text_color="gray55",
            )

    def _on_llm_id_only_toggle(self, *_):
        if self.llm_id_only_var.get():
            self.llm_id_only_hint.configure(text=LLM_ID_ONLY_HINT, text_color="#A78BFA")
            self.batch_var.set(False)
            self.batch_checkbox.configure(state="disabled")
            self.batch_hint_label.configure(
                text="Batch disabled in Extraction-Only mode.", text_color="gray55",
            )
        else:
            self.llm_id_only_hint.configure(text="", text_color="gray55")
            if self.tier_var.get() == "paid":
                self.batch_checkbox.configure(state="normal")
                self.batch_hint_label.configure(text=BATCH_OFF_HINT, text_color="gray70")

    def _on_id_model_change(self, model: str):
        """Update the hint label to show which API key is needed."""
        env_var = _env_for_id_model(model)
        if not env_var:
            self.id_model_hint.configure(text="", text_color="gray60")
            return
        if os.environ.get(env_var, "").strip():
            self.id_model_hint.configure(
                text=f"🟢 needs {env_var} — set ✓",
                text_color="#3FB950",
            )
        else:
            self.id_model_hint.configure(
                text=f"🔴 needs {env_var} — NOT set!",
                text_color="#F87171",
            )

    def _update_key_status(self, prov: str):
        env_var = PROVIDER_ENV.get(prov, "")
        if not env_var:
            self.key_status_label.configure(text="")
            return
        if os.environ.get(env_var, "").strip():
            self.key_status_label.configure(
                text=f"🟢 {env_var} is set.", text_color="#3FB950",
            )
        else:
            self.key_status_label.configure(
                text=f"🔴 {env_var} NOT set — pipeline will fail.",
                text_color="#F87171",
            )

    def _open_log_viewer(self):
        LogViewerWindow(self)

    # ════════════════════════════════════════════════════════════════════
    # PIPELINE LAUNCH
    # ════════════════════════════════════════════════════════════════════
    def run_pipeline(self):
        provider           = self.provider_var.get()
        model              = self.model_var.get()
        tier               = self.tier_var.get()
        use_batch          = self.batch_var.get()
        coa_mapping_folder = COA_MAPPING
        id_model           = self.id_model_var.get()   # ← selected ID model

        if provider == "claude":
            tier = "paid"
            self.tier_var.set("paid")

        errors = []
        if not os.path.isdir(REPORTS):
            errors.append(f"Reports folder not found: {REPORTS}")
        if not os.path.isdir(PROMPTS):
            errors.append(f"Prompts folder not found: {PROMPTS}")
        if not os.path.isfile(XLSX):
            errors.append(f"COA master xlsx not found: {XLSX}")
        if provider in PROVIDER_ENV and not os.environ.get(PROVIDER_ENV[provider], "").strip():
            errors.append(f"{PROVIDER_ENV[provider]} env var not set.")
        if use_batch and tier != "paid":
            errors.append("Batch Processing requires Paid tier.")
        if coa_mapping_folder and not os.path.isdir(coa_mapping_folder):
            errors.append(f"COA Mapping folder not found: {coa_mapping_folder}")

        # Validate the ID model's API key when LLM page ID is enabled
        if self.llm_id_var.get():
            id_env = _env_for_id_model(id_model)
            if id_env and not os.environ.get(id_env, "").strip():
                errors.append(
                    f"ID model '{id_model}' needs {id_env} — not set in environment."
                )

        if errors:
            self.log_box.delete("1.0", "end")
            for e in errors:
                self._log(f"❌ {e}")
            return

        Path(OUTPUT).mkdir(parents=True, exist_ok=True)
        Path(MANUAL_OUTPUT).mkdir(parents=True, exist_ok=True)
        Path(LOG_DIR).mkdir(parents=True, exist_ok=True)

        pipeline_script = os.path.join(BASE_DIR, "pipeline.py")
        cmd = [
            sys.executable, "-u", pipeline_script,
            "--folder",        REPORTS,
            "--xlsx",          XLSX,
            "--prompts",       PROMPTS,
            "--output",        OUTPUT,
            "--manual-output", MANUAL_OUTPUT,
            "--provider",      provider,
            "--model",         model,
            "--tier",          tier,
        ]
        if coa_mapping_folder:
            cmd += ["--coa-mapping", coa_mapping_folder]
        if use_batch:
            cmd.append("--batch")
        if self.llm_id_var.get():
            cmd.append("--llm-page-id")
            cmd += ["--id-model", id_model]   # ← use selected model, not hardcoded
            if self.llm_id_only_var.get():
                cmd.append("--skip-normalization")
        cmd += ["--max-tokens", "65536" if provider == "gemini" else "16000"]

        self.log_box.delete("1.0", "end")
        # mode = "⚡ BATCH (async, ≤24 h, 50% cost)" if use_batch else f"🔄 SYNC (tier={tier})"
        # self._log("=" * 60)
        # self._log(f"  Provider  : {provider.upper()}")
        # self._log(f"  Model     : {model}")
        # self._log(f"  Tier      : {tier}")
        # self._log(f"  Mode      : {mode}")
        # self._log(f"  ID Model  : {id_model if self.llm_id_var.get() else 'Python (no LLM)'}")
        # self._log(f"  COA Map   : {coa_mapping_folder}" if coa_mapping_folder
        #           else "⚠ No COA Mapping folder set.")
        # self._log("=" * 60)

        self.log_status_label.configure(
            text="📝 Waiting for pipeline to start logging…",
            text_color="gray55",
        )

        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")

        threading.Thread(
            target=self._run_subprocess, args=(cmd,), daemon=True,
        ).start()

    # ════════════════════════════════════════════════════════════════════
    # SUBPROCESS
    # ════════════════════════════════════════════════════════════════════
    def _run_subprocess(self, cmd: list):
        try:
            self.proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                cwd=BASE_DIR,
                env=os.environ.copy(),
            )

            for line in self.proc.stdout:
                self.log_box.insert("end", line)
                self.log_box.see("end")
                self.log_box.update_idletasks()

                if "[LOG] Parquet saved" in line:
                    self.log_status_label.configure(
                        text="📝 Log saved → logs/pipeline_logs.parquet  ✅",
                        text_color="gray70",
                    )

            self.proc.wait()
            ec = self.proc.returncode
            self._log("─" * 60)
            self._log(
                "✅ Pipeline finished successfully."
                if ec == 0
                else f"❌ Pipeline exited with code {ec}."
            )

        except FileNotFoundError:
            self._log(f"❌ pipeline.py not found at: {os.path.join(BASE_DIR, 'pipeline.py')}")
        except Exception as e:
            self._log(f"❌ Unexpected error: {e}")
        finally:
            self.proc = None
            self.run_button.configure(state="normal")
            self.stop_button.configure(state="disabled")

    # ════════════════════════════════════════════════════════════════════
    # STOP + HELPERS
    # ════════════════════════════════════════════════════════════════════
    def stop_pipeline(self):
        if self.proc is not None and self.proc.poll() is None:
            self.proc.terminate()
            self._log("\n[UI] ⏹ Pipeline stopped by user.")
        self.run_button.configure(state="normal")
        self.stop_button.configure(state="disabled")

    def _log(self, msg: str):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.update_idletasks()


if __name__ == "__main__":
    app = App()
    app.mainloop()