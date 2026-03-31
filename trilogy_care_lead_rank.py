#!/usr/bin/env python3
"""
╔═══════════════════════════════════════════════════════════════╗
║        💅  TRILOGY CARE  ·  SLAY RANKER  💅          ║
║              She scores. She slays. She converts.            ║
╚═══════════════════════════════════════════════════════════════╝
Run:  python trilogy_care_lead_rank.py
Deps: pip install pandas   (optional but recommended)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv, os, re, webbrowser
from datetime import datetime, date
from typing import Optional, List, Dict

try:
    import pandas as pd
    _PANDAS = True
except ImportError:
    _PANDAS = False

# ═════════════════════════════════════════════════════════════
#  BARBIE / SLAY THEME 💅
# ═════════════════════════════════════════════════════════════
T = {
    # -- Backgrounds
    'bg':           '#FFF0F7',
    'bg2':          '#FFE4F2',
    'card':         '#FFFFFF',
    'header_bg':    '#FF007F',
    'border':       '#FFB3D9',
    # -- Text
    'text':         '#4A0030',
    'text_muted':   '#C2607A',
    'text_bright':  '#FFFFFF',
    # -- Tier: SLAY 💅 (Hot pink)
    'legendary':    '#FF007F',
    'leg_bg':       '#FFF0F7',
    'leg_alt':      '#FFE4F2',
    # -- Tier: GLAM 💖 (Deep pink / magenta)
    'epic':         '#E91E8C',
    'epic_bg':      '#FDE8F4',
    'epic_alt':     '#FAD4EC',
    # -- Tier: CUTE 🌸 (Soft pink)
    'rare':         '#FF80AB',
    'rare_bg':      '#FFF5FA',
    'rare_alt':     '#FFECF5',
    # -- Tier: BASIC 🎀 (Muted pink-grey)
    'common':       '#C2607A',
    'common_bg':    '#FFF8FB',
    'common_alt':   '#FFF2F7',
    # -- Buttons / Accents
    'btn_green':    '#FF007F',
    'btn_green_h':  '#CC0066',
    'btn_orange':   '#FFD700',
    'btn_orange_h': '#FFC200',
    'xp_fill':      '#FFD700',
    'xp_empty':     '#FFB3D9',
    'white':        '#FFFFFF',
}

FONTS = {
    'title':    ('Segoe UI', 16, 'bold'),
    'heading':  ('Segoe UI', 11, 'bold'),
    'subhead':  ('Segoe UI', 10, 'bold'),
    'body':     ('Segoe UI', 10),
    'small':    ('Segoe UI', 9),
    'detail':   ('Consolas', 9),
    'xp':       ('Segoe UI', 14, 'bold'),
}

TIER_META = {
    'LEGENDARY': {'emoji': '💅', 'label': 'Slay',  'color': 'legendary', 'bg': 'leg_bg',  'alt': 'leg_alt'},
    'EPIC':      {'emoji': '💖', 'label': 'Glam',  'color': 'epic',      'bg': 'epic_bg', 'alt': 'epic_alt'},
    'RARE':      {'emoji': '🌸', 'label': 'Cute',  'color': 'rare',      'bg': 'rare_bg', 'alt': 'rare_alt'},
    'COMMON':    {'emoji': '🎀', 'label': 'Basic', 'color': 'common',    'bg': 'common_bg','alt': 'common_alt'},
}


# ═════════════════════════════════════════════════════════════
#  SCORING ENGINE  (based on Trilogy Care Zoho CRM fields)
# ═════════════════════════════════════════════════════════════
class ScoringEngine:
    """
    Scores leads 0–100 XP across 5 dimensions tuned to
    Trilogy Care's real Zoho CRM export:

      Care Readiness   (max 35 XP)  — Journey Stage + Stage self-report
      Lead Quality     (max 20 XP)  — Lead Status
      Source Quality   (max 20 XP)  — Lead Source + Channel Attribution
      Engagement       (max 15 XP)  — Total Notes (touches)
      Recency          (max 10 XP)  — Last Activity Time
    """

    def score(self, row: Dict, mapping: Dict) -> Dict:
        sections, total = [], 0
        for label, cap, fn in [
            ('Care Readiness', 35, self._care_readiness),
            ('Lead Quality',   20, self._lead_quality),
            ('Source Quality', 20, self._source_quality),
            ('Engagement',     15, self._engagement),
            ('Recency',        10, self._recency),
        ]:
            pts, msg = fn(row, mapping)
            total += pts
            sections.append({'label': label, 'pts': pts, 'cap': cap, 'msg': msg})

        xp = min(round(total), 100)
        if   xp >= 75: tier = 'LEGENDARY'
        elif xp >= 50: tier = 'EPIC'
        elif xp >= 25: tier = 'RARE'
        else:          tier = 'COMMON'

        name = self._field(row, mapping, 'lead_name') or 'Unknown Lead'
        summary = self._narrative(name, xp, tier, sections)
        return {'score': xp, 'tier': tier, 'summary': summary,
                'sections': sections, '_row': row}

    # ── Helpers ─────────────────────────────────────────────
    def _field(self, row, mapping, key) -> Optional[str]:
        col = mapping.get(key)
        if not col or col == '(skip)':
            return None
        val = str(row.get(col, '')).strip()
        return val if val and val.lower() not in ('', 'nan', 'none', 'n/a', '-', '#n/a', 'false') else None

    # ── Care Readiness (35 XP) ──────────────────────────────
    def _care_readiness(self, row, mapping):
        journey = (self._field(row, mapping, 'journey_stage') or '').lower()
        stage   = (self._field(row, mapping, 'stage')         or '').lower()
        desc    = (self._field(row, mapping, 'description')   or '').lower()
        combo   = f"{journey} {stage} {desc}"

        # Switching provider = highest urgency
        if 'switch' in combo:
            return 35, "Actively switching providers — extremely high urgency"

        # Allocated + has code or package assigned
        if 'allocated' in journey:
            if 'referral code' in stage or 'assigned' in stage:
                return 33, "HCP allocated & package/code confirmed — ready to onboard"
            return 28, "HCP allocated — package pending confirmation"

        # Has referral code (active HCP)
        if 'referral code' in stage:
            return 30, "Has referral code — ready to activate with Trilogy Care"

        # Assigned package
        if 'assigned' in stage or 'assigned' in combo:
            return 27, "Package assigned — ready to start services"

        # Completed ACAT assessment
        if 'completed' in combo and ('assessment' in combo or 'acat' in combo):
            return 25, "Completed aged care assessment — awaiting funding"

        # Active HCP (already with a provider)
        if 'active' in journey:
            return 20, "Active HCP — receiving care, potential switch opportunity"

        # Waitlist
        if 'waitlist' in combo or 'waiting' in combo:
            return 15, "On HCP waitlist — monitor for package activation"

        # CHSP (lower level support)
        if 'chsp' in combo:
            return 10, "On CHSP — potential future HCP upgrade"

        return 5, "Care journey stage unclear — needs qualification"

    # ── Lead Quality (20 XP) ────────────────────────────────
    def _lead_quality(self, row, mapping):
        status = (self._field(row, mapping, 'lead_status') or '').lower()
        if not status:
            return 0, None
        if 'high value' in status:
            return 20, "Contacted – High Value lead status"
        if 'contacted' in status:
            return 12, "Contacted and engaged"
        if 'new' in status:
            return 5, "New — not yet contacted"
        return 5, f"Status: {status.title()}"

    # ── Source Quality (20 XP) ──────────────────────────────
    def _source_quality(self, row, mapping):
        source  = (self._field(row, mapping, 'lead_source')          or '').lower()
        channel = (self._field(row, mapping, 'channel_attribution')   or '').lower()
        combo   = f"{source} {channel}"

        if 'ptr' in combo or 'partner referral' in channel or 'ptr' in source:
            return 20, "Partner referral (PTR) — highest quality source"
        if 'phone' in combo:
            return 16, "Phone enquiry — high intent, direct contact"
        if 'website' in combo:
            return 12, "Website enquiry — organic interest"
        if 'facebook' in combo or 'instagram' in combo or 'fb' in source:
            return 8, "Facebook/Instagram ad — marketing generated"
        if 'google' in combo:
            return 10, "Google search — actively searching for care"
        if source:
            return 5, f"Source: {source.title()}"
        return 0, None

    # ── Engagement (15 XP) ──────────────────────────────────
    def _engagement(self, row, mapping):
        raw = self._field(row, mapping, 'total_notes') or '0'
        try:
            n = int(float(raw))
        except (ValueError, TypeError):
            return 0, None
        if n >= 6: return 15, f"{n} notes logged — highly engaged"
        if n >= 4: return 12, f"{n} notes logged — well engaged"
        if n >= 3: return  9, f"{n} notes logged — moderately engaged"
        if n >= 2: return  6, f"{n} notes logged — some engagement"
        if n >= 1: return  3, f"{n} note logged — minimal engagement"
        return 0, "No notes recorded yet"

    # ── Recency (10 XP) ─────────────────────────────────────
    def _recency(self, row, mapping):
        raw = (self._field(row, mapping, 'last_activity') or
               self._field(row, mapping, 'created_time'))
        if not raw:
            return 0, None
        dt = self._parse_date(raw)
        if dt is None:
            return 0, None
        days = (datetime.now() - dt).days
        if days <= 3:  return 10, f"Active {days}d ago — red hot 🔥"
        if days <= 7:  return  8, f"Active {days}d ago — very recent"
        if days <= 14: return  6, f"Active {days}d ago — recent"
        if days <= 30: return  4, f"Active {days}d ago"
        if days <= 60: return  2, f"Active {days}d ago — follow up needed"
        return 1, f"Last active {days}d ago — going cold"

    # ── Date parser ─────────────────────────────────────────
    def _parse_date(self, raw) -> Optional[datetime]:
        clean = raw.strip().replace('T', ' ').replace('Z', '').strip()
        fmts = [
            '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
            '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%d/%m/%Y',
            '%d-%m-%Y', '%d %b %Y', '%d %B %Y', '%m/%d/%Y',
        ]
        for fmt in fmts:
            try:
                return datetime.strptime(clean, fmt)
            except ValueError:
                try:
                    sample = datetime(2000, 1, 1).strftime(fmt)
                    return datetime.strptime(clean[:len(sample)], fmt)
                except ValueError:
                    continue
        return None

    # ── Narrative ────────────────────────────────────────────
    def _narrative(self, name, xp, tier, sections):
        hits = [s['msg'] for s in sections if s['pts'] > 0 and s['msg']]
        openers = {
            'LEGENDARY': f"💅 {name} is an absolute SLAY — maximum onboarding potential, she's ready!",
            'EPIC':      f"💖 {name} is giving GLAM — strong conversion energy, follow up now!",
            'RARE':      f"🌸 {name} is CUTE potential — worth nurturing, don't sleep on her!",
            'COMMON':    f"🎀 {name} is giving BASIC energy — lower priority at this stage.",
        }
        factors = '; '.join(hits[:3]) if hits else 'Limited data available'
        return f"{openers[tier]} Key signals: {factors}. Total XP: {xp}/100."


# ═════════════════════════════════════════════════════════════
#  COLUMN MAPPING DIALOG
# ═════════════════════════════════════════════════════════════
class MappingDialog(tk.Toplevel):
    FIELDS = [
        ('lead_name',           'Lead Name',                ['lead name', 'name', 'full name', 'first name', 'contact']),
        ('journey_stage',       'Journey Stage (HCP Status)',['journey stage', 'journey', 'hcp level', 'hcp', 'package level']),
        ('stage',               'Stage (Self-Reported)',    ['stage', 'home care status', 'funding status', 'hcp stage']),
        ('lead_status',         'Lead Status',              ['lead status', 'status']),
        ('lead_source',         'Lead Source',              ['lead source', 'source', 'channel', 'origin']),
        ('channel_attribution', 'Channel Attribution',      ['channel attribution', 'attribution', 'attribution (text)', 'channel']),
        ('total_notes',         'Total Notes (Touches)',    ['total notes', 'number of notes', 'notes', 'note count']),
        ('last_activity',       'Last Activity Time',       ['last activity time', 'last activity', 'modified time', 'last contact']),
        ('created_time',        'Created Time',             ['created time', 'created_time', 'lead created time', 'created date']),
        ('description',         'Description / Notes',      ['description', 'desc', 'notes', 'comments', 'journey summary']),
    ]

    def __init__(self, parent, columns):
        super().__init__(parent)
        self.title("💅  Map Your CSV Columns — Slay Ranker")
        self.configure(bg=T['bg2'])
        self.resizable(False, False)
        self.result = None
        self.vars = {}
        self._build(list(columns))
        self.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{px+(pw-w)//2}+{py+(ph-h)//2}")
        self.grab_set()
        self.wait_window()

    def _best_guess(self, hints, columns):
        cols_lower = {c.lower().strip(): c for c in columns}
        for hint in hints:
            if hint in cols_lower:
                return cols_lower[hint]
            for col_l, col in cols_lower.items():
                if hint in col_l or col_l in hint:
                    return col
        return '(skip)'

    def _build(self, columns):
        opts = ['(skip)'] + columns

        hdr = tk.Frame(self, bg=T['header_bg'], pady=14, padx=20)
        hdr.pack(fill='x')
        tk.Label(hdr, text="💅  Map Your CSV Columns",
                 font=FONTS['heading'], bg=T['header_bg'],
                 fg=T['white']).pack(side='left')

        body = tk.Frame(self, bg=T['bg2'], padx=28, pady=16)
        body.pack(fill='both', expand=True)

        tk.Label(body, text="Match each scoring field to a column in your CSV.\nSet to '(skip)' if a field doesn't exist.",
                 font=FONTS['small'], bg=T['bg2'],
                 fg=T['text_muted'], justify='left').grid(
                 row=0, column=0, columnspan=2, sticky='w', pady=(0, 14))

        for i, (key, label, hints) in enumerate(self.FIELDS):
            guess = self._best_guess(hints, columns)
            var = tk.StringVar(value=guess)
            self.vars[key] = var

            dot_color = T['xp_fill'] if guess != '(skip)' else T['text_muted']
            tk.Label(body, text='●', font=FONTS['body'],
                     bg=T['bg2'], fg=dot_color).grid(
                     row=i+1, column=0, sticky='e', padx=(0, 8), pady=4)

            fr = tk.Frame(body, bg=T['bg2'])
            fr.grid(row=i+1, column=1, sticky='w', pady=4)
            tk.Label(fr, text=label, font=FONTS['small'],
                     bg=T['bg2'], fg=T['text'], width=28, anchor='w').pack(side='left')

            style = ttk.Style()
            style.configure('Dark.TCombobox',
                            fieldbackground=T['card'],
                            background=T['card'],
                            foreground=T['text'],
                            selectbackground=T['btn_green'],
                            selectforeground=T['white'])
            cb = ttk.Combobox(fr, textvariable=var, values=opts,
                              width=32, state='readonly')
            cb.pack(side='left')

        btn_row = tk.Frame(self, bg=T['bg2'], padx=28, pady=14)
        btn_row.pack(fill='x')

        tk.Button(btn_row, text="💅  LET'S SLAY →",
                  font=FONTS['subhead'], bg=T['btn_green'], fg=T['white'],
                  relief='flat', padx=18, pady=8, cursor='hand2',
                  activebackground=T['btn_green_h'],
                  command=self._confirm).pack(side='right')
        tk.Button(btn_row, text="Cancel",
                  font=FONTS['small'], bg=T['bg'], fg=T['text_muted'],
                  relief='flat', padx=12, pady=8, cursor='hand2',
                  command=self.destroy).pack(side='right', padx=(0, 8))

    def _confirm(self):
        self.result = {k: v.get() for k, v in self.vars.items()}
        self.destroy()


# ═════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("💅 Trilogy Care — Slay Ranker")
        self.geometry("1200x760")
        self.minsize(960, 640)
        self.configure(bg=T['bg'])

        self.engine        = ScoringEngine()
        self.scored_data : List[Dict] = []
        self.mapping     : Dict       = {}
        self.sort_col    = 'score'
        self.sort_rev    = True
        self._current_zoho_url: str  = ''   # URL of the currently selected lead

        self._apply_styles()
        self._build_ui()

    def _apply_styles(self):
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('Jungle.Treeview',
                    font=FONTS['body'], rowheight=30,
                    background=T['card'], fieldbackground=T['card'],
                    foreground=T['text'], borderwidth=0,
                    relief='flat')
        s.configure('Jungle.Treeview.Heading',
                    font=FONTS['subhead'],
                    background=T['header_bg'], foreground=T['white'],
                    relief='flat', borderwidth=0, padding=(6, 8))
        s.map('Jungle.Treeview.Heading',
              background=[('active', T['btn_green_h'])])
        s.map('Jungle.Treeview',
              background=[('selected', '#FFD6EB')],
              foreground=[('selected', T['header_bg'])])
        s.configure('Jungle.Vertical.TScrollbar',
                    background=T['border'], troughcolor=T['bg2'],
                    bordercolor=T['border'], arrowcolor=T['header_bg'])
        s.configure('Jungle.Horizontal.TScrollbar',
                    background=T['border'], troughcolor=T['bg2'],
                    bordercolor=T['border'], arrowcolor=T['header_bg'])

    # ── UI Build ───────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._build_toolbar()
        self._build_stats_bar()
        self._build_table()
        self._build_detail_panel()

    def _build_header(self):
        bar = tk.Frame(self, bg=T['header_bg'], height=62)
        bar.pack(fill='x')
        bar.pack_propagate(False)

        # Left title
        left = tk.Frame(bar, bg=T['header_bg'])
        left.pack(side='left', padx=16, fill='y')
        tk.Label(left, text="💅  SLAY RANKER",
                 font=('Segoe UI', 18, 'bold'),
                 bg=T['header_bg'], fg=T['white']).pack(side='left', pady=14)
        tk.Label(left, text="  ✨ by Trilogy Care",
                 font=('Segoe UI', 12, 'italic'),
                 bg=T['header_bg'], fg='#FFD700').pack(side='left', pady=14)

        # Right badge
        tk.Label(bar, text="she scores. she slays. she converts. 💖  ",
                 font=('Segoe UI', 9, 'italic'),
                 bg=T['header_bg'], fg='#FFD700').pack(side='right', pady=14)

    def _build_toolbar(self):
        bar = tk.Frame(self, bg=T['bg2'], pady=8, padx=14,
                       highlightbackground=T['border'], highlightthickness=1)
        bar.pack(fill='x')

        self._btn_import = tk.Button(
            bar, text="👛  IMPORT CSV",
            font=FONTS['subhead'], bg=T['btn_green'], fg=T['white'],
            relief='flat', padx=14, pady=6, cursor='hand2',
            activebackground=T['btn_green_h'],
            command=self._import_csv)
        self._btn_import.pack(side='left', padx=(0, 8))

        self._btn_export = tk.Button(
            bar, text="✨  EXPORT RANKED",
            font=FONTS['subhead'], bg=T['btn_orange'], fg=T['text'],
            relief='flat', padx=14, pady=6, cursor='hand2',
            state='disabled',
            activebackground=T['btn_orange_h'],
            command=self._export_csv)
        self._btn_export.pack(side='left', padx=(0, 16))

        tk.Frame(bar, bg=T['border'], width=1, height=26).pack(side='left', padx=8)

        self._file_var = tk.StringVar(value="No leads loaded yet bestie — hit Import CSV 💅")
        tk.Label(bar, textvariable=self._file_var,
                 font=FONTS['small'], bg=T['bg2'], fg=T['text_muted']).pack(side='left')

        # Search
        self._search_var = tk.StringVar()
        self._search_var.trace_add('write', self._on_search)
        tk.Label(bar, text="🔍", font=FONTS['body'],
                 bg=T['bg2'], fg=T['text_muted']).pack(side='right', padx=(6, 2))
        e = tk.Entry(bar, textvariable=self._search_var,
                     font=FONTS['body'], width=22,
                     bg=T['card'], fg=T['text'], insertbackground=T['xp_fill'],
                     relief='flat', bd=0,
                     highlightbackground=T['border'], highlightthickness=1)
        e.pack(side='right', padx=(0, 4))

    def _build_stats_bar(self):
        outer = tk.Frame(self, bg=T['bg'], pady=8, padx=14)
        outer.pack(fill='x')
        self._stat_labels: Dict[str, tk.Label] = {}

        items = [
            ('total',     T['header_bg'], '💅 Total Leads'),
            ('LEGENDARY', T['legendary'], '💅 Slay'),
            ('EPIC',      T['epic'],      '💖 Glam'),
            ('RARE',      T['rare'],      '🌸 Cute'),
            ('COMMON',    T['common'],    '🎀 Basic'),
        ]
        for key, color, label in items:
            card = tk.Frame(outer, bg=T['card'], padx=18, pady=8,
                            highlightbackground=T['border'], highlightthickness=1)
            card.pack(side='left', padx=(0, 8))
            count = tk.Label(card, text='—', font=('Segoe UI', 15, 'bold'),
                             bg=T['card'], fg=color)
            count.pack()
            tk.Label(card, text=label, font=FONTS['small'],
                     bg=T['card'], fg=T['text_muted']).pack()
            self._stat_labels[key] = count

    def _build_table(self):
        container = tk.Frame(self, bg=T['bg'])
        container.pack(fill='both', expand=True, padx=14, pady=(2, 0))

        cols    = ('rank', 'name',        'xp',      'tier',     'journey',          'status',      'source',       'notes', 'last_active')
        headers = ('Rank', 'Lead Name',  '✨ Score', 'Vibe',     'Journey Stage',    'Lead Status', 'Source',       'Notes', 'Last Active')
        widths  = {'rank': 52, 'name': 180, 'xp': 80, 'tier': 120,
                   'journey': 175, 'status': 160, 'source': 140, 'notes': 60, 'last_active': 120}
        anchors = {'rank': 'center', 'xp': 'center', 'tier': 'center', 'notes': 'center'}

        self.tree = ttk.Treeview(container, columns=cols, show='headings',
                                 style='Jungle.Treeview', selectmode='browse')

        for col, hdr in zip(cols, headers):
            self.tree.heading(col, text=hdr,
                              command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=widths.get(col, 100),
                             anchor=anchors.get(col, 'w'),
                             stretch=(col == 'journey'))

        for tier, meta in TIER_META.items():
            self.tree.tag_configure(tier,          background=T[meta['bg']],  foreground=T[meta['color']])
            self.tree.tag_configure(tier + '_ALT', background=T[meta['alt']], foreground=T[meta['color']])

        vsb = ttk.Scrollbar(container, orient='vertical',   style='Jungle.Vertical.TScrollbar',   command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient='horizontal', style='Jungle.Horizontal.TScrollbar', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self._empty_lbl = tk.Label(
            container,
            text="💅  Import a CSV to start slayyy-ing\n✨  Your leads are waiting bestie",
            font=FONTS['heading'], fg=T['text_muted'], bg=T['card'],
            justify='center')
        self._empty_lbl.grid(row=0, column=0)
        self.tree.bind('<<TreeviewSelect>>', self._on_select)
        self.tree.bind('<Double-1>', self._on_double_click)

    def _build_detail_panel(self):
        self._detail_outer = tk.Frame(self, bg=T['header_bg'],
                                      highlightbackground=T['border'],
                                      highlightthickness=1)
        self._detail_outer.pack(fill='x', padx=14, pady=(4, 10))

        hdr = tk.Frame(self._detail_outer, bg=T['header_bg'], padx=12, pady=4)
        hdr.pack(fill='x')
        tk.Label(hdr, text="✨  Lead Score Breakdown — Spill the Tea",
                 font=FONTS['small'], bg=T['header_bg'], fg=T['white']).pack(side='left')

        # Zoho CRM link button (right side of header)
        self._btn_zoho = tk.Button(
            hdr, text="🔗  Open in Zoho CRM",
            font=FONTS['small'], bg=T['btn_orange'], fg=T['text'],
            relief='flat', padx=10, pady=2, cursor='hand2',
            state='disabled',
            activebackground=T['btn_green_h'],
            command=self._open_zoho)
        self._btn_zoho.pack(side='right', padx=4)

        self._detail_text = tk.Text(
            self._detail_outer, height=5, font=FONTS['detail'],
            bg=T['card'], fg=T['text'], relief='flat', bd=0,
            wrap='word', padx=12, pady=8,
            state='disabled', cursor='arrow',
            insertbackground=T['xp_fill'])
        self._detail_text.pack(fill='x')

        # Colour tags for the detail panel
        self._detail_text.tag_configure('bold',   font=('Consolas', 9, 'bold'))
        self._detail_text.tag_configure('muted',  foreground=T['text_muted'])
        self._detail_text.tag_configure('LEGENDARY', foreground=T['legendary'])
        self._detail_text.tag_configure('EPIC',      foreground=T['epic'])
        self._detail_text.tag_configure('RARE',      foreground=T['rare'])
        self._detail_text.tag_configure('COMMON',    foreground=T['common'])
        self._detail_text.tag_configure('xp',        foreground=T['xp_fill'])

        self._set_detail("Select a lead to spill the tea on their score ✨")

    # ── Import / Export ────────────────────────────────────
    def _import_csv(self):
        path = filedialog.askopenfilename(
            title="Select Leads CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            rows, columns = self._read_csv(path)
        except Exception as exc:
            messagebox.showerror("Import Error", f"Could not read CSV:\n\n{exc}")
            return
        if not rows:
            messagebox.showwarning("Empty File", "CSV file has no data rows.")
            return

        dlg = MappingDialog(self, columns)
        if dlg.result is None:
            return

        self.mapping     = dlg.result
        self.scored_data = [self.engine.score(r, self.mapping) for r in rows]

        self._file_var.set(f"📄  {os.path.basename(path)}    ({len(rows)} leads loaded)")
        self._btn_export.config(state='normal')
        self._update_stats()
        self._refresh_table()
        self._set_detail("Select a lead to spill the tea on their score ✨")

    def _read_csv(self, path):
        if _PANDAS:
            df = pd.read_csv(path, dtype=str, encoding='utf-8-sig').fillna('')
            return df.to_dict('records'), list(df.columns)
        with open(path, newline='', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            cols = list(reader.fieldnames or [])
        return rows, cols

    @staticmethod
    def _zoho_link(row: dict) -> str:
        """Build a direct Zoho CRM link from the Record Id field."""
        record_id = str(row.get('Record Id', '')).strip()
        if not record_id or record_id.lower() in ('', 'nan', 'none'):
            return ''
        numeric_id = record_id.replace('zcrm_', '')
        return f"https://crm.zoho.com/crm/org724596924/tab/Leads/{numeric_id}"

    def _export_csv(self):
        if not self.scored_data:
            return
        path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[("CSV files", "*.csv")],
            initialfile="Trilogy_Care_Ranked_Leads.csv",
            title="Export Ranked Leads")
        if not path:
            return

        ranked = sorted(self.scored_data, key=lambda x: x['score'], reverse=True)
        extras = ['Rank', 'XP_Score', 'Tier', 'AI_Summary', 'Zoho_CRM_Link']
        orig   = list(ranked[0]['_row'].keys()) if ranked else []
        fields = extras + [c for c in orig if c not in extras]

        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=fields, extrasaction='ignore')
            w.writeheader()
            for rank, item in enumerate(ranked, 1):
                out = dict(item['_row'])
                out.update({
                    'Rank':          rank,
                    'XP_Score':      item['score'],
                    'Tier':          item['tier'],
                    'AI_Summary':    item['summary'],
                    'Zoho_CRM_Link': self._zoho_link(item['_row']),
                })
                w.writerow(out)

        messagebox.showinfo("Slay! 💅✨", f"Ranked leads exported — you ate that bestie!\n\n{path}")

    # ── Table ──────────────────────────────────────────────
    def _refresh_table(self, filter_text: str = ''):
        self._empty_lbl.grid_remove()
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        data = list(self.scored_data)
        if filter_text:
            ft = filter_text.lower()
            data = [d for d in data if any(ft in str(v).lower() for v in d['_row'].values())]

        # Sort
        tier_order = {'LEGENDARY': 4, 'EPIC': 3, 'RARE': 2, 'COMMON': 1}
        if self.sort_col == 'xp':
            data.sort(key=lambda x: x['score'], reverse=self.sort_rev)
        elif self.sort_col == 'tier':
            data.sort(key=lambda x: tier_order.get(x['tier'], 0), reverse=self.sort_rev)
        else:
            data.sort(key=lambda x: x['score'], reverse=True)

        def _get(row, key, maxlen=35):
            col = self.mapping.get(key, '')
            if not col or col == '(skip)':
                return '—'
            v = str(row.get(col, '')).strip()
            if not v or v.lower() in ('nan', 'none', 'false'):
                return '—'
            return v[:maxlen]

        for rank, item in enumerate(data, 1):
            row  = item['_row']
            tier = item['tier']
            meta = TIER_META[tier]
            tag  = tier if rank % 2 == 1 else tier + '_ALT'

            name     = _get(row, 'lead_name', 30)
            journey  = _get(row, 'journey_stage', 25)
            stage_v  = _get(row, 'stage', 30)
            status   = _get(row, 'lead_status', 22)
            source   = _get(row, 'lead_source', 20)
            notes    = _get(row, 'total_notes', 5)
            last_act = _get(row, 'last_activity', 20)

            # Build journey display
            journey_display = journey
            if stage_v != '—':
                journey_display = f"{journey} · {stage_v}" if journey != '—' else stage_v

            self.tree.insert('', 'end', iid=str(id(item)),
                             values=(rank,
                                     name if name != '—' else '(no name)',
                                     f"{item['score']} XP",
                                     f"{meta['emoji']}  {meta['label']}",
                                     journey_display,
                                     status, source, notes,
                                     last_act[:16] if last_act != '—' else '—'),
                             tags=(tag,))

        if not data:
            self._empty_lbl.grid(row=0, column=0)

    def _update_stats(self):
        counts = {'total': len(self.scored_data),
                  'LEGENDARY': 0, 'EPIC': 0, 'RARE': 0, 'COMMON': 0}
        for item in self.scored_data:
            counts[item['tier']] = counts.get(item['tier'], 0) + 1
        for key, lbl in self._stat_labels.items():
            lbl.config(text=str(counts.get(key, 0)))

    def _sort_by(self, col):
        self.sort_rev = not self.sort_rev if col == self.sort_col else True
        self.sort_col = col
        self._refresh_table(self._search_var.get())

    def _on_search(self, *_):
        self._refresh_table(self._search_var.get())

    def _on_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        item = next((x for x in self.scored_data if str(id(x)) == sel[0]), None)
        if not item:
            return

        # Build and store Zoho URL
        self._current_zoho_url = self._zoho_link(item['_row'])
        if self._current_zoho_url:
            self._btn_zoho.config(state='normal')
        else:
            self._btn_zoho.config(state='disabled')

        tier = item['tier']
        meta = TIER_META[tier]
        lines = [
            ('bold',  f"{meta['emoji']}  "),
            (tier,    f"{item['summary']}\n"),
            ('muted', '\n  XP Breakdown:\n'),
        ]
        for s in item['sections']:
            if not s['msg']:
                continue
            filled = int(round(s['pts'] / s['cap'] * 12)) if s['cap'] else 0
            empty  = 12 - filled
            bar    = '▓' * filled + '░' * empty
            tag    = tier if s['pts'] > 0 else 'muted'
            lines.append((tag, f"  {s['label']:<18}  {s['pts']:2}/{s['cap']:2} XP  {bar}   {s['msg']}\n"))

        self._set_detail(lines)

    def _open_zoho(self):
        """Open the selected lead's Zoho CRM record in the default browser."""
        if self._current_zoho_url:
            webbrowser.open(self._current_zoho_url)

    def _on_double_click(self, _event=None):
        """Double-clicking a row opens the lead directly in Zoho CRM."""
        if self._current_zoho_url:
            webbrowser.open(self._current_zoho_url)

    def _set_detail(self, content):
        self._detail_text.config(state='normal')
        self._detail_text.delete('1.0', 'end')
        if isinstance(content, str):
            self._detail_text.insert('end', content, 'muted')
        else:
            for tag, text in content:
                self._detail_text.insert('end', text, tag)
        self._detail_text.config(state='disabled')


# ═════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═════════════════════════════════════════════════════════════
def main():
    app = App()
    app.mainloop()

if __name__ == '__main__':
    main()
