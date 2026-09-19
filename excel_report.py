"""Génération du livrable Excel client. Charte WPP Media. Dépendance : openpyxl."""
from __future__ import annotations
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker

# Charte WPP
NAVY = "191D63"      # WPP blue
BLUE = "3A3F8F"
LIGHT = "E4E5F1"
GREEN = "C6EFCE"
RED = "F4C7C3"
ORANGE = "FCE4A6"
GREY = "F2F2F2"
WHITE = "FFFFFF"

HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(bold=True, color=WHITE, size=11)
TITLE_FONT = Font(bold=True, color=NAVY, size=16)
SUB_FONT = Font(bold=True, color=BLUE, size=12)
THIN = Side(style="thin", color="C9CBE0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

QUAD_FILL = {"Q1": PatternFill("solid", fgColor=RED), "Q2": PatternFill("solid", fgColor=LIGHT),
             "Q3": PatternFill("solid", fgColor=ORANGE), "Q4": PatternFill("solid", fgColor=GREY)}
GRAVITE_FILL = {"Élevée": PatternFill("solid", fgColor=RED), "Moyenne": PatternFill("solid", fgColor=ORANGE),
                "Faible": PatternFill("solid", fgColor=GREEN)}


def _safe_cell(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, float):
        return "" if v != v else round(v, 3)
    if isinstance(v, (int, str)):
        return v
    try:
        import numpy as _np
        if isinstance(v, _np.generic):
            v2 = v.item()
            return "" if (isinstance(v2, float) and v2 != v2) else v2
    except Exception:
        pass
    if isinstance(v, (list, tuple, set)):
        return ", ".join(str(x) for x in v)
    try:
        import pandas as _pd
        if v is getattr(_pd, "NA", None) or _pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v)


def _style_header(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def _autosize(ws, df, start_col=1, max_w=60, min_w=10):
    for i, col in enumerate(df.columns):
        try:
            vals = df[col].astype(str).head(200).tolist()
            body = max((len(str(v)) for v in vals), default=0)
        except Exception:
            body = 12
        width = max(len(str(col)), body) + 2
        ws.column_dimensions[get_column_letter(start_col + i)].width = max(min_w, min(max_w, width))


def _write_df(ws, df, start_row=1, start_col=1, autofilter=True, header=True, band=True):
    r = start_row
    if header:
        for j, col in enumerate(df.columns):
            ws.cell(row=r, column=start_col + j, value=str(col))
        _style_header(ws, r, len(df.columns), start_col)
        r += 1
    for idx, (_, row) in enumerate(df.iterrows()):
        for j, col in enumerate(df.columns):
            cell = ws.cell(row=r, column=start_col + j, value=_safe_cell(row[col]))
            cell.border = BORDER
            cell.alignment = LEFT if j == 0 else CENTER
            if band and idx % 2 == 1:
                cell.fill = PatternFill("solid", fgColor=GREY)
        r += 1
    if autofilter and header:
        last = get_column_letter(start_col + len(df.columns) - 1)
        ws.auto_filter.ref = f"{get_column_letter(start_col)}{start_row}:{last}{r-1}"
    ws.freeze_panes = ws.cell(row=start_row + 1, column=start_col)
    _autosize(ws, df, start_col)
    return r


# ==========================================================================
def build_excel(results: dict, output_path: str, site_name: str = "Client"):
    p = results["pages"].copy()
    # Ancres réellement utilisées par page cible (liens en contenu), pour les onglets Ancres / Analyse par page.
    _lk = results.get("links")
    if _lk is not None and "anchor" in getattr(_lk, "columns", []):
        _c = _lk[_lk["position"] == "Content"].copy()
        _c["anchor"] = _c["anchor"].fillna("").astype(str).str.strip().replace("", "(vide)")
        _amap = {t: " ; ".join(f"{a} ({n})" for a, n in _c_g["anchor"].value_counts().head(4).items())
                 for t, _c_g in _c.groupby("target")}
    else:
        _amap = {}
    p["_ancres_txt"] = p["url"].map(_amap).fillna("(aucune)")
    cfg = results["config"]
    wb = Workbook()
    notes = []

    def _safe(fn, *a):
        try:
            fn(*a)
        except Exception as e:
            notes.append(f"{fn.__name__}: {e}")

    _safe(_sheet_legende, wb, cfg, results, p)
    _safe(_sheet_plan_action, wb, results, p)
    _safe(_sheet_synthese_ecrite, wb, results, p, site_name, notes)
    _safe(_sheet_dashboard, wb, results, p, site_name)
    _safe(_sheet_problemes, wb, results, p)
    _safe(_sheet_prescriptions, wb, results)
    _safe(_sheet_analyse_page, wb, p)
    _safe(_sheet_ancres, wb, p)
    _safe(_sheet_silotage, wb, results)

    order = ["1. Méthodologie & Légende", "Plan d'action", "2. Synthèse écrite", "3. Tableau de bord",
             "4. Problèmes identifiés", "5. Prescription liens", "6. Analyse par page",
             "7. Analyse ancres", "8. Silotage clusters"]
    try:
        wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
    except Exception:
        pass
    try:
        wb.save(output_path)
    except Exception:
        try:
            for ws in list(wb.worksheets):
                try:
                    ws._charts = []
                except Exception:
                    pass
            if "_chartdata" in wb.sheetnames:
                del wb["_chartdata"]
        except Exception:
            pass
        wb.save(output_path)
    return output_path


# 1. Méthodologie & Légende ------------------------------------------------
def _sheet_legende(wb, cfg, results, p):
    ws = wb.active
    ws.title = "1. Méthodologie & Légende"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 95

    r = [1]
    def title(t):
        c = ws.cell(row=r[0], column=2, value=t); c.font = TITLE_FONT; r[0] += 2
    def h2(t):
        c = ws.cell(row=r[0], column=2, value=t); c.font = SUB_FONT; r[0] += 1
    def line(k, v):
        a = ws.cell(row=r[0], column=2, value=k); a.font = Font(bold=True)
        a.alignment = Alignment(vertical="top", wrap_text=True)
        b = ws.cell(row=r[0], column=3, value=v); b.alignment = Alignment(vertical="top", wrap_text=True)
        r[0] += 1
    def gap():
        r[0] += 1

    title("Comment lire ce fichier")
    h2("Les onglets, dans l'ordre")
    line("2. Synthèse écrite", "L'analyse rédigée : ce qui est sous-maillé, sur-maillé, et les actions prioritaires.")
    line("3. Tableau de bord", "La matrice de diagnostic en graphique + les chiffres clés + top pages.")
    line("4. Problèmes identifiés", "Chaque type de problème avec son volume, son % de pages et un exemple.")
    line("5. Prescription liens", "Le livrable actionnable : quels liens créer, avec quelle ancre et où.")
    line("6. Analyse par page", "Le détail page par page, toutes les métriques (pour aller au fond).")
    line("7. Analyse ancres", "La qualité des ancres, en pourcentages.")
    line("8. Silotage clusters", "Comment les catégories se lient entre elles.")
    gap()

    h2("Le diagnostic repose sur 2 axes (pas un score unique)")
    line("Axe X — Potentiel business (0 à 10)", "Mesure la valeur business d'une page. Combine : volume de recherche "
         "des mots-clés visés + clics GSC + impressions GSC + conversions GA4. Chaque ingrédient est classé en "
         "rang (percentile) puis moyenné et ramené sur 10. 10 = page à très fort enjeu, 0 = enjeu faible.")
    line("Axe Y — Déficit de maillage (0 à 10)", "Mesure à quel point une page est mal maillée. Combine (à l'inverse) : "
         "PageRank interne (InRank) + nombre de liens contextuels entrants + qualité des ancres + profondeur de clic. "
         "10 = fortement sous-maillée, 0 = très bien maillée.")
    gap()

    h2("Les 4 quadrants (croisement des 2 axes)")
    line("Q1 — Prioritaire", "Fort potentiel ET fort déficit. Ce sont les pages à renforcer en priorité : "
         "on leur prescrit des liens. « Top 20 Q1 » = les 20 pages Q1 les plus prioritaires.")
    line("Q2 — Hors scope maillage", "Fort potentiel mais déjà bien maillé. Si elles ne performent pas, le problème "
         "est ailleurs (contenu, technique, concurrence), pas dans le maillage.")
    line("Q3 — Sur-maillage", "Faible potentiel mais très bien maillé : elles captent du PageRank pour peu de valeur. "
         "« Top 20 Q3 » = celles dont on peut réallouer les liens vers les Q1.")
    line("Q4 — Ne rien faire", "Faible potentiel et fort déficit : pas d'enjeu, on ne dépense pas d'effort.")
    gap()

    h2("Définitions utiles")
    line("Lien contextuel", "Lien situé dans le contenu de la page (emplacement « Content »), par opposition au menu, "
         "header, footer ou sidebar. C'est le lien qui compte le plus pour le SEO.")
    line("Page orpheline", "Page sans aucun lien interne entrant : invisible dans la navigation.")
    line("Ancre", "Le texte cliquable d'un lien. Une bonne ancre est descriptive et variée.")
    line("Ancre générique", "Ancre vide ou vague (« cliquez ici », « en savoir plus »…) : peu utile pour le SEO.")
    line("PageRank interne (InRank)", "Popularité d'une page au sein du site, calculée par OnCrawl (échelle 0-10).")
    gap()

    h2("Seuils utilisés pour ce rapport")
    line("Reclassement navigation", f"un lien présent sur plus de {cfg['nav_repeat_threshold']:.0%} des pages est considéré comme de la navigation")
    line("Cluster fermé", f"plus de {cfg['closed_cluster_threshold']:.0%} de liens à l'intérieur du même cluster")
    line("Ancre sur-homogène", f"une seule ancre représente plus de {cfg['homogeneous_anchor_threshold']:.0%} des liens d'une page")
    line("Ancre pauvre", f"plus de {cfg['poor_anchor_threshold']:.0%} d'ancres génériques ou vides sur une page Q1")
    line("Seuils des axes (X / Y)", f"{cfg['quadrant_x_split']} / {cfg['quadrant_y_split']} sur 10")
    gap()
    line("Sources de données", "Crawl OnCrawl (structure + InRank), Search Console (performance), GA4 (conversion, optionnel), "
         "étude sémantique (volumes et clusters, optionnel).")


# Plan d'action (vue principale, actionnable) --------------------------------
def _sheet_plan_action(wb, results, p):
    import numpy as _np
    ws = wb.create_sheet("Plan d'action")
    ws["A1"] = "Plan d'action — quoi faire, par page, du plus important au moins important"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("3 natures d'action : AJOUTER des liens (page peu/pas maillée) · OPTIMISER les ancres "
                "(varier le texte des liens) · CORRIGER la navigation (lien menu/pied à revoir). Trié par clics GSC.")
    ws["A2"].font = Font(italic=True, color=BLUE)
    ws["A2"].alignment = Alignment(wrap_text=True)

    flags = results.get("flags")
    ftypes = {}
    if flags is not None and len(flags):
        for u, g in flags.groupby("url"):
            ftypes[u] = set(g["type"])

    has_clics = "clics_gsc" in p.columns
    d = p.copy()
    if has_clics:
        d = d.sort_values("clics_gsc", ascending=False)

    rows = []
    for _, r in d.iterrows():
        url = r["url"]
        ft = ftypes.get(url, set())
        ctx = int(pd.to_numeric(r.get("liens_ctx_entrants", 0), errors="coerce") or 0)
        tot = int(pd.to_numeric(r.get("liens_entrants_total", 0), errors="coerce") or 0)
        quad = r.get("quadrant_code", "")
        anc = r.get("_ancres_txt", "")
        nbkw = int(pd.to_numeric(r.get("nb_kw", 0), errors="coerce") or 0)
        anchor_issue = any("Ancre" in t for t in ft)

        if tot == 0:
            action, emp, quoi = ("① Ajouter des liens", "Page invisible (0 lien entrant)",
                                 "Créer des liens de contenu vers cette page depuis des pages fortes du même thème.")
        elif ctx == 0:
            action, emp, quoi = ("① Ajouter du lien de contenu", "Menu / pied seulement",
                                 "Reliée uniquement par des blocs répétés. Ajouter de vrais liens éditoriaux dans le corps des pages.")
        elif anchor_issue:
            suff = f" (page positionnée sur {nbkw} mots-clés)" if nbkw >= 2 else ""
            action, emp, quoi = ("② Optimiser / varier les ancres", "Corps de page",
                                 f"Ancres en place : {anc}. Varier le texte des liens vers les mots-clés cibles{suff}.")
        elif quad == "Q1":
            action, emp, quoi = ("① Densifier le maillage de contenu", "Corps de page",
                                 "Fort potentiel mais peu de liens contextuels. Ajouter des liens de contenu.")
        else:
            continue

        clics = int(pd.to_numeric(r.get("clics_gsc", 0), errors="coerce") or 0) if has_clics else ""
        posv = pd.to_numeric(r.get("position_gsc"), errors="coerce") if has_clics else None
        pos = round(float(posv), 1) if posv is not None and pd.notna(posv) else ""
        rows.append({"Page": url, "Clics GSC": clics, "Position GSC (moy.)": pos,
                     "Mots-clés cibles (volume · position)": r.get("_kw_detail", ""),
                     "Action": action, "Emplacement": emp, "Quoi faire": quoi})

    df = pd.DataFrame(rows, columns=["Page", "Clics GSC", "Position GSC (moy.)",
                                     "Mots-clés cibles (volume · position)", "Action", "Emplacement", "Quoi faire"])
    df = df.head(80)
    if has_clics and len(df):
        c = pd.to_numeric(df["Clics GSC"], errors="coerce").fillna(0)
        q80, q40 = c.quantile(0.8), c.quantile(0.4)
        df.insert(0, "Priorité", _np.where(c >= q80, 1, _np.where(c >= q40, 2, 3)))

    end = _write_df(ws, df.reset_index(drop=True), start_row=4)
    # coloration priorité
    if "Priorité" in df.columns:
        pcol = list(df.columns).index("Priorité") + 1
        pf = {1: RED, 2: ORANGE, 3: GREEN}
        for row in ws.iter_rows(min_row=5, max_row=end - 1, min_col=pcol, max_col=pcol):
            for cell in row:
                if cell.value in pf:
                    cell.fill = PatternFill("solid", fgColor=pf[cell.value])
    # bloc "vos pages fortes en maillage"
    strong = p.copy()
    if "liens_ctx_entrants" in strong.columns:
        cols = [c for c in ["url", "cluster", "clics_gsc", "liens_ctx_entrants", "pr_interne"] if c in strong.columns]
        strong = strong.sort_values("liens_ctx_entrants", ascending=False).head(10)[cols].rename(columns={
            "url": "Page", "cluster": "Cluster", "clics_gsc": "Clics GSC",
            "liens_ctx_entrants": "Liens de contenu entrants", "pr_interne": "PageRank interne"})
        r2 = end + 2
        ws.cell(row=r2, column=1, value="Vos pages les mieux maillées (référence)").font = SUB_FONT
        end = _write_df(ws, strong.reset_index(drop=True), start_row=r2 + 1)

    # ③ Navigation à corriger : liens menu/pied qui pointent vers une redirection
    nav = (results.get("report") or {}).get("_nav_a_corriger")
    if nav is not None and len(nav):
        nv = nav.copy()
        nv["Action"] = "Changer ce lien de navigation : le faire pointer vers l'URL finale (200)"
        nv = nv.rename(columns={"target": "URL cible (redirige)", "emplacement": "Emplacement",
                                "nb_liens": "Nb liens (≈ nb pages)", "statut_cible": "Statut cible"})
        keep = [c for c in ["URL cible (redirige)", "Emplacement", "Nb liens (≈ nb pages)",
                            "Statut cible", "Action"] if c in nv.columns]
        nv = nv[keep].head(60)
        r3 = end + 2
        ws.cell(row=r3, column=1,
                value="③ Navigation à corriger — liens menu/pied vers une redirection (à recibler)").font = SUB_FONT
        _write_df(ws, nv.reset_index(drop=True), start_row=r3 + 1)


# 2. Synthèse écrite -------------------------------------------------------
def _sheet_synthese_ecrite(wb, results, p, site_name, notes=None):
    ws = wb.create_sheet("2. Synthèse écrite")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 118

    rows = [("title", f"Synthèse de l'audit de maillage interne — {site_name}")]
    try:
        n = len(p); n_orphan = int(p["orpheline"].sum())
        quad = p["quadrant_code"].value_counts().to_dict()
        rows.append(("h2", "Vue d'ensemble"))
        rows.append(("body", f"{n} pages analysées, dont {n_orphan} orphelines ({n_orphan/n*100:.0f}%) sans aucun lien interne entrant."))
        rows.append(("body", f"Répartition : {quad.get('Q1',0)} pages prioritaires (Q1), {quad.get('Q2',0)} bien maillées (Q2), "
                             f"{quad.get('Q3',0)} en sur-maillage (Q3), {quad.get('Q4',0)} sans enjeu (Q4)."))
    except Exception as e:
        notes and notes.append(f"syn_vue:{e}")
    try:
        q1 = p[p["quadrant_code"] == "Q1"]
        if len(q1):
            g = q1.groupby("cluster").agg(n=("url", "size"), deficit=("deficit_maillage", "mean"),
                                          pot=("potentiel_business", "mean")).sort_values("n", ascending=False).head(4)
            rows.append(("h2", "Priorité 1 : pages à fort potentiel, sous-maillées"))
            rows.append(("body", "Ces types de pages ont de la valeur business mais reçoivent trop peu de liens internes. Cibles n°1 de prescription."))
            for c, x in g.iterrows():
                rows.append(("bullet", f"« {c} » : {int(x['n'])} pages, déficit {x['deficit']:.1f}/10 pour un potentiel de {x['pot']:.1f}/10."))
    except Exception as e:
        notes and notes.append(f"syn_q1:{e}")
    try:
        q3 = p[p["quadrant_code"] == "Q3"]
        if len(q3):
            g = q3.groupby("cluster").agg(n=("url", "size"), pr=("pr_interne", "mean"),
                                          inl=("liens_ctx_entrants", "mean")).sort_values("pr", ascending=False).head(4)
            rows.append(("h2", "Sur-maillage : pages sur-abreuvées en PageRank pour un faible potentiel"))
            rows.append(("body", "Elles concentrent du PageRank et des liens pour peu de valeur. Réallouer ce maillage vers les pages Q1."))
            for c, x in g.iterrows():
                rows.append(("bullet", f"« {c} » : {int(x['n'])} pages, PageRank {x['pr']:.1f}/10, {x['inl']:.0f} liens entrants en moyenne."))
    except Exception as e:
        notes and notes.append(f"syn_q3:{e}")
    try:
        depth = pd.to_numeric(p["depth"], errors="coerce")
        rows.append(("h2", "Profondeur de crawl"))
        rows.append(("body", f"{(depth>3).mean()*100:.0f}% des pages sont à plus de 3 clics de l'accueil (max {int(depth.max())}). "
                             f"Une profondeur excessive dilue le PageRank : remonter les pages stratégiques."))
    except Exception as e:
        notes and notes.append(f"syn_depth:{e}")
    try:
        flags = results.get("flags")
        if flags is not None and len(flags):
            nc = int((flags["type"] == "Cannibalisation").sum())
            if nc:
                rows.append(("h2", "Cannibalisation"))
                rows.append(("body", f"{nc} cas de pages d'un même cluster positionnées sur le même mot-clé. À consolider ou différencier."))
    except Exception as e:
        notes and notes.append(f"syn_can:{e}")
    try:
        presc = results.get("prescriptions")
        n_presc = len(presc) if presc is not None else 0
        n_p1 = int((presc["priorite"] == 1).sum()) if (presc is not None and len(presc) and "priorite" in presc.columns) else 0
        rows.append(("h2", "Recommandations, par ordre de priorité"))
        rows.append(("bullet", f"1. Déployer les {n_presc} liens prescrits (onglet « Prescription liens »), en commençant par les {n_p1} de priorité 1."))
        rows.append(("bullet", "2. Réallouer le maillage des pages Q3 (sur-maillées) vers les pages Q1."))
        rows.append(("bullet", "3. Donner au moins un lien contextuel entrant aux pages orphelines."))
        rows.append(("bullet", "4. Remonter les pages profondes à fort potentiel dans l'arborescence."))
        rows.append(("bullet", "5. Corriger les ancres pauvres ou trop homogènes sur les pages Q1."))
    except Exception as e:
        notes and notes.append(f"syn_reco:{e}")

    r = 1
    for style, text in rows:
        if style == "h2":
            r += 1
            c = ws.cell(row=r, column=2, value=text); c.font = SUB_FONT
        elif style == "title":
            c = ws.cell(row=r, column=2, value=text); c.font = TITLE_FONT
        elif style == "bullet":
            c = ws.cell(row=r, column=2, value="•  " + text); c.font = Font(size=11)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        else:
            c = ws.cell(row=r, column=2, value=text); c.font = Font(size=11)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    if notes:
        r += 1
        ws.cell(row=r, column=2, value="Notes techniques : " + " | ".join(notes)).font = Font(italic=True, color="999999", size=9)


# 3. Tableau de bord -------------------------------------------------------
def _sheet_dashboard(wb, results, p, site_name):
    ws = wb.create_sheet("3. Tableau de bord")
    ws.sheet_view.showGridLines = False
    ws["B2"] = f"Tableau de bord — {site_name}"
    ws["B2"].font = TITLE_FONT
    ws.column_dimensions["A"].width = 2

    rep = results["report"]
    quad = p["quadrant_code"].value_counts().to_dict()
    kpis = [("Pages analysées", rep.get("pages_total", len(p))),
            ("Liens internes valides", rep.get("liens_valides", "")),
            ("Pages orphelines", rep.get("pages_orphelines", int(p["orpheline"].sum()))),
            ("Q1 — Prioritaires", quad.get("Q1", 0)),
            ("Q3 — Sur-maillage", quad.get("Q3", 0)),
            ("Prescriptions générées", rep.get("prescriptions_generees", 0))]
    r = 4
    ws.cell(row=r, column=2, value="Chiffres clés").font = SUB_FONT
    r += 1
    for label, val in kpis:
        a = ws.cell(row=r, column=2, value=label); a.font = Font(bold=True); a.border = BORDER
        a.alignment = LEFT
        b = ws.cell(row=r, column=3, value=_safe_cell(val)); b.alignment = CENTER; b.border = BORDER
        if label.startswith("Q1"):
            a.fill = QUAD_FILL["Q1"]; b.fill = QUAD_FILL["Q1"]
        r += 1
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 16

    # Matrice de diagnostic en TABLEAU (pas de graphe à points)
    r += 1
    ws.cell(row=r, column=2, value="Matrice de diagnostic (nombre de pages)").font = SUB_FONT
    r += 1

    def _mx(row, col, val, fill=None, white=False):
        c = ws.cell(row=row, column=col, value=val)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if fill is not None:
            c.fill = fill
        c.font = Font(bold=True, color="FFFFFF" if white else "191D63")
        return c

    _mx(r, 2, "")
    _mx(r, 3, "Faible potentiel business", HEADER_FILL, white=True)
    _mx(r, 4, "Fort potentiel business", HEADER_FILL, white=True)
    r += 1
    _mx(r, 2, "Fort déficit de maillage", HEADER_FILL, white=True)
    _mx(r, 3, f"Q4 — Ne rien faire\n{quad.get('Q4', 0)} pages", QUAD_FILL["Q4"])
    _mx(r, 4, f"Q1 — PRIORITAIRE\n{quad.get('Q1', 0)} pages", QUAD_FILL["Q1"])
    ws.row_dimensions[r].height = 32
    r += 1
    _mx(r, 2, "Bon maillage", HEADER_FILL, white=True)
    _mx(r, 3, f"Q3 — Sur-maillage\n{quad.get('Q3', 0)} pages", QUAD_FILL["Q3"])
    _mx(r, 4, f"Q2 — Hors scope\n{quad.get('Q2', 0)} pages", QUAD_FILL["Q2"])
    ws.row_dimensions[r].height = 32
    ws.column_dimensions["D"].width = 26

    cols = [c for c in ["url", "cluster", "potentiel_business", "deficit_maillage", "pr_interne",
                        "liens_ctx_entrants", "clics_gsc", "position_gsc"] if c in p.columns]
    r2 = r + 2
    ws.cell(row=r2, column=2, value="Top 20 Q1 — pages prioritaires (à renforcer)").font = SUB_FONT
    q1 = p[p["quadrant_code"] == "Q1"].sort_values("score_composite", ascending=False).head(20)[cols]
    end = _write_df(ws, q1.reset_index(drop=True), start_row=r2 + 1, start_col=2, autofilter=False)
    r3 = end + 2
    ws.cell(row=r3, column=2, value="Top 20 Q3 — sur-maillées (PageRank à réallouer)").font = SUB_FONT
    q3 = p[p["quadrant_code"] == "Q3"].sort_values("pr_interne", ascending=False).head(20)[cols]
    _write_df(ws, q3.reset_index(drop=True), start_row=r3 + 1, start_col=2, autofilter=False)


# 4. Problèmes identifiés --------------------------------------------------
def _sheet_problemes(wb, results, p):
    ws = wb.create_sheet("4. Problèmes identifiés")
    ws["A1"] = "Problèmes identifiés"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Pour chaque type : sa gravité, son volume, sa part des pages, et un exemple concret."
    ws["A2"].font = Font(italic=True, color=BLUE)

    flags = results["flags"].copy()
    n_pages = max(len(p), 1)
    url2cluster = dict(zip(p["url"], p["cluster"])) if "cluster" in p.columns else {}

    if len(flags):
        recap = []
        order = {"Élevée": 0, "Moyenne": 1, "Faible": 2}
        for typ, g in flags.groupby("type"):
            vol = len(g)
            grav = g["gravite"].iloc[0]
            ex_url = g["url"].iloc[0]
            ex_clu = url2cluster.get(ex_url, "—")
            recap.append({"Type de problème": typ, "Gravité": grav, "Volume": vol,
                          "% des pages": f"{vol/n_pages*100:.1f}%",
                          "Exemple (URL)": ex_url, "Cluster de l'exemple": ex_clu})
        recap = pd.DataFrame(recap).sort_values(
            by="Gravité", key=lambda s: s.map(order).fillna(3)).reset_index(drop=True)
    else:
        recap = pd.DataFrame(columns=["Type de problème", "Gravité", "Volume", "% des pages",
                                      "Exemple (URL)", "Cluster de l'exemple"])
    ws.cell(row=4, column=1, value="Récapitulatif par type").font = SUB_FONT
    end = _write_df(ws, recap, start_row=5)
    if "Gravité" in recap.columns:
        gcol = list(recap.columns).index("Gravité") + 1
        for row in ws.iter_rows(min_row=6, max_row=end - 1, min_col=gcol, max_col=gcol):
            for cell in row:
                if cell.value in GRAVITE_FILL:
                    cell.fill = GRAVITE_FILL[cell.value]

    # détail avec cluster + importance business, trié par gravité puis clics GSC
    if len(flags):
        det = flags.copy()
        det["cluster"] = det["url"].map(url2cluster).fillna("—")
        imp_cols = [c for c in ["clics_gsc", "impressions_gsc", "position_gsc", "position_sem_best",
                                "kw_opportunite", "volume_semantique", "nb_kw", "_ancres_txt"] if c in p.columns]
        if imp_cols:
            det = det.merge(p[["url"] + imp_cols], on="url", how="left")
        for _pc in ["position_gsc", "position_sem_best"]:
            if _pc in det.columns:
                det[_pc] = pd.to_numeric(det[_pc], errors="coerce").round(1)
        det = det.rename(columns={"type": "Type", "gravite": "Gravité", "url": "URL", "detail": "Détail",
                                  "cluster": "Cluster", "clics_gsc": "Clics GSC",
                                  "impressions_gsc": "Impressions GSC", "position_gsc": "Position GSC",
                                  "position_sem_best": "Position sém. (best)", "kw_opportunite": "Kw en opportunité (4-20)",
                                  "volume_semantique": "Volume sém.", "nb_kw": "Nb mots-clés",
                                  "_ancres_txt": "Ancres en place (nb liens)"})
        keep = [c for c in ["Type", "Gravité", "Clics GSC", "Position GSC", "Position sém. (best)",
                            "Kw en opportunité (4-20)", "Volume sém.", "Nb mots-clés",
                            "Ancres en place (nb liens)", "Cluster", "URL", "Détail"] if c in det.columns]
        det["_g"] = det["Gravité"].map({"Élevée": 0, "Moyenne": 1, "Faible": 2}).fillna(3)
        sort_by = ["_g"] + [c for c in ["Clics GSC", "Impressions GSC"] if c in det.columns]
        det = det.sort_values(sort_by, ascending=[True] + [False] * (len(sort_by) - 1))[keep]
    else:
        det = pd.DataFrame(columns=["Type", "Gravité", "Cluster", "URL", "Détail"])
    r2 = end + 2
    ws.cell(row=r2, column=1, value="Détail complet (trié par gravité puis importance)").font = SUB_FONT
    d_end = _write_df(ws, det.reset_index(drop=True), start_row=r2 + 1)
    if "Gravité" in det.columns:
        gcol = list(det.columns).index("Gravité") + 1
        for row in ws.iter_rows(min_row=r2 + 2, max_row=d_end - 1, min_col=gcol, max_col=gcol):
            for cell in row:
                if cell.value in GRAVITE_FILL:
                    cell.fill = GRAVITE_FILL[cell.value]


# 5. Prescription liens ----------------------------------------------------
def _sheet_prescriptions(wb, results):
    ws = wb.create_sheet("5. Prescription liens")
    ws["A1"] = "Prescription de liens internes — livrable central"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Une ligne = un lien à créer, de « Page source » vers « Page cible », avec l'ancre indiquée, en contenu. "
                "Trié pour traiter d'abord les pages cibles à fort trafic / fort volume.")
    ws["A2"].font = Font(italic=True, color=BLUE)
    presc = results["prescriptions"].copy()
    # Enrichir avec l'importance business de la page cible (clics / impressions / volume) et trier dessus.
    pg = results.get("pages")
    if pg is not None and len(presc) and "page_cible" in presc.columns:
        imp_cols = [c for c in ["clics_gsc", "impressions_gsc", "position_gsc", "volume_semantique",
                                "_kw_detail"] if c in pg.columns]
        if imp_cols:
            imp = pg[["url"] + imp_cols].rename(columns={"url": "page_cible", "clics_gsc": "clics_cible",
                    "impressions_gsc": "impr_cible", "position_gsc": "pos_cible", "volume_semantique": "vol_cible",
                    "_kw_detail": "kw_cible"})
            if "pos_cible" in imp.columns:
                imp["pos_cible"] = pd.to_numeric(imp["pos_cible"], errors="coerce").round(1)
            presc = presc.merge(imp, on="page_cible", how="left")
            sort_by = [c for c in ["clics_cible", "vol_cible", "impr_cible"] if c in presc.columns]
            if sort_by:
                presc = presc.sort_values(sort_by, ascending=False)
    ren = {"page_source": "Page source", "page_cible": "Page cible", "cluster_cible": "Cluster cible",
           "ancre_recommandee": "Ancre recommandée", "emplacement_suggere": "Emplacement",
           "priorite": "Priorité", "pr_source": "PR source", "match_kw": "Match mot-clé",
           "potentiel_cible": "Potentiel cible", "deficit_cible": "Déficit cible",
           "clics_cible": "Clics GSC cible", "impr_cible": "Impressions cible",
           "pos_cible": "Position GSC cible (moy.)", "vol_cible": "Volume sém. cible (somme mots-clés)",
           "kw_cible": "Mots-clés cibles (volume · position)"}
    order = ["priorite", "page_cible", "clics_cible", "pos_cible", "kw_cible", "vol_cible", "cluster_cible",
             "page_source", "ancre_recommandee", "emplacement_suggere", "pr_source", "match_kw",
             "potentiel_cible", "deficit_cible"]
    order = [c for c in order if c in presc.columns]
    presc = presc[order].rename(columns=ren) if len(presc) else presc.rename(columns=ren)
    end = _write_df(ws, presc, start_row=4)
    if "Priorité" in presc.columns:
        pcol = list(presc.columns).index("Priorité") + 1
        pf = {1: RED, 2: ORANGE, 3: GREEN}
        for row in ws.iter_rows(min_row=5, max_row=end - 1, min_col=pcol, max_col=pcol):
            for cell in row:
                if cell.value in pf:
                    cell.fill = PatternFill("solid", fgColor=pf[cell.value])


# 6. Analyse par page ------------------------------------------------------
def _sheet_analyse_page(wb, p):
    ws = wb.create_sheet("6. Analyse par page")
    ws["A1"] = "Analyse par page — triée par importance (clics GSC puis volume sémantique)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Toutes les pages, les plus importantes en haut. « Orpheline » = aucun lien interne entrant valide."
    ws["A2"].font = Font(italic=True, color=BLUE)
    d = p.copy()
    if "_ancres_txt" not in d.columns:
        d["_ancres_txt"] = "(aucune)"
    d["_orph"] = d["orpheline"].map({True: "OUI", False: ""}) if "orpheline" in d.columns else ""
    sort_cols = [c for c in ["clics_gsc", "volume_semantique", "impressions_gsc"] if c in d.columns] or ["score_composite"]
    d = d.sort_values(sort_cols, ascending=False)
    cols = [c for c in ["url", "cluster", "quadrant_code", "clics_gsc", "impressions_gsc", "volume_semantique",
            "pr_interne", "depth", "liens_ctx_entrants", "liens_entrants_total", "_orph", "_ancres_txt",
            "potentiel_business", "deficit_maillage"] if c in d.columns]
    ren = {"url": "URL", "cluster": "Cluster", "quadrant_code": "Quadrant",
           "clics_gsc": "Clics GSC", "impressions_gsc": "Impressions GSC", "volume_semantique": "Volume sém.",
           "pr_interne": "PageRank interne", "depth": "Profondeur", "liens_ctx_entrants": "Liens ctx entrants",
           "liens_entrants_total": "Liens entrants (tous)", "_orph": "Orpheline",
           "_ancres_txt": "Ancres utilisées (nb liens)",
           "potentiel_business": "Potentiel", "deficit_maillage": "Déficit"}
    df = d[cols].reset_index(drop=True).rename(columns=ren)
    end = _write_df(ws, df, start_row=4)
    if "Quadrant" in df.columns:
        qcol = list(df.columns).index("Quadrant") + 1
        for row in ws.iter_rows(min_row=5, max_row=end - 1, min_col=qcol, max_col=qcol):
            for cell in row:
                if cell.value in QUAD_FILL:
                    cell.fill = QUAD_FILL[cell.value]


# 7. Analyse ancres (en %) -------------------------------------------------
def _sheet_ancres(wb, p):
    ws = wb.create_sheet("7. Analyse ancres")
    ws["A1"] = "Analyse des ancres (en pourcentages)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("« Ancres utilisées » = les ancres réellement en place vers la page (liens en contenu), avec le nombre "
                "de liens entre parenthèses. Trié par importance (clics GSC). Objectif : diversité haute, "
                "dominante et génériques basses.")
    ws["A2"].font = Font(italic=True, color=BLUE)
    ws["A2"].alignment = Alignment(wrap_text=True)

    d = p[p["liens_ctx_entrants"] > 0].copy()
    if "_ancres_txt" not in d.columns:
        d["_ancres_txt"] = "(aucune)"
    d["Diversité %"] = (pd.to_numeric(d["ratio_diversite_ancres"], errors="coerce") * 100).round(0)
    d["Ancre dominante %"] = (pd.to_numeric(d["part_ancre_dominante"], errors="coerce") * 100).round(0)
    d["Génériques %"] = (pd.to_numeric(d["part_ancres_generiques"], errors="coerce") * 100).round(0)
    sort_cols = [c for c in ["clics_gsc", "impressions_gsc"] if c in d.columns] or ["Génériques %"]
    d = d.sort_values(sort_cols, ascending=False)
    cols = ["url", "cluster", "quadrant_code", "clics_gsc", "impressions_gsc", "volume_semantique",
            "liens_ctx_entrants", "_ancres_txt", "Diversité %", "Ancre dominante %", "Génériques %"]
    cols = [c for c in cols if c in d.columns]
    ren = {"url": "URL", "cluster": "Cluster", "quadrant_code": "Quadrant",
           "clics_gsc": "Clics GSC", "impressions_gsc": "Impressions GSC", "volume_semantique": "Volume sém.",
           "liens_ctx_entrants": "Liens entrants", "_ancres_txt": "Ancres utilisées (nb liens)"}
    df = d[cols].reset_index(drop=True).rename(columns=ren)
    end = _write_df(ws, df, start_row=4)
    # coloration lecture
    if "Génériques %" in df.columns:
        gc = list(df.columns).index("Génériques %") + 1
        dc = list(df.columns).index("Ancre dominante %") + 1
        for row in ws.iter_rows(min_row=5, max_row=end - 1):
            gv, dv = row[gc - 1].value, row[dc - 1].value
            if isinstance(gv, (int, float)) and gv > 50:
                row[gc - 1].fill = PatternFill("solid", fgColor=RED)
            if isinstance(dv, (int, float)) and dv > 70:
                row[dc - 1].fill = PatternFill("solid", fgColor=ORANGE)


# 8. Silotage clusters -----------------------------------------------------
def _sheet_silotage(wb, results):
    ws = wb.create_sheet("8. Silotage clusters")
    ws["A1"] = "Silotage inter-clusters"
    ws["A1"].font = TITLE_FONT
    mat = results["silo_matrix_pct"].copy()
    ws.cell(row=3, column=1, value="Matrice source → cible (% des liens contextuels sortants)").font = SUB_FONT
    m = mat.reset_index()
    m.columns = ["source \\ cible"] + list(mat.columns)
    end = _write_df(ws, m, start_row=4, autofilter=False, band=False)
    for row in ws.iter_rows(min_row=5, max_row=end - 1, min_col=2, max_col=1 + len(mat.columns)):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                if cell.value >= 0.85:
                    cell.fill = PatternFill("solid", fgColor=RED)
                elif cell.value >= 0.5:
                    cell.fill = PatternFill("solid", fgColor=ORANGE)
                elif cell.value >= 0.1:
                    cell.fill = PatternFill("solid", fgColor=LIGHT)
    flags = results["silo_flags"]
    r2 = end + 2
    ws.cell(row=r2, column=1, value="Diagnostic par cluster").font = SUB_FONT
    _write_df(ws, flags, start_row=r2 + 1, autofilter=False)
