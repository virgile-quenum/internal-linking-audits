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
    # Autofilter et gel des volets : UNIQUEMENT pour un tableau ancré en haut de feuille
    # (start_row <= 6) et une seule fois par feuille. Sinon, sur les onglets multi-tableaux
    # (Plan d'action, Tableau de bord, Silotage), le freeze se posait au milieu de la feuille
    # (ligne 115...) et bloquait tout le scroll.
    top_anchored = start_row <= 6
    if autofilter and header and top_anchored and ws.auto_filter.ref is None:
        last = get_column_letter(start_col + len(df.columns) - 1)
        ws.auto_filter.ref = f"{get_column_letter(start_col)}{start_row}:{last}{r-1}"
    if header and top_anchored and ws.freeze_panes is None:
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
        # Ancre vide = lien porté par une image ou un bouton (pas de texte cliquable). On le dit
        # explicitement plutôt que d'afficher "(vide)", qui laissait penser à un bug (retour consultant).
        _c["anchor"] = _c["anchor"].fillna("").astype(str).str.strip().replace("", "(image/bouton)")
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

    # Onglets (méthodologie à la fin). L'onglet Mots-clés n'apparaît que si l'étude sémantique
    # est fournie : c'est la seule source fiable pour relier un terme à une URL.
    _safe(_sheet_synthese_ecrite, wb, results, p, site_name, notes)   # 1
    _safe(_sheet_plan_action, wb, results, p)                          # 2
    _safe(_sheet_nav_fix, wb, results)                                 # 3 menu & footer
    _safe(_sheet_prescriptions, wb, results)                           # 4
    _safe(_sheet_motscles, wb, results)                                # 5 (si sémantique)
    _safe(_sheet_detail, wb, results, p)                               # 6
    _safe(_sheet_legende, wb, cfg, results, p)                         # 7

    # Texte d'intro en A2 (onglets tableaux) : fusionné sur la largeur du tableau, aligné EN HAUT,
    # ligne assez haute pour être lu d'un bloc (retour : texte trop bas / cellule trop étroite).
    for _ws in wb.worksheets:
        if not _ws.title[:1].isdigit():
            continue  # ne pas toucher la feuille par défaut (sinon elle n'est plus vue comme vide)
        try:
            _v = _ws["A2"].value
            if not _v or _ws.title.startswith(("1.", "7.")):
                continue
            _last = max(8, min(_ws.max_column, 12))
            _ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=_last)
            _ws["A2"].alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            _ws.row_dimensions[2].height = max(34, min(90, 15 * (len(str(_v)) // 140 + 1) + 6))
        except Exception as _e:
            notes.append(f"intro A2 {_ws.title}: {_e}")

    # Retirer la feuille "Sheet" vide créée par défaut par openpyxl (plus repeuplée).
    for _dft in ("Sheet", "Feuille", "Feuil1"):
        if _dft in wb.sheetnames and wb[_dft].max_row <= 1 and wb[_dft].max_column <= 1:
            del wb[_dft]

    order = ["1. Synthèse", "2. Plan d'action", "3. Liens à corriger", "4. Prescription liens",
             "5. Mots-clés", "6. Détail par page", "7. Méthodologie"]
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
    ws = wb.create_sheet("7. Méthodologie")
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
    line("1. Synthèse", "L'analyse rédigée + les chiffres clés + la matrice de diagnostic. À lire en premier.")
    line("2. Plan d'action", "Le livrable actionnable : quoi faire, par page, du plus important au moins important.")
    line("3. Liens à corriger", "Les liens existants à modifier ou supprimer (menu/footer, puis dans le contenu page par page), avec l'URL à mettre à la place.")
    line("4. Prescription liens", "Les liens à créer : source, cible, ancre, emplacement, priorité.")
    line("5. Mots-clés", "Un terme par ligne, trié par opportunité : sur quels mots-clés remonter (si étude sémantique fournie).")
    line("6. Détail par page", "Le détail page par page (métriques, ancres, problèmes, silotage) pour aller au fond.")
    line("7. Méthodologie", "Ce document : comment lire, définitions, seuils.")
    gap()

    h2("Le diagnostic repose sur 2 axes (pas un score unique)")
    line("Axe X — Opportunité (0 à 10)", "Mesure l'upside réel d'une page. Croise la DEMANDE (volume de "
         "recherche + impressions GSC) avec le GAP DE POSITION : une page déjà en top 3 pèse peu (déjà gagnée), "
         "le maximum est en position 4-10 (bas de page 1 / haut de page 2). Quand la sémantique est fournie, "
         "le calcul est fait mot-clé par mot-clé. 10 = fort upside, 0 = peu à aller chercher.")
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
    ws = wb.create_sheet("2. Plan d'action")
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
        # Rôle (jalon C) : les pages "Exclue" (FAQ, légal, tunnel... selon la config) sortent du plan d'action.
        if str(r.get("role", "Cible + Source")) == "Exclue":
            continue

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
        rows.append({"Page": url, "Typologie": r.get("typologie", ""), "Clics GSC": clics,
                     "Position GSC (moy.)": pos,
                     "Mots-clés cibles (volume · position)": r.get("_kw_detail", ""),
                     "Action": action, "Emplacement": emp, "Quoi faire": quoi})

    df = pd.DataFrame(rows, columns=["Page", "Typologie", "Clics GSC", "Position GSC (moy.)",
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

    # Maillage par typologie de page (jalon C)
    if "typologie" in p.columns:
        akw = dict(nb_pages=("url", "size"),
                   role=("role", lambda s: s.astype(str).mode().iat[0] if len(s) else ""),
                   liens_ctx_moy=("liens_ctx_entrants", "mean"),
                   orphelines=("orpheline", "sum"))
        if "clics_gsc" in p.columns:
            akw["clics_gsc"] = ("clics_gsc", "sum")
        g = p.groupby("typologie").agg(**akw).reset_index()
        g["liens_ctx_moy"] = pd.to_numeric(g["liens_ctx_moy"], errors="coerce").round(1)
        g["orphelines"] = pd.to_numeric(g["orphelines"], errors="coerce").fillna(0).astype(int)
        g = g.sort_values("nb_pages", ascending=False).rename(columns={
            "typologie": "Typologie", "role": "Rôle (maillage)", "nb_pages": "Nb pages",
            "liens_ctx_moy": "Liens contenu entrants (moy.)", "orphelines": "Orphelines", "clics_gsc": "Clics GSC"})
        rT = end + 2
        ws.cell(row=rT, column=1,
                value="Maillage par typologie de page (rôle défini à l'étape segmentation, éditable)").font = SUB_FONT
        end = _write_df(ws, g.reset_index(drop=True), start_row=rT + 1)

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

    # ③ Liens à corriger : désormais dans leur propre onglet « 3. Liens à corriger » (plus lisible).
    _rp = results.get("report") or {}
    _nv = _rp.get("_nav_a_corriger")
    _nn = len(_nv) if _nv is not None else 0
    _nc = _rp.get("liens_contenu_a_corriger", 0)
    if _nn or _nc:
        r3 = end + 2
        ws.cell(row=r3, column=1,
                value=f"③ Liens à corriger : {_nn} dans le menu/footer, {_nc} dans le contenu "
                      f"(dont {_rp.get('liens_contenu_a_supprimer', 0)} à supprimer), "
                      f"voir l'onglet « 3. Liens à corriger ».").font = SUB_FONT
        end = r3 + 1

    # Pages orphelines "hors crawl" : vues par GSC/sémantique mais jamais atteintes par le crawl
    hc = (results.get("report") or {}).get("_pages_hors_crawl")
    if hc is not None and len(hc):
        r4 = end + 2
        ws.cell(row=r4, column=1,
                value="Pages hors crawl — trafic/volume mais aucun lien interne trouvé (orphelines de fait)").font = SUB_FONT
        _cn = ws.cell(row=r4 + 1, column=1,
                      value="Ces pages existent (clics GSC ou volume sémantique) mais le crawler ne les a jamais "
                            "atteintes : aucune page du site ne pointe vers elles. À reconnecter par un lien interne.")
        _cn.font = Font(italic=True, color=BLUE)
        _write_df(ws, hc.head(80).reset_index(drop=True), start_row=r4 + 2)


# 2. Synthèse écrite -------------------------------------------------------
def _sheet_synthese_ecrite(wb, results, p, site_name, notes=None):
    ws = wb.create_sheet("1. Synthèse")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 92
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 30

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
            _inl_col = "liens_entrants_total" if "liens_entrants_total" in q3.columns else "liens_ctx_entrants"
            g = q3.groupby("cluster").agg(n=("url", "size"), pr=("pr_interne", "mean"),
                                          inl=(_inl_col, "mean"))
            # On surface les clusters réellement sur-maillés : beaucoup de liens entrants (nav incluse)
            # pour peu de valeur. On ignore les segments à 1 page et quasi sans liens (bruit type légal isolé).
            g = g[g["inl"] >= 2].sort_values("inl", ascending=False).head(4)
            if len(g):
                rows.append(("h2", "Sur-maillage : pages très liées en interne pour un faible potentiel"))
                rows.append(("body", "Elles reçoivent beaucoup de liens internes (souvent via le menu ou le pied de page) pour peu de valeur business. Réallouer ce maillage vers les pages Q1."))
                for c, x in g.iterrows():
                    rows.append(("bullet", f"« {c} » : {int(x['n'])} page(s), {x['inl']:.0f} liens internes entrants en moyenne (PageRank {x['pr']:.1f}/10)."))
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
        _nav = (results.get("report") or {}).get("_nav_a_corriger")
        n_nav = len(_nav) if _nav is not None else 0
        rows.append(("h2", "Recommandations, par ordre de priorité"))
        k = 1
        if n_nav:
            rows.append(("bullet", f"{k}. Corriger les {n_nav} liens de menu / footer à problème (onglet « Liens à corriger ») : une correction dans le gabarit règle toutes les pages. Effort faible, gain immédiat."))
            k += 1
        _rp = results.get("report") or {}
        if _rp.get("liens_contenu_a_corriger", 0):
            rows.append(("bullet", f"{k}. Nettoyer les liens de contenu cassés ou redirigés : {_rp.get('liens_contenu_a_supprimer', 0)} à supprimer (404) et {_rp['liens_contenu_a_corriger'] - _rp.get('liens_contenu_a_supprimer', 0)} à mettre à jour, page par page (onglet « Liens à corriger »)."))
            k += 1
        rows.append(("bullet", f"{k}. Déployer les {n_presc} liens prescrits (onglet « Prescription liens »), en commençant par les {n_p1} de priorité 1 (produits et catégories d'abord)."))
        rows.append(("bullet", f"{k+1}. Cibler en priorité les pages à fort upside : forte demande mais position 4 à 20. Les pages déjà en top 3 sont volontairement écartées, elles ont peu à gagner."))
        rows.append(("bullet", f"{k+2}. Donner au moins un lien contextuel entrant aux pages orphelines."))
        rows.append(("bullet", f"{k+3}. Corriger les ancres pauvres, trop homogènes ou portées par une image (sans texte) sur les pages Q1."))
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
    # --- Chiffres clés + matrice de diagnostic (ex-onglet Tableau de bord, fondu ici) ---
    rep = results.get("report", {})
    quad = p["quadrant_code"].value_counts().to_dict()
    r += 2
    ws.cell(row=r, column=2, value="Chiffres clés").font = SUB_FONT
    r += 1
    kpis = [("Pages analysées", rep.get("pages_total", len(p))),
            ("Liens internes valides", rep.get("liens_valides", "")),
            ("Pages orphelines", rep.get("pages_orphelines", int(p["orpheline"].sum()))),
            ("Q1 — Prioritaires", quad.get("Q1", 0)),
            ("Q3 — Sur-maillage", quad.get("Q3", 0)),
            ("Liens menu / footer à corriger", rep.get("urls_nav_a_corriger", 0)),
            ("Prescriptions générées", rep.get("prescriptions_generees", 0))]
    for label, val in kpis:
        a = ws.cell(row=r, column=2, value=label); a.font = Font(bold=True); a.border = BORDER; a.alignment = LEFT
        b = ws.cell(row=r, column=3, value=_safe_cell(val)); b.alignment = CENTER; b.border = BORDER
        if label.startswith("Q1"):
            a.fill = QUAD_FILL["Q1"]; b.fill = QUAD_FILL["Q1"]
        r += 1

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
    _mx(r, 3, "Faible opportunité", HEADER_FILL, white=True)
    _mx(r, 4, "Forte opportunité", HEADER_FILL, white=True)
    r += 1
    _mx(r, 2, "Fort déficit de maillage", HEADER_FILL, white=True)
    _mx(r, 3, f"Q4 — Ne rien faire\n{quad.get('Q4', 0)} pages", QUAD_FILL["Q4"])
    _mx(r, 4, f"Q1 — PRIORITAIRE\n{quad.get('Q1', 0)} pages", QUAD_FILL["Q1"])
    ws.row_dimensions[r].height = 30
    r += 1
    _mx(r, 2, "Bon maillage", HEADER_FILL, white=True)
    _mx(r, 3, f"Q3 — Sur-maillage\n{quad.get('Q3', 0)} pages", QUAD_FILL["Q3"])
    _mx(r, 4, f"Q2 — Hors scope\n{quad.get('Q2', 0)} pages", QUAD_FILL["Q2"])
    ws.row_dimensions[r].height = 30

    if notes:
        r += 2
        ws.cell(row=r, column=2, value="Notes techniques : " + " | ".join(notes)).font = Font(italic=True, color="999999", size=9)


# (ex-onglet Tableau de bord : fondu dans la Synthèse) ---------------------
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


# Section Problèmes (fondue dans l'onglet Détail) --------------------------
def _section_problemes(ws, results, p, start_row):
    ws.cell(row=start_row, column=1, value="Problèmes identifiés").font = TITLE_FONT
    ws.cell(row=start_row + 1, column=1,
            value="Pour chaque type : sa gravité, son volume, sa part des pages, et un exemple concret.").font = Font(italic=True, color=BLUE)

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
    ws.cell(row=start_row + 2, column=1, value="Récapitulatif par type").font = SUB_FONT
    end = _write_df(ws, recap, start_row=start_row + 3, autofilter=False)
    if "Gravité" in recap.columns:
        gcol = list(recap.columns).index("Gravité") + 1
        for row in ws.iter_rows(min_row=start_row + 4, max_row=end - 1, min_col=gcol, max_col=gcol):
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
    return d_end


# 3. Liens à corriger : menu & footer (gabarit) + contenu (page par page) ----
def _color_actions(ws, first_row, last_row, col=1):
    for row in ws.iter_rows(min_row=first_row, max_row=last_row, min_col=col, max_col=col):
        for cell in row:
            v = str(cell.value or "")
            if v.startswith("MODIFIER"):
                cell.fill = PatternFill("solid", fgColor=ORANGE)
            elif v.startswith(("SUPPRIMER", "RETIRER")):
                cell.fill = PatternFill("solid", fgColor=RED)
            elif v.startswith("VÉRIFIER"):
                cell.fill = PatternFill("solid", fgColor=LIGHT)


def _fill_repl(d):
    if "remplacer_par" in d.columns:
        d["remplacer_par"] = [
            r if (isinstance(r, str) and r) else
            ("URL finale à identifier (absente de l'export)" if str(a).startswith("MODIFIER") else "")
            for r, a in zip(d["remplacer_par"], d["action"])]
    return d


def _sheet_nav_fix(wb, results):
    rep = results.get("report") or {}
    nav = rep.get("_nav_a_corriger")
    cfix = rep.get("_contenu_a_corriger")
    ws = wb.create_sheet("3. Liens à corriger")
    ws["A1"] = "Liens à corriger : à modifier ou supprimer"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Liens existants qui pointent vers une page à problème. MODIFIER = remplacer l'URL du lien par "
                "celle de la colonne « Remplacer par ». SUPPRIMER = lien cassé (404), à enlever ou remplacer. "
                "VÉRIFIER = cas souvent voulu (compte, panier, redirection temporaire), à trancher. "
                "Bloc 1 : menu & footer, une correction dans le gabarit règle toutes les pages. "
                "Bloc 2 : liens dans le contenu, à corriger page par page dans le CMS. "
                "Les URL bloquées par robots.txt (compte client, avis…) sont volontairement ignorées.")
    ws["A2"].font = Font(italic=True, color=BLUE)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")

    # --- Bloc 1 : menu / header / footer ---
    r = 4
    n_nav = len(nav) if nav is not None else 0
    ws.cell(row=r, column=1, value=f"1. Menu, header & footer ({n_nav} lien(s)) — à corriger dans le gabarit").font = SUB_FONT
    if n_nav:
        d = _fill_repl(nav.copy())
        ren = {"action": "Action", "emplacement": "Emplacement", "lien_actuel": "Lien actuel (URL pointée)",
               "remplacer_par": "Remplacer par", "probleme": "Problème", "nb_pages": "Nb pages concernées",
               "nb_liens": "Nb liens"}
        cols = [c for c in ["action", "emplacement", "lien_actuel", "remplacer_par", "probleme",
                            "nb_pages", "nb_liens"] if c in d.columns]
        end = _write_df(ws, d[cols].rename(columns=ren).reset_index(drop=True), start_row=r + 1, autofilter=False)
        _color_actions(ws, r + 2, end - 1)
    else:
        ws.cell(row=r + 1, column=1, value="Aucun : tous les liens de navigation pointent vers des pages valides.").font = Font(italic=True)
        end = r + 2

    # --- Bloc 2 : liens dans le contenu (page à éditer) ---
    r2 = end + 2
    n_c = len(cfix) if cfix is not None else 0
    n_sup = int(cfix["action"].str.startswith("SUPPRIMER").sum()) if n_c else 0
    ws.cell(row=r2, column=1,
            value=f"2. Dans le contenu ({n_c} lien(s), dont {n_sup} à supprimer) — à corriger page par page").font = SUB_FONT
    if n_c:
        d = _fill_repl(cfix.copy())
        pg = results.get("pages")
        if pg is not None and "clics_gsc" in pg.columns:
            d["clics_page"] = d["page_a_editer"].map(dict(zip(pg["url"], pg["clics_gsc"]))).fillna(0).astype(int)
        else:
            d["clics_page"] = 0
        # 404 d'abord (à supprimer), puis redirections/canoniques ; dans chaque groupe, les pages
        # à plus fort trafic en premier.
        d = d.sort_values(["_rang", "clics_page", "page_a_editer"], ascending=[True, False, True])
        ren = {"action": "Action", "page_a_editer": "Page à éditer", "clics_page": "Clics GSC page",
               "lien_actuel": "Lien actuel (URL pointée)", "ancre": "Ancre du lien",
               "remplacer_par": "Remplacer par", "probleme": "Problème", "nb_occurrences": "Occurrences"}
        cols = [c for c in ["action", "page_a_editer", "clics_page", "lien_actuel", "ancre", "remplacer_par",
                            "probleme", "nb_occurrences"] if c in d.columns]
        df = d[cols].rename(columns=ren).reset_index(drop=True)
        end2 = _write_df(ws, df, start_row=r2 + 1, autofilter=False)
        # filtre sur ce bloc (le plus long) pour trier / isoler une page ou une action
        ws.auto_filter.ref = f"A{r2 + 1}:{get_column_letter(len(df.columns))}{end2 - 1}"
        _color_actions(ws, r2 + 2, end2 - 1)
    else:
        ws.cell(row=r2 + 1, column=1, value="Aucun lien de contenu à corriger.").font = Font(italic=True)
    ws.column_dimensions["A"].width = 46


# 4. Prescription liens ----------------------------------------------------
def _sheet_prescriptions(wb, results):
    ws = wb.create_sheet("4. Prescription liens")
    ws["A1"] = "Prescription de liens internes — livrable central"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Une ligne = un lien à créer, de « Page source » vers « Page cible », en contenu. "
                "Trié par PRIORITÉ (1 = à faire d'abord : produits/catégories à fort upside). "
                "Les pages déjà en top 3 ne reçoivent pas de prescription (peu à gagner). "
                "« Ancre à affiner » = ancre dérivée du titre, à retravailler vers un mot-clé.")
    ws["A2"].font = Font(italic=True, color=BLUE)
    ws["A2"].alignment = Alignment(wrap_text=True)
    presc = results["prescriptions"].copy()
    # Enrichir avec l'importance business de la page cible (clics / volume) et trier PRIORITÉ puis clics.
    pg = results.get("pages")
    if pg is not None and len(presc) and "page_cible" in presc.columns:
        imp_cols = [c for c in ["clics_gsc", "volume_semantique", "_kw_detail"] if c in pg.columns]
        if imp_cols:
            imp = pg[["url"] + imp_cols].rename(columns={"url": "page_cible", "clics_gsc": "clics_cible",
                    "volume_semantique": "vol_cible", "_kw_detail": "kw_cible"})
            presc = presc.merge(imp, on="page_cible", how="left")
        sort_by = [c for c in ["priorite", "clics_cible", "vol_cible"] if c in presc.columns]
        if sort_by:
            presc = presc.sort_values(sort_by, ascending=[True] + [False] * (len(sort_by) - 1))
    ren = {"page_source": "Page source", "page_cible": "Page cible", "typologie_cible": "Typologie cible",
           "cluster_cible": "Cluster cible", "position_cible": "Position actuelle (meilleure)",
           "ancre_recommandee": "Ancre (ce lien)", "ancres_suggerees": "Ancres suggérées (varier)",
           "ancre_qualite": "Qualité ancre",
           "emplacement_suggere": "Emplacement", "priorite": "Priorité", "pr_source": "PR source",
           "match_kw": "Match mot-clé", "potentiel_cible": "Opportunité cible", "deficit_cible": "Déficit cible",
           "clics_cible": "Clics GSC cible", "vol_cible": "Volume sém. cible (somme mots-clés)",
           "kw_cible": "Mots-clés cibles (volume · position)"}
    order = ["priorite", "page_cible", "typologie_cible", "position_cible", "clics_cible", "kw_cible",
             "vol_cible", "cluster_cible", "page_source", "ancre_recommandee", "ancres_suggerees",
             "ancre_qualite", "emplacement_suggere", "pr_source", "match_kw", "potentiel_cible", "deficit_cible"]
    order = [c for c in order if c in presc.columns]
    presc = presc[order].rename(columns=ren) if len(presc) else presc.rename(columns=ren)
    end = _write_df(ws, presc, start_row=4, band=False)  # banding géré par cible ci-dessous
    if "Priorité" in presc.columns:
        pcol = list(presc.columns).index("Priorité") + 1
        pf = {1: RED, 2: ORANGE, 3: GREEN}
        for row in ws.iter_rows(min_row=5, max_row=end - 1, min_col=pcol, max_col=pcol):
            for cell in row:
                if cell.value in pf:
                    cell.fill = PatternFill("solid", fgColor=pf[cell.value])
    # Regroupement visuel par PAGE CIBLE : une même cible reçoit plusieurs liens (plusieurs sources).
    # On alterne une bande de couleur à chaque changement de cible pour distinguer les blocs d'un coup d'œil.
    if "Page cible" in presc.columns and len(presc):
        ccol = list(presc.columns).index("Page cible") + 1
        band_fill = PatternFill("solid", fgColor=LIGHT)
        prev, shade = None, False
        ncols = len(presc.columns)
        for i, ridx in enumerate(range(5, end)):
            cur = ws.cell(row=ridx, column=ccol).value
            if cur != prev:
                shade = not shade
                prev = cur
            if shade:
                for cc in range(1, ncols + 1):
                    if cc != pcol:  # ne pas écraser la couleur de priorité
                        ws.cell(row=ridx, column=cc).fill = band_fill


# 4. Mots-clés (croisement direct des termes à pousser) --------------------
def _sheet_motscles(wb, results):
    kw = results.get("kw_table")
    ws = wb.create_sheet("5. Mots-clés")
    ws["A1"] = "Mots-clés — sur quels termes remonter"
    ws["A1"].font = TITLE_FONT
    if kw is None or not len(kw):
        ws["A2"] = ("Cet onglet se remplit avec l'étude sémantique (mot-clé, volume, position, URL cible). "
                    "Aucune sémantique fournie sur ce run : ajoute-la à l'étape 1 pour croiser les termes à pousser. "
                    "La Search Console seule donne les clics par page, mais ne relie pas un mot-clé précis à une URL.")
        ws["A2"].font = Font(italic=True, color=BLUE)
        ws["A2"].alignment = Alignment(wrap_text=True)
        ws.column_dimensions["A"].width = 120
        return
    ws["A2"] = ("Un mot-clé par ligne, trié par OPPORTUNITÉ (volume x gap de position). "
                "« À pousser » = position 4 à 20 : proche de la 1re page, c'est là que le maillage paie. "
                "Filtre la colonne pour isoler les termes à travailler, puis croise avec la page cible.")
    ws["A2"].font = Font(italic=True, color=BLUE)
    ws["A2"].alignment = Alignment(wrap_text=True)

    d = kw.copy()
    if "position" in d.columns:
        d["position"] = pd.to_numeric(d["position"], errors="coerce").round(1)
    if "a_pousser" in d.columns:
        d["a_pousser"] = d["a_pousser"].map({True: "OUI", False: ""})
    ren = {"keyword": "Mot-clé", "volume": "Volume", "position": "Position actuelle",
           "opp_kw": "Opportunité", "a_pousser": "À pousser (pos 4-20)", "url": "Page cible (arbitrée)",
           "nb_urls_candidates": "Nb URLs candidates", "cluster": "Cluster", "typologie": "Typologie",
           "quadrant_code": "Quadrant", "liens_ctx_entrants": "Liens ctx entrants", "pr_interne": "PageRank interne"}
    cols = [c for c in ["keyword", "volume", "position", "opp_kw", "a_pousser", "url", "nb_urls_candidates",
                        "cluster", "typologie", "quadrant_code", "liens_ctx_entrants", "pr_interne"] if c in d.columns]
    df = d[cols].rename(columns=ren)
    end = _write_df(ws, df, start_row=4)
    if "À pousser (pos 4-20)" in df.columns:
        acol = list(df.columns).index("À pousser (pos 4-20)") + 1
        for row in ws.iter_rows(min_row=5, max_row=end - 1, min_col=acol, max_col=acol):
            for cell in row:
                if cell.value == "OUI":
                    cell.fill = PatternFill("solid", fgColor=GREEN)


# 5. Détail par page (page + ancres + problèmes + silotage, fondus) --------
def _sheet_detail(wb, results, p):
    ws = wb.create_sheet("6. Détail par page")
    ws["A1"] = "Détail par page — triée par importance (clics GSC puis volume sémantique)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Toutes les pages, les plus importantes en haut. « Orpheline » = aucun lien interne entrant. "
                "« Position » = meilleure position connue (GSC/sémantique). Colonnes ancres en % : viser diversité haute, "
                "dominante et génériques basses. Plus bas : problèmes détaillés et silotage inter-clusters.")
    ws["A2"].font = Font(italic=True, color=BLUE)
    ws["A2"].alignment = Alignment(wrap_text=True)
    d = p.copy()
    if "_ancres_txt" not in d.columns:
        d["_ancres_txt"] = "(aucune)"
    d["_orph"] = d["orpheline"].map({True: "OUI", False: ""}) if "orpheline" in d.columns else ""
    for _src, _dst in [("ratio_diversite_ancres", "_div"), ("part_ancre_dominante", "_dom"),
                       ("part_ancres_generiques", "_gen")]:
        d[_dst] = (pd.to_numeric(d.get(_src), errors="coerce") * 100).round(0) if _src in d.columns else ""
    if "_best_position" in d.columns:
        d["_pos"] = pd.to_numeric(d["_best_position"], errors="coerce").round(1)
    sort_cols = [c for c in ["clics_gsc", "volume_semantique", "impressions_gsc"] if c in d.columns] or ["score_composite"]
    d = d.sort_values(sort_cols, ascending=False)
    cols = [c for c in ["url", "cluster", "typologie", "quadrant_code", "_pos", "clics_gsc", "impressions_gsc",
            "volume_semantique", "pr_interne", "depth", "liens_ctx_entrants", "liens_entrants_total", "_orph",
            "_ancres_txt", "_div", "_dom", "_gen", "potentiel_business", "deficit_maillage"] if c in d.columns]
    ren = {"url": "URL", "cluster": "Cluster", "typologie": "Typologie", "quadrant_code": "Quadrant",
           "_pos": "Position (meilleure)", "clics_gsc": "Clics GSC", "impressions_gsc": "Impressions GSC",
           "volume_semantique": "Volume sém.", "pr_interne": "PageRank interne", "depth": "Profondeur",
           "liens_ctx_entrants": "Liens ctx entrants", "liens_entrants_total": "Liens entrants (tous)",
           "_orph": "Orpheline", "_ancres_txt": "Ancres utilisées (nb liens)", "_div": "Diversité %",
           "_dom": "Ancre dominante %", "_gen": "Génériques %",
           "potentiel_business": "Opportunité", "deficit_maillage": "Déficit"}
    df = d[cols].reset_index(drop=True).rename(columns=ren)
    end = _write_df(ws, df, start_row=4)
    if "Quadrant" in df.columns:
        qcol = list(df.columns).index("Quadrant") + 1
        for row in ws.iter_rows(min_row=5, max_row=end - 1, min_col=qcol, max_col=qcol):
            for cell in row:
                if cell.value in QUAD_FILL:
                    cell.fill = QUAD_FILL[cell.value]

    end = _section_problemes(ws, results, p, end + 2)
    _section_silotage(ws, results, end + 2)


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


# Section Silotage (fondue dans l'onglet Détail) ---------------------------
def _section_silotage(ws, results, start_row):
    ws.cell(row=start_row, column=1, value="Silotage inter-clusters").font = TITLE_FONT
    mat = results["silo_matrix_pct"].copy()
    ws.cell(row=start_row + 1, column=1,
            value="Matrice source → cible (% des liens contextuels sortants)").font = SUB_FONT
    m = mat.reset_index()
    m.columns = ["source \\ cible"] + list(mat.columns)
    end = _write_df(ws, m, start_row=start_row + 2, autofilter=False, band=False)
    for row in ws.iter_rows(min_row=start_row + 3, max_row=end - 1, min_col=2, max_col=1 + len(mat.columns)):
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
    return _write_df(ws, flags, start_row=r2 + 1, autofilter=False)
