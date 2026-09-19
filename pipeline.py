"""
Pipeline d'audit de maillage interne.
Source de crawl canonique : OnCrawl (InRank = PageRank interne natif, 0-10).
Entrées attendues sous forme de DataFrames aux colonnes CANONIQUES
(le mapping depuis les noms réels OnCrawl/GSC/GA4 est fait en amont par l'app).

Colonnes canoniques :
  pages : url, inrank, depth, status_code, indexable, word_count, title, h1
  links : source, target, anchor, follow, position
  gsc   : url, clicks, impressions, ctr, position
  ga4   : url, sessions, conversions, value        (optionnel)
  sem   : keyword, volume, target_url, cluster      (optionnel)

Tous les seuils sont dans DEFAULT_CONFIG et surchargés depuis l'app.
"""
from __future__ import annotations
import re
import logging
from urllib.parse import urlsplit, urlunsplit

import numpy as np
import pandas as pd

log = logging.getLogger("maillage")

# --------------------------------------------------------------------------
# CONFIG : tous les seuils modifiables, aucun en dur dans le code
# --------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "nav_repeat_threshold": 0.80,      # lien présent sur >X% des pages -> reclassé navigation
    "closed_cluster_threshold": 0.85,  # >X% de liens intra -> cluster fermé
    "homogeneous_anchor_threshold": 0.70,  # une ancre > X% des occurrences -> sur-homogène
    "poor_anchor_threshold": 0.50,     # >X% d'ancres génériques/vides -> ancre pauvre
    "quadrant_x_split": 5.0,           # seuil axe X (potentiel business) 0-10
    "quadrant_y_split": 5.0,           # seuil axe Y (déficit maillage) 0-10
    "incoherent_link_threshold": 0.15, # similarité sous ce seuil -> lien incohérent (si TF-IDF actif)
    "strip_query_params": True,        # retirer les paramètres d'URL à la normalisation
    "keep_params": [],                 # paramètres à conserver malgré tout
}

GENERIC_ANCHORS = {
    "", "cliquez ici", "clique ici", "ici", "en savoir plus", "savoir plus",
    "lire la suite", "lire l'article", "voir plus", "plus", "en lire plus",
    "click here", "read more", "learn more", "voir", "découvrir", "detail",
    "détail", "details", "lien", "link", "accueil", "home",
}

POSITION_MAP = {
    "content": "Content", "contenu": "Content", "body": "Content", "main": "Content",
    "navigation": "Navigation", "nav": "Navigation", "menu": "Navigation",
    "header": "Header", "entete": "Header", "en-tete": "Header", "top": "Header",
    "footer": "Footer", "pied": "Footer", "bottom": "Footer",
    "sidebar": "Sidebar", "aside": "Sidebar", "lateral": "Sidebar",
}


# --------------------------------------------------------------------------
# Normalisation
# --------------------------------------------------------------------------
def normalize_url(u, cfg):
    if not isinstance(u, str) or not u.strip():
        return None
    u = u.strip()
    try:
        parts = urlsplit(u)
    except ValueError:
        return None
    scheme = "https"  # on unifie le protocole
    netloc = parts.netloc.lower()
    if netloc.startswith("www."):
        netloc_key = netloc  # on garde www tel quel pour rester cohérent avec le crawl
    path = parts.path or "/"
    if len(path) > 1 and path.endswith("/"):
        path = path.rstrip("/")
    query = parts.query
    if cfg.get("strip_query_params", True):
        if cfg.get("keep_params"):
            kept = [kv for kv in query.split("&")
                    if kv.split("=")[0] in cfg["keep_params"]]
            query = "&".join(kept)
        else:
            query = ""
    return urlunsplit((scheme, netloc, path, query, ""))


def normalize_position(p):
    if not isinstance(p, str):
        return "Content"
    key = p.strip().lower()
    return POSITION_MAP.get(key, "Content" if key in ("", "nan") else p.strip().capitalize())


def _pct_rank(s: pd.Series) -> pd.Series:
    """Normalisation robuste 0-1 par rang percentile (insensible aux outliers)."""
    s = pd.to_numeric(s, errors="coerce").fillna(0.0)
    if s.nunique() <= 1:
        return pd.Series(np.zeros(len(s)), index=s.index)
    return s.rank(pct=True)


# --------------------------------------------------------------------------
# Clustering : proposition automatique par pattern d'URL
# --------------------------------------------------------------------------
def _detect_skip_levels(urls, min_share=0.9, max_levels=4):
    """Nb de segments de path initiaux constants (préfixes de langue type /fr/fr/)
    à ignorer avant le segment discriminant."""
    splits = []
    for u in urls:
        path = urlsplit(u).path.strip("/")
        splits.append(path.split("/") if path else [])
    skip = 0
    for level in range(max_levels):
        vals = [s[level] for s in splits if len(s) > level]
        if not vals:
            break
        vc = pd.Series(vals).value_counts(normalize=True)
        if vc.iloc[0] >= min_share and len(vc) <= 2:
            skip += 1
        else:
            break
    return skip


def _make_segmenter(urls):
    """Retourne une fonction url -> pattern de cluster, calée sur le jeu d'URLs."""
    skip = _detect_skip_levels(urls)

    def seg(url):
        path = urlsplit(url).path.strip("/")
        s = path.split("/") if path else []
        if not s:
            return "(home)"
        if len(s) <= skip:
            return "/" + s[-1] + "/"
        return "/" + s[skip] + "/"
    return seg


def propose_clusters(pages: pd.DataFrame) -> pd.DataFrame:
    """Segmente par 1er segment de path discriminant (saute les préfixes de langue).
    Retourne un mapping éditable : colonnes [pattern, cluster, nb_pages, exemples]."""
    seg = _make_segmenter(pages["url"].tolist())
    tmp = pages.copy()
    tmp["pattern"] = tmp["url"].map(seg)
    agg = (tmp.groupby("pattern")
              .agg(nb_pages=("url", "size"),
                   exemples=("url", lambda s: " | ".join(s.head(2))))
              .reset_index()
              .sort_values("nb_pages", ascending=False))
    agg["cluster"] = agg["pattern"].str.strip("/").replace("", "home")
    return agg[["pattern", "cluster", "nb_pages", "exemples"]]


def apply_clusters(pages: pd.DataFrame, mapping: pd.DataFrame) -> pd.DataFrame:
    """Applique le mapping pattern->cluster validé aux pages."""
    seg = _make_segmenter(pages["url"].tolist())
    pat2cluster = dict(zip(mapping["pattern"], mapping["cluster"]))
    pages = pages.copy()
    pages["_pattern"] = pages["url"].map(seg)
    pages["cluster"] = pages["_pattern"].map(pat2cluster).fillna("autre")
    return pages.drop(columns=["_pattern"])


# --------------------------------------------------------------------------
# Filtrage et qualification des liens
# --------------------------------------------------------------------------
def prepare_links(links: pd.DataFrame, pages: pd.DataFrame, cfg, report: dict):
    links = links.copy()
    links["source"] = links["source"].map(lambda u: normalize_url(u, cfg))
    links["target"] = links["target"].map(lambda u: normalize_url(u, cfg))
    n0 = len(links)

    links = links.dropna(subset=["source", "target"])
    links = links[links["source"] != links["target"]]
    report["liens_bruts"] = n0
    report["liens_self_ou_vides"] = n0 - len(links)

    # follow / nofollow
    if "follow" in links.columns:
        follow_norm = links["follow"].astype(str).str.lower()
        is_nofollow = follow_norm.str.contains("nofollow") | (follow_norm == "false") | (follow_norm == "no")
        report["liens_nofollow_exclus"] = int(is_nofollow.sum())
        links = links[~is_nofollow]

    # position normalisée
    links["position"] = links.get("position", "content").map(normalize_position)

    # cibles valides : présentes dans pages, indexables, statut 200
    page_status = pages.set_index("url")["status_code"].to_dict()
    if "indexable" in pages.columns:
        page_index = pages.set_index("url")["indexable"].astype(str).str.lower().to_dict()
    else:
        page_index = {}  # colonne absente de l'export Pages : indexabilité jugée côté links
    valid_set = set(pages["url"].tolist())

    before = len(links)
    links = links[links["target"].isin(valid_set)]
    report["liens_cible_inconnue"] = before - len(links)

    def target_ok(u):
        st = page_status.get(u)
        ix = page_index.get(u, "")
        st_ok = st is None or (isinstance(st, (int, float)) and 200 <= st < 300) or str(st).startswith("2")
        ix_ok = "non" not in ix and "false" not in ix  # exclut Non-Indexable (si connu)
        return st_ok and ix_ok
    before = len(links)
    mask = links["target"].map(target_ok)
    # filtrage complémentaire par l'indexabilité portée par le fichier links (OnCrawl)
    if "target_indexable" in links.columns:
        ti = links["target_indexable"].astype(str).str.lower()
        link_idx_ok = ~ti.isin(["false", "no", "0", "noindex", "non-indexable", "nan"])
        mask = mask & link_idx_ok
    report["liens_cible_non_indexable_ou_3xx4xx5xx"] = int((~mask).sum())

    # Capture des liens de NAVIGATION (menu/pied/nav) pointant vers une redirection (3xx),
    # avant de les écarter : ce sont les "liens menu à changer" (recibler vers l'URL finale).
    dropped = links[~mask].copy()
    report["_nav_a_corriger"] = None
    report["liens_nav_vers_redirection"] = 0
    if len(dropped):
        dropped["_tstat"] = dropped["target"].map(page_status)
        def _is3xx(s):
            return str(s).startswith("3") if pd.notna(s) else False
        redir_nav = dropped[dropped["_tstat"].map(_is3xx)
                            & dropped["position"].isin(["Header", "Footer", "Navigation"])]
        report["liens_nav_vers_redirection"] = int(len(redir_nav))
        if len(redir_nav):
            agg = (redir_nav.groupby("target")
                   .agg(nb_liens=("source", "size"),
                        emplacement=("position", lambda s: s.value_counts().index[0]),
                        statut_cible=("_tstat", "first"))
                   .reset_index()
                   .sort_values("nb_liens", ascending=False))
            report["_nav_a_corriger"] = agg

    links = links[mask]

    report["liens_valides"] = len(links)
    return links


def reclassify_nav(links: pd.DataFrame, n_pages_total: int, cfg, report: dict):
    """Reclasse en Navigation les cibles reçues en Content depuis > seuil% des pages
    (bloc de liens répété non identifié comme navigation par le crawler)."""
    thr = cfg["nav_repeat_threshold"]
    content = links[links["position"] == "Content"]
    # nb de sources distinctes pointant chaque cible en content
    src_per_target = content.groupby("target")["source"].nunique()
    repeated = src_per_target[src_per_target > thr * n_pages_total].index.tolist()
    report["cibles_reclassees_navigation"] = len(repeated)
    links["reclassified_nav"] = False
    if repeated:
        mask = (links["position"] == "Content") & (links["target"].isin(repeated))
        links.loc[mask, "position"] = "Navigation"
        links.loc[mask, "reclassified_nav"] = True
    return links, repeated


# --------------------------------------------------------------------------
# Métriques par page
# --------------------------------------------------------------------------
def compute_page_metrics(pages, links, cfg):
    p = pages.copy()

    # InRank normalisé sur 10 (OnCrawl est déjà 0-10 mais on sécurise)
    inr = pd.to_numeric(p["inrank"], errors="coerce").fillna(0.0)
    mx = inr.max() if inr.max() > 0 else 1
    p["pr_interne"] = (inr / mx * 10).round(2)

    content = links[links["position"] == "Content"]

    # liens contextuels entrants / sortants
    in_ctx = content.groupby("target").size()
    out_ctx = content.groupby("source").size()
    in_all = links.groupby("target").size()  # tous positions confondues
    p["liens_ctx_entrants"] = p["url"].map(in_ctx).fillna(0).astype(int)
    p["liens_ctx_sortants"] = p["url"].map(out_ctx).fillna(0).astype(int)
    p["liens_entrants_total"] = p["url"].map(in_all).fillna(0).astype(int)
    p["orpheline"] = p["liens_entrants_total"] == 0

    # Ancres : diversité, dominante, génériques, pondération par PR source
    pr_by_url = dict(zip(p["url"], p["pr_interne"]))
    anchor_stats = _anchor_metrics(content, pr_by_url)
    p = p.merge(anchor_stats, on="url", how="left")
    for c in ["nb_ancres_distinctes", "ratio_diversite_ancres", "part_ancre_dominante",
              "part_ancres_generiques", "pr_ancre_pondere"]:
        p[c] = p[c].fillna(0)

    return p


def _anchor_metrics(content_links, pr_by_url):
    rows = []
    grp = content_links.groupby("target")
    for target, g in grp:
        anchors = g["anchor"].fillna("").astype(str).str.strip()
        n = len(anchors)
        distinct = anchors.nunique()
        counts = anchors.value_counts()
        dominant_share = counts.iloc[0] / n if n else 0
        generic_mask = anchors.str.lower().isin(GENERIC_ANCHORS) | (anchors == "")
        generic_share = generic_mask.sum() / n if n else 0
        # pondération : somme des PR des pages sources, normalisée par nb liens
        weights = g["source"].map(pr_by_url).fillna(0)
        pr_weighted = weights.sum()
        rows.append({
            "url": target,
            "nb_ancres_distinctes": distinct,
            "ratio_diversite_ancres": round(distinct / n, 3) if n else 0,
            "part_ancre_dominante": round(dominant_share, 3),
            "part_ancres_generiques": round(generic_share, 3),
            "pr_ancre_pondere": round(pr_weighted, 2),
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["url", "nb_ancres_distinctes", "ratio_diversite_ancres",
                 "part_ancre_dominante", "part_ancres_generiques", "pr_ancre_pondere"])


# --------------------------------------------------------------------------
# Jointures data business
# --------------------------------------------------------------------------
def join_business_data(p, gsc, ga4, sem, cfg):
    if gsc is not None and len(gsc):
        g = gsc.copy()
        g["url"] = g["url"].map(lambda u: normalize_url(u, cfg))
        for col in ["clicks", "impressions", "position"]:
            if col in g.columns:
                g[col] = pd.to_numeric(g[col], errors="coerce")
        g = g.groupby("url", as_index=False).agg(
            clics_gsc=("clicks", "sum"),
            impressions_gsc=("impressions", "sum"),
            position_gsc=("position", "mean"))
        p = p.merge(g, on="url", how="left")
    else:
        p["clics_gsc"] = 0
        p["impressions_gsc"] = 0
        p["position_gsc"] = np.nan

    if ga4 is not None and len(ga4):
        a = ga4.copy()
        a["url"] = a["url"].map(lambda u: normalize_url(u, cfg))
        for col in ["sessions", "conversions", "value"]:
            if col in a.columns:
                a[col] = pd.to_numeric(a[col], errors="coerce")
        a = a.groupby("url", as_index=False).agg(
            sessions_ga4=("sessions", "sum"),
            conversions_ga4=("conversions", "sum"),
            valeur_ga4=("value", "sum"))
        p = p.merge(a, on="url", how="left")
        p["_has_ga4"] = True
    else:
        p["sessions_ga4"] = 0
        p["conversions_ga4"] = 0
        p["valeur_ga4"] = 0
        p["_has_ga4"] = False

    if sem is not None and len(sem):
        s = sem.copy()
        s["target_url"] = s["target_url"].map(lambda u: normalize_url(u, cfg))
        s["volume"] = pd.to_numeric(s["volume"], errors="coerce").fillna(0)
        agg_kwargs = dict(volume_semantique=("volume", "sum"),
                          kw_principal=("keyword", "first"),
                          nb_kw=("keyword", "size"))
        # Positions issues de l'étude sémantique (2e source, par mot-clé, plus granulaire que la GSC).
        has_pos = "position" in s.columns
        if has_pos:
            s["position"] = pd.to_numeric(s["position"], errors="coerce")
            s.loc[(s["position"] <= 0) | (s["position"] > 100), "position"] = pd.NA
            agg_kwargs["position_sem_best"] = ("position", "min")
            agg_kwargs["position_sem_med"] = ("position", "median")
            # nb de mots-clés en opportunité (positions 4 à 20 : proches de la 1re page, à pousser)
            s["_opp"] = s["position"].between(4, 20)
            agg_kwargs["kw_opportunite"] = ("_opp", "sum")
        vol = s.groupby("target_url", as_index=False).agg(**agg_kwargs)
        vol = vol.rename(columns={"target_url": "url"})
        p = p.merge(vol, on="url", how="left")
        if has_pos:
            p["kw_opportunite"] = pd.to_numeric(p.get("kw_opportunite"), errors="coerce").fillna(0).astype(int)
        # kw par page (liste) pour la prescription d'ancres
        kw_map = s.groupby("target_url")["keyword"].apply(list).to_dict()
        p["_kw_list"] = p["url"].map(kw_map)
        # détail par mot-clé (top 3 par volume) : "mot (vol X · pos Y)" -> savoir SUR QUEL MOT agir
        def _kwdetail(g):
            g2 = g.sort_values("volume", ascending=False).head(3)
            parts = []
            for _, x in g2.iterrows():
                pv = x.get("position")
                postxt = f" · pos {int(pv)}" if (has_pos and pd.notna(pv)) else ""
                parts.append(f"{x['keyword']} (vol {int(x['volume'])}{postxt})")
            return " ; ".join(parts)
        p["_kw_detail"] = p["url"].map(s.groupby("target_url").apply(_kwdetail).to_dict()).fillna("")
        p["_has_sem"] = True
    else:
        p["volume_semantique"] = 0
        p["kw_principal"] = ""
        p["nb_kw"] = 0
        p["_kw_list"] = [[] for _ in range(len(p))]
        p["_kw_detail"] = ""
        p["_has_sem"] = False

    for c in ["clics_gsc", "impressions_gsc", "sessions_ga4", "conversions_ga4",
              "valeur_ga4", "volume_semantique", "nb_kw"]:
        p[c] = pd.to_numeric(p[c], errors="coerce").fillna(0)
    p["kw_principal"] = p["kw_principal"].fillna("")
    return p


# --------------------------------------------------------------------------
# Matrice 2 axes : potentiel business (X) x déficit maillage (Y)
# --------------------------------------------------------------------------
def compute_two_axis_matrix(p, cfg):
    # Axe X : potentiel business (uniquement les signaux réellement présents)
    def _has_signal(col):
        return pd.to_numeric(p[col], errors="coerce").fillna(0).nunique() > 1
    comps_x = []
    for col in ["volume_semantique", "clics_gsc", "impressions_gsc"]:
        if col in p.columns and _has_signal(col):
            comps_x.append(_pct_rank(p[col]))
    if p["_has_ga4"].iloc[0] and _has_signal("conversions_ga4"):
        comps_x.append(_pct_rank(p["conversions_ga4"]))
    if not comps_x:
        comps_x = [pd.Series(np.zeros(len(p)), index=p.index)]
    p["potentiel_business"] = (sum(comps_x) / len(comps_x) * 10).round(2)

    # Qualité de maillage (0-1) -> déficit = (1 - qualité) * 10
    anchor_quality = _anchor_quality_score(p, cfg)
    depth = pd.to_numeric(p["depth"], errors="coerce").fillna(p["depth"].max())
    quality = pd.concat([
        _pct_rank(p["pr_interne"]),
        _pct_rank(p["liens_ctx_entrants"]),
        anchor_quality,
        (1 - _pct_rank(depth)),   # faible profondeur = bon
    ], axis=1).mean(axis=1)
    p["_maillage_quality"] = quality.round(3)
    p["deficit_maillage"] = ((1 - quality) * 10).round(2)

    # score composite interne (tri seulement, pas le diagnostic présenté)
    p["score_composite"] = (p["potentiel_business"] * 0.5 +
                            p["deficit_maillage"] * 0.5).round(2)

    # Quadrants
    xs, ys = cfg["quadrant_x_split"], cfg["quadrant_y_split"]
    def quad(row):
        hi_pot = row["potentiel_business"] >= xs
        hi_def = row["deficit_maillage"] >= ys
        if hi_pot and hi_def:
            return "Q1 - Prioritaire (fort potentiel / fort déficit)"
        if hi_pot and not hi_def:
            return "Q2 - Hors scope maillage (fort potentiel / bon maillage)"
        if not hi_pot and not hi_def:
            return "Q3 - Sur-maillage (faible potentiel / bon maillage)"
        return "Q4 - Ne rien faire (faible potentiel / fort déficit)"
    p["quadrant"] = p.apply(quad, axis=1)
    p["quadrant_code"] = p["quadrant"].str[:2]
    return p


def _anchor_quality_score(p, cfg):
    dom = p["part_ancre_dominante"].fillna(0)
    gen = p["part_ancres_generiques"].fillna(0)
    div = p["ratio_diversite_ancres"].fillna(0)
    # pénalise dominance excessive et génériques ; récompense un minimum de diversité
    q = 1 - 0.5 * gen - 0.5 * (dom.clip(lower=cfg["homogeneous_anchor_threshold"]) - cfg["homogeneous_anchor_threshold"]) / (1 - cfg["homogeneous_anchor_threshold"])
    q = q * (0.5 + 0.5 * div.clip(0, 1))
    # orphelines : qualité nulle
    q = q.where(~p["orpheline"], 0.0)
    return q.clip(0, 1)


# --------------------------------------------------------------------------
# Silotage inter-clusters
# --------------------------------------------------------------------------
def compute_siloing(p, links, cfg):
    url2cluster = dict(zip(p["url"], p["cluster"]))
    content = links[links["position"] == "Content"].copy()
    content["c_src"] = content["source"].map(url2cluster)
    content["c_tgt"] = content["target"].map(url2cluster)
    content = content.dropna(subset=["c_src", "c_tgt"])

    matrix = pd.crosstab(content["c_src"], content["c_tgt"])
    matrix_pct = matrix.div(matrix.sum(axis=1).replace(0, np.nan), axis=0).fillna(0).round(3)

    clusters = sorted(set(p["cluster"]))
    flags = []
    for c in clusters:
        out_total = matrix.loc[c].sum() if c in matrix.index else 0
        intra = matrix.loc[c, c] if (c in matrix.index and c in matrix.columns) else 0
        in_total = matrix[c].sum() if c in matrix.columns else 0
        intra_share = intra / out_total if out_total else 0
        # aspirateur : reçoit beaucoup, renvoie peu
        ratio_in_out = (in_total / out_total) if out_total else np.inf
        flags.append({
            "cluster": c,
            "liens_sortants": int(out_total),
            "liens_entrants": int(in_total),
            "part_intra_cluster": round(intra_share, 3),
            "cluster_ferme": intra_share > cfg["closed_cluster_threshold"],
            "cluster_aspirateur": (ratio_in_out > 3) and (in_total > 10),
        })
    silo_flags = pd.DataFrame(flags)

    # ventilation Content/Nav/Header/Footer/Sidebar par cluster cible (%)
    links2 = links.copy()
    links2["c_tgt"] = links2["target"].map(url2cluster)
    vent = pd.crosstab(links2["c_tgt"], links2["position"], normalize="index").round(3) * 100
    return matrix, matrix_pct, silo_flags, vent


# --------------------------------------------------------------------------
# Flags
# --------------------------------------------------------------------------
def compute_flags(p, links, cfg):
    flags = []

    def add(url, typ, gravite, detail):
        flags.append({"url": url, "type": typ, "gravite": gravite, "detail": detail})

    for _, r in p.iterrows():
        if r["orpheline"]:
            add(r["url"], "Page orpheline", "Élevée", "0 lien interne entrant")
        if r["quadrant_code"] == "Q1" and r["part_ancres_generiques"] > cfg["poor_anchor_threshold"]:
            add(r["url"], "Ancre pauvre", "Élevée",
                f"{r['part_ancres_generiques']:.0%} d'ancres génériques/vides sur page Q1")
        if r["part_ancre_dominante"] > cfg["homogeneous_anchor_threshold"]:
            add(r["url"], "Ancre sur-homogène", "Moyenne",
                f"1 ancre = {r['part_ancre_dominante']:.0%} des occurrences")
        # page positionnée sur plusieurs mots-clés mais ancres concentrées sur un seul -> à diversifier
        if (r.get("nb_kw", 0) >= 2 and r["part_ancre_dominante"] >= cfg["homogeneous_anchor_threshold"]
                and r["liens_ctx_entrants"] >= 2):
            add(r["url"], "Ancres à diversifier (page multi-requêtes)", "Élevée",
                f"positionnée sur {int(r.get('nb_kw', 0))} mots-clés, mais 1 ancre = "
                f"{r['part_ancre_dominante']:.0%} des liens entrants")

    # cluster isolé (depuis silo)
    _, _, silo_flags, _ = compute_siloing(p, links, cfg)
    for _, r in silo_flags.iterrows():
        if r["cluster_ferme"]:
            add(f"[cluster] {r['cluster']}", "Cluster isolé", "Moyenne",
                f"{r['part_intra_cluster']:.0%} de liens intra-cluster")

    # cannibalisation : 2+ pages d'un même cluster sur le même kw principal en GSC
    if p["_has_sem"].iloc[0]:
        canib = p[(p["kw_principal"] != "") & (p["clics_gsc"] > 0)]
        dup = canib.groupby(["cluster", "kw_principal"]).filter(lambda g: len(g) > 1)
        for kw, g in dup.groupby(["cluster", "kw_principal"]):
            urls = g.sort_values("clics_gsc", ascending=False)["url"].tolist()
            for u in urls:
                add(u, "Cannibalisation", "Élevée",
                    f"{len(urls)} pages du cluster '{kw[0]}' sur '{kw[1]}'")

    df = pd.DataFrame(flags) if flags else pd.DataFrame(columns=["url", "type", "gravite", "detail"])
    return df


# --------------------------------------------------------------------------
# Prescription de liens (livrable central)
# --------------------------------------------------------------------------
def build_prescriptions(p, links, cfg, near_clusters=None):
    near_clusters = near_clusters or {}
    url2cluster = dict(zip(p["url"], p["cluster"]))
    pr_by_url = dict(zip(p["url"], p["pr_interne"]))
    median_pr = p["pr_interne"].median()

    # liens content existants (source->target) pour éviter les doublons
    content = links[links["position"] == "Content"]
    existing = set(zip(content["source"], content["target"]))

    targets = p[p["quadrant_code"] == "Q1"].copy()
    # index title/h1 des sources pour l'heuristique "contient le mot-clé"
    text_by_url = {r["url"]: f"{r.get('title','')} {r.get('h1','')}".lower()
                   for _, r in p.iterrows()}

    rows = []
    for _, t in targets.iterrows():
        tgt = t["url"]
        tcluster = t["cluster"]
        allowed_clusters = {tcluster} | set(near_clusters.get(tcluster, []))
        kw_variants = _anchor_variants(t)
        kw_tokens = _kw_tokens(t)

        # candidats sources
        cand = p[
            (p["cluster"].isin(allowed_clusters)) &
            (p["url"] != tgt) &
            (p["pr_interne"] > median_pr)
        ].copy()

        scored = []
        for _, s in cand.iterrows():
            src = s["url"]
            if (src, tgt) in existing:
                continue
            stext = text_by_url.get(src, "")
            kw_hit = any(tok in stext for tok in kw_tokens) if kw_tokens else False
            same_cluster = s["cluster"] == tcluster
            # règle 4 : contient le kw cible OU variante sémantique (proxy title/h1 + cluster)
            if not (kw_hit or same_cluster):
                continue
            score = s["pr_interne"] + (2 if kw_hit else 0) + (1 if same_cluster else 0)
            scored.append((score, src, s["pr_interne"], kw_hit))

        scored.sort(reverse=True)
        for rank, (score, src, spr, kw_hit) in enumerate(scored[:8]):  # max 8 sources / cible
            priority = 1 if spr >= p["pr_interne"].quantile(0.75) else (2 if spr >= median_pr else 3)
            rows.append({
                "page_source": src,
                "page_cible": tgt,
                "cluster_cible": tcluster,
                "ancre_recommandee": kw_variants[rank % len(kw_variants)] if kw_variants else (t["kw_principal"] or "ancre descriptive"),
                "emplacement_suggere": "Content",
                "priorite": priority,
                "pr_source": spr,
                "match_kw": "oui" if kw_hit else "cluster",
                "potentiel_cible": t["potentiel_business"],
                "deficit_cible": t["deficit_maillage"],
            })

    df = pd.DataFrame(rows)
    if len(df):
        df = df.sort_values(["priorite", "deficit_cible", "pr_source"],
                            ascending=[True, False, False]).reset_index(drop=True)
    else:
        df = pd.DataFrame(columns=["page_source", "page_cible", "cluster_cible",
                                   "ancre_recommandee", "emplacement_suggere", "priorite"])
    return df


def _as_list(v):
    if isinstance(v, list):
        return v
    return []


def _kw_tokens(page_row):
    toks = set()
    kw_list = _as_list(page_row.get("_kw_list"))
    for kw in kw_list:
        toks.update(str(kw).lower().split())
    if page_row.get("kw_principal"):
        toks.update(str(page_row["kw_principal"]).lower().split())
    # tokens du slug de l'URL en secours
    slug = urlsplit(page_row["url"]).path.strip("/").split("/")[-1]
    toks.update(re.split(r"[-_]", slug.lower()))
    return {t for t in toks if len(t) > 3}


def _anchor_variants(page_row):
    """3 à 5 variantes d'ancres : mot-clé exact, longue traîne, formulation naturelle."""
    variants = []
    kw = str(page_row.get("kw_principal", "")).strip()
    kw_list = _as_list(page_row.get("_kw_list"))
    if kw:
        variants.append(kw)                              # exact
        variants.append(f"{kw} : notre guide")           # naturelle
    for k in kw_list:
        k = str(k).strip()
        if k and k not in variants:
            variants.append(k)                            # variantes longue traîne
        if len(variants) >= 5:
            break
    if not variants:
        # secours : dérivé du H1 / slug
        h1 = str(page_row.get("h1", "")).strip()
        if h1:
            variants.append(h1)
        slug = urlsplit(page_row["url"]).path.strip("/").split("/")[-1].replace("-", " ")
        variants.append(slug)
    # nettoyage
    variants = [v for v in dict.fromkeys(variants) if v][:5]
    return variants or ["ancre descriptive"]


def infer_near_clusters(matrix_pct, threshold=0.10):
    """Clusters 'proches' = ceux qui s'inter-linkent déjà notablement."""
    near = {}
    for c in matrix_pct.index:
        row = matrix_pct.loc[c]
        near[c] = [other for other in row.index
                   if other != c and row[other] >= threshold]
    return near


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------
def run_pipeline(pages, links, gsc=None, ga4=None, sem=None,
                 cluster_mapping=None, cfg=None):
    cfg = {**DEFAULT_CONFIG, **(cfg or {})}
    report = {}

    # Phase 1 - préparation
    pages = pages.copy()
    pages["url"] = pages["url"].map(lambda u: normalize_url(u, cfg))
    pages = pages.dropna(subset=["url"])
    pages["status_code"] = pd.to_numeric(pages.get("status_code", 200), errors="coerce").fillna(200)
    # Dédup par URL normalisée en gardant le MEILLEUR statut (200 prioritaire sur 301/302/4xx).
    # Sinon /machines/ (301) et /machines (200), qui fusionnent après normalisation, peuvent
    # laisser la variante 301 gagner -> les liens vers ces pages sont exclus à tort et les
    # pages orphelines explosent artificiellement (bug identifié sur crawl OnCrawl).
    _n_before = len(pages)
    pages = pages.sort_values("status_code", kind="stable").drop_duplicates(subset=["url"], keep="first")
    report["pages_doublons_url_fusionnees"] = _n_before - len(pages)
    report["pages_total"] = len(pages)

    if cluster_mapping is None:
        cluster_mapping = propose_clusters(pages)
    pages = apply_clusters(pages, cluster_mapping)

    links = prepare_links(links, pages, cfg, report)
    links, _ = reclassify_nav(links, len(pages), cfg, report)

    # Phase 2 - metriques par page
    p = compute_page_metrics(pages, links, cfg)
    p = join_business_data(p, gsc, ga4, sem, cfg)

    # Phase 4 - matrice 2 axes
    p = compute_two_axis_matrix(p, cfg)

    # Phase 3 - silotage
    matrix, matrix_pct, silo_flags, ventilation = compute_siloing(p, links, cfg)

    # Phase 5 - flags
    flags = compute_flags(p, links, cfg)

    # Phase 6 - prescriptions
    near = infer_near_clusters(matrix_pct)
    prescriptions = build_prescriptions(p, links, cfg, near_clusters=near)

    report["quadrants"] = p["quadrant_code"].value_counts().to_dict()
    report["pages_orphelines"] = int(p["orpheline"].sum())
    report["prescriptions_generees"] = len(prescriptions)

    return {
        "pages": p,
        "links": links,
        "silo_matrix": matrix,
        "silo_matrix_pct": matrix_pct,
        "silo_flags": silo_flags,
        "ventilation": ventilation,
        "flags": flags,
        "prescriptions": prescriptions,
        "cluster_mapping": cluster_mapping,
        "report": report,
        "config": cfg,
    }
