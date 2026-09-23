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
import unicodedata
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
    "max_liens_par_source": 5,         # nb max de NOUVEAUX liens prescrits sur une même page source
}

GENERIC_ANCHORS = {
    "", "cliquez ici", "clique ici", "ici", "en savoir plus", "savoir plus",
    "lire la suite", "lire l'article", "voir plus", "plus", "en lire plus",
    "click here", "read more", "learn more", "voir", "découvrir", "detail",
    "détail", "details", "lien", "link", "accueil", "home",
}

POSITION_MAP = {
    # anglais + OnCrawl
    "content": "Content", "contenu": "Content", "body": "Content", "main": "Content",
    "navigation": "Navigation", "nav": "Navigation", "menu": "Navigation",
    "header": "Header", "entete": "Header", "en tete": "Header", "top": "Header", "tete": "Header",
    "footer": "Footer", "pied": "Footer", "bottom": "Footer", "pied de page": "Footer",
    "sidebar": "Sidebar", "aside": "Sidebar", "lateral": "Sidebar", "auxiliaire": "Sidebar",
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


def _strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def normalize_position(p):
    if not isinstance(p, str):
        return "Content"
    key = _strip_accents(p.strip().lower()).replace("-", " ").replace("_", " ").strip()
    key = re.sub(r"\s+", " ", key)
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


# --------------------------------------------------------------------------
# Typologie de page (Produit / Listing / FAQ / Éditorial / Home / Légal / Compte)
# et rôle dans le maillage (Cible / Source / Cible + Source / Exclue).
# --------------------------------------------------------------------------
TYPOLOGIES = ["Produit", "Listing/Catégorie", "Éditorial", "Home", "FAQ", "Légal", "Compte/Tunnel", "Autre"]
ROLES = ["Cible + Source", "Cible", "Source", "Exclue"]
# Rôle par défaut : on cible Produit/Listing, on exclut FAQ/Légal/Compte, l'éditorial et la home émettent.
ROLE_DEFAUT = {
    "Produit": "Cible + Source", "Listing/Catégorie": "Cible + Source", "Autre": "Cible + Source",
    "Éditorial": "Source", "Home": "Source",
    "FAQ": "Exclue", "Légal": "Exclue", "Compte/Tunnel": "Exclue",
}


def _guess_typologie(cluster, examples):
    s = (str(cluster) + " " + str(examples)).lower()
    def has(words):
        return any(w in s for w in words)
    if str(cluster).strip().lower() in ("(home)", "home", ""):
        return "Home"
    if has(["faq", "/aide", "support", "question", "help"]):
        return "FAQ"
    if has(["mention", "cgv", "cgu", "politique", "confidential", "cookie", "/legal", "accessibilit",
            "retractation", "sanitaire", "plan-du-site", "sitemap", "rgpd"]):
        return "Légal"
    if has(["compte", "login", "connexion", "panier", "cart", "checkout", "commande", "reorder",
            "wishlist", "favoris", "mon-compte", "inscription"]):
        return "Compte/Tunnel"
    if has(["blog", "actualite", "conseil", "guide", "univers", "coffee-shop", "article", "recette",
            "magazine", "inspiration", "/news", "dossier", "lexique"]):
        return "Éditorial"
    # Produit vs Listing selon la profondeur des exemples d'URL
    seg = 0
    for u in str(examples).split(" | "):
        path = u.split("://")[-1]
        path = path.split("/", 1)[1] if "/" in path else ""
        seg = max(seg, len([x for x in path.strip("/").split("/") if x]))
    if seg >= 3:
        return "Produit"
    if seg >= 1:
        return "Listing/Catégorie"
    return "Autre"


def propose_clusters(pages: pd.DataFrame) -> pd.DataFrame:
    """Segmente par 1er segment de path discriminant (saute les préfixes de langue).
    Retourne un mapping éditable : colonnes [pattern, cluster, typologie, role, nb_pages, exemples]."""
    seg = _make_segmenter(pages["url"].tolist())
    tmp = pages.copy()
    tmp["pattern"] = tmp["url"].map(seg)
    agg = (tmp.groupby("pattern")
              .agg(nb_pages=("url", "size"),
                   exemples=("url", lambda s: " | ".join(s.head(2))))
              .reset_index()
              .sort_values("nb_pages", ascending=False))
    agg["cluster"] = agg["pattern"].str.strip("/").replace("", "home")
    agg["typologie"] = agg.apply(lambda r: _guess_typologie(r["cluster"], r["exemples"]), axis=1)
    agg["role"] = agg["typologie"].map(ROLE_DEFAUT).fillna("Cible + Source")
    return agg[["pattern", "cluster", "typologie", "role", "nb_pages", "exemples"]]


def apply_clusters(pages: pd.DataFrame, mapping: pd.DataFrame) -> pd.DataFrame:
    """Applique le mapping pattern->cluster/typologie/role validé aux pages."""
    seg = _make_segmenter(pages["url"].tolist())
    pat2cluster = dict(zip(mapping["pattern"], mapping["cluster"]))
    pages = pages.copy()
    pages["_pattern"] = pages["url"].map(seg)
    pages["cluster"] = pages["_pattern"].map(pat2cluster).fillna("autre")
    if "typologie" in mapping.columns:
        pat2typo = dict(zip(mapping["pattern"], mapping["typologie"]))
        pages["typologie"] = pages["_pattern"].map(pat2typo).fillna("Autre")
    else:
        pages["typologie"] = "Autre"
    if "role" in mapping.columns:
        pat2role = dict(zip(mapping["pattern"], mapping["role"]))
        pages["role"] = pages["_pattern"].map(pat2role).fillna("Cible + Source")
    else:
        pages["role"] = pages["typologie"].map(ROLE_DEFAUT).fillna("Cible + Source")
    return pages.drop(columns=["_pattern"])


# --------------------------------------------------------------------------
# Filtrage et qualification des liens
# --------------------------------------------------------------------------
NAV_POSITIONS = {"Header", "Footer", "Navigation", "Sidebar"}


def _build_raw_lookup(pages_raw, cfg):
    """Toutes les URL du crawl AVANT filtrage (redirections, 404, noindex, canonisées compris),
    avec statut, URL de redirection et URL canonique si l'export les fournit (Screaming Frog oui)."""
    r = pages_raw.copy()
    r["url"] = r["url"].map(lambda u: normalize_url(u, cfg))
    r = r.dropna(subset=["url"])
    r["status_code"] = (pd.to_numeric(r["status_code"], errors="coerce")
                        if "status_code" in r.columns else np.nan)
    for c in ["redirect_url", "canonical_url"]:
        if c in r.columns:
            r[c] = r[c].map(lambda u: normalize_url(u, cfg) if isinstance(u, str) and u.strip() else None)
        else:
            r[c] = None
    r["indexable"] = r["indexable"].astype(str) if "indexable" in r.columns else ""
    r = r.sort_values("status_code", kind="stable").drop_duplicates(subset=["url"], keep="first")
    return r.set_index("url")[["status_code", "redirect_url", "canonical_url", "indexable"]].to_dict("index")


def _bad_target_filter(valid_set, target_ok, raw_lookup):
    known = set(raw_lookup.keys())

    def _is_bad(u):
        if u in valid_set:
            return not target_ok(u)
        return u in known
    return _is_bad


def _classify_target(tgt, raw_lookup, page_status):
    """Diagnostic d'une URL cible à problème -> (problème, action, remplacer_par, statut) ou None.
    None = on ignore (URL non crawlée, bloquée par robots.txt : choix volontaire du site)."""
    def _final_redirect(u):
        seen, cur = set(), u
        for _ in range(6):  # suit les chaînes de redirection
            info = raw_lookup.get(cur) or {}
            nxt = info.get("redirect_url")
            st = info.get("status_code")
            if not nxt or nxt in seen or not (pd.notna(st) and 300 <= st < 400):
                break
            seen.add(cur)
            cur = nxt
        return cur if cur != u else None

    info = raw_lookup.get(tgt) or {}
    st = info.get("status_code")
    if st is None or (isinstance(st, float) and np.isnan(st)):
        st = page_status.get(tgt)
    st_num = pd.to_numeric(st, errors="coerce")
    if pd.isna(st_num) or st_num == 0:
        return None
    canon = info.get("canonical_url")
    if st_num in (302, 303, 307):
        # Redirection TEMPORAIRE : souvent voulue (connexion requise, géo, A/B). On ne tranche pas.
        return (f"Redirection temporaire ({int(st_num)})",
                "VÉRIFIER : redirection temporaire souvent voulue (ex. connexion requise), sinon pointer vers l'URL finale",
                _final_redirect(tgt) or "", int(st_num))
    if 300 <= st_num < 400:
        repl = _final_redirect(tgt) or ""
        # si l'URL finale est elle-même canonisée ailleurs, on propose directement la canonique
        _c2 = (raw_lookup.get(repl) or {}).get("canonical_url") if repl else None
        if _c2 and _c2 != repl:
            repl = _c2
        return (f"Redirection ({int(st_num)})", "MODIFIER : pointer directement vers l'URL finale", repl, int(st_num))
    if st_num >= 400:
        return (f"Lien cassé ({int(st_num)})", "SUPPRIMER le lien (ou le remplacer par une page équivalente)",
                "", int(st_num))
    if canon and canon != tgt:
        return ("Page canonisée vers une autre URL", "MODIFIER : pointer vers l'URL canonique", canon, int(st_num))
    return ("Page non indexable (noindex)",
            "VÉRIFIER : normal pour compte/panier, sinon retirer le lien ou rendre la page indexable",
            "", int(st_num))


def _content_links_to_fix(links, valid_set, target_ok, raw_lookup, page_status):
    """Liens DANS LE CONTENU qui pointent vers une URL à problème. Une ligne par couple
    (page à éditer, lien) : la correction se fait page par page dans le CMS, donc on donne la page source."""
    lk = links[links["position"] == "Content"]
    if not len(lk):
        return None
    lk = lk[lk["target"].map(_bad_target_filter(valid_set, target_ok, raw_lookup))]
    if not len(lk):
        return None
    diag = {t: _classify_target(t, raw_lookup, page_status) for t in lk["target"].unique()}
    rows = []
    has_anchor = "anchor" in lk.columns
    for (src, tgt), g in lk.groupby(["source", "target"]):
        d = diag.get(tgt)
        if d is None:
            continue
        pb, action, repl, st = d
        anc = ""
        if has_anchor:
            _a = g["anchor"].fillna("").astype(str).str.strip()
            anc = _a[_a != ""].iloc[0] if (_a != "").any() else "(image/bouton)"
        rang = 0 if action.startswith("SUPPRIMER") else (1 if action.startswith("MODIFIER") else 2)
        rows.append({"_rang": rang, "page_a_editer": src, "lien_actuel": tgt, "ancre": anc,
                     "probleme": pb, "action": action, "remplacer_par": repl,
                     "nb_occurrences": int(len(g)), "statut_cible": st})
    if not rows:
        return None
    return pd.DataFrame(rows).sort_values(["_rang", "lien_actuel"]).reset_index(drop=True)


def _nav_links_to_fix(links, valid_set, target_ok, raw_lookup, page_status):
    """Liens de MENU / HEADER / FOOTER / SIDEBAR qui pointent vers une URL à problème.
    Une ligne par URL cible : où est le lien, sur combien de pages, le problème, l'action,
    et l'URL à mettre à la place quand on la connaît (redirection ou canonique)."""
    lk = links[links["position"].isin(NAV_POSITIONS)]
    if not len(lk):
        return None
    lk = lk[lk["target"].map(_bad_target_filter(valid_set, target_ok, raw_lookup))]
    if not len(lk):
        return None

    rows = []
    for tgt, g in lk.groupby("target"):
        d = _classify_target(tgt, raw_lookup, page_status)
        if d is None:
            continue
        pb, action, repl, st = d
        rows.append({"_rang": 1 if action.startswith("VÉRIFIER") else 0,
                     "emplacement": " / ".join(sorted(g["position"].unique())),
                     "lien_actuel": tgt, "probleme": pb, "action": action, "remplacer_par": repl,
                     "nb_pages": int(g["source"].nunique()), "nb_liens": int(len(g)),
                     "statut_cible": st})
    if not rows:
        return None
    # D'abord les actions certaines (MODIFIER / SUPPRIMER), puis les VÉRIFIER ; chaque groupe trié
    # par nombre de pages concernées (un lien de footer sur 1 300 pages passe avant un lien sur 5).
    df = (pd.DataFrame(rows).sort_values(["_rang", "nb_pages", "nb_liens"], ascending=[True, False, False])
          .drop(columns=["_rang"]).reset_index(drop=True))
    return df


def prepare_links(links: pd.DataFrame, pages: pd.DataFrame, cfg, report: dict, raw_lookup=None):
    links = links.copy()
    links["source"] = links["source"].map(lambda u: normalize_url(u, cfg))
    links["target"] = links["target"].map(lambda u: normalize_url(u, cfg))
    n0 = len(links)

    links = links.dropna(subset=["source", "target"])
    links = links[links["source"] != links["target"]]
    report["liens_bruts"] = n0
    report["liens_self_ou_vides"] = n0 - len(links)

    valid_set = set(pages["url"].tolist())

    # Reachability (pour l'orphelinat) : nb de liens internes vers chaque page connue EN INCLUANT
    # le nofollow. Une page reliée seulement en nofollow n'est PAS orpheline au sens navigation
    # (Screaming Frog la voit reliée). Le nofollow reste exclu de l'équité et du maillage contextuel.
    report["_in_all_reachable"] = links[links["target"].isin(valid_set)].groupby("target").size().to_dict()

    # Statut / indexabilité des pages retenues (utilisé pour juger une cible valide)
    page_status = pages.set_index("url")["status_code"].to_dict()
    if "indexable" in pages.columns:
        page_index = pages.set_index("url")["indexable"].astype(str).str.lower().to_dict()
    else:
        page_index = {}

    def target_ok(u):
        st = page_status.get(u)
        ix = page_index.get(u, "")
        st_ok = st is None or (isinstance(st, (int, float)) and 200 <= st < 300) or str(st).startswith("2")
        ix_ok = "non" not in ix and "false" not in ix
        return st_ok and ix_ok

    # Liens de menu / footer à corriger (calculé AVANT l'exclusion du nofollow : un lien de footer
    # cassé reste à corriger même s'il est en nofollow).
    links["position"] = links.get("position", "content").map(normalize_position)
    _fix = _nav_links_to_fix(links, valid_set, target_ok, raw_lookup or {}, page_status)
    report["_nav_a_corriger"] = _fix
    report["liens_nav_a_corriger"] = int(_fix["nb_liens"].sum()) if _fix is not None else 0
    report["urls_nav_a_corriger"] = int(len(_fix)) if _fix is not None else 0
    # Idem pour les liens DANS LE CONTENU (404 à supprimer, redirections à mettre à jour).
    _cfix = _content_links_to_fix(links, valid_set, target_ok, raw_lookup or {}, page_status)
    report["_contenu_a_corriger"] = _cfix
    report["liens_contenu_a_corriger"] = int(len(_cfix)) if _cfix is not None else 0
    report["liens_contenu_a_supprimer"] = (int(_cfix["action"].str.startswith("SUPPRIMER").sum())
                                           if _cfix is not None else 0)

    # follow / nofollow (exclus pour l'équité et le maillage contextuel)
    if "follow" in links.columns:
        follow_norm = links["follow"].astype(str).str.lower()
        is_nofollow = follow_norm.str.contains("nofollow") | (follow_norm == "false") | (follow_norm == "no")
        report["liens_nofollow_exclus"] = int(is_nofollow.sum())
        links = links[~is_nofollow]

    # cibles valides : présentes dans pages, indexables, statut 200
    before = len(links)
    links = links[links["target"].isin(valid_set)]
    report["liens_cible_inconnue"] = before - len(links)

    before = len(links)
    mask = links["target"].map(target_ok)
    # filtrage complémentaire par l'indexabilité portée par le fichier links (OnCrawl)
    if "target_indexable" in links.columns:
        ti = links["target_indexable"].astype(str).str.lower()
        link_idx_ok = ~ti.isin(["false", "no", "0", "noindex", "non-indexable", "nan"])
        mask = mask & link_idx_ok
    report["liens_cible_non_indexable_ou_3xx4xx5xx"] = int((~mask).sum())

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

    content = links[links["position"] == "Content"]

    # liens contextuels entrants / sortants
    in_ctx = content.groupby("target").size()
    out_ctx = content.groupby("source").size()
    in_all = links.groupby("target").size()  # tous positions confondues
    p["liens_ctx_entrants_tmp"] = p["url"].map(in_ctx).fillna(0).astype(int)

    # PageRank interne : on prend celui du crawl (PageRank OnCrawl ou Link Score Screaming Frog).
    # On ne le RECALCULE PAS. S'il est absent, on retombe sur un proxy « popularité par liens
    # entrants » (à défaut), et on le signale : l'idéal est d'exporter la colonne PageRank/Link Score.
    inr = pd.to_numeric(p["inrank"], errors="coerce") if "inrank" in p.columns else None
    if inr is not None and inr.notna().sum() > 0 and inr.dropna().nunique() > 1:
        inr = inr.fillna(0.0)
        mx = inr.max() if inr.max() > 0 else 1
        p["pr_interne"] = (inr / mx * 10).round(2)
        p["_pr_source"] = "PageRank / Link Score (crawl)"
    else:
        _pop = p["url"].map(in_all).fillna(0)
        p["pr_interne"] = (_pct_rank(_pop) * 10).round(2)
        p["_pr_source"] = "proxy liens entrants (colonne PageRank/Link Score absente du crawl)"

    p["liens_ctx_entrants"] = p.pop("liens_ctx_entrants_tmp")
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
            # Opportunité pondérée PAR MOT-CLÉ : somme(volume x facteur de position).
            # Un mot-clé déjà top 3 pèse peu, un mot-clé en 4-10 pèse au max (voir _position_factor).
            s["_opp_pond"] = s["volume"] * s["position"].map(_position_factor)
            agg_kwargs["opp_sem"] = ("_opp_pond", "sum")
        vol = s.groupby("target_url", as_index=False).agg(**agg_kwargs)
        vol = vol.rename(columns={"target_url": "url"})
        p = p.merge(vol, on="url", how="left")
        if has_pos:
            p["kw_opportunite"] = pd.to_numeric(p.get("kw_opportunite"), errors="coerce").fillna(0).astype(int)
            p["opp_sem"] = pd.to_numeric(p.get("opp_sem"), errors="coerce").fillna(0.0)
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
        p["opp_sem"] = 0.0
        p["_has_sem"] = False
    if "opp_sem" not in p.columns:
        p["opp_sem"] = 0.0

    for c in ["clics_gsc", "impressions_gsc", "sessions_ga4", "conversions_ga4",
              "valeur_ga4", "volume_semantique", "nb_kw"]:
        p[c] = pd.to_numeric(p[c], errors="coerce").fillna(0)
    p["kw_principal"] = p["kw_principal"].fillna("")
    return p


# --------------------------------------------------------------------------
# Facteur de position : combien d'upside reste-t-il selon le rang actuel ?
# Paliers fins (retour consultant) : une page déjà top 3 n'a quasi rien à gagner,
# le sweet spot est en 4-10 (bas de page 1 / haut de page 2), puis ça décroît.
# --------------------------------------------------------------------------
def _position_factor(pos):
    try:
        p = float(pos)
    except (TypeError, ValueError):
        return 0.35  # position inconnue : upside moyen par défaut, ni survalorisé ni ignoré
    if p != p:  # NaN
        return 0.35
    if p <= 1:
        return 0.05   # déjà 1er : rien à aller chercher
    if p <= 3:
        return 0.20   # top 3 : upside faible
    if p <= 10:
        return 1.00   # 4-10 : upside maximal, le maillage fait bouger
    if p <= 15:
        return 0.75   # 11-15
    if p <= 20:
        return 0.50   # 16-20
    if p <= 30:
        return 0.30   # 21-30
    return 0.12       # au-delà : le maillage seul ne suffira pas


def _best_position(p):
    """Meilleure position connue par page : min(GSC moyenne, meilleure position sémantique)."""
    cols = []
    if "position_gsc" in p.columns:
        cols.append(pd.to_numeric(p["position_gsc"], errors="coerce"))
    if "position_sem_best" in p.columns:
        cols.append(pd.to_numeric(p["position_sem_best"], errors="coerce"))
    if not cols:
        return pd.Series(np.nan, index=p.index)
    return pd.concat(cols, axis=1).min(axis=1)


# --------------------------------------------------------------------------
# Matrice 2 axes : opportunité (X, demande x gap de position) x déficit maillage (Y)
# --------------------------------------------------------------------------
def compute_two_axis_matrix(p, cfg):
    def _has_signal(col):
        return pd.to_numeric(p[col], errors="coerce").fillna(0).nunique() > 1

    p["_best_position"] = _best_position(p)

    # Axe X : OPPORTUNITÉ = demande latente ponderée par le gap de position.
    # 1) Piste sémantique (la plus juste) : somme(volume_kw x facteur_position_kw), déjà calculée.
    # 2) Piste GSC (fallback) : impressions x facteur(position moyenne GSC). Les impressions = la
    #    demande captable ; la position dit combien d'upside il reste.
    # On ne récompense plus une page qui cartonne déjà (clics élevés + top 3) : c'est le fond du
    # retour consultant (Shiba/Pro Plan sortaient à tort en priorité).
    opp_components = []
    if p.get("_has_sem", pd.Series([False])).iloc[0] and _has_signal("opp_sem"):
        opp_components.append(_pct_rank(p["opp_sem"]))
    if "impressions_gsc" in p.columns and _has_signal("impressions_gsc"):
        posfac_gsc = pd.to_numeric(p.get("position_gsc"), errors="coerce").map(_position_factor)
        opp_gsc = pd.to_numeric(p["impressions_gsc"], errors="coerce").fillna(0) * posfac_gsc.fillna(0.35)
        opp_components.append(_pct_rank(opp_gsc))
    if p["_has_ga4"].iloc[0] and _has_signal("conversions_ga4"):
        opp_components.append(_pct_rank(p["conversions_ga4"]))
    if not opp_components:
        # Aucune donnée de position/perf : on retombe sur l'ancienne demande brute.
        for col in ["volume_semantique", "clics_gsc", "impressions_gsc"]:
            if col in p.columns and _has_signal(col):
                opp_components.append(_pct_rank(p[col]))
    if not opp_components:
        opp_components = [pd.Series(np.zeros(len(p)), index=p.index)]
    p["potentiel_business"] = (sum(opp_components) / len(opp_components) * 10).round(2)

    # Qualité de maillage (0-1) -> déficit = (1 - qualité) * 10
    anchor_quality = _anchor_quality_score(p, cfg)
    depth = pd.to_numeric(p["depth"], errors="coerce")
    depth = depth.fillna(depth.max() if depth.notna().any() else 0)  # robuste si profondeur en texte / manquante
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

    # même ancre pointant vers plusieurs pages différentes (ambiguïté d'ancre)
    _c = links[links["position"] == "Content"].copy()
    _c["anchor"] = _c["anchor"].fillna("").astype(str).str.strip()
    _c = _c[(_c["anchor"] != "") & (~_c["anchor"].str.lower().isin(GENERIC_ANCHORS))]
    if len(_c):
        amb = _c.groupby("anchor").agg(n_cibles=("target", "nunique"), n_liens=("target", "size"))
        amb = amb[(amb["n_cibles"] >= 2) & (amb["n_liens"] >= 3)]
        for anc, row in amb.iterrows():
            cibles = _c[_c["anchor"] == anc]["target"].value_counts().index.tolist()
            for u in cibles[:12]:
                add(u, "Ancre ambiguë (même ancre, plusieurs pages)", "Moyenne",
                    f"l'ancre « {anc} » pointe vers {int(row['n_cibles'])} pages différentes")

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

    # Rôles (jalon C) : on ne prescrit QUE vers les pages "Cible", depuis les pages "Source".
    role = (p["role"].astype(str) if "role" in p.columns
            else pd.Series("Cible + Source", index=p.index))
    is_target_role = role.str.contains("Cible")
    is_source_role = role.str.contains("Source")
    # Garde-fou : si aucune page n'a de rôle "Cible" (ex. champ "types à cibler" laissé vide dans
    # l'app -> tout devient "Source"), on ne doit PAS renvoyer 0 prescription en silence. On retombe
    # alors sur "toute page non exclue est une cible / une source".
    if not bool(is_target_role.any()):
        is_target_role = ~role.str.contains("Exclue")
    if not bool(is_source_role.any()):
        is_source_role = ~role.str.contains("Exclue")

    targets = p[(p["quadrant_code"] == "Q1") & is_target_role.values].copy()
    # On ne prescrit PAS vers les pages déjà en top 3 (retour consultant : Shiba/Pro Plan).
    # Elles ont déjà gagné, un lien interne de plus n'a quasi pas d'upside. Position inconnue = on garde.
    if "_best_position" in targets.columns:
        _bp = pd.to_numeric(targets["_best_position"], errors="coerce")
        targets = targets[~(_bp <= 3)]
    sources_pool = p[is_source_role.values]
    # index title/h1 des sources pour l'heuristique "contient le mot-clé"
    text_by_url = {r["url"]: f"{r.get('title','')} {r.get('h1','')}".lower()
                   for _, r in p.iterrows()}

    # Priorité PILOTÉE PAR LA CIBLE (et non par la source) : c'est l'importance de la page à
    # renforcer qui prime. On croise l'opportunité (position + demande) et la typologie
    # (Produit / Listing d'abord, retour consultant).
    _typo_rank = {"Produit": 0, "Listing/Catégorie": 1, "Autre": 2,
                  "Éditorial": 3, "Home": 3, "FAQ": 4, "Légal": 4, "Compte/Tunnel": 4}
    if len(targets):
        q75 = targets["potentiel_business"].quantile(0.75)
        q50 = targets["potentiel_business"].quantile(0.50)
    else:
        q75 = q50 = 0

    def _target_priority(t):
        op = t["potentiel_business"]
        base = 1 if op >= q75 else (2 if op >= q50 else 3)
        typo = str(t.get("typologie", ""))
        if typo in ("Produit", "Listing/Catégorie") and base > 1:
            base -= 1  # les pages produit/catégorie remontent d'un cran
        return base

    rows = []
    for _, t in targets.iterrows():
        tgt = t["url"]
        tcluster = t["cluster"]
        allowed_clusters = {tcluster} | set(near_clusters.get(tcluster, []))
        kw_variants = _anchor_variants(t)
        anchors_suggested = " ; ".join(kw_variants)
        kw_tokens = _kw_tokens(t)
        tgt_prio = _target_priority(t)
        typo = str(t.get("typologie", ""))
        # Qualité de l'ancre proposée : "mot-clé" si issue de la sémantique, sinon "à affiner"
        # (dérivée du H1/slug, souvent une ancre de marque ou générique -> à retravailler).
        has_kw = bool(t.get("_has_sem")) and len(_as_list(t.get("_kw_list"))) > 0
        ancre_qualite = "mot-clé (sémantique)" if has_kw else "à affiner (dérivée du titre)"
        best_pos = t.get("_best_position")
        best_pos = round(float(best_pos), 1) if pd.notna(best_pos) else ""

        cand = sources_pool[
            (sources_pool["cluster"].isin(allowed_clusters)) &
            (sources_pool["url"] != tgt) &
            (sources_pool["pr_interne"] > median_pr)
        ].copy()

        scored = []
        for _, s in cand.iterrows():
            src = s["url"]
            if (src, tgt) in existing:
                continue
            stext = text_by_url.get(src, "")
            kw_hit = any(tok in stext for tok in kw_tokens) if kw_tokens else False
            same_cluster = s["cluster"] == tcluster
            if not (kw_hit or same_cluster):
                continue
            score = s["pr_interne"] + (2 if kw_hit else 0) + (1 if same_cluster else 0)
            scored.append((score, src, s["pr_interne"], kw_hit))

        scored.sort(reverse=True)
        for rank, (score, src, spr, kw_hit) in enumerate(scored[:8]):  # max 8 sources / cible
            rows.append({
                "page_source": src,
                "page_cible": tgt,
                "typologie_cible": typo,
                "cluster_cible": tcluster,
                "position_cible": best_pos,
                "ancre_recommandee": kw_variants[rank % len(kw_variants)] if kw_variants else (t["kw_principal"] or "ancre descriptive"),
                "ancres_suggerees": anchors_suggested,
                "ancre_qualite": ancre_qualite,
                "emplacement_suggere": "Content",
                "priorite": tgt_prio,
                "_typo_rank": _typo_rank.get(typo, 5),
                "pr_source": spr,
                "match_kw": "oui" if kw_hit else "cluster",
                "potentiel_cible": t["potentiel_business"],
                "deficit_cible": t["deficit_maillage"],
            })

    df = pd.DataFrame(rows)
    if len(df):
        df = df.sort_values(["priorite", "_typo_rank", "potentiel_cible", "deficit_cible", "pr_source"],
                            ascending=[True, True, False, False, False]).reset_index(drop=True)
        df = df.drop(columns=["_typo_rank"])
        # Plafond par page source : au-delà de N nouveaux liens, on dilue le jus de la source et on
        # recrée le problème qu'on corrige. On garde les N liens les plus prioritaires par source.
        cap = int(cfg.get("max_liens_par_source", 5) or 0)
        if cap > 0:
            df = df[df.groupby("page_source").cumcount() < cap].reset_index(drop=True)
    else:
        df = pd.DataFrame(columns=["page_source", "page_cible", "typologie_cible", "cluster_cible",
                                   "position_cible", "ancre_recommandee", "ancres_suggerees",
                                   "ancre_qualite", "emplacement_suggere", "priorite"])
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


def _clean_title(t):
    """Nettoie un title pour en faire une ancre : coupe le suffixe de marque, retire ® ™ et séparateurs."""
    s = str(t or "").strip()
    if not s or s.lower() == "nan":
        return ""
    for sep in ["|", "–", "—", " - ", "·", "»", ":"]:
        if sep in s:
            s = s.split(sep)[0]
            break
    s = re.sub(r"[®™©]", "", s)
    s = re.sub(r"\s+", " ", s).strip(" -–—:·|")
    # trop long (phrase de meta) -> on garde les premiers mots
    words = s.split()
    if len(words) > 7:
        s = " ".join(words[:7])
    return s.strip()


def _anchor_variants(page_row):
    """Sélection de 3 à 5 ancres à proposer pour DÉCRIRE la page cible : mot-clé sémantique en
    priorité, sinon dérivées du title (nettoyé), du H1, du slug et du couple typologie+slug.
    Objectif : donner un vrai choix d'ancres descriptives, pas une seule ancre pauvre."""
    variants = []
    kw = str(page_row.get("kw_principal", "")).strip()
    kw_list = _as_list(page_row.get("_kw_list"))
    if kw:
        variants.append(kw)                              # exact
    for k in kw_list:                                     # variantes longue traîne (sémantique)
        k = str(k).strip()
        if k and k.lower() not in [v.lower() for v in variants]:
            variants.append(k)
        if len(variants) >= 5:
            break

    # Dérivées du crawl (title/H1/slug) : servent de secours SANS sémantique, et complètent sinon.
    title = _clean_title(page_row.get("title"))
    h1 = _clean_title(page_row.get("h1"))
    slug = urlsplit(page_row["url"]).path.strip("/").split("/")[-1].replace("-", " ").replace("_", " ").strip()
    cluster = str(page_row.get("cluster", "")).strip().replace("-", " ")
    cat = cluster[:-1] if cluster.endswith("s") and len(cluster) > 3 else cluster  # machines -> machine
    cand = [title, h1, slug]
    if cat and slug and cat.lower() not in slug.lower():
        cand.append(f"{cat} {slug}")                     # ex. "machine piccolo"
    for c in cand:
        c = str(c).strip()
        if (c and len(c) > 2 and c.lower() not in [v.lower() for v in variants]
                and c.lower() not in GENERIC_ANCHORS):
            variants.append(c)

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


def _build_kw_table(sem, p, cfg):
    """Table par MOT-CLÉ (croisement direct des termes à pousser), issue de l'étude sémantique.
    Chaque ligne = un mot-clé sur sa page cible, avec volume, position, score d'opportunité
    (volume x facteur de position) et le contexte maillage de la page (cluster, quadrant, liens).
    Retourne None si pas de sémantique : c'est le seul export qui relie un mot-clé à une URL."""
    if sem is None or not len(sem):
        return None
    s = sem.copy()
    if "keyword" not in s.columns or "target_url" not in s.columns:
        return None
    s["url"] = s["target_url"].map(lambda u: normalize_url(u, cfg))
    s = s.dropna(subset=["url"])
    s["volume"] = pd.to_numeric(s.get("volume"), errors="coerce").fillna(0)
    has_pos = "position" in s.columns
    if has_pos:
        s["position"] = pd.to_numeric(s["position"], errors="coerce")
        s.loc[(s["position"] <= 0) | (s["position"] > 100), "position"] = pd.NA
        s["opp_kw"] = (s["volume"] * s["position"].map(_position_factor)).round(0)
        s["a_pousser"] = s["position"].between(4, 20)
    else:
        s["position"] = pd.NA
        s["opp_kw"] = (s["volume"] * 0.35).round(0)
        s["a_pousser"] = False
    attrs = [c for c in ["url", "cluster", "typologie", "quadrant_code",
                         "liens_ctx_entrants", "pr_interne", "potentiel_business", "title", "h1"] if c in p.columns]
    s = s.merge(p[attrs], on="url", how="left")

    # ARBITRAGE : un même mot-clé peut apparaître sur plusieurs URL cibles. On n'en garde qu'UNE,
    # celle qui fait le plus de sens, en croisant :
    #  1) match crawl : le terme est-il dans le title/H1/slug de la page ? (signal le plus fort)
    #  2) meilleure position, 3) plus fort potentiel de page, 4) plus fort PageRank interne.
    def _match_score(row):
        toks = [t for t in re.split(r"[\s\-_/]+", str(row.get("keyword", "")).lower()) if len(t) > 2]
        if not toks:
            return 0
        hay = f"{row.get('title','')} {row.get('h1','')} {row.get('url','')}".lower()
        return sum(1 for t in toks if t in hay)
    s["_match"] = s.apply(_match_score, axis=1)
    s["_posrank"] = pd.to_numeric(s["position"], errors="coerce").fillna(999)
    s["_pot"] = pd.to_numeric(s.get("potentiel_business"), errors="coerce").fillna(0)
    s["_pr"] = pd.to_numeric(s.get("pr_interne"), errors="coerce").fillna(0)
    # Clé insensible casse/accents : "café" et "cafe" = même intention -> une seule ligne arbitrée.
    s["_kwkey"] = s["keyword"].astype(str).map(lambda k: _strip_accents(k.lower().strip()))
    s = s.sort_values(["_match", "_posrank", "_pot", "_pr"], ascending=[False, True, False, False])
    # nb d'URL candidates par mot-clé (pour signaler l'arbitrage) puis on garde la meilleure
    s["nb_urls_candidates"] = s.groupby("_kwkey")["url"].transform("nunique")
    s = s.drop_duplicates(subset=["_kwkey"], keep="first")

    cols = ["keyword", "volume", "position", "opp_kw", "a_pousser", "url",
            "cluster", "typologie", "quadrant_code", "liens_ctx_entrants", "pr_interne", "nb_urls_candidates"]
    cols = [c for c in cols if c in s.columns]
    out = s[cols].sort_values(["opp_kw", "volume"], ascending=False).reset_index(drop=True)
    return out


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------
def _pages_hors_crawl(p, gsc, sem, cfg):
    """Pages présentes dans la GSC ou la sémantique mais absentes du crawl (même domaine).
    Orphelines de fait : le crawler ne les a jamais atteintes, donc aucun lien interne n'y mène."""
    crawl_urls = set(p["url"])
    hosts = p["url"].map(lambda u: urlsplit(u).netloc)
    ref_host = hosts.mode().iat[0] if len(hosts) else ""
    rows = {}

    def _touch(u):
        return rows.setdefault(u, {"URL": u, "Clics GSC": 0, "Volume sém.": 0, "_src": set()})

    if gsc is not None and len(gsc) and "url" in gsc.columns:
        g = gsc.copy()
        g["_u"] = g["url"].map(lambda u: normalize_url(u, cfg))
        g["_c"] = pd.to_numeric(g["clicks"], errors="coerce").fillna(0) if "clicks" in g.columns else 0
        g = g.dropna(subset=["_u"])
        for u, sub in g.groupby("_u"):
            if u in crawl_urls or (ref_host and urlsplit(u).netloc != ref_host):
                continue
            e = _touch(u); e["Clics GSC"] = int(sub["_c"].sum()); e["_src"].add("GSC")

    if sem is not None and len(sem) and "target_url" in sem.columns:
        s = sem.copy()
        s["_u"] = s["target_url"].map(lambda u: normalize_url(u, cfg))
        s["_v"] = pd.to_numeric(s["volume"], errors="coerce").fillna(0) if "volume" in s.columns else 0
        s = s.dropna(subset=["_u"])
        for u, sub in s.groupby("_u"):
            if u in crawl_urls or (ref_host and urlsplit(u).netloc != ref_host):
                continue
            e = _touch(u); e["Volume sém."] = int(sub["_v"].sum()); e["_src"].add("Sémantique")

    out = [{"URL": e["URL"], "Clics GSC": e["Clics GSC"], "Volume sém.": e["Volume sém."],
            "Vue par": " + ".join(sorted(e["_src"]))} for e in rows.values()]
    df = pd.DataFrame(out, columns=["URL", "Clics GSC", "Volume sém.", "Vue par"])
    if len(df):
        df = df.sort_values(["Clics GSC", "Volume sém."], ascending=False).reset_index(drop=True)
    return df


def run_pipeline(pages, links, gsc=None, ga4=None, sem=None,
                 cluster_mapping=None, cfg=None):
    cfg = {**DEFAULT_CONFIG, **(cfg or {})}
    report = {}

    # Phase 0 - filtrage source de crawl (OnCrawl a une colonne 'intern' ; Screaming Frog a
    # une colonne 'link_type' = Hyperlien/Image/CSS/JS...). On ne garde que les vrais hyperliens
    # internes. Les liens externes restants sont écartés plus loin (cible absente des pages).
    links = links.copy()
    report["_liens_deposes"] = len(links)
    if "link_type" in links.columns:
        lt = links["link_type"].astype(str).map(_strip_accents).str.lower()
        keep = lt.str.contains("hyperl", na=False) | lt.isin(["", "nan", "hyperlink", "hyperlien"])
        links = links[keep]
    if "intern" in links.columns:
        it = links["intern"].astype(str).str.lower()
        links = links[~it.isin(["external", "externe", "0", "false", "no", "non"])]
    report["liens_hyperliens_internes"] = len(links)

    # Phase 1 - préparation
    pages = pages.copy()
    # Photo de TOUTES les URL crawlées avant filtrage (redirections, 404, noindex, canonisées) :
    # sert à lister les liens de menu/footer à corriger et à dire par quoi les remplacer.
    try:
        raw_lookup = _build_raw_lookup(pages, cfg)
    except Exception as _e:
        log.warning("raw_lookup: %s", _e)
        raw_lookup = {}
    # Screaming Frog "Interne : Tous" liste TOUTES les ressources (images, CSS, JS...).
    # On ne garde que les pages HTML si la colonne type de contenu est fournie.
    if "content_type" in pages.columns:
        _ct = pages["content_type"].astype(str).str.lower()
        pages = pages[_ct.str.contains("html", na=False) | _ct.isin(["", "nan", "none"])]
    # Ne garder que les pages indexables si l'info est fournie (écarte URL à paramètres, noindex,
    # redirections... surtout utile pour Screaming Frog dont l'export liste tout).
    if "indexable" in pages.columns:
        _ix = pages["indexable"].astype(str).map(_strip_accents).str.lower()
        _known = _ix.str.contains("index", na=False)
        _bad = _ix.str.contains("non", na=False) | _ix.isin(["false", "no", "0", "noindex"])
        pages = pages[(~_known) | (_known & ~_bad)]
    pages["url"] = pages["url"].map(lambda u: normalize_url(u, cfg))
    pages = pages.dropna(subset=["url"])
    pages["status_code"] = pd.to_numeric(pages.get("status_code", 200), errors="coerce").fillna(200)
    if "depth" not in pages.columns:  # profondeur optionnelle
        pages["depth"] = np.nan
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

    links = prepare_links(links, pages, cfg, report, raw_lookup=raw_lookup)
    links, _ = reclassify_nav(links, len(pages), cfg, report)

    # Phase 2 - metriques par page
    p = compute_page_metrics(pages, links, cfg)
    # Orphelinat basé sur TOUS les hyperliens internes (nofollow inclus), comme Screaming Frog :
    # une page reliée seulement en nofollow n'est pas orpheline.
    _reach = report.get("_in_all_reachable") or {}
    if _reach:
        p["liens_entrants_total"] = p["url"].map(_reach).fillna(0).astype(int)
        p["orpheline"] = p["liens_entrants_total"] == 0
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

    # Orphelines "hors crawl" : pages vues par la GSC ou la sémantique mais absentes du crawl.
    # Le crawler ne les a jamais atteintes -> aucun lien interne n'y mène -> orphelines de fait.
    hors = _pages_hors_crawl(p, gsc, sem, cfg)
    report["_pages_hors_crawl"] = hors if len(hors) else None
    report["pages_hors_crawl"] = int(len(hors))

    # Table par mot-clé (onglet dédié) : croisement des termes à pousser, si sémantique fournie.
    kw_table = _build_kw_table(sem, p, cfg)

    return {
        "pages": p,
        "links": links,
        "kw_table": kw_table,
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
