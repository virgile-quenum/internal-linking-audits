"""
Audit de maillage interne — application Streamlit pour consultants SEO.
Lancer :  streamlit run app.py

Flux : source des exports (upload OU chemin local) → mapping colonnes →
validation → segmentation clusters (éditable) → analyse → livrable Excel.

Gros volumes (millions de liens) : utiliser le champ "chemin local du fichier".
La lecture ne charge que les colonnes utiles (usecols) pour limiter la RAM.
"""
import io
import os
import unicodedata
import datetime as dt

import pandas as pd
import streamlit as st

import pipeline as pl
from excel_report import build_excel
from deck_report import build_deck

st.set_page_config(page_title="Audit maillage interne", layout="wide", page_icon="🔗")

# --------------------------------------------------------------------------
# Aliases pour le mapping automatique des colonnes
# --------------------------------------------------------------------------
def _norm(s):
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s.replace("_", " ").replace("-", " ").strip()

ALIASES = {
    "pages": {
        # OnCrawl + Screaming Frog (FR/EN). Les intitulés sont comparés après normalisation (accents/casse retirés).
        "url": ["url", "address", "adresse", "page", "full url"],
        "inrank": ["inrank", "in rank", "internal pagerank", "pagerank interne", "pagerank", "link score", "popularity", "prank"],
        "depth": ["depth", "crawl depth", "crawl profondeur", "profondeur", "niveau", "depth level", "profondeur du dossier"],
        "status_code": ["status code", "status", "http code", "code http", "code", "http status", "response code", "code de statut"],
        "indexable": ["indexable", "indexability", "indexabilite", "index status", "is indexable", "statut d indexabilite"],
        "word_count": ["word count", "words", "nb words", "wordcount", "mots", "nombre de mots"],
        "title": ["title", "title 1", "titre", "page title", "meta title"],
        "h1": ["h1", "h1 1", "h1 tag", "premier h1"],
        "content_type": ["content type", "type de contenu", "contenttype", "mime type"],
    },
    "links": {
        # OnCrawl : origin/target/anchor/follow/origin_position/intern.
        # Screaming Frog (export All Inlinks / Tous les liens entrants) : Source/Destination/Ancrage/Suivre/Position du lien/Type.
        "source": ["origin", "source", "from", "source url", "url source", "origin url", "src"],
        "target": ["target", "destination", "to", "target url", "url cible", "dest", "cible"],
        "anchor": ["anchor", "anchor text", "ancre", "ancrage", "texte ancre", "texte de lien", "link text", "texte alt"],
        "follow": ["follow", "nofollow", "suivre", "rel", "follow type", "link follow", "dofollow"],
        "position": ["origin position", "position du lien", "link position", "position", "emplacement", "link location", "location"],
        "link_type": ["type", "type de lien", "lien type", "link type"],
        "intern": ["intern", "internal", "interne"],
        "target_indexable": ["target meta robots index", "target indexable", "target index", "target is index"],
    },
    "gsc": {
        "url": ["url", "page", "adresse", "landing page", "pages", "top pages"],
        "clicks": ["clicks", "clics", "click"],
        "impressions": ["impressions", "impression"],
        "ctr": ["ctr", "taux de clic", "click through rate"],
        "position": ["position", "avg position", "average position", "position moyenne"],
    },
    "ga4": {
        "url": ["url", "page", "page path", "landing page", "chemin de page", "page path and screen class"],
        "sessions": ["sessions", "session"],
        "conversions": ["conversions", "key events", "conversion", "evenements cles"],
        "value": ["value", "valeur", "revenue", "revenu", "total revenue", "event value"],
    },
    "sem": {
        "keyword": ["keyword", "mot cle", "query", "requete", "mot-cle", "terme"],
        "volume": ["volume", "search volume", "volume de recherche", "sv", "recherches"],
        "target_url": ["url cible", "target url", "url", "landing", "page cible"],
        "cluster": ["cluster", "groupe", "thematique", "categorie", "segment"],
    },
}
REQUIRED = {
    "pages": ["url", "inrank", "depth", "status_code"],
    "links": ["source", "target", "position"],
    "gsc": ["url", "clicks", "impressions"],
    "ga4": ["url", "sessions", "conversions"],
    "sem": ["keyword", "volume", "target_url"],
}
LABELS = {"pages": "Crawl Pages", "links": "Crawl Links", "gsc": "GSC", "ga4": "GA4", "sem": "Sémantique"}


def guess_mapping(columns, kind):
    normcols = {c: _norm(c) for c in columns}
    mapping = {}
    for canon, aliases in ALIASES[kind].items():
        found = None
        for c, nc in normcols.items():
            if nc in aliases:
                found = c
                break
        if not found:
            for c, nc in normcols.items():
                if any(a in nc for a in aliases):
                    found = c
                    break
        mapping[canon] = found
    return mapping


# --------------------------------------------------------------------------
# Lecture : supporte fichier uploadé OU chemin local ; lecture par colonnes
# --------------------------------------------------------------------------
def _sep_from_line(line):
    return ";" if line.count(";") > line.count(",") else ","

def _is_excel(name):
    return name.lower().endswith((".xlsx", ".xls"))

def load_df(source, usecols=None, nrows=None):
    """source = ('path', chemin) ou ('upload', fichier). usecols=None -> tout.
    Lecture robuste : on lit toutes les colonnes en texte puis on sous-sélectionne.
    Cela évite un bug pandas (usecols + lignes malformées -> IndexError) sur les
    gros exports dont les ancres contiennent des guillemets/retours ligne cassés."""
    kind, obj = source

    def _finish(df):
        if usecols:
            keep = [c for c in usecols if c in df.columns]
            return df[keep]
        return df

    if kind == "path":
        p = obj
        if _is_excel(p):
            return _finish(pd.read_excel(p, nrows=nrows))
        with open(p, "rb") as fh:
            first = fh.readline().decode("utf-8-sig", errors="replace")
        sep = _sep_from_line(first)
        df = pd.read_csv(p, sep=sep, nrows=nrows, encoding="utf-8-sig",
                         dtype=str, engine="c", on_bad_lines="skip")
        return _finish(df)
    else:
        f = obj
        raw = f.getvalue()
        if _is_excel(f.name):
            return _finish(pd.read_excel(io.BytesIO(raw), nrows=nrows))
        text = None
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        sep = _sep_from_line(text.split("\n", 1)[0])
        df = pd.read_csv(io.StringIO(text), sep=sep, nrows=nrows,
                         dtype=str, engine="c", on_bad_lines="skip")
        return _finish(df)


# --------------------------------------------------------------------------
# État
# --------------------------------------------------------------------------
ss = st.session_state
ss.setdefault("mapped", {})
ss.setdefault("cluster_mapping", None)
ss.setdefault("results", None)


def _check_password():
    """Gate optionnel : actif seulement si 'app_password' est défini dans les secrets
    (déploiement hébergé). En local sans secret, accès libre."""
    try:
        pwd = st.secrets.get("app_password", None)
    except Exception:
        pwd = None
    if not pwd:
        pwd = os.environ.get("APP_PASSWORD")  # Railway / Docker : variable d'env
    if not pwd:
        return True
    if ss.get("auth_ok"):
        return True
    st.title("🔗 Audit de maillage interne")
    with st.form("login"):
        entered = st.text_input("Mot de passe", type="password")
        ok = st.form_submit_button("Entrer")
    if ok and entered == pwd:
        ss["auth_ok"] = True
        st.rerun()
    elif ok:
        st.error("Mot de passe incorrect.")
    return False


if not _check_password():
    st.stop()

st.title("🔗 Audit de maillage interne")
st.caption("Crawl (OnCrawl ou Screaming Frog) → diagnostic 2 axes → prescription de liens → livrables Excel + PPT. "
           "PageRank interne = InRank (OnCrawl) ou Link Score (Screaming Frog).")

# --------------------------------------------------------------------------
# Sidebar : seuils
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Seuils")
    cfg = dict(pl.DEFAULT_CONFIG)
    cfg["nav_repeat_threshold"] = st.slider("Reclassement navigation (part des pages)", 0.5, 0.99, 0.80, 0.01)
    cfg["closed_cluster_threshold"] = st.slider("Cluster fermé (part intra)", 0.5, 0.99, 0.85, 0.01)
    cfg["homogeneous_anchor_threshold"] = st.slider("Ancre sur-homogène", 0.5, 0.95, 0.70, 0.05)
    cfg["poor_anchor_threshold"] = st.slider("Ancre pauvre (génériques/vides)", 0.3, 0.9, 0.50, 0.05)
    cfg["quadrant_x_split"] = st.slider("Seuil axe X (potentiel)", 3.0, 7.0, 5.0, 0.5)
    cfg["quadrant_y_split"] = st.slider("Seuil axe Y (déficit)", 3.0, 7.0, 5.0, 0.5)
    cfg["strip_query_params"] = st.checkbox("Retirer les paramètres d'URL", value=True)
    st.divider()
    site_name = st.text_input("Nom du client (titre du livrable)", "Client")
    if st.button("🔄 Réinitialiser"):
        for k in ["mapped", "cluster_mapping", "results"]:
            ss.pop(k, None)
        st.rerun()

# --------------------------------------------------------------------------
# ÉTAPE 1 — Source des fichiers (upload OU chemin local)
# --------------------------------------------------------------------------
st.subheader("1. Indiquer les exports")
st.caption("Upload pour les fichiers légers. Pour un export volumineux (> 500 Mo, millions de liens), "
           "colle plutôt le chemin local du fichier : lecture directe depuis le disque, sans upload navigateur.")

with st.expander("📄 Quels fichiers fournir ? (OnCrawl ou Screaming Frog, FR/EN)"):
    st.markdown(
        "**2 fichiers de crawl obligatoires** (Pages + Liens), au choix OnCrawl **ou** Screaming Frog. "
        "GSC recommandée, sémantique et GA4 optionnelles. Les colonnes sont détectées automatiquement "
        "(français ou anglais) et corrigeables à l'écran.\n\n"
        "**Crawl (obligatoire)**\n\n"
        "| Source | Fichier à exporter | Ce que l'outil y lit |\n"
        "|---|---|---|\n"
        "| OnCrawl | Export **Pages** | url, InRank, profondeur, code HTTP |\n"
        "| OnCrawl | Export **Links** (tous les liens, sans plafond) | origin, target, ancre, follow, emplacement, intern |\n"
        "| Screaming Frog | Onglet **Interne → Tous** (Internal: All) | Adresse, Link Score, Crawl profondeur, Code HTTP, Indexabilité |\n"
        "| Screaming Frog | Export en masse → **Tous les liens entrants** (All Inlinks) | Source, Destination, Ancrage, Suivre, Position du lien, Type |\n\n"
        "_Exporter l'intégralité des liens, jamais un échantillon. Sur Screaming Frog, l'outil ne garde que "
        "les hyperliens vers des pages HTML indexables ; images, CSS, redirections et noindex sont écartés._\n\n"
        "**Enrichissement**\n\n"
        "| Fichier | Statut | Source / export | Colonnes utiles |\n"
        "|---|---|---|---|\n"
        "| Search Console | Recommandé | Performances → Pages, 3 mois | URL, Clics, Impressions, Position |\n"
        "| Étude sémantique | Optionnel | Export mots-clés | Mot-clé, Volume, URL cible, Position |\n"
        "| GA4 | Optionnel | Rapport Pages, 3 mois | URL, Sessions, Conversions |\n\n"
        "**Bonnes pratiques** : sur Screaming Frog, activer le rendu JavaScript si le site en dépend ; "
        "exporter sans filtre appliqué sur le tableau."
    )

def source_input(kind, label):
    c1, c2 = st.columns([1, 1])
    up = c1.file_uploader(label, type=["csv", "xlsx"], key=f"u_{kind}")
    path = c2.text_input(f"…ou chemin local ({LABELS[kind]})", key=f"p_{kind}",
                         placeholder=r"C:\exports\oncrawl_links.csv")
    if path and path.strip():
        p = path.strip().strip('"')
        if not os.path.exists(p):
            st.warning(f"Chemin introuvable : {p}")
            return None
        return ("path", p)
    if up is not None:
        return ("upload", up)
    return None

sources = {}
sources["pages"] = source_input("pages", "Pages — OnCrawl (export Pages) ou Screaming Frog (Interne : Tous) — obligatoire")
sources["links"] = source_input("links", "Liens — OnCrawl (export Links) ou Screaming Frog (Tous les liens entrants) — obligatoire")
sources["gsc"] = source_input("gsc", "Search Console — Pages (recommandé)")
with st.expander("Sources optionnelles (GA4, étude sémantique)"):
    sources["ga4"] = source_input("ga4", "GA4 — Pages (optionnel)")
    sources["sem"] = source_input("sem", "Étude sémantique (optionnel)")

sources = {k: v for k, v in sources.items() if v is not None}

if "pages" not in sources or "links" not in sources:
    st.info("Fournis au minimum les deux exports de crawl (Pages + Liens), OnCrawl ou Screaming Frog, pour démarrer.")
    st.stop()

# --------------------------------------------------------------------------
# Aperçus (échantillon) pour le mapping
# --------------------------------------------------------------------------
previews = {}
for kind, src in sources.items():
    try:
        previews[kind] = load_df(src, nrows=2000)
    except Exception as e:
        st.error(f"Lecture impossible ({LABELS[kind]}) : {e}")
        st.stop()

# --------------------------------------------------------------------------
# ÉTAPE 2 — Mapping des colonnes
# --------------------------------------------------------------------------
st.subheader("2. Vérifier le mapping des colonnes")
st.caption("Auto-détecté sur un échantillon. Corrige un menu si une colonne est mal reconnue.")

mapping_ok = True
for kind, df in previews.items():
    with st.expander(f"{LABELS[kind]} — {df.shape[1]} colonnes détectées", expanded=(kind in ("pages", "links"))):
        guess = guess_mapping(list(df.columns), kind)
        cols = st.columns(min(4, len(ALIASES[kind])))
        mapping = {}
        options = ["(aucune)"] + list(df.columns)
        for i, canon in enumerate(ALIASES[kind].keys()):
            with cols[i % len(cols)]:
                default = guess.get(canon)
                idx = options.index(default) if default in options else 0
                sel = st.selectbox(f"{canon}", options, index=idx, key=f"map_{kind}_{canon}")
                mapping[canon] = None if sel == "(aucune)" else sel
        ss.mapped[kind] = mapping
        missing = [c for c in REQUIRED[kind] if not mapping.get(c)]
        if missing:
            st.warning(f"Colonnes requises non mappées : {', '.join(missing)}")
            mapping_ok = False
        else:
            st.success("Colonnes requises OK")
        st.dataframe(df.head(3), use_container_width=True)

if not mapping_ok:
    st.error("Complète le mapping des colonnes requises avant de continuer.")
    st.stop()

# --------------------------------------------------------------------------
# ÉTAPE 3 — Segmentation en clusters (sur échantillon de pages, éditable)
# --------------------------------------------------------------------------
st.subheader("3. Valider la segmentation en clusters")

# Pour proposer les clusters, lire la colonne url complète des pages (légère)
url_col = ss.mapped["pages"]["url"]
pages_urls = load_df(sources["pages"], usecols=[url_col]).rename(columns={url_col: "url"})
pages_urls["url"] = pages_urls["url"].map(lambda u: pl.normalize_url(u, cfg))
pages_urls = pages_urls.dropna(subset=["url"]).drop_duplicates(subset=["url"])

if ss.cluster_mapping is None or st.button("↺ Re-proposer la segmentation"):
    ss.cluster_mapping = pl.propose_clusters(pages_urls)

st.caption("Segmentation par pattern d'URL (les préfixes de langue type /fr/fr/ sont ignorés automatiquement). "
           "Renomme ou fusionne des clusters dans la colonne « cluster », puis valide.")
edited = st.data_editor(
    ss.cluster_mapping, use_container_width=True, num_rows="fixed",
    column_config={
        "pattern": st.column_config.TextColumn("Pattern URL", disabled=True),
        "cluster": st.column_config.TextColumn("Cluster (éditable)"),
        "nb_pages": st.column_config.NumberColumn("Nb pages", disabled=True),
        "exemples": st.column_config.TextColumn("Exemples", disabled=True),
    }, key="cluster_editor")

# --------------------------------------------------------------------------
# ÉTAPE 4 — Analyse (lecture complète, colonnes utiles seulement)
# --------------------------------------------------------------------------
st.subheader("4. Lancer l'analyse")
if st.button("🚀 Analyser", type="primary"):
    try:
        with st.spinner("Lecture des fichiers complets (colonnes utiles)…"):
            canon = {}
            for kind, src in sources.items():
                usecols = [c for c in ss.mapped[kind].values() if c]
                df = load_df(src, usecols=usecols)
                inv = {v: k for k, v in ss.mapped[kind].items() if v}
                canon[kind] = df.rename(columns=inv)
        with st.spinner(f"Analyse ({len(canon['links'])} liens, {len(canon['pages'])} pages)…"):
            ss.results = pl.run_pipeline(
                pages=canon["pages"], links=canon["links"],
                gsc=canon.get("gsc"), ga4=canon.get("ga4"), sem=canon.get("sem"),
                cluster_mapping=edited, cfg=cfg)
        st.success("Analyse terminée.")
    except Exception as e:
        st.exception(e)
        st.stop()

# --------------------------------------------------------------------------
# ETAPE 5 & 6 - Resultats et export
# --------------------------------------------------------------------------
if ss.results:
    res = ss.results
    p = res["pages"]
    rep = res["report"]

    st.subheader("5. Diagnostic")
    quad = p["quadrant_code"].value_counts().to_dict()
    k = st.columns(5)
    k[0].metric("Pages", rep.get("pages_total", len(p)))
    k[1].metric("Liens valides", rep.get("liens_valides", 0))
    k[2].metric("Orphelines", rep.get("pages_orphelines", 0))
    k[3].metric("Q1 prioritaires", quad.get("Q1", 0))
    k[4].metric("Prescriptions", rep.get("prescriptions_generees", 0))

    warn = []
    try:
        if not p["_has_ga4"].iloc[0]:
            warn.append("GA4 absent.")
        if not p["_has_sem"].iloc[0]:
            warn.append("Etude semantique absente.")
    except Exception:
        pass
    if rep.get("cibles_reclassees_navigation", 0):
        warn.append(f"{rep['cibles_reclassees_navigation']} cible(s) reclassee(s).")
    for w in warn:
        st.info(w)

    tabs = st.tabs(["Matrice 2 axes", "Prescription liens", "Problemes", "Silotage", "Rapport technique"])
    with tabs[0]:
        quad = p["quadrant_code"].value_counts().to_dict()
        mat = pd.DataFrame(
            {"Faible potentiel business": [f"Q4 · Ne rien faire · {quad.get('Q4', 0)}",
                                           f"Q3 · Sur-maillage · {quad.get('Q3', 0)}"],
             "Fort potentiel business": [f"Q1 · PRIORITAIRE · {quad.get('Q1', 0)}",
                                         f"Q2 · Hors scope · {quad.get('Q2', 0)}"]},
            index=["Fort déficit de maillage", "Bon maillage"])
        st.table(mat)
        st.caption("Q1 = prioritaire (prescrire des liens). Q3 = sur-maillage (réallouer le PageRank vers les Q1).")
    with tabs[1]:
        presc = res["prescriptions"]
        st.write(f"**{len(presc)} liens prescrits**.")
        st.dataframe(presc.head(500), use_container_width=True)
    with tabs[2]:
        flags = res["flags"]
        if len(flags):
            st.bar_chart(flags["type"].value_counts())
            st.dataframe(flags, use_container_width=True)
        else:
            st.write("Aucun flag.")
    with tabs[3]:
        st.dataframe(res["silo_matrix_pct"], use_container_width=True)
        st.dataframe(res["silo_flags"], use_container_width=True)
    with tabs[4]:
        st.json(rep)

    st.subheader("6. Télécharger les livrables")
    import tempfile
    safe = site_name.replace(" ", "_")
    c_xl, c_pp = st.columns(2)
    with c_xl:
        st.caption("Fichier Excel détaillé (8 onglets).")
        if st.button("📊 Générer l'Excel", key="gen_xl"):
            with st.spinner("Génération de l'Excel…"):
                tmp = os.path.join(tempfile.gettempdir(), "audit_maillage.xlsx")
                build_excel(res, tmp, site_name=site_name)
                with open(tmp, "rb") as fh:
                    ss["xl_data"] = fh.read()
        if ss.get("xl_data"):
            st.download_button("⬇️ Télécharger l'Excel", data=ss["xl_data"],
                               file_name=f"audit_maillage_{safe}_{dt.date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c_pp:
        st.caption("Deck PowerPoint explicatif, éditable, à intégrer dans un deck client.")
        if st.button("🖼️ Générer le PPT", key="gen_pp"):
            with st.spinner("Génération du PPT…"):
                tmp = os.path.join(tempfile.gettempdir(), "audit_maillage.pptx")
                build_deck(res, tmp, site_name=site_name)
                with open(tmp, "rb") as fh:
                    ss["pp_data"] = fh.read()
        if ss.get("pp_data"):
            st.download_button("⬇️ Télécharger le PPT", data=ss["pp_data"],
                               file_name=f"audit_maillage_{safe}_{dt.date.today()}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
