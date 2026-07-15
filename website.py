"""
Mediris eFact Analyse Dashboard
--------------------------------
Streamlit-applicatie om Mediris eFact exports te analyseren.
Er wordt NIETS opgeslagen: alle data leeft enkel in het geheugen van de
actieve browsersessie en verdwijnt zodra de sessie/tab wordt gesloten of
ververst. Er worden geen bestanden weggeschreven op de server.
"""

import io
import re
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go  # <-- Deze toevoegen bovenaan je script

import streamlit as st


st.set_page_config(
    page_title="Mediris eFact Analyse",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Codetabel: Prestatiecode -> leesbare naam
# ---------------------------------------------------------------------------
CODE_MAP = {
    "106315": "RPL onco",
    "102631": "RPL",
    "350055": "HAO",
    "384230": "NO",
    "471273": "Spiro met",
    "471251": "Spiro zonder",
    "471310": "RV",
    "471354": "Dlco",
    "471376": "Raw",
}


def code_to_label(code: str) -> str:
    code = str(code).strip()
    return CODE_MAP.get(code, f"Onbekende code ({code})")


def doctor_label(verstrekker: str, email: str) -> str:
    """Maak een leesbare dokternaam op basis van e-mail, met Verstrekker-ID als fallback."""
    if isinstance(email, str) and "@" in email:
        local = email.split("@")[0]
        parts = re.split(r"[._\-]+", local)
        naam = " ".join(p.capitalize() for p in parts if p)
        
        if naam:
            # --- OVERRIDE FOR BRUNO ---
            # Als de gegenereerde naam 'Brunodebelie' (of 'Bruno De Belie') is,
            # forceer deze dan direct naar 'Bruno'.
            clean_naam = naam.replace(" ", "").lower()
            if "brunodebelie" in clean_naam:
                return "Bruno"
            
            # Voor de andere artsen (Eva, Erika, Jelle, Alexandra) pakken we de voornaam:
            return parts[0].capitalize()
            
    return str(verstrekker)


# ---------------------------------------------------------------------------
# Data inladen
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=1)
def load_files(file_bytes_list):
    """Leest een lijst van (bestandsnaam, bytes) in en combineert alle sheets."""
    all_consult = []
    all_prest = []

    for fname, fbytes in file_bytes_list:
        xls = pd.ExcelFile(io.BytesIO(fbytes))

        if "Consultaties" in xls.sheet_names:
            cdf = xls.parse("Consultaties")
            cdf["__bronbestand"] = fname
            all_consult.append(cdf)

        if "Prestaties" in xls.sheet_names:
            pdf = xls.parse("Prestaties")
            pdf["__bronbestand"] = fname
            all_prest.append(pdf)

    consult = pd.concat(all_consult, ignore_index=True) if all_consult else pd.DataFrame()
    prest = pd.concat(all_prest, ignore_index=True) if all_prest else pd.DataFrame()

    # Datum parsen
    for df in (consult, prest):
        if not df.empty and "Datum" in df.columns:
            df["Datum"] = pd.to_datetime(df["Datum"], format="%d-%m-%Y %H:%M", errors="coerce")

    # Dokternaam toevoegen
    if not consult.empty:
        consult["Dokter"] = consult.apply(
            lambda r: doctor_label(r.get("Verstrekker"), r.get("Email")), axis=1
        )
    if not prest.empty:
        # Prestaties-sheet heeft geen Email-kolom -> koppel via Verstrekker-ID
        if not consult.empty:
            verstrekker_naam = (
                consult.drop_duplicates("Verstrekker")
                .set_index("Verstrekker")["Dokter"]
                .to_dict()
            )
            prest["Dokter"] = prest["Verstrekker"].map(verstrekker_naam).fillna(
                prest["Verstrekker"]
            )
        else:
            prest["Dokter"] = prest["Verstrekker"]
        prest["Ingreep"] = prest["Code"].apply(code_to_label)

    # Als er geen aparte Prestaties-sheet is, reconstrueer op basis van de
    # 'Prestaties'-kolom (codes gescheiden door "-") in Consultaties. Hierbij
    # is enkel een telling mogelijk, geen exacte bedragen per code.
    if prest.empty and not consult.empty and "Prestaties" in consult.columns:
        rows = []
        for _, r in consult.iterrows():
            codes = str(r.get("Prestaties") or "").split("-")
            for c in codes:
                c = c.strip()
                if not c or c.lower() == "nan":
                    continue
                rows.append(
                    {
                        "Contact Id": r.get("Id"),
                        "Datum": r.get("Datum"),
                        "Verstrekker": r.get("Verstrekker"),
                        "Patiënt": r.get("Patiënt"),
                        "Code": c,
                        "Dokter": r.get("Dokter"),
                        "Ingreep": code_to_label(c),
                        "Status": r.get("Status"),
                        "__bronbestand": r.get("__bronbestand"),
                    }
                )
        prest = pd.DataFrame(rows)

    return consult, prest


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📊 Mediris eFact Analyse")
st.caption(
    "Upload één of meerdere Mediris eFact export-bestanden (.xlsx). "
    "Er wordt niets bewaard op de server — alle verwerking gebeurt enkel "
    "in het geheugen van deze sessie."
)

uploaded_files = st.file_uploader(
    "Kies één of meerdere Mediris eFact export-bestanden",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Upload minstens één bestand om te starten.")
    st.stop()

file_bytes_list = [(f.name, f.getvalue()) for f in uploaded_files]
consult, prest = load_files(file_bytes_list)

if consult.empty:
    st.error("Kon geen 'Consultaties'-sheet terugvinden in de geüploade bestanden.")
    st.stop()

# ---------------------------------------------------------------------------
# Filters (sidebar)
# ---------------------------------------------------------------------------
st.sidebar.header("Filters")

statuses = sorted(consult["Status"].dropna().unique().tolist())
gekozen_status = st.sidebar.multiselect("Status", statuses, default=statuses)

doctors_all = sorted(consult["Dokter"].dropna().unique().tolist())
gekozen_doctors = st.sidebar.multiselect("Dokter(s)", doctors_all, default=doctors_all)

min_date = consult["Datum"].min()
max_date = consult["Datum"].max()
if pd.notna(min_date) and pd.notna(max_date):
    date_range = st.sidebar.date_input(
        "Periode",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
    )
else:
    date_range = None

mask = consult["Status"].isin(gekozen_status) & consult["Dokter"].isin(gekozen_doctors)
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = date_range
    mask &= consult["Datum"].dt.date.between(start, end)

cdf = consult[mask].copy()

if not prest.empty:
    pmask = prest["Status"].isin(gekozen_status) & prest["Dokter"].isin(gekozen_doctors)
    if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
        pmask &= prest["Datum"].dt.date.between(start, end)
    pdf = prest[pmask].copy()
else:
    pdf = prest

st.sidebar.markdown("---")
st.sidebar.caption(f"📁 {len(uploaded_files)} bestand(en) geladen")
st.sidebar.caption(f"📄 {len(cdf)} consultaties na filtering")

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_overzicht, tab_dokters, tab_patienten, tab_ingrepen, tab_data = st.tabs(
    ["📊 Overzicht", "👨‍⚕️ Per dokter", "🔁 Patiënten", "🩺 Ingrepen", "🗂️ Ruwe data"]
)

# --- Overzicht -------------------------------------------------------------
with tab_overzicht:
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Consultaties", f"{len(cdf):,}".replace(",", "."))
    c2.metric("Unieke patiënten", f"{cdf['Patiënt'].nunique():,}".replace(",", "."))
    c3.metric("Unieke dokters", cdf["Dokter"].nunique())
    c4.metric("Totale omzet (€)", f"{cdf['Totaal'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    c5.metric("Gem. per consultatie (€)", f"{cdf['Totaal'].mean():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    st.markdown("### Omzet en volume doorheen de tijd")
    if cdf["Datum"].notna().any():
        tijd = cdf.dropna(subset=["Datum"]).copy()
        tijd["Maand"] = tijd["Datum"].dt.to_period("M").dt.to_timestamp()
        per_maand = tijd.groupby("Maand").agg(
            Consultaties=("Id", "count"), Omzet=("Totaal", "sum")
        ).reset_index()
        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(per_maand, x="Maand", y="Consultaties", title="Aantal consultaties per maand")
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            fig = px.bar(per_maand, x="Maand", y="Omzet", title="Omzet per maand (€)")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Status van de consultaties")
    status_counts = cdf["Status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Aantal"]
    fig = px.pie(status_counts, names="Status", values="Aantal", title="Verdeling per status")
    st.plotly_chart(fig, use_container_width=True)

# --- Per dokter --------------------------------------------------------
# with tab_dokters:
#     st.markdown("### Overzicht per dokter")
#     per_dokter = cdf.groupby("Dokter").agg(
#         Consultaties=("Id", "count"),
#         Unieke_patienten=("Patiënt", "nunique"),
#         Omzet=("Totaal", "sum"),
#         Gem_per_consult=("Totaal", "mean"),
#     ).reset_index().sort_values("Consultaties", ascending=False)
#     per_dokter.columns = ["Dokter", "Consultaties", "Unieke patiënten", "Omzet (€)", "Gem. per consult (€)"]
#     st.dataframe(
#         per_dokter.style.format({"Omzet (€)": "{:,.2f}", "Gem. per consult (€)": "{:,.2f}"}),
#         use_container_width=True,
#         hide_index=True,
#     )

#     col1, col2 = st.columns(2)
#     with col1:
#         fig = px.bar(per_dokter, x="Dokter", y="Consultaties", title="Aantal consultaties per dokter")
#         st.plotly_chart(fig, use_container_width=True)
#     with col2:
#         fig = px.bar(per_dokter, x="Dokter", y="Omzet (€)", title="Omzet per dokter (€)")
#         st.plotly_chart(fig, use_container_width=True)

#     st.markdown("### Consultaties per dokter doorheen de tijd")
#     if cdf["Datum"].notna().any():
#         tijd = cdf.dropna(subset=["Datum"]).copy()
#         tijd["Maand"] = tijd["Datum"].dt.to_period("M").dt.to_timestamp()
#         per_maand_dokter = tijd.groupby(["Maand", "Dokter"]).size().reset_index(name="Consultaties")
#         fig = px.line(per_maand_dokter, x="Maand", y="Consultaties", color="Dokter", markers=True)
#         st.plotly_chart(fig, use_container_width=True)

        # --- Per dokter --------------------------------------------------------
with tab_dokters:
    st.markdown("### Overzicht per dokter")
    
    # 1. Definieer vaste, herkenbare kleuren voor elke dokter
    DOCTOR_COLORS = {
        "Bruno": "#1f77b4",       # Blauw
        "Eva": "#e377c2",         # Roze
        "Erika": "#2ca02c",       # Groen
        "Jelle": "#ff7f0e",       # Oranje
        "Alexandra": "#9467bd"    # Paars
    }
    
    per_dokter = cdf.groupby("Dokter").agg(
        Consultaties=("Id", "count"),
        Unieke_patienten=("Patiënt", "nunique"),
        Omzet=("Totaal", "sum"),
        Gem_per_consult=("Totaal", "mean"),
    ).reset_index().sort_values("Consultaties", ascending=False)
    per_dokter.columns = ["Dokter", "Consultaties", "Unieke patiënten", "Omzet (€)", "Gem. per consult (€)"]
    
    st.dataframe(
        per_dokter.style.format({"Omzet (€)": "{:,.2f}", "Gem. per consult (€)": "{:,.2f}"}),
        use_container_width=True,
        hide_index=True,
    )

    st.info("💡 **Tip:** Klik op de naam van een dokter in de legende van de grafieken om deze aan of uit te zetten.")

    col1, col2 = st.columns(2)
    with col1:
        # Toegevoegd: color="Dokter" en color_discrete_map voor individuele kleuren en toggle-functionaliteit
        fig = px.bar(
            per_dokter, 
            x="Dokter", 
            y="Consultaties", 
            color="Dokter",
            color_discrete_map=DOCTOR_COLORS,
            title="Aantal consultaties per dokter"
        )
        # Zorgt ervoor dat de legende altijd klikbaar is om te filteren
        fig.update_layout(clickmode='event+select') 
        st.plotly_chart(fig, use_container_width=True)
        
    with col2:
        # Toegevoegd: color="Dokter" en color_discrete_map
        fig = px.bar(
            per_dokter, 
            x="Dokter", 
            y="Omzet (€)", 
            color="Dokter",
            color_discrete_map=DOCTOR_COLORS,
            title="Omzet per dokter (€)"
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Consultaties per dokter doorheen de tijd")
    if cdf["Datum"].notna().any():
        tijd = cdf.dropna(subset=["Datum"]).copy()
        tijd["Maand"] = tijd["Datum"].dt.to_period("M").dt.to_timestamp()
        per_maand_dokter = tijd.groupby(["Maand", "Dokter"]).size().reset_index(name="Consultaties")
        
        # Toegevoegd: color_discrete_map voor consistente lijnkleuren per dokter
        fig = px.line(
            per_maand_dokter, 
            x="Maand", 
            y="Consultaties", 
            color="Dokter", 
            color_discrete_map=DOCTOR_COLORS,
            markers=True
        )
        st.plotly_chart(fig, use_container_width=True)

# --- Patiënten: terugkerend vs nieuw -----------------------------------
with tab_patienten:
    st.markdown("### Terugkerende versus eenmalige patiënten")

    bezoeken = cdf.groupby("Patiënt").agg(
        Aantal_bezoeken=("Id", "count"),
        Eerste_bezoek=("Datum", "min"),
        Laatste_bezoek=("Datum", "max"),
        Aantal_dokters=("Dokter", "nunique"),
        Dokters=("Dokter", lambda s: ", ".join(sorted(s.dropna().unique()))),
        Omzet=("Totaal", "sum"),
    ).reset_index()
    bezoeken["Terugkerend"] = bezoeken["Aantal_bezoeken"] > 1

    n_terug = int(bezoeken["Terugkerend"].sum())
    n_eenmalig = int((~bezoeken["Terugkerend"]).sum())
    n_totaal = len(bezoeken)

    c1, c2, c3 = st.columns(3)
    c1.metric("Totaal aantal patiënten", n_totaal)
    c2.metric("Terugkerende patiënten", f"{n_terug} ({n_terug / n_totaal:.1%})")
    c3.metric("Eenmalige patiënten", f"{n_eenmalig} ({n_eenmalig / n_totaal:.1%})")

    fig = px.pie(
        names=["Terugkerend (≥2 bezoeken)", "Eenmalig (1 bezoek)"],
        values=[n_terug, n_eenmalig],
        title="Verdeling terugkerende vs eenmalige patiënten",
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Lijst terugkerende patiënten")
    terug_df = bezoeken[bezoeken["Terugkerend"]].sort_values("Aantal_bezoeken", ascending=False)
    st.dataframe(
        terug_df[["Patiënt", "Aantal_bezoeken", "Aantal_dokters", "Dokters", "Eerste_bezoek", "Laatste_bezoek", "Omzet"]]
        .rename(columns={
            "Patiënt": "Patiënt",
            "Aantal_bezoeken": "Aantal bezoeken",
            "Aantal_dokters": "Aantal verschillende dokters",
            "Eerste_bezoek": "Eerste bezoek",
            "Laatste_bezoek": "Laatste bezoek",
            "Omzet": "Totale omzet (€)",
        })
        .style.format({"Totale omzet (€)": "{:,.2f}"}),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("#### Lijst eenmalige patiënten")
    eenmalig_df = bezoeken[~bezoeken["Terugkerend"]].sort_values("Eerste_bezoek", ascending=False)
    st.dataframe(
        eenmalig_df[["Patiënt", "Eerste_bezoek", "Dokters", "Omzet"]]
        .rename(columns={"Patiënt": "Patiënt", "Eerste_bezoek": "Bezoekdatum", "Dokters": "Dokter", "Omzet": "Omzet (€)"})
        .style.format({"Omzet (€)": "{:,.2f}"}),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("### Terugkeerpercentage per dokter")
    st.caption("Percentage van de patiënten van een dokter dat bij diezelfde dokter is teruggekomen.")
    rows = []
    for dokter, g in cdf.groupby("Dokter"):
        per_pat = g.groupby("Patiënt").size()
        terug = int((per_pat > 1).sum())
        totaal = int(per_pat.shape[0])
        rows.append({
            "Dokter": dokter,
            "Unieke patiënten": totaal,
            "Terugkerend bij deze dokter": terug,
            "Terugkeerpercentage": terug / totaal if totaal else 0,
        })
    retentie_df = pd.DataFrame(rows).sort_values("Terugkeerpercentage", ascending=False)
    st.dataframe(
        retentie_df.style.format({"Terugkeerpercentage": "{:.1%}"}),
        use_container_width=True,
        hide_index=True,
    )
    fig = px.bar(retentie_df, x="Dokter", y="Terugkeerpercentage", title="Terugkeerpercentage per dokter")
    fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Patiënten die meerdere dokters bezochten")
    multi_doc = bezoeken[bezoeken["Aantal_dokters"] > 1].sort_values("Aantal_dokters", ascending=False)
    st.caption(f"{len(multi_doc)} patiënt(en) zagen meer dan één dokter binnen de gekozen periode.")
    if not multi_doc.empty:
        st.dataframe(
            multi_doc[["Patiënt", "Aantal_dokters", "Dokters", "Aantal_bezoeken"]]
            .rename(columns={"Aantal_dokters": "Aantal dokters", "Dokters": "Dokters bezocht", "Aantal_bezoeken": "Totaal bezoeken"}),
            use_container_width=True,
            hide_index=True,
        )

# --- Ingrepen / Prestaties ----------------------------------------------
with tab_ingrepen:
    st.markdown("### Ingrepen (prestaties) per dokter")
    if pdf.empty:
        st.warning("Geen prestatiedata beschikbaar.")
    else:
        st.caption(
            "Gebaseerd op de kolom 'Prestaties'/'Code', waarbij codes zijn omgezet naar "
            "leesbare namen: " + ", ".join(f"{k} = {v}" for k, v in CODE_MAP.items())
        )

        totaal_per_ingreep = pdf["Ingreep"].value_counts().reset_index()
        totaal_per_ingreep.columns = ["Ingreep", "Aantal"]
        fig = px.bar(totaal_per_ingreep, x="Ingreep", y="Aantal", title="Totaal aantal per type ingreep (alle dokters)")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Ingrepen per dokter (tabel)")
        pivot = pd.crosstab(pdf["Dokter"], pdf["Ingreep"])
        pivot["Totaal"] = pivot.sum(axis=1)
        pivot = pivot.sort_values("Totaal", ascending=False)
        st.dataframe(pivot, use_container_width=True)

        st.markdown("#### Ingrepen per dokter (grafiek)")
        per_dokter_ingreep = pdf.groupby(["Dokter", "Ingreep"]).size().reset_index(name="Aantal")
        fig = px.bar(
            per_dokter_ingreep, x="Dokter", y="Aantal", color="Ingreep",
            title="Verdeling van ingrepen per dokter", barmode="stack",
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Meest uitgevoerde ingreep per dokter")
        top_per_dokter = (
            per_dokter_ingreep.sort_values("Aantal", ascending=False)
            .groupby("Dokter")
            .first()
            .reset_index()
            .rename(columns={"Ingreep": "Meest uitgevoerde ingreep", "Aantal": "Aantal keer"})
        )
        st.dataframe(top_per_dokter, use_container_width=True, hide_index=True)

        if "Totaal" in pdf.columns and pdf["Totaal"].notna().any():
            st.markdown("#### Omzet per type ingreep")
            omzet_ingreep = pdf.groupby("Ingreep")["Totaal"].sum().reset_index().sort_values("Totaal", ascending=False)
            fig = px.bar(omzet_ingreep, x="Ingreep", y="Totaal", title="Omzet per type ingreep (€)")
            st.plotly_chart(fig, use_container_width=True)

# --- Ruwe data ------------------------------------------------------------
with tab_data:
    st.markdown("### Consultaties (gefilterd)")
    st.dataframe(cdf, use_container_width=True)
    st.download_button(
        "⬇️ Download gefilterde consultaties als CSV",
        data=cdf.to_csv(index=False).encode("utf-8"),
        file_name="consultaties_gefilterd.csv",
        mime="text/csv",
    )

    if not pdf.empty:
        st.markdown("### Prestaties (gefilterd)")
        st.dataframe(pdf, use_container_width=True)
        st.download_button(
            "⬇️ Download gefilterde prestaties als CSV",
            data=pdf.to_csv(index=False).encode("utf-8"),
            file_name="prestaties_gefilterd.csv",
            mime="text/csv",
        )

st.markdown("---")
st.caption(
    "🔒 Privacy: deze applicatie bewaart niets. Alle geüploade data blijft enkel "
    "in het geheugen van je browsersessie en wordt niet naar schijf of een database "
    "geschreven. Sluit of ververs de pagina om alles te wissen."
)