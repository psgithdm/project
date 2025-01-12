import pandas as pd
#import matplotlib.pyplot as plt
import plotly.express as px
import streamlit as st
from fpdf import FPDF
from datetime import datetime
import plotly.express as px
import plotly.io as pio
import os
import warnings

warnings.filterwarnings("ignore", message="missing ScriptRunContext!")

# Sidebar
# Wide Mode aktivieren
st.set_page_config(layout="wide")

# Einlesen der Excel-Daten
file_path = "../data/raw/Liefertreue_Daten_2024_final_liefertreue.xlsx"
df = pd.read_excel(file_path)

# Datenquelle Informationen
data_source = {
    "Dateiname": os.path.basename(file_path),
    "Letzte Bearbeitung": pd.to_datetime(os.path.getmtime(file_path), unit='s').strftime("%Y-%m-%d %H:%M:%S")
}

# Datenqualität analysieren
duplicates_count = df.duplicated().sum()
missing_values_count = df.isnull().sum()
missing_percentages = (missing_values_count / len(df) * 100).round(2)

# Duplikate speichern
df_duplicate_data = df[df.duplicated()]  # Enthält nur die Duplikate

# Datenbereinigung
df_cleaned = df.drop_duplicates()  # Duplikate entfernen
df_cleaned.fillna({"WE-Menge": 0, "Soll-Menge": 0}, inplace=True)  # Fehlende Werte füllen

# Berechnung der Liefertreue und Mengenabweichung
df_cleaned["Verspätung (Tage)"] = (
    pd.to_datetime(df_cleaned["Wareneingangsdatum (WE)"]) - pd.to_datetime(df_cleaned["Lieferdatum (Soll)"])
).dt.days
df_cleaned["Liefertreue (Ja/Nein)"] = df_cleaned["Verspätung (Tage)"].apply(lambda x: "Ja" if x <= 0 else "Nein")
df_cleaned["Mengenabweichung"] = df_cleaned["WE-Menge"] - df_cleaned["Soll-Menge"]

# Spalte Datenqualität
df_cleaned["Datenqualität"] = df_cleaned.apply(
    lambda row: "Fehlende Werte" if row.isnull().any() else "OK", axis=1
)

# Tabs erstellen
tabs = st.tabs(["Dashboard Übersicht", "Analyse Lieferant", "Analyse Material", "PDF-Report", "Datenqualität", "Datenquelle", "Kontakt"])
df = df_cleaned.copy()

# Sidebar-Filter
st.sidebar.header("Filteroptionen")

# Reihenfolge der Filter in Sidebar: Land, Jahr, Monat, Liefertreue, Mengenabweichung, Lieferant
# Länderauswahl
selected_country = st.sidebar.multiselect(
    "Selektion Länder:", options=df["Land"].unique(), default=df["Land"].unique()
)

# Jahr-Auswahl
min_date = pd.to_datetime(df["Lieferdatum (Soll)"].min()).date()
max_date = pd.to_datetime(df["Lieferdatum (Soll)"].max()).date()
selected_year = st.sidebar.selectbox(
    "Selektion Jahr:", options=range(min_date.year, max_date.year + 1), index=max_date.year - min_date.year
)

# Monat-Auswahl
month_names = ["Januar", "Februar", "März", "April", "Mai", "Juni",
               "Juli", "August", "September", "Oktober", "November", "Dezember"]
selected_months = st.sidebar.multiselect(
    "Selektion Monate:", options=range(1, 13), format_func=lambda x: month_names[x - 1], default=range(1, 13)
)

# Liefertreue-Filter
liefertreue_options = ["Alle", "Ja", "Nein"]
selected_liefertreue = st.sidebar.multiselect(
    "Selektion Liefertreue:", options=liefertreue_options, default=["Alle"]
)

# Mengenabweichungsfilter (min und max)
min_abweichung, max_abweichung = st.sidebar.slider(
    "Filter nach Mengenabweichung (Ist - Soll):",
    min_value=int(df["Mengenabweichung"].min()),
    max_value=int(df["Mengenabweichung"].max()),
    value=(int(df["Mengenabweichung"].min()), int(df["Mengenabweichung"].max())),
    step=1
)

# Lieferantenauswahl
sorted_suppliers = sorted(df["Lieferantenbezeichnung"].unique())
supplier_options = ["Alle"] + sorted_suppliers
selected_suppliers = st.sidebar.multiselect(
    "Selektion Lieferanten:", options=supplier_options, default=["Alle"]
)

# Filterdaten anwenden
filtered_df = df.copy()

if selected_country:
    filtered_df = filtered_df[filtered_df["Land"].isin(selected_country)]

filtered_df = filtered_df[
    (pd.to_datetime(filtered_df["Lieferdatum (Soll)"]).dt.year == selected_year) &
    (pd.to_datetime(filtered_df["Lieferdatum (Soll)"]).dt.month.isin(selected_months))
]

if "Alle" not in selected_liefertreue:
    filtered_df = filtered_df[filtered_df["Liefertreue (Ja/Nein)"].isin(selected_liefertreue)]

filtered_df = filtered_df[
    (filtered_df["Mengenabweichung"] >= min_abweichung) &
    (filtered_df["Mengenabweichung"] <= max_abweichung)
]

if "Alle" not in selected_suppliers:
    filtered_df = filtered_df[filtered_df["Lieferantenbezeichnung"].isin(selected_suppliers)]

# Anzeigen der gefilterten Daten
st.sidebar.markdown(f"### Gefilterte Daten: {len(filtered_df)} Einträge")

# Tab 0: Dashboard Übersicht
with tabs[0]:
    st.title("Dashboard Übersicht")
    
    # Kennzahlen
    total_deliveries = len(filtered_df)
    on_time = len(filtered_df[filtered_df["Liefertreue (Ja/Nein)"] == "Ja"])
    delayed = len(filtered_df[filtered_df["Liefertreue (Ja/Nein)"] == "Nein"])
    
    # Berechnung des Anteils Liefertreue = Nein
    total_rows = len(filtered_df)
    no_delivery_reliability = filtered_df[filtered_df["Liefertreue (Ja/Nein)"] == "Nein"]
    no_delivery_reliability_count = len(no_delivery_reliability)
    reliability_no_percentage = round((no_delivery_reliability_count / total_rows * 100), 2)

    
    # Zusätzliche Kennzahlen
    unique_suppliers = filtered_df["Lieferantenbezeichnung"].nunique()
    unique_materials = filtered_df["Materialnummer"].nunique()
    unique_invoices = filtered_df["Lieferscheinnummer"].nunique()
    unique_countries = filtered_df["Land"].nunique()

    # Einheitliches Design für die Kennzahlen
    def styled_metric(label, value, background_color="#1976D2", text_color="white"):
        """
        Zeigt eine Kennzahl mit einem einheitlichen farbigen Hintergrund an.
        
        Parameters:
            label (str): Beschriftung der Kennzahl.
            value (str/int/float): Wert der Kennzahl.
            background_color (str): Hintergrundfarbe (Standard: "#1976D2" für grün).
            text_color (str): Schriftfarbe (Standard: "white").
        """
        return f"""
        <div style="background-color: {background_color}; 
                    padding: 15px; 
                    border-radius: 15px; 
                    text-align: center; 
                    color: {text_color};">
            <h5>{label}</h5>
            <h3>{value}</h3>
        </div>
        """
    
    # Erste Reihe von Kennzahlen
    st.markdown("### Kennzahlen")
    col1, col2, col3, col4 = st.columns(4)

    col1.markdown(styled_metric("Anzahl Lieferungen", total_deliveries), unsafe_allow_html=True)
    col2.markdown(styled_metric("Pünktliche Lieferungen", on_time), unsafe_allow_html=True)
    col3.markdown(styled_metric("Verspätete Lieferungen", delayed), unsafe_allow_html=True)
    col4.markdown(styled_metric("Nicht termingerechte", f"{reliability_no_percentage:.2f}%"), unsafe_allow_html=True)

    # Leerzeichen zwischen den Reihen
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Zweite Reihe von Kennzahlen
    col5, col6, col7, col8 = st.columns(4)

    col5.markdown(styled_metric("Anzahl Lieferanten", unique_suppliers), unsafe_allow_html=True)
    col6.markdown(styled_metric("Anzahl Materialien", unique_materials), unsafe_allow_html=True)
    col7.markdown(styled_metric("Anzahl Lieferscheine", unique_invoices), unsafe_allow_html=True)
    col8.markdown(styled_metric("Anzahl Länder", unique_countries), unsafe_allow_html=True)

    # Leerzeichen zwischen den Reihen
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Diagramme
    st.markdown("### Auswertungen")
    col1, col2, col3 = st.columns(3)

    # Daten vorbereiten: Liefertreue zählen
    liefertreue_counts = (
        filtered_df["Liefertreue (Ja/Nein)"]
        .value_counts()  # Anteil berechnen
        .reset_index()
        #.rename(columns={"index": "Liefertreue", "Liefertreue (Ja/Nein)": "Anzahl"})
    )

    # Überprüfen, ob die Daten korrekt sind
    #st.write("Liefertreue-Daten (für Debugging):", liefertreue_counts)

    # Horizontales Balkendiagramm erstellen
    fig_bar = px.bar(
        liefertreue_counts,
        x="count", 
        y="Liefertreue (Ja/Nein)",
        orientation='h',  # Horizontal
        title="Anteil der Liefertreue",
        labels={"Anzahl": "Anzahl der Lieferungen", "Liefertreue": "Liefertreue (Ja/Nein)"},
        color="Liefertreue (Ja/Nein)",
        color_discrete_sequence=["#1976D2", "#63B2EE"],
        text="count"
    )

    fig_bar.update_layout(
        xaxis_title="Anzahl der Lieferungen",
        yaxis_title="Liefertreue (Ja/Nein)",
        showlegend=False,  # Keine Legende
        plot_bgcolor="rgba(0,0,0,0)"  # Transparenter Hintergrund
    )

    # Anzeige im Streamlit-Dashboard
    col1.plotly_chart(fig_bar, use_container_width=True)
    pio.write_image(fig_bar, "liefertreue_chart.png", width=794, height=400)

    # Zeitverlauf: Liefertreue
    daily_deliveries = (
        filtered_df.groupby(pd.to_datetime(filtered_df["Lieferdatum (Soll)"]).dt.date)["Liefertreue (Ja/Nein)"]
        .value_counts()
        .unstack(fill_value=0)
    )
    daily_deliveries["Datum"] = daily_deliveries.index
    fig_line = px.area(
        daily_deliveries, x="Datum", y=["Ja", "Nein"],
        title="Liefertreue über die Zeit",
        labels={"value": "Anzahl", "variable": "Status"}
    )
    col2.plotly_chart(fig_line, use_container_width=True)
    
    # Abweichungen pro Land berechnen
    abweichung_nach_land = filtered_df.groupby("Land").agg({
        "Soll-Menge": "sum",
        "WE-Menge": "sum"
    }).reset_index()

    # Abweichung berechnen: Ist - Soll
    abweichung_nach_land["Abweichung"] = abweichung_nach_land["WE-Menge"] - abweichung_nach_land["Soll-Menge"]

    # Separate Spalten für Über- und Unterlieferung
    abweichung_nach_land["Überlieferung"] = abweichung_nach_land["Abweichung"].apply(lambda x: x if x > 0 else 0)
    abweichung_nach_land["Unterlieferung"] = abweichung_nach_land["Abweichung"].apply(lambda x: x if x < 0 else 0)

    # Bar-Chart erstellen
    fig_stacked_bar = px.bar(
        abweichung_nach_land,
        x="Land",
        y=["Überlieferung", "Unterlieferung"],  # Zwei separate Balken: Über- und Unterlieferung
        title="Über- und Unterlieferungen nach Land",
        labels={"value": "Abweichung (Ist - Soll)", "variable": "Typ", "Land": "Land"},
        text_auto=True,  # Automatische Anzeige der Werte
        color_discrete_map={"Überlieferung": "#1976D2", "Unterlieferung": "#63B2EE"}# Farben für die beiden Kategorien
    )

    # Layout anpassen
    fig_stacked_bar.update_layout(
        barmode="relative",  # Balken gestapelt (relativ)
        xaxis_title="Land",
        yaxis_title="Abweichung (Ist - Soll)",
        plot_bgcolor="rgba(0,0,0,0)",  # Hintergrund transparent
        showlegend=True  # Legende für Über- und Unterlieferung anzeigen
    )

    # Anzeige im Streamlit-Dashboard
    col3.plotly_chart(fig_stacked_bar, use_container_width=True)

    # CSS für breitere Scrollbar hinzufügen
    st.markdown(
        """
        <style>
        /* Breitere Scrollbar */
        ::-webkit-scrollbar {
            width: 12px;
            height: 12px;
        }
        /* Scrollbar-Farbe */
        ::-webkit-scrollbar-thumb {
            background: #1976D2;  /* Farbe der Scrollbar */
            border-radius: 10px;  /* Abgerundete Kanten */
        }
        /* Hintergrund der Scrollbar */
        ::-webkit-scrollbar-track {
            background: #f1f1f1;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    
# Tab 1: Analyse Lieferant
with tabs[1]:
    st.title("Analyse Lieferant")

    # Spalten auswählen und Datumsformat anpassen
    filtered_supplier_data = filtered_df.copy()
    filtered_supplier_data["Bestelldatum"] = pd.to_datetime(filtered_supplier_data["Bestelldatum"]).dt.strftime('%d.%m.%Y')
    filtered_supplier_data["Lieferdatum (Soll)"] = pd.to_datetime(filtered_supplier_data["Lieferdatum (Soll)"]).dt.strftime('%d.%m.%Y')
    filtered_supplier_data["Wareneingangsdatum (WE)"] = pd.to_datetime(filtered_supplier_data["Wareneingangsdatum (WE)"]).dt.strftime('%d.%m.%Y')

    # Tabelle erstellen mit den ausgewählten Spalten
    table_columns = [
        "Lieferantennummer",
        "Lieferantenbezeichnung",
        "Land",
        "Lieferscheinnummer",
        "Materialnummer",
        "Bestelldatum",
        "Lieferdatum (Soll)",
        "Wareneingangsdatum (WE)",
        "Soll-Menge",
        "WE-Menge",
        "Verspätung (Tage)"
    ]

    supplier_table = filtered_supplier_data[table_columns]

    # --- Diagramme und Analysen ---
    st.markdown("### Analysen")
    col1 = st.container()
    
        # 3. Lieferperformance der letzten 6 Monate
    filtered_supplier_data["Lieferdatum (Soll)"] = pd.to_datetime(
        filtered_supplier_data["Lieferdatum (Soll)"], format="%d.%m.%Y", dayfirst=True
    )
    max_date = filtered_supplier_data["Lieferdatum (Soll)"].max()
    
    last_six_months = pd.date_range(end=max_date, periods=6, freq="M").to_period("M")

    lieferperformance = (
        filtered_supplier_data[
            pd.to_datetime(filtered_supplier_data["Lieferdatum (Soll)"]).dt.to_period("M").isin(last_six_months)
        ]
        .groupby([
            pd.to_datetime(filtered_supplier_data["Lieferdatum (Soll)"]).dt.to_period("M"),
            "Lieferantenbezeichnung"
        ])
        .agg({
            "Lieferscheinnummer": "count",
            "Liefertreue (Ja/Nein)": lambda x: round((x == "Ja").mean() * 100, 2)  # Anteil mit 2 Nachkommastellen
        })
        .reset_index()
        .rename(columns={"Lieferdatum (Soll)": "Monat", "Liefertreue (Ja/Nein)": "Zuverlässigkeit"})
    )

    lieferperformance_pivot = lieferperformance.pivot(index="Monat", columns="Lieferantenbezeichnung", values="Zuverlässigkeit").fillna(0)
    lieferperformance_pivot.index = lieferperformance_pivot.index.astype(str)
    df_lieferperformance = lieferperformance_pivot.reset_index().melt(id_vars=["Monat"], var_name="Lieferant", value_name="Zuverlässigkeit")

    fig = px.line(
        df_lieferperformance,
        x="Monat",
        y="Zuverlässigkeit",
        color="Lieferant",
        title=f"Lieferperformance der letzten 6 Monate (bis {max_date.strftime('%B %Y')})",
        labels={"Monat": "Monat", "Zuverlässigkeit": "Zuverlässigkeit (%)", "Lieferant": "Lieferant"}
    )

    fig.update_layout(
        yaxis=dict(ticksuffix="%", range=[0, 100]),
        xaxis=dict(showgrid=True),
        plot_bgcolor="white",
    )
    fig.update_traces(mode="lines+markers")
    fig.update_layout(hovermode="x unified")

    col1.plotly_chart(fig, use_container_width=True)

    col2, col3 = st.columns(2)
    
    # 1. Liefertreue Verteilung (Gestapeltes Balkendiagramm)
    liefertreue_summary = (
        filtered_supplier_data.groupby(["Lieferantenbezeichnung", "Liefertreue (Ja/Nein)"])["Lieferscheinnummer"]
        .count()
        .reset_index()
    )

    # Berechnung des Anteils von "Nein" für jeden Lieferanten
    total_counts = (
        liefertreue_summary.groupby("Lieferantenbezeichnung")["Lieferscheinnummer"]
        .sum()
        .reset_index()
        .rename(columns={"Lieferscheinnummer": "Total"})
    )

    liefertreue_summary = liefertreue_summary.merge(total_counts, on="Lieferantenbezeichnung")
    liefertreue_summary["Anteil Nein"] = liefertreue_summary.apply(
        lambda row: (row["Lieferscheinnummer"] / row["Total"] * 100) if row["Liefertreue (Ja/Nein)"] == "Nein" else 0,
        axis=1
    )

    # Sortierung basierend auf dem Anteil "Nein"
    top_10_lieferanten = (
        liefertreue_summary[liefertreue_summary["Liefertreue (Ja/Nein)"] == "Nein"]
        .sort_values(by="Anteil Nein", ascending=False)
        .head(10)["Lieferantenbezeichnung"]
    )

    # Filterung und Berechnung der Prozentwerte
    filtered_top_data = liefertreue_summary[
        liefertreue_summary["Lieferantenbezeichnung"].isin(top_10_lieferanten)
    ]

    filtered_top_data["Prozent"] = (
        filtered_top_data.groupby("Lieferantenbezeichnung")["Lieferscheinnummer"]
        .transform(lambda x: round(100 * x / x.sum(), 2))  # Prozent mit 2 Nachkommastellen
    )

    # Sortieren der gefilterten Daten nach Anteil Nein
    filtered_top_data = filtered_top_data.sort_values(by="Anteil Nein", ascending=False)

    # Gestapeltes Balkendiagramm erstellen
    liefertreue_barchart = px.bar(
        filtered_top_data,
        x="Lieferantenbezeichnung",
        y="Lieferscheinnummer",
        color="Liefertreue (Ja/Nein)",
        text=filtered_top_data["Prozent"].astype(str) + "%",  # Prozent als Text-Label
        title="Top 10 Lieferanten mit den höchsten Anteilen an Liefertreue = Nein",
        labels={
            "Lieferscheinnummer": "Anzahl Lieferungen",
            "Lieferantenbezeichnung": "Lieferant",
            "Liefertreue (Ja/Nein)": "Liefertreue"
        },
        color_discrete_map={"Ja": "#1976D2", "Nein": "#63B2EE"}
    )

    liefertreue_barchart.update_layout(barmode="stack", plot_bgcolor="rgba(0,0,0,0)")
    liefertreue_barchart.update_traces(textposition="inside")
    col2.plotly_chart(liefertreue_barchart, use_container_width=True)

    # 2. Mengenabweichung nach Lieferant
    top_10_mengeabweichung = (
        filtered_supplier_data.groupby("Lieferantenbezeichnung")["Mengenabweichung"]
        .sum()
        .abs()
        .nlargest(10)
        .reset_index()
        .sort_values("Mengenabweichung", ascending=False)
    )

    mengeabweichung_bar = px.bar(
        top_10_mengeabweichung,
        x="Lieferantenbezeichnung",
        y="Mengenabweichung",
        title="Top 10 Lieferanten mit Mengenabweichung (Ist - Soll)",
        text="Mengenabweichung",
        color_discrete_sequence=["#1976D2"],
        labels={"Mengenabweichung": "Mengenabweichung", "Lieferantenbezeichnung": "Lieferant"}
    )
    col3.plotly_chart(mengeabweichung_bar, use_container_width=True)

    # Lieferantentabelle erstellen
    supplier_table = filtered_supplier_data[[
        "Lieferantennummer",
        "Lieferantenbezeichnung",
        "Land",
        "Lieferscheinnummer",
        "Materialnummer",
        "Bestelldatum",
        "Lieferdatum (Soll)",
        "Wareneingangsdatum (WE)",
        "Soll-Menge",
        "WE-Menge",
        "Verspätung (Tage)"
    ]]

    # CSV-Download für Lieferantentabelle
    def convert_df_to_csv(df):
        return df.to_csv(index=False).encode("utf-8")

    supplier_csv = convert_df_to_csv(supplier_table)

    st.download_button(
        label="Tabelle als CSV herunterladen",
        data=supplier_csv,
        file_name="lieferantendaten.csv",
        mime="text/csv"
    )
    
    # Tabelle anzeigen
    st.markdown(f"### Gefilterte Daten ({len(supplier_table)} Datensätze)")
    st.dataframe(supplier_table.reset_index(drop=True), height=300, use_container_width=True)

# Tab 2: Analyse Material
with tabs[2]:
    st.title("Analyse Material")

    # Materialtabelle erstellen
    material_risks = (
        filtered_supplier_data.groupby(["Materialnummer", "Materialbezeichnung"])
        .agg({
            "Mengenabweichung": lambda x: (x != 0).sum(),
            "Liefertreue (Ja/Nein)": lambda x: (x == "Nein").sum()
        })
        .rename(columns={
            "Mengenabweichung": "Anzahl Mengenabweichungen",
            "Liefertreue (Ja/Nein)": "Anzahl Verspätungen"
        })
        .reset_index()
    )
    
   # Diagramme zur Visualisierung
    st.markdown("### Diagramme zur Analyse")
    col1, col2 = st.columns(2)

    top_10_delays = material_risks.sort_values("Anzahl Verspätungen", ascending=False).head(10)
    top_10_deviations = material_risks.sort_values("Anzahl Mengenabweichungen", ascending=False).head(10)
    # Diagramm: Anzahl Verspätungen nach Materialnummer
    delays_bar = px.bar(
        top_10_delays,
        x="Materialnummer",
        y="Anzahl Verspätungen",
        text="Anzahl Verspätungen",
        title="Top 10 Materialien nach Anzahl Verspätungen",
        labels={"Materialnummer": "Materialnummer", "Anzahl Verspätungen": "Anzahl Verspätungen"}
    )
    delays_bar.update_traces(textposition="outside")
    col1.plotly_chart(delays_bar, use_container_width=True)

    # Diagramm: Anzahl Mengenabweichungen nach Materialnummer
    deviation_bar = px.bar(
        top_10_deviations,
        x="Materialnummer",
        y="Anzahl Mengenabweichungen",
        text="Anzahl Mengenabweichungen",
        title="Top 10 Materialien nach Anzahl Mengenabweichungen",
        labels={"Materialnummer": "Materialnummer", "Anzahl Mengenabweichungen": "Anzahl Mengenabweichungen"}
    )
    deviation_bar.update_traces(textposition="outside")
    col2.plotly_chart(deviation_bar, use_container_width=True)
    
    # CSV-Download für Materialtabelle
    material_csv = convert_df_to_csv(material_risks)

    st.download_button(
        label="Tabelle als CSV herunterladen",
        data=material_csv,
        file_name="materialdaten.csv",
        mime="text/csv"
    )
    
    # Tabelle anzeigen
    st.markdown(f"### Übersicht Materialien ({len(material_risks)} Datensätze)")
    st.dataframe(material_risks, height=300, use_container_width=True)

# Tab 3: PDF-Report
with tabs[3]:
    st.title("PDF-Report generieren")
    
    # Spaltenauswahl für den Export
    st.markdown("### Wähle Spalten für den PDF-Export:")
    selected_columns = st.multiselect(
        "Spalten auswählen:", options=list(df_cleaned.columns), default=list(df_cleaned.columns)
    )

    # Spaltenauswahl für Sortierung
    st.markdown("### Wähle Spalte für die Sortierung:")
    sort_column = st.selectbox("Sortieren nach:", options=selected_columns)

    # Sortierreihenfolge festlegen
    sort_ascending = st.checkbox("Aufsteigend sortieren", value=True)

    # PDF-Generierung
    def generate_pdf(dataframe, columns, sort_column, ascending):
        # Daten sortieren
        sorted_dataframe = dataframe.sort_values(by=sort_column, ascending=ascending)

        # PDF erstellen
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Report", ln=True, align="C")

        # Hinzufügen der Tabelle
        pdf.set_font("Arial", size=10)
        for col in columns:
            pdf.cell(40, 10, txt=col, border=1)
        pdf.ln()
        for _, row in sorted_dataframe[columns].iterrows():
            for col in columns:
                pdf.cell(40, 10, txt=str(row[col]), border=1)
            pdf.ln()

        return pdf

    # PDF generieren und herunterladen
    if st.button("PDF-Report generieren"):
        if selected_columns:
            pdf = generate_pdf(filtered_df, selected_columns, sort_column, sort_ascending)
            pdf_output_path = "report.pdf"
            pdf.output(pdf_output_path)
            with open(pdf_output_path, "rb") as pdf_file:
                st.download_button(
                    label="PDF herunterladen",
                    data=pdf_file,
                    file_name="report.pdf",
                    mime="application/pdf"
                )
        else:
            st.warning("Bitte wähle mindestens eine Spalte aus.")
   
# Tab 4: Datenqualität
with tabs[4]:
    st.title("Datenqualität")

    # Datenqualitätsübersicht
    st.markdown("### Datenqualitätsübersicht")
    # Kennzahl für Duplikate
    col1, col2 = st.columns(2)
    col1.metric(label="Anzahl Duplikate", value=duplicates_count)
    col2.metric(label="Fehlende Werte (Gesamt)", value=missing_values_count.sum())

    # Fehlende Werte pro Spalte als Tabelle und/oder Diagramm
    st.markdown("### Fehlende Werte pro Spalte (Prozent)")
    if missing_percentages.sum() > 0:
        # Tabelle und Diagramm kombiniert
        st.markdown("#### Tabelle:")
        st.table(missing_percentages.reset_index().rename(columns={"index": "Spalte", 0: "Fehlende Werte (%)"}))

        st.markdown("#### Diagramm:")
        st.bar_chart(missing_percentages)
    else:
        st.info("Es gibt keine fehlenden Werte in den Daten.")
    
    # Duplikate extrahieren
    df_duplicate_head = df_duplicate_data.head()  # Erste 5 Zeilen der Duplikate

    # Hinweis zur Datenverarbeitung
    st.markdown("### Hinweis")
    st.info("Die Daten wurden bereinigt, wobei Duplikate entfernt wurden. Diese Duplikate könnten jedoch in den Quellsystemen noch vorhanden sein und sollten dort weiter analysiert werden. Unten finden Sie eine Übersicht der identifizierten Duplikate. Diese können bei Bedarf als CSV-Datei heruntergeladen werden.") 

    st.markdown("### Vorschau - Daten Duplikate")
    if not df_duplicate_data.empty:
        st.dataframe(df_duplicate_head, height=300, use_container_width=True)

        # Funktion zum Herunterladen der Duplikate als CSV
        def convert_df_to_csv(df):
            return df.to_csv(index=False).encode('utf-8')

        csv_data = convert_df_to_csv(df_duplicate_data)

        st.download_button(
            label="Duplikate als CSV herunterladen",
            data=csv_data,
            file_name="duplikate.csv",
            mime="text/csv",
        )
    else:
        st.info("Es wurden keine Duplikate in den Daten gefunden.")
        
# Tab 5: Details Datenquelle
with tabs[5]:
    
    # Datenquelle Informationen
    data_source_description = """
    Die Daten-Logik für diese Analyse stammen aus dem **SAP-basierten Logistiksystem Automotive Supply**. 
    Sie wurden zuvor aus den relevanten Tabellen (z. B. **EKKO**, **EKPO**, **EKBE**, **LIKP**, **LIPS**, **MSEG**, **MARA**, **MARC**, **MAKT**, **LFA1**) extrahiert. 
    Diese Tabellen enthalten Informationen zu Bestellungen, Lieferungen, Material- und Lieferantenstammdaten, die essenziell für die Untersuchung der Liefertermintreue sind.
    """

    data_source_metadata = {
        "Dateiname": os.path.basename(file_path),
        "Letzte Bearbeitung": pd.to_datetime(os.path.getmtime(file_path), unit='s').strftime("%Y-%m-%d %H:%M:%S")
    }
    
    st.title("Datenquelle")
    
    # Beschreibung der Datenlogik
    st.markdown("### Beschreibung der Datenlogik")
    st.markdown(data_source_description)

    # Technische Daten zur Quelle
    st.markdown("### Technische Informationen")
    st.write("**Dateiname:**", data_source_metadata["Dateiname"])
    st.write("**Letzte Bearbeitung:**", data_source_metadata["Letzte Bearbeitung"])

    # Beispielhafte Tabellen aus dem SAP-System
    st.markdown("### Beispielhafte Tabellen aus dem SAP-System")
    st.write("""
    - **EKKO**: Kopfdaten der Bestellungen  
    - **EKPO**: Positionsdaten der Bestellungen  
    - **EKBE**: Bestellhistorie (Wareneingänge, Rechnungen, etc.)  
    - **LIKP**: Kopfdaten der Lieferungen  
    - **LIPS**: Positionsdaten der Lieferungen  
    - **MSEG**: Bewegungsdaten (Materialbewegungen)  
    - **MARA**: Materialstammdaten (allgemeine Daten)  
    - **MARC**: Materialstammdaten (werkspezifisch)  
    - **MAKT**: Materialbeschreibungen  
    - **LFA1**: Lieferantenstammdaten  
    """)

    # Hinweis zur Datenverarbeitung
    st.markdown("### Hinweis")
    st.info("Die Daten wurden für diese Analyse bereinigt und standardisiert, um eine konsistente Untersuchung der Liefertermintreue zu ermöglichen.")
    
# Tab 6: Kontakt
with tabs[6]:
    st.title("Kontakt")
    st.markdown("""
        **Kontaktperson:**  
        Product Owner: Priscila Strömsdörfer  
        E-Mail: ps178@hdm-stuttgart.de  
        Telefon: +49 1778026982  
    """)
    
    st.markdown("""
        **Support:**  
        IT Hotline: 0800 123456  
        08:00 - 18:00 Uhr
    """)  