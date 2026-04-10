# ============================================================
# PROJET DATA ANALYST - SUPERSTORE FINANCE DASHBOARD
# Script de nettoyage + calcul des KPIs pour Power BI
# ============================================================

import pandas as pd
import os

# ── 1. CHARGEMENT ──────────────────────────────────────────
print("Chargement du fichier...")
df = pd.read_csv("Sample - Superstore.csv", encoding="latin-1")
print(f"  {len(df)} lignes chargées, {df.shape[1]} colonnes")


# ── 2. NETTOYAGE ───────────────────────────────────────────
print("\nNettoyage...")

# Conversion des dates
df["Order Date"] = pd.to_datetime(df["Order Date"])
df["Ship Date"]  = pd.to_datetime(df["Ship Date"])

# Colonnes de temps utiles
df["Year"]  = df["Order Date"].dt.year
df["Month"] = df["Order Date"].dt.month
df["Month Name"] = df["Order Date"].dt.strftime("%B")
df["Quarter"] = "Q" + df["Order Date"].dt.quarter.astype(str)

# Nettoyage des noms de colonnes (espaces → underscores)
df.columns = df.columns.str.strip().str.replace(" ", "_").str.replace("-", "_")

# Suppression des colonnes inutiles pour Power BI
df.drop(columns=["Row_ID", "Country"], inplace=True)

print(f"  Colonnes après nettoyage : {list(df.columns)}")


# ── 3. CALCUL DES KPIs ─────────────────────────────────────
print("\nCalcul des KPIs...")

# Marge brute (%)
df["Profit_Margin_%"] = (df["Profit"] / df["Sales"] * 100).round(2)

# Chiffre d'affaires après remise
df["Sales_After_Discount"] = (df["Sales"] * (1 - df["Discount"])).round(2)

# Durée de livraison (jours)
df["Delivery_Days"] = (df["Ship_Date"] - df["Order_Date"]).dt.days


# ── 4. TABLES AGRÉGÉES ─────────────────────────────────────
print("\nGénération des tables agrégées...")

# -- KPIs annuels (pour courbe de croissance YoY)
kpis_annuels = df.groupby("Year").agg(
    CA_Total=("Sales", "sum"),
    Profit_Total=("Profit", "sum"),
    Nb_Commandes=("Order_ID", "nunique"),
    Nb_Clients=("Customer_ID", "nunique"),
).reset_index()

kpis_annuels["Marge_%"] = (kpis_annuels["Profit_Total"] / kpis_annuels["CA_Total"] * 100).round(2)
kpis_annuels["Croissance_CA_%"] = kpis_annuels["CA_Total"].pct_change().mul(100).round(2)

# -- Performance par Catégorie
perf_categorie = df.groupby("Category").agg(
    CA=("Sales", "sum"),
    Profit=("Profit", "sum"),
    Quantite=("Quantity", "sum"),
    Nb_Commandes=("Order_ID", "nunique"),
).reset_index()
perf_categorie["Marge_%"] = (perf_categorie["Profit"] / perf_categorie["CA"] * 100).round(2)
perf_categorie.sort_values("CA", ascending=False, inplace=True)

# -- Performance par Sous-Catégorie
perf_sous_cat = df.groupby(["Category", "Sub_Category"]).agg(
    CA=("Sales", "sum"),
    Profit=("Profit", "sum"),
).reset_index()
perf_sous_cat["Marge_%"] = (perf_sous_cat["Profit"] / perf_sous_cat["CA"] * 100).round(2)
perf_sous_cat.sort_values("Profit", ascending=True, inplace=True)

# -- Performance par Région + État
perf_region = df.groupby(["Region", "State"]).agg(
    CA=("Sales", "sum"),
    Profit=("Profit", "sum"),
    Nb_Clients=("Customer_ID", "nunique"),
).reset_index()
perf_region["Marge_%"] = (perf_region["Profit"] / perf_region["CA"] * 100).round(2)

# -- Performance par Segment client
perf_segment = df.groupby("Segment").agg(
    CA=("Sales", "sum"),
    Profit=("Profit", "sum"),
    Nb_Clients=("Customer_ID", "nunique"),
).reset_index()
perf_segment["Marge_%"] = (perf_segment["Profit"] / perf_segment["CA"] * 100).round(2)

# -- Évolution mensuelle
perf_mensuelle = df.groupby(["Year", "Month", "Month_Name"]).agg(
    CA=("Sales", "sum"),
    Profit=("Profit", "sum"),
).reset_index().sort_values(["Year", "Month"])


# ── 5. EXPORT CSV ──────────────────────────────────────────
print("\nExport des fichiers CSV...")

os.makedirs("powerbi_data", exist_ok=True)

exports = {
    "powerbi_data/01_donnees_principales.csv": df,
    "powerbi_data/02_kpis_annuels.csv":        kpis_annuels,
    "powerbi_data/03_perf_categorie.csv":       perf_categorie,
    "powerbi_data/04_perf_sous_categorie.csv":  perf_sous_cat,
    "powerbi_data/05_perf_region.csv":          perf_region,
    "powerbi_data/06_perf_segment.csv":         perf_segment,
    "powerbi_data/07_perf_mensuelle.csv":       perf_mensuelle,
}

for path, table in exports.items():
    table.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"  ✓ {path}  ({len(table)} lignes)")


# ── 6. RÉSUMÉ DES INSIGHTS ─────────────────────────────────
print("\n" + "="*55)
print("RÉSUMÉ EXÉCUTIF — À mettre dans ton rapport Power BI")
print("="*55)

ca_total   = df["Sales"].sum()
profit_tot = df["Profit"].sum()
marge_moy  = profit_tot / ca_total * 100

print(f"  CA Total         : ${ca_total:,.0f}")
print(f"  Profit Total     : ${profit_tot:,.0f}")
print(f"  Marge globale    : {marge_moy:.1f}%")

print("\n  Top catégories par CA :")
for _, row in perf_categorie.iterrows():
    print(f"    {row['Category']:20s}  CA=${row['CA']:>10,.0f}  Marge={row['Marge_%']:>6.1f}%")

print("\n  Sous-catégories les moins rentables :")
worst = perf_sous_cat.head(3)
for _, row in worst.iterrows():
    print(f"    {row['Sub_Category']:20s}  Profit=${row['Profit']:>9,.0f}  Marge={row['Marge_%']:>6.1f}%")

print("\n  Croissance CA année par année :")
for _, row in kpis_annuels.iterrows():
    croiss = f"{row['Croissance_CA_%']:+.1f}%" if pd.notna(row["Croissance_CA_%"]) else "—"
    print(f"    {int(row['Year'])}  CA=${row['CA_Total']:>10,.0f}  Croissance={croiss}")

print("\n  Fichiers prêts dans le dossier /powerbi_data/")
print("  → Importe-les dans Power BI Desktop via 'Obtenir des données > CSV'")
print("="*55)
