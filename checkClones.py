import pandas as pd

# -------------------------
# Βήμα 1: Φόρτωση των δύο sheets
# -------------------------
excel_file = "contracts.xlsx"

# df_data: το πρώτο sheet με τα δεδομένα
# df_mapping: το δεύτερο sheet με το mapping
df_data = pd.read_excel(excel_file, sheet_name="Data")
df_mapping = pd.read_excel(excel_file, sheet_name="Mapping")

# -------------------------
# Βήμα 2: Δημιουργία side-by-side για όλα τα ζευγάρια
# -------------------------
final_side_by_side = []

for _, row in df_mapping.iterrows():
    old_contract = row["SAMMELN_OLD"]
    new_contract = row["SAMMELN_NEW"]
    
    # Φιλτράρουμε γραμμές για κάθε συμβόλαιο
    df_old = df_data[df_data["SAMMELN"] == old_contract].copy()
    df_new = df_data[df_data["SAMMELN"] == new_contract].copy()
    
    # Προσθέτουμε __id για σωστό side-by-side merge
    df_old["__id"] = df_old.index
    df_new["__id"] = df_new.index
    
    merged = df_old.merge(
        df_new,
        how="outer",
        on="__id",
        suffixes=("_old", "_new"),
        indicator=True
    )
    
    # Προσθέτουμε στήλες με τα συμβόλαια για αναφορά
    merged["SAMMELN_OLD"] = old_contract
    merged["SAMMELN_NEW"] = new_contract
    
    # Αφαιρούμε __id
    merged.drop(columns="__id", inplace=True)
    
    final_side_by_side.append(merged)

# -------------------------
# Βήμα 3: Συνένωση όλων των ζευγαριών
# -------------------------
final_df = pd.concat(final_side_by_side, ignore_index=True)

# -------------------------
# Βήμα 4: Αποθήκευση σε τρίτο sheet του Excel
# -------------------------
with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a") as writer:
    final_df.to_excel(writer, sheet_name="Side_by_Side", index=False)