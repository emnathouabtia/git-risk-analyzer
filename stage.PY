import git
import re
from collections import defaultdict
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.tree import DecisionTreeClassifier, plot_tree
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report, confusion_matrix
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ===================================
# 1. EXTRACTION DES DONNÉES GIT
# ===================================

repo_path = r"C:\Users\-PC-\Desktop\stage\requests-repo"
repo = git.Repo(repo_path)

bug_keywords = ["fix", "bug", "issue", "error", "patch"]
function_regex = re.compile(r"def (\w+)\(")
function_stats = defaultdict(lambda: {"touched": 0, "bugs": 0, "lines_changed": 0})

print("Analyse des commits en cours...")

for commit in repo.iter_commits(max_count=500):
    parent = commit.parents[0] if commit.parents else git.NULL_TREE
    for diff in commit.diff(parent, create_patch=True):
        if diff.b_path and diff.b_path.endswith(".py"):
            try:
                content = diff.diff.decode(errors="ignore")
            except:
                continue
            touched_functions = set(function_regex.findall(content))
            lines_changed = content.count("\n+")
            for func in touched_functions:
                function_stats[func]["touched"] += 1
                function_stats[func]["lines_changed"] += lines_changed
                if any(keyword in commit.message.lower() for keyword in bug_keywords):
                    function_stats[func]["bugs"] += 1

# ===================================
# 2. PRÉPARATION DU DATASET
# ===================================

data = []
labels = []
func_names = []

for func, stats in function_stats.items():
    X = [stats["touched"], stats["bugs"], stats["lines_changed"]]
    data.append(X)
    func_names.append(func)
    if stats["bugs"] >= 3:
        labels.append(2)
    elif stats["bugs"] >= 1:
        labels.append(1)
    else:
        labels.append(0)

df = pd.DataFrame(data, columns=["touched", "bugs", "lines_changed"])
df["label"] = labels
df["function"] = func_names

print("===== Aperçu du dataset =====")
print(df.head())

# ===================================
# 3. DECISION TREE
# ===================================

X = df[["touched", "bugs", "lines_changed"]]
y = df["label"]

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=42)

clf = DecisionTreeClassifier(max_depth=3, random_state=42)
clf.fit(X_train, y_train)

y_pred = clf.predict(X_test)
print("\n===== Rapport de classification (Decision Tree) =====")
print(classification_report(y_test, y_pred, zero_division=0))

print("\n===== Matrice de confusion (Decision Tree) =====")
print(confusion_matrix(y_test, y_pred))

plt.figure(figsize=(12, 8))
plot_tree(clf, feature_names=["touched", "bugs", "lines_changed"],
          class_names=["Safe", "At Risk", "Risky"], filled=True)
plt.title("Arbre de décision")
plt.tight_layout()
plt.savefig(r"C:\Users\-PC-\Desktop\stage\decision_tree.png", dpi=150)
plt.show()

df["predicted_risk"] = clf.predict(X)

risk_colors = {0: "🟢 Sans risque", 1: "🟡 À risque", 2: "🔴 Risquée"}
print("\n===== Classement final (Decision Tree) =====")
for row in df.itertuples():
    print(f"{row.function}: Touches={row.touched}, Bugs={row.bugs}, "
          f"Lignes={row.lines_changed} -> {risk_colors[row.predicted_risk]}")

# ===================================
# 4. RANDOM FOREST
# ===================================

print("\n===== Random Forest =====")
clf_rf = RandomForestClassifier(n_estimators=100, max_depth=5, random_state=42, class_weight="balanced")
clf_rf.fit(X_train, y_train)

y_pred_rf = clf_rf.predict(X_test)
print("\n--- Rapport de classification (Random Forest) ---")
print(classification_report(y_test, y_pred_rf, zero_division=0))

print("\n--- Matrice de confusion (Random Forest) ---")
print(confusion_matrix(y_test, y_pred_rf))

feature_importance = pd.Series(
    clf_rf.feature_importances_,
    index=["touched", "bugs", "lines_changed"]
).sort_values(ascending=False)
print("\n--- Importance des features ---")
print(feature_importance)

df["predicted_risk_rf"] = clf_rf.predict(X)

# ===================================
# 5. TOP 10 FONCTIONS RISQUÉES
# ===================================

print("\nGénération du graphique top 10...")

top10 = df.nlargest(10, "bugs")
color_map = {0: "#4ade80", 1: "#facc15", 2: "#f87171"}
bar_colors = [color_map[r] for r in top10["predicted_risk_rf"]]

fig, ax = plt.subplots(figsize=(12, 6))
bars = ax.barh(top10["function"], top10["bugs"], color=bar_colors, edgecolor="white", linewidth=0.8)
ax.set_xlabel("Nombre de bugs", fontsize=12)
ax.set_title("Top 10 fonctions les plus risquées", fontsize=14, fontweight="bold")
ax.invert_yaxis()
ax.set_facecolor("#FAFAFA")
fig.patch.set_facecolor("white")

for bar, val in zip(bars, top10["bugs"]):
    ax.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height() / 2,
            str(val), va="center", fontsize=10)

from matplotlib.patches import Patch
legend_elements = [
    Patch(facecolor="#4ade80", label="Sans risque"),
    Patch(facecolor="#facc15", label="À risque"),
    Patch(facecolor="#f87171", label="Risquée"),
]
ax.legend(handles=legend_elements, loc="lower right")
plt.tight_layout()
plt.savefig(r"C:\Users\-PC-\Desktop\stage\top10_risques.png", dpi=150, bbox_inches="tight")
plt.show()
print("Graphique sauvegardé : top10_risques.png")

# ===================================
# 6. EXPORT EXCEL
# ===================================

print("\nExport Excel en cours...")

wb = Workbook()
ws = wb.active
ws.title = "Analyse des Risques"

VIOLET = "7C3AED"
WHITE  = "FFFFFF"

headers = ["Fonction", "Touches", "Bugs", "Lignes modifiées", "Risque (DT)", "Risque (RF)"]
ws.append(headers)

for col_num in range(1, len(headers) + 1):
    cell = ws.cell(row=1, column=col_num)
    cell.fill = PatternFill("solid", start_color=VIOLET)
    cell.font = Font(bold=True, color=WHITE, name="Calibri", size=11)
    cell.alignment = Alignment(horizontal="center", vertical="center")

ws.row_dimensions[1].height = 25

risk_fill = {
    0: PatternFill("solid", start_color="D1FAE5"),
    1: PatternFill("solid", start_color="FEF3C7"),
    2: PatternFill("solid", start_color="FEE2E2"),
}
risk_text = {0: "Sans risque", 1: "À risque", 2: "Risquée"}

for _, row in df.iterrows():
    ws.append([
        row["function"],
        int(row["touched"]),
        int(row["bugs"]),
        int(row["lines_changed"]),
        risk_text[row["predicted_risk"]],
        risk_text[row["predicted_risk_rf"]],
    ])
    row_num = ws.max_row
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=row_num, column=col_num)
        cell.font = Font(name="Calibri", size=10)
        cell.alignment = Alignment(horizontal="center" if col_num > 1 else "left")
        if row["predicted_risk_rf"] != 0:
            cell.fill = risk_fill[row["predicted_risk_rf"]]

col_widths = [30, 10, 8, 18, 16, 16]
for col_num, width in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(col_num)].width = width

ws.freeze_panes = "A2"

ws2 = wb.create_sheet("Résumé")
ws2.append(["Statistique", "Valeur"])
ws2.append(["Total fonctions",  len(df)])
ws2.append(["Sans risque",      len(df[df["predicted_risk_rf"] == 0])])
ws2.append(["À risque",         len(df[df["predicted_risk_rf"] == 1])])
ws2.append(["Risquées",         len(df[df["predicted_risk_rf"] == 2])])
ws2.append(["Total bugs",       int(df["bugs"].sum())])

for row_num in range(1, ws2.max_row + 1):
    ws2.cell(row=row_num, column=1).font = Font(bold=(row_num == 1), name="Calibri")
    ws2.cell(row=row_num, column=2).font = Font(name="Calibri")
    if row_num == 1:
        for col in [1, 2]:
            ws2.cell(1, col).fill = PatternFill("solid", start_color=VIOLET)
            ws2.cell(1, col).font = Font(bold=True, color=WHITE, name="Calibri")

ws2.column_dimensions["A"].width = 25
ws2.column_dimensions["B"].width = 15

excel_path = r"C:\Users\-PC-\Desktop\stage\analyse_risques.xlsx"
wb.save(excel_path)
print(f"Excel sauvegardé : {excel_path}")

# ===================================
# 7. SAUVEGARDE CSV POUR LE DASHBOARD
# ===================================

csv_path = r"C:\Users\-PC-\Desktop\stage\results.csv"

df_export = df[["function", "touched", "bugs", "lines_changed", "predicted_risk", "predicted_risk_rf"]].copy()
df_export.columns = ["function", "touched", "bugs", "lines_changed", "predicted_risk", "predicted_risk_rf"]

print("\nColonnes exportées :", df_export.columns.tolist())
print(df_export.head())

df_export.to_csv(csv_path, index=False)
print(f"\nCSV sauvegardé : {csv_path}")
print("Ouvre maintenant dashboard.html dans ton navigateur !")