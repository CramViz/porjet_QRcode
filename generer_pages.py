import pandas as pd
import json
import re
from pathlib import Path

# === NOMS DE FICHIERS ===
EXCEL_FILE = "eleves.xlsx"  # adapte EXACTEMENT au nom de ton fichier
TEMPLATE_FILE = "presence_eleve_firebase_complet.html"    # celui que tu as déjà
OUTPUT_FILE = "presence_eleve_firebase_complet.html"      # on écrase le même fichier

def charger_listes_depuis_excel(path_excel):
    xls = pd.ExcelFile(path_excel)
    listes = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(path_excel, sheet_name=sheet_name, header=None)
        df = df.dropna(how="all").fillna("")
        fullnames = []
        for _, row in df.iterrows():
            nom = str(row[0]).strip()
            prenom = str(row[1]).strip()
            if nom or prenom:
                fullnames.append((nom + " " + prenom).strip())
        listes[sheet_name] = fullnames
    return listes

def construire_js_listes(listes):
    parts = []
    for sheet_name, names in listes.items():
        js_array = json.dumps(names, ensure_ascii=False)
        parts.append(f'      "{sheet_name}": {js_array}')
    return "{\n" + ",\n".join(parts) + "\n    }"

def construire_js_alias(listes):
    alias_parts = []
    for sheet_name in listes.keys():
        alias = ""
        for token in sheet_name.split():
            if any(ch.isdigit() for ch in token):
                alias = token
                break
        alias_parts.append(f'      "{alias or sheet_name}": "{sheet_name}"')
    return "{\n" + ",\n".join(alias_parts) + "\n    }"

def construire_options_html(listes):
    return "\n".join(
        f'      <option value="{name}">{name}</option>'
        for name in listes.keys()
    )

def main():
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        print(f"❌ Fichier Excel introuvable : {EXCEL_FILE}")
        return

    listes = charger_listes_depuis_excel(excel_path)
    js_listes = construire_js_listes(listes)
    js_alias = construire_js_alias(listes)
    options_html = construire_options_html(listes)

    template = Path(TEMPLATE_FILE).read_text(encoding="utf-8")

    # Remplace le bloc JS LISTES
    template = re.sub(
        r"const LISTES = \{[\s\S]*?};",
        f"const LISTES = {js_listes};",
        template,
    )

    # Remplace le bloc JS ALIAS
    template = re.sub(
        r"const ALIAS_CRENEAU = \{[\s\S]*?};",
        f"const ALIAS_CRENEAU = {js_alias};",
        template,
    )

    # Remplace les options du <select> des créneaux
    template = re.sub(
        r"<select id=\"creneau\">[\s\S]*?</select>",
        '<select id="creneau">\n' + options_html + '\n    </select>',
        template,
    )

    Path(OUTPUT_FILE).write_text(template, encoding="utf-8")
    print("✅ Nouveau fichier généré :", OUTPUT_FILE)

if __name__ == "__main__":
    main()
