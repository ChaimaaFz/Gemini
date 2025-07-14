import sys
import io
import os
import fitz  # PyMuPDF
import pandas as pd
import google.generativeai as genai
from io import StringIO
# --- Fonction extraction texte depuis PDF ---
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    return full_text


# --- Fonction lecture matrice Excel ---
def read_excel_matrix(excel_path):
    df = pd.read_excel(excel_path)
    return df

def create_prompt_general(matrice_json, texte_pdf):
    prompt = f"""
Tu es un assistant intelligent spécialisé dans l’analyse de rapports statistiques.

Voici une matrice de données incomplète extraite d’un fichier Excel. Elle contient des lignes avec différents indicateurs (exemple : taux de chômage...), et des colonnes décrivant des dimensions comme l'année, la région, etc.

Chaque ligne représente un cas unique à compléter.

Voici la matrice :

{matrice_json}

Et voici le contenu textuel extrait de plusieurs fichiers PDF de rapports annuels ou trimestriels, non structurés :

{texte_pdf}

Ta tâche est de :
1. Comprendre la structure de la matrice (colonnes et lignes).
2. Rechercher dans le texte les données correspondant à chaque ligne.
3. Remplir chaque cellule vide si la valeur peut être trouvée dans le texte.
4. Si une valeur est absente du texte ou incertaine, laisse la cellule vide.

Réponds uniquement au format CSV, avec la même structure que la matrice d’origine, remplie avec les valeurs complétées si disponibles.
**Utilise le point-virgule (;) comme séparateur de colonnes dans le CSV.**
Ne mets aucune explication, uniquement le contenu CSV.
"""
    return prompt


# --- Fonction appel API Gemini ---

def call_gemini_api(prompt):

    genai.configure(api_key="AIzaSyBLQfllqi4UiLpjt6HUqUCQ0iBMykvWelc")  # Remplace par ta vraie clé

    model = genai.GenerativeModel(model_name="gemini-2.5-pro")
    response = model.generate_content(prompt)
    return response.text


def dataframe_to_json_list(df):
    # Convertit chaque ligne du DataFrame en dictionnaire
    return df.fillna("").to_dict(orient="records")

def update_excel_with_results(df_original, df_resultat):
    for i in range(len(df_original)):
        for col in df_original.columns:
            if col in df_resultat.columns:
                val_originale = df_original.at[i, col]
                val_extraite = df_resultat.at[i, col]
                if (pd.isna(val_originale) or val_originale == "") and str(val_extraite).strip() not in ["", "Non trouvé", None]:
                    df_original.at[i, col] = val_extraite
    return df_original

# --- Parseur Gemini robuste ---
def parse_gemini_csv_response(response_text):
    try:
        df = pd.read_csv(io.StringIO(response_text), sep=';', quotechar='"', encoding='utf-8')
        # Nettoyage des colonnes
        df.columns = df.columns.str.strip()
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        return df
    except Exception as e:
        print("Erreur lors du parsing CSV :", e)
        print("Réponse brute :", response_text)
        return pd.DataFrame()
def main(pdf_files, excel_path, output_excel_path):
    # Ton code actuel ici (à adapter un peu)

    # Exemple (simplifié) :


    df = pd.read_excel(excel_path)
    full_text = ""
    for pdf in pdf_files:
        full_text += extract_text_from_pdf(pdf) + "\n"

    matrice_csv = df.to_csv(index=False)
    prompt = create_prompt_general(matrice_csv, full_text)
    response_text = call_gemini_api(prompt)
    df_result = parse_gemini_csv_response(response_text)

    if df_result.empty:
        print("Résultat vide, arrêt du script.")
        sys.exit(1)

    df_result.to_excel(output_excel_path, index=False)
    print(f"Fichier sauvegardé dans {output_excel_path}")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python traitement.py <pdfs_csv> <excel_path> <output_excel_path>")
        sys.exit(1)

    pdfs_csv = sys.argv[1]
    excel_path = sys.argv[2]
    output_excel_path = sys.argv[3]

    pdf_files = pdfs_csv.split(",")

    main(pdf_files, excel_path, output_excel_path)
