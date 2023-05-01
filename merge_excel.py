import pandas as pd

def merge_excel_files(file1, file2):
    # Lire les fichiers Excel
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    
    # Fusionner les deux DataFrames
    merged_df = pd.concat([df1, df2], ignore_index=True)
    
    return merged_df


def remove_duplicates(df, column_name):
    # Supprimer les doublons de numéros de téléphone
    df.drop_duplicates(subset=column_name, keep='first', inplace=True)
    
    return df


def remove_numbers_by_prefix(df, column_name, prefix):
    # Filtrer les numéros de téléphone qui ne commencent pas par le préfixe
    df_filtered = df[~df[column_name].astype(str).str.startswith(prefix)]
    
    return df_filtered


def save_to_excel(df, output_file):
    # Trier les données par numéro de téléphone
    df.sort_values(by='Numéro de téléphone', inplace=True)
    
    # Écrire le résultat dans un nouveau fichier Excel
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    file1 = "chemin/vers/le/fichier1.xlsx"
    file2 = "chemin/vers/le/fichier2.xlsx"
    output_file = "chemin/vers/le/fichier_fusionné.xlsx"
    
    merged_df = merge_excel_files(file1, file2)
    unique_df = remove_duplicates(merged_df, 'Numéro de téléphone')
    save_to_excel(unique_df, output_file)
