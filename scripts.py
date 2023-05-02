import pandas as pd
import math


def merge_excel_files(file1, file2):
    # Lire les fichiers Excel
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    # Fusionner les deux DataFrames
    merged_df = pd.concat([df1, df2], ignore_index=True)
    return merged_df


def read_column_from_excel(file, column_name):
    # Lire le fichier Excel
    df = pd.read_excel(file)
    
    # Extraire les données de la colonne spécifiée
    column_data = df[column_name]
    
    return column_data


def remove_duplicates(series): #(df, column_name)
    # Supprimer les doublons de numéros de téléphone
    # df.drop_duplicates(subset=column_name, keep='first', inplace=True)
    # return df
    # Supprimer les doublons de la Series
    unique_series = series.drop_duplicates(keep='first')

    return unique_series


def add_string_to_column(df, column_name, string_to_add):
    # Appliquer la modification sur la colonne spécifiée
    df[column_name] = df[column_name].apply(lambda x: f"{string_to_add}{x}")

    # Renvoyer le DataFrame modifié
    return df


def excel_to_csv(input_file, output_file, sheet_name=0):
    # Charger le fichier Excel en utilisant pandas
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # Enregistrer le DataFrame dans un fichier CSV
    df.to_csv(output_file, index=False)

def remove_numbers_by_prefix(df, column_name, prefix):
    # Filtrer les numéros de téléphone qui ne commencent pas par le préfixe
    df_filtered = df[~df[column_name].astype(str).str.startswith(prefix)]
    
    return df_filtered


def count_duplicates_in_excel(file, column_name):
    # Lire le fichier Excel
    df = pd.read_excel(file)

    # Compter les doublons dans la colonne spécifiée
    duplicate_count = df.duplicated(subset=column_name, keep='first').sum()

    return duplicate_count

def save_to_file(data, output_file, file_format='excel'):
    # Vérifier si les données sont une Series ou un DataFrame
    if isinstance(data, (pd.Series, pd.DataFrame)):
        # Enregistrer les données dans un fichier Excel ou CSV
        if file_format.lower() == 'excel':
            data.to_excel(output_file, index=False)
        elif file_format.lower() == 'csv':
            data.to_csv(output_file, index=False)
        else:
            raise ValueError("Le format de fichier spécifié n'est pas pris en charge. Utilisez 'excel' ou 'csv'.")
    else:
        raise ValueError("Les données fournies ne sont ni une Series ni un DataFrame.")
    
    
def split_series_to_excel_files(series, lines_per_file, output_prefix):
    total_rows = len(series)
    total_files = math.ceil(total_rows / lines_per_file)

    for i in range(total_files):
        start_index = i * lines_per_file
        end_index = (i + 1) * lines_per_file

        # Extraire un sous-ensemble de la Series
        subset = series[start_index:end_index]

        # Créer un nouveau DataFrame à partir de la Series
        df = pd.DataFrame(subset).reset_index(drop=True)

        # Enregistrer le DataFrame dans un fichier Excel
        output_file = f"{output_prefix}_part_{i + 1}.xlsx"
        df.to_excel(output_file, index=False)

    return total_files



