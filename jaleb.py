import pandas as pd
import numpy as np
import scipy.stats as stats
from statsmodels.stats.multitest import multipletests
import os
import sys

titleStyle = '\033[93m'  # Amarillo
normalStyle = "\033[0m"  # Restaurar el estilo normal

sys.stdout.write(
    titleStyle + "\n ~~~~~~~~~~~~Jaleb (Cuniculus paca): Chi-square test of independence at multilevel~~~~~~~~~~~~" +
    "\n This program creates automatically the contingency tables for each variable" +
    "\n Each contingency table is applied a chi-square test or Fisher test" +
    "\n Bonferroni correction is applied when necessary." +
    "\n Yates's correction for continuity is applied when necessary." +
    "\n Python and statistics: automation of data analysis. (2024) Moo-Millan Joel I." +
    "\n Please cite this program as:" +
    "\n Moo-Millan Joel I. (2024). Python and statistics: automation of data analysis." +
    "\n https://github.com/JOEL-ISRAEL-MOO-MILLAN/jaleb.py.git" +
    "\n ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + 
    normalStyle + "\n" + "\n")


def chisq_colum(df, col_indices):
    vertebrate_col = df.columns[col_indices[0]]
    negative_col = df.columns[col_indices[1]]
    positive_col = df.columns[col_indices[2]]
    
    results = []
    contingency_tables = []  
    
    # Recorrer cada vertebrado
    for vertebrate_name in df[vertebrate_col].unique():
        specific_vertebrate = df[df[vertebrate_col] == vertebrate_name]
        non_specific_vertebrate = df[df[vertebrate_col] != vertebrate_name]
        non_specific_summarized = pd.DataFrame({
            vertebrate_col: [f"Non-{vertebrate_name}"],
            negative_col: [non_specific_vertebrate[negative_col].sum()],
            positive_col: [non_specific_vertebrate[positive_col].sum()]
        })

        combined_data = pd.concat([specific_vertebrate, non_specific_summarized], ignore_index=True)

        contingency_table = combined_data[[vertebrate_col, negative_col, positive_col]]
        contingency_tables.append(contingency_table)  # Guardar la tabla de contingencia
        
        contingency_table_np = contingency_table[[negative_col, positive_col]].to_numpy()

        # Calcular valores esperados
        chi2, p, dof, expected = stats.chi2_contingency(contingency_table_np, correction=False)

        # Determinar si se aplica Fisher o Chi-cuadrado
        if np.any(expected < 5):
            test_result = stats.fisher_exact(contingency_table_np)
            results.append({
                vertebrate_col: vertebrate_name,
                'Test': "Fisher",
                'p-value': test_result[1],
                'conf_int': np.nan,
                'odds_ratio': test_result[0]
            })
        else:
            # Aplicar corrección de Yates si la tabla es 2x2
            if contingency_table_np.shape == (2, 2):
                test_result = stats.chi2_contingency(contingency_table_np, correction=True)
            else:
                test_result = stats.chi2_contingency(contingency_table_np, correction=False)
            results.append({
                vertebrate_col: vertebrate_name,
                'Test': "Chi-squared",
                'p-value': test_result[1],
                'Chi square': test_result[0],
                'df': test_result[2]
            })

    # Extraer los valores p para la corrección de Bonferroni
    p_values = [result['p-value'] for result in results]
    adjusted_p_values = multipletests(p_values, method='bonferroni')[1]

    # Crear una tabla con los resultados y redondear a 4 dígitos
    results_table = pd.DataFrame(results)
    results_table['Adjusted p-value'] = np.round(adjusted_p_values, 4)

    # Añadir la columna de significancia
    results_table['Signif'] = np.where(results_table['Adjusted p-value'] < 0.001, '***',
                                       np.where(results_table['Adjusted p-value'] < 0.01, '**',
                                                np.where(results_table['Adjusted p-value'] < 0.05, '*',
                                                         'NS')))
    print(results_table)
    print("\ngithub:https://github.com/JOEL-ISRAEL-MOO-MILLAN\nDesigned by Moo-Millan Joel Israel\n")  
    print("\nSignif. codes: ***= 0.0001; **=0.001; *=0.01 ; 0.05=NS; 0.1=NS\n")

    # Reordenar columnas y manejar la presencia o ausencia de 'conf_int' y 'odds_ratio'
    columns_to_include = [vertebrate_col, 'Test', 'p-value', 'Adjusted p-value', 'Signif']
    if 'Chi square' in results_table.columns:
        columns_to_include.insert(4, 'Chi square')
        columns_to_include.insert(5, 'df')
    if 'conf_int' in results_table.columns:
        columns_to_include.append('conf_int')
    if 'odds_ratio' in results_table.columns:
        columns_to_include.append('odds_ratio')
        
    results_table = results_table[columns_to_include]

    # Añadir líneas de separación entre tablas
    blank_line = pd.DataFrame([[''] * len(contingency_tables[0].columns)] * 2, columns=contingency_tables[0].columns)
    separated_tables = []

    for table in contingency_tables:
        table = table[[vertebrate_col, negative_col, positive_col]]  # Reorder columns
        separated_tables.append(table)
        separated_tables.append(blank_line)  # Añadir líneas en blanco entre tablas

    combined_contingency_tables = pd.concat(separated_tables, ignore_index=True)

    return results_table, combined_contingency_tables

def main():
    input_file = input("Type the name of the excel file (without extension): ")
    input_file += ".xlsx"
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return

    if df.shape[1] < 3:
        print("El DataFrame debe tener al menos 3 columnas.")
        return
    
    print("Las columnas disponibles en el archivo son:")
    for i, col in enumerate(df.columns):
        print(f"{i+1}. {col}")

    # Primero, pedir la columna para la variable categórica
    col_index = int(input("Type the number which corresponds to the Categorical variable: ")) - 1
    col_indices = [col_index]

    # Luego, pedir las columnas para las variables numéricas
    for i in range(2):  # Pedimos dos columnas numéricas
        col_index = int(input(f"Type the number which corresponds to Numeric variable {i+1}: ")) - 1
        col_indices.append(col_index)

    output_file = input("Type the name of the table of results (without extension): ") or "Chis-square results"
    output_file += ".xlsx"
    contingency_file = input("Type the name of the file of the contigency table created (without extension): ") or "Contingency_tables"
    contingency_file += ".xlsx"
    
    try:
        results_table, combined_contingency_tables = chisq_colum(df, col_indices)
        
        # Guardar los resultados y ajustar el ancho de las columnas
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            results_table.to_excel(writer, index=False, sheet_name='Results')
            worksheet = writer.sheets['Results']
            for i, col in enumerate(results_table.columns):
                column_length = max(results_table[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_length)
        print(f"Resultados guardados en {output_file}")

        # Guardar las tablas de contingencia y ajustar el ancho de las columnas
        with pd.ExcelWriter(contingency_file, engine='xlsxwriter') as writer:
            combined_contingency_tables.to_excel(writer, sheet_name='Contingency_Tables', index=False)
            worksheet = writer.sheets['Contingency_Tables']
            for i, col in enumerate(combined_contingency_tables.columns):
                column_length = max(combined_contingency_tables[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_length)
        print(f"Tablas de contingencia guardadas en {contingency_file}")
        
    except Exception as e:
        print(f"Error durante el cálculo o guardado: {e}")

if __name__ == "__main__":
    main()
    
input("Press Enter to exit")
