from flask import Flask, request, send_from_directory, jsonify, render_template
import os
import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
import shutil
import os
import openpyxl
import zipfile
import ternary
import matplotlib.pyplot as plt
from PIL import Image, ImageTk
import tempfile
import matplotlib.pyplot as plt
plt.switch_backend('Agg')
import gc

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
PLOTS_FOLDER = 'ternary_plots'
ASSETS_FOLDER = 'assets'  # Adicionando a pasta de assets
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/create_excel', methods=['POST'])
def create_excel():
    data = request.json
    soil_code = data['soil_code']
    num_points = data['num_points']
    h_values = data['h_values']
    theta_values = data['theta_values']

    # Criação do DataFrame
    df = pd.DataFrame({
        'SampleID': [soil_code] * num_points,
        'code': [1] * num_points,
        'h': h_values,
        'theta': theta_values
    })

    # Salvando o arquivo Excel
    filepath = os.path.join(DOWNLOAD_FOLDER, 'YourSoil.xlsx')
    df.to_excel(filepath, index=False)

    # Criar SoilFull
    soil_full_path = os.path.join(DOWNLOAD_FOLDER, 'SoilFull.xlsx')
    df.to_excel(soil_full_path, index=False)  # Salvando todos os dados sem alterações

    # Criar SoilComplete
    soil_complete = df[df['h'] >= 30]
    soil_complete_path = os.path.join(DOWNLOAD_FOLDER, 'SoilComplete.xlsx')
    soil_complete.to_excel(soil_complete_path, index=False)

    # Criar SoilSorted
    soil_sorted = pd.DataFrame(columns=['code', 'h', 'theta', 'PT'])

    # Primeiro ponto
    if ((df['h'] >= 0) & (df['h'] <= 1)).any():  # Verifica se existe pelo menos um valor entre 0 e 1
        first_point = df[(df['h'] >= 0) & (df['h'] <= 1)]
        closest_to_zero = first_point.iloc[(first_point['h']).abs().argsort()[:1]][['code', 'h', 'theta']]
        soil_sorted = pd.concat([soil_sorted, closest_to_zero], ignore_index=True)
    else:
        return jsonify({'success': False, 'message': 'No points with h between 0 and 1.'})

    # Segundo ponto
    second_point = df[(df['h'] >= 30) & (df['h'] <= 80)]
    if not second_point.empty:
        second_point = second_point.iloc[(second_point['h'] - 60).abs().argsort()[:1]][['code', 'h', 'theta']]
        soil_sorted = pd.concat([soil_sorted, second_point], ignore_index=True)

    # Terceiro ponto
    third_point = df[(df['h'] >= 250) & (df['h'] <= 500)]
    if not third_point.empty:
        third_point = third_point.iloc[(third_point['h'] - 330).abs().argsort()[:1]][['code', 'h', 'theta']]
        soil_sorted = pd.concat([soil_sorted, third_point], ignore_index=True)

    # Quarto ponto
    fourth_point = df[(df['h'] >= 9000) & (df['h'] <= 18000)]
    if not fourth_point.empty:
        fourth_point = fourth_point.iloc[(fourth_point['h'] - 15000).abs().argsort()[:1]][['code', 'h', 'theta']]
        soil_sorted = pd.concat([soil_sorted, fourth_point], ignore_index=True)

    # Verificação de pontos encontrados
    if len(soil_sorted) < 4:
        return jsonify({'success': False, 'message': 'SoilSorted requires at least 4 points.'})

    # Criar a nova coluna PT
    soil_sorted['PT'] = [soil_sorted['theta'].iloc[0]] + [''] * (len(soil_sorted) - 1)

    # Remover a coluna SampleID
    soil_sorted = soil_sorted.drop(columns=['SampleID'], errors='ignore')

    soil_sorted_path = os.path.join(DOWNLOAD_FOLDER, 'SoilSorted.xlsx')
    soil_sorted.to_excel(soil_sorted_path, index=False)

    # Calculations for Bounds.xlsx
    calculate_bounds2()

    # Generate Parameters.xlsx
    calculate_parameters2()

    return jsonify({'success': True, 'message': 'Excel files created successfully.'})

def calculate_bounds2():
    df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted.xlsx'))

    # Calculate the alpha and T15000 values
    codes = []
    t15000_values = []
    alpha_values = []

    for code in df['code'].unique():
        df_temp = df[df['code'] == code]

        third_theta = df_temp.iloc[2]['theta']
        fourth_theta = df_temp.iloc[3]['theta']
        first_theta = df_temp.iloc[0]['theta']
        h_third_line = df_temp.iloc[2]['h']

        alpha = (((third_theta - (fourth_theta - 0.01)) / (first_theta - (fourth_theta - 0.01))) ** (-1 / 0.5) - 1) ** (
                    1 - 0.5) / h_third_line
        t15000_temp = df_temp.iloc[3]['theta'] - 0.01

        codes.append(code)
        t15000_values.append(t15000_temp)
        alpha_values.append(alpha)

    df_temp = pd.DataFrame({'code': codes, 'T15000': t15000_values, 'alpha': alpha_values})
    df_temp.to_csv(os.path.join(DOWNLOAD_FOLDER,'InitialParam.txt'), header=False, index=False, sep='\t')

    # Calculate bounds
    med0 = df[df['h'] == 0]['theta'].mean()
    med15000 = df[df['h'] >= 9000]['theta'].mean() - 0.01
    min0 = df[df['h'] == 0]['theta'].min()
    max0 = df[df['h'] == 0]['theta'].max()
    max15000 = df[df['h'] >= 9000]['theta'].max()

    data = {
        'Med0': [med0],
        'Med15000': [med15000],
        'Min0': [min0],
        'Max0': [max0],
        'Max15000': [max15000]
    }

    df_result = pd.DataFrame(data)
    df_result.to_excel(os.path.join(DOWNLOAD_FOLDER,'Bounds.xlsx'), index=False)

    # Importar a biblioteca openpyxl
    import openpyxl
    def insert_row_excel():
        # Carregar o arquivo Excel
        wb = openpyxl.load_workbook(os.path.join(DOWNLOAD_FOLDER,"SoilSorted.xlsx"))
        # Selecionar a primeira planilha
        ws = wb.active

        row = 2  # Começando pela linha 2
        while ws.cell(row=row, column=1).value is not None:  # Enquanto houver dados na coluna A
            if ws.cell(row=row, column=1).value != ws.cell(row=row + 1, column=1).value:
                # Se o valor da célula atual for diferente da célula abaixo
                ws.insert_rows(row + 1)  # Inserir uma linha abaixo da linha atual
                row += 1  # Mover para a próxima linha

            row += 1  # Mover para a próxima linha

        # Salvar as alterações no arquivo Excel
        wb.save(os.path.join(DOWNLOAD_FOLDER,"SoilSorted.xlsx"))

    # Chamar a função para inserir linhas
    insert_row_excel()

    # Inicializar uma lista para armazenar os dados transpostos
    transposed_data = []

    # Loop para transpor os dados para cada grupo de dados na planilha original
    for site_id, group in df.groupby("code"):
        # Transpor os dados do Head
        head = group["h"].head(16).reset_index(drop=True)
        matric_potential_values = list(head.combine_first(pd.Series([0]*16)))

        # Transpor os dados do Theta
        theta = group["theta"].head(16).reset_index(drop=True)
        water_content_values = list(theta.combine_first(pd.Series([0]*16)))

        # Adicionar os dados transpostos à lista
        transposed_data.append([site_id, len(group), *matric_potential_values, *water_content_values, group["PT"].iloc[0]])

    # Salvar os dados em um arquivo de texto
    with open(os.path.join(DOWNLOAD_FOLDER,"ret.txt"), "w") as f:
        for line in transposed_data:
            formatted_line = "\t".join(map(lambda x: '{:.3f}'.format(x) if isinstance(x, float) else str(x), line))
            f.write(formatted_line + "\n")

def calculate_parameters2():
    import numpy as np
    from scipy.optimize import curve_fit
    import pandas as pd
       # Define the Van Genuchten model function with four parameters, with par[1] fixed
    def vg4(par, h, par1):
        return np.abs(par[0]) + (par1 - np.abs(par[0])) / (1 + (par[1] * h)**par[2])**(1 - 1 / par[2])

    # Define the function that calculates the sum of squares of the residuals
    def vg4ssq(par, press, theta, par1):
        theta1 = vg4(par, press, par1)
        diff = theta - theta1
        ssq = np.dot(diff, diff)
        return ssq

    # Define a curve fitting function to be used with curve_fit
    def vg4ssq2(press, par0, par2, par3, par1):
        par = [par0, par2, par3]
        return vg4(par, press, par1)
    df_bounds = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'Bounds.xlsx'))
    med15000_value = df_bounds['Med15000'].iloc[0]
    med0_value = df_bounds['Med0'].iloc[0]
    min0_value = df_bounds['Min0'].iloc[0]
    max15000_value = df_bounds['Max15000'].iloc[0]
    max0_value = df_bounds['Max0'].iloc[0]
    df_iniciais = pd.read_csv(os.path.join(DOWNLOAD_FOLDER,'InitialParam.txt'), header=None, sep='\t', names=['code', 'T15000', 'alpha'])
    ret = np.loadtxt(os.path.join(DOWNLOAD_FOLDER,'ret.txt'))
    MAXSAMP = ret.shape[0] // 22  # Assuming 22 columns per sample

    # Define the maximum number of points
    MAXPNT = 3  # Updated to reflect the use of only 3 points

    # Initialize arrays to store the results
    res2 = np.zeros((MAXSAMP, 4))
    rmse2 = np.zeros(MAXSAMP)

    # Initialize arrays to store the pressure and theta data
    hdata = np.zeros((MAXSAMP, MAXPNT))
    tdata = np.zeros((MAXSAMP, MAXPNT))
    nret = np.zeros(MAXSAMP)

    # Loop over each sample
    for i in range(MAXSAMP):
        start_index = i * 22
        end_index = (i + 1) * 22
        sample_data = ret[start_index:end_index]

        # Get the ID and the number of points for the current sample
        id = int(sample_data[0])
        npoints = int(sample_data[1])
        nret[i] = MAXPNT  # Set to 3 since we are using only 3 points

        # Get the theta and pressure values for the second, third, and fourth points
        if npoints >= 4:
            theta = sample_data[19:22]  # Second, third, and fourth theta values
            press = sample_data[3:6]  # Corresponding pressures

            # Retrieve the T15000 and alpha values from the Iniciais.txt file
            t15000_value = df_iniciais.loc[df_iniciais['code'] == id, 'T15000'].values[0]
            alpha_value = df_iniciais.loc[df_iniciais['code'] == id, 'alpha'].values[0]

            # Ensure initial parameters are within bounds
            t15000_value = np.clip(t15000_value, 0, max15000_value)
            alpha_value = np.clip(alpha_value, 0.001, 10000)

            # Set the initial parameters using the values from Iniciais.txt
            par1_value = sample_data[18]
            par = np.array([t15000_value, alpha_value, 2])

            # Define the bounds for the curve fitting without upper bounds for alpha and n
            bounds = ([0, 0.001, 1], [max15000_value, 10000, 10])

            # Perform the curve fitting with absolute_sigma=True
            res2a, _ = curve_fit(lambda press, par0, par2, par3: vg4ssq2(press, par0, par2, par3, par1_value),
                                 press, theta, p0=par, bounds=bounds, method='trf', maxfev=100000, absolute_sigma=True)

            # Calculate the sum of squares of the residuals
            ssq2a = vg4ssq(res2a, press, theta, par1_value)

            # Store the results
            res2[i, 0] = res2a[0]
            res2[i, 1] = par1_value
            res2[i, 2] = res2a[1]
            res2[i, 3] = res2a[2]
            ssq2 = ssq2a

            # Calculate the RMSE (Root Mean Square Error)
            rmse2[i] = np.sqrt(ssq2 / nret[i])
        else:
            # If there are fewer than 4 points for the sample, assign default values
            res2[i, :] = -9.9
            rmse2[i] = -9.9
    with open(os.path.join(DOWNLOAD_FOLDER,'res2.txt'), 'wt+') as fres2:
        for i in range(MAXSAMP):
            fres2.write(f'{i + 1} {res2[i, 0]:.5f} {res2[i, 1]:.5f} {res2[i, 2]:.5f} {res2[i, 3]:.5f} {rmse2[i]:.5f}\n')
    columns = ['ThetaR', 'ThetaS', 'alpha', 'n']
    df_results = pd.DataFrame(res2, columns=columns)
    df_results['m'] = 1 - 1 / df_results['n']
    df_results['RMSE'] = rmse2
    df_results.insert(0, 'code', range(1, len(df_results) + 1))
    output_filename = "Parameters.xlsx"
    df_results.to_excel(os.path.join(DOWNLOAD_FOLDER,output_filename), index=False)
    # Continue processing with the provided code
    import pandas as pd

    # Load the Excel file
    df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted.xlsx'))

    # Drop the 'PT' column
    df = df.drop(columns=['PT'])

    # Define a function to transpose the 'h' and 'theta' data for each code
    def transpose_h_theta(df):
        new_df = pd.DataFrame()
        for code in df['code'].unique():
            code_data = df[df['code'] == code].reset_index(drop=True)
            h_values = code_data['h']
            theta_values = code_data['theta']
            for i, (h, theta) in enumerate(zip(h_values, theta_values)):
                new_df.at[code, f'h{i}'] = h
                new_df.at[code, f'theta{i}'] = theta
                new_df.at[code, 'code'] = code  # Include the 'code' column
        return new_df

    # Apply the function and reorder the columns
    new_df = transpose_h_theta(df)
    new_df = new_df[['code', 'h0', 'theta0', 'h1', 'theta1', 'h2', 'theta2', 'h3', 'theta3']]

    # Save the new Excel file
    new_df.to_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted_Head.xlsx'), index=False)

    import pandas as pd
    import numpy as np

    # Load the Excel files
    df_soil_complete = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilComplete.xlsx'))
    df_parameters = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'Parameters.xlsx'))

    # Merge the dataframes based on the 'code' column
    df_merged = pd.merge(df_soil_complete, df_parameters, on='code')

    # Calculate the new column 'theta_calculated'
    def calculate_theta_calculated(row):
        h = row['h']
        thetaR = row['ThetaR']
        thetaS = row['ThetaS']
        alpha = row['alpha']
        m = row['m']
        n = row['n']
        return thetaR + (thetaS - thetaR) / (1 + (alpha * h) ** n) ** (m)
    df_merged['theta_calculated'] = df_merged.apply(calculate_theta_calculated, axis=1)

    # Calculate the new column 'RMSETOTAL'
    def calculate_RMSETOTAL(row):
        code = row['code']
        theta = row['theta']
        theta_calculated = row['theta_calculated']

        # Select all rows with the same code
        subset = df_merged[df_merged['code'] == code]

        # Calculate RMSE
        num_lines = len(subset)
        if num_lines > 3:
            divisor = num_lines - 3
        else:
            divisor = 3  # Avoid division by zero
        sum_squared_diff = np.sum((subset['theta'] - subset['theta_calculated']) ** 2)
        TOTALRMSE = np.sqrt(sum_squared_diff / divisor)
        return TOTALRMSE
    df_merged['TOTALRMSE'] = df_merged.apply(calculate_RMSETOTAL, axis=1)
    df_merged.to_excel(os.path.join(DOWNLOAD_FOLDER,'RMSE.xlsx'), index=False)

    import pandas as pd
    import numpy as np

    # Load the Excel files
    df_soil = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted.xlsx'))
    df_parameters = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'Parameters.xlsx'))

    # Remove rows where h is equal to 0 from the SoloSorted file
    df_soil = df_soil[df_soil['h'] != 0]

    # Merge the dataframes using 'code' as the key
    df_merged = pd.merge(df_soil, df_parameters, on='code', how='inner')

    # Calculate theta_calculated using the provided formula
    df_merged['theta_calculated'] = df_merged['ThetaR'] + (df_merged['ThetaS'] - df_merged['ThetaR']) / (1 + (df_merged['alpha'] * df_merged['h']) ** df_merged['n']) ** (df_merged['m'])

    # Calculate the differences between theta and theta_calculated
    df_merged['difference'] = abs(df_merged['theta'] - df_merged['theta_calculated'])

    # Group by 'code' and calculate the maximum absolute value of 'difference' for each group
    df_errormax = df_merged.groupby('code')['difference'].max().reset_index()

    # Rename the column to 'ERRORMAX'
    df_errormax.rename(columns={'difference': 'ERRORMAX'}, inplace=True)

    # Merge the dataframes to add the 'ERRORMAX' column to the original dataframe
    df_final = pd.merge(df_merged, df_errormax, on='code', how='inner')

    # Select only the columns 'code', 'theta', 'theta_calculated', and 'ERRORMAX'
    df_final = df_final[['code', 'theta', 'theta_calculated', 'ERRORMAX']]

    # Save the final dataframe to an Excel file
    df_final.to_excel(os.path.join(DOWNLOAD_FOLDER,'ERRORMAX.xlsx'), index=False)
    import pandas as pd
    import numpy as np

    # Load the RMSE and ERROMAX Excel files
    df_rmse = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'RMSE.xlsx'))
    df_errormax = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'ERRORMAX.xlsx'))

    # Create the DataFrame 'df' based on the RMSE DataFrame
    df = df_rmse.copy()

    # Map the first ERROMAX value for each code
    first_errormax = df_errormax.groupby('code')['ERRORMAX'].first().reset_index()

    # Rename the column to avoid conflicts
    first_errormax.rename(columns={'ERRORMAX': 'First_ERRORMAX'}, inplace=True)

    # Merge DataFrames based on the 'code' column
    df = pd.merge(df, first_errormax, on='code')

    # Create the new 'ERRORMAX' column based on the first ERRORMAX value for each code
    df['ERRORMAX'] = df['First_ERRORMAX']

    # Drop the 'First_ERROMAX' column as it is no longer needed
    df.drop(columns=['First_ERRORMAX'], inplace=True)

    # Adding new columns
    df['W60'] = ((1 + (df['alpha'] * 60) ** df['n']) ** -df['m'])
    df['W15000'] = ((1 + (df['alpha'] * 15000) ** df['n']) ** -df['m'])
    df['A60'] = 1 - df['W60']
    df['W60-W15000'] = df['W60'] - df['W15000']
    df['W60%'] = df['W60'] * 100
    df['W15000%'] = df['W15000'] * 100
    df['A60%'] = 100 - df['W60%']
    df['W60%-W15000%'] = df['W60%'] - df['W15000%']

    # Remove 'h' and 'theta' columns
    df.drop(columns=['h', 'theta', 'theta_calculated'], inplace=True)

    # Keep only the first row for each 'code'
    df = df.groupby('code').first().reset_index()

    # Add the 'Adherence' column to the DataFrame
    df['Adherence'] = np.where((df['ERRORMAX'] <= 0.02) & (df['TOTALRMSE'] <= 0.035), 'High',
                                np.where((df['ERRORMAX'] > 0.04) | (df['TOTALRMSE'] > 0.07), 'Low', 'Medium'))

    # Add the 'Classification' column to the DataFrame
    df['Classification'] = np.where(df['Adherence'] == 'High', 'Genuine Soil',
                                     np.where(df['Adherence'] == 'Medium', 'Adopted Soil', 'Rejected Soil'))

    # Load the SoloSorted_Head data
    df_head = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted_Head.xlsx'))

    # Merge DataFrames using the 'code' column as the joining key, with df_head coming first
    df = pd.merge(df_head, df, on='code', how='inner')

    # Save the merged DataFrame to a new Excel file
    df.to_excel(os.path.join(DOWNLOAD_FOLDER,'Adherence.xlsx'), index=False)

    # Additional processing and classification
    import pandas as pd

    # Load the Adherence.xlsx file
    df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'Adherence.xlsx'))
        # Function to classify the subordinate order
    def classify_suborder(row):
        diff = row['theta0'] - row['ThetaR']
        if 0 < diff <= 0.20:
            return 1
        elif 0.20 < diff <= 0.40:
            return 2
        elif 0.40 < diff <= 0.60:
            return 3
        elif diff > 0.60:
            return 4
        else:
            return None

    # Function to classify the main order
    def classify_order(row):
        if 2/3 <= row["A60"] < 1 and 0 < row['W60-W15000'] < 1/3 and 0 < row['W15000'] < 1/3:
            return "A"
        elif 1/3 <= row["A60"] < 2/3 and 0 < row['W60-W15000'] < 1/3 and 0 < row['W15000'] < 1/3:
            return "D"
        elif 1/3 <= row["A60"] < 2/3 and 1/3 <= row['W60-W15000'] < 2/3 and 0 < row['W15000'] < 1/3:
            return "B"
        elif 0 < row["A60"] < 1/3 and 2/3 <= row['W60-W15000'] < 1 and 0 < row['W15000'] < 1/3:
            return "C"
        elif 0 < row["A60"] < 1/3 and 1/3 <= row['W60-W15000'] < 2/3 and 0 < row['W15000'] < 1/3:
            return "E"
        elif 0 < row["A60"] < 1/3 and 1/3 <= row['W60-W15000'] < 2/3 and 1/3 <= row['W15000'] < 2/3:
            return "G"
        elif 0 < row["A60"] < 1/3 and 0 < row['W60-W15000'] < 1/3 and 2/3 <= row['W15000'] < 1:
            return "I"
        elif 0 < row["A60"] < 1/3 and 0 < row['W60-W15000'] < 1/3 and 1/3 <= row['W15000'] < 2/3:
            return "H"
        elif 1/3 <= row["A60"] < 2/3 and 0 < row['W60-W15000'] < 1/3 and 1/3 <= row['W15000'] < 2/3:
            return "F"
        else:
            return "Error"

    # Apply the functions to classify order and subordinate order
    df["Order"] = df.apply(classify_order, axis=1)
    df["Suborder"] = df.apply(classify_suborder, axis=1)

    # Ensure Suborder is an integer
    df["Suborder"] = df["Suborder"].fillna(0).astype(int)

    # Create the 'Family' column
    df["Family"] = df["Order"] + df["Suborder"].astype(str)

    # Dicionário para Suborder
    suborder_dict = {
        1: "Low Effective Porosity",
        2: "Moderate Effective Porosity",
        3: "High Effective Porosity",
        4: "Very High Effective Porosity"
    }

    # Dicionário para Order
    order_dict = {
        "A": "Highly Macrospacious Soil",
        "C": "Highly Mesospacious Soil",
        "I": "Highly Microspacious Soil",
        "D": "Macrospacious Soil",
        "E": "Mesospacious Soil",
        "H": "Microspacious Soil",
        "B": "Macro-Mesospacious Soil",
        "F": "Macro-Microspacious Soil",
        "G": "Meso-Microspacious Soil"
    }

    # Função para criar a nomenclatura da família
    def get_family_nomenclature(row):
        porosity = suborder_dict.get(row['Suborder'], "Unknown Porosity")
        order_name = order_dict.get(row['Order'], "Unknown Soil Type")
        return f"{porosity} - {order_name}"

    # Aplicar a função para criar a nova coluna
    df["Family Nomenclature"] = df.apply(get_family_nomenclature, axis=1)

    # Criar coluna 'a (cm³/cm³)' = theta0 - theta2
    df['a (cm³/cm³)'] = df['theta0'] - df['theta2']

    # Criar coluna 'w (cm³/cm³)' = theta2 - theta3
    df['w (cm³/cm³)'] = df['theta2'] - df['theta3']

    # Criar coluna 'Ksat (cm/d)' = 1931 * (a ^ 1.948)
    df['Ksat (cm/d)'] = 1931 * (df['a (cm³/cm³)'] ** 1.948)


    def classify_a_value(a):
        if a < 0.10:
            return "Low"
        elif 0.10 <= a <= 0.20:
            return "Moderate"
        elif a > 0.20:
            return "High"
        else:
            return "Undefined"

    df['a value'] = df['a (cm³/cm³)'].apply(classify_a_value)

    def classify_w_value(w):
        if w < 0.06:
            return "Low"
        elif 0.06 <= w <= 0.12:
            return "Moderate"
        elif w > 0.12:
            return "High"
        else:
            return "Undefined"

    df['w value'] = df['w (cm³/cm³)'].apply(classify_w_value)

    def classify_ksat_value(ksat):
        if ksat <= 12:
            return "Slow"
        elif 12 < ksat <= 120:
            return "Moderate"
        elif ksat > 120:
            return "Rapid"
        else:
            return "Undefined"

    df['ksat value'] = df['Ksat (cm/d)'].apply(classify_ksat_value)
    
    def classify_hydraulic_class(row):
        a_val = row['a value']
        w_val = row['w value']
        ksat_val = row['ksat value']

        # Primeira parte - restrição
        if (a_val == "Low" and w_val in ["Low", "Moderate", "High"]):
            restriction = "Soil with air restriction"
        elif (a_val in ["High", "Moderate"] and w_val == "Low"):
            restriction = "Soil with water restriction"
        elif ((a_val == "Moderate" and w_val == "High") or
              (a_val == "High" and w_val in ["Moderate", "High"]) or
              (a_val == "Moderate" and w_val == "Moderate")):
            restriction = "Soil without air and water restriction"
        else:
            restriction = "Undefined Restriction"

        # Segunda parte - permeabilidade
        if ksat_val == "Slow":
            permeability = "slow permeability"
        elif ksat_val == "Moderate":
            permeability = "moderate permeability"
        elif ksat_val == "Rapid":
            permeability = "rapid permeability"
        else:
            permeability = "undefined permeability"

        return f"{restriction} and {permeability}"

    # Aplicar a classificação para cada linha
    df['Hydraulic Class'] = df.apply(classify_hydraulic_class, axis=1)

    df.drop(columns=['a value', 'w value', 'ksat value'], inplace=True)

    # Rename the columns as requested
    df.rename(columns={"RMSE": "RMSE_3points", "TOTALRMSE": "RMSE_30cm-18000cm"}, inplace=True)

    # Reorder the columns to have 'Sample_ID' as the first column
    columns_order = ['SampleID'] + [col for col in df.columns if col != 'SampleID']
    df = df[columns_order]

    # Convert the column SAMPLE_ID to text
    df['SampleID'] = df['SampleID'].astype(str)

    
    # Save the DataFrame to an Excel file
    df.to_excel(os.path.join(DOWNLOAD_FOLDER,'YourSoilClassified.xlsx'), index=False)


    soilfull_df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilFull.xlsx'))
    adherence_df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER, 'Adherence.xlsx'))

    # Convertendo colunas relevantes para log10
    soilfull_df['log_h'] = np.log10(soilfull_df['h'].replace(0, np.nan))  # Transformando h e substituindo 0 por NaN
    soilfull_df['Theta'] = soilfull_df['theta']  # Assumindo que a coluna se chama 'theta'

    unique_codes = soilfull_df['code'].unique()

    # Criar diretório para armazenar os plots
    os.makedirs(PLOTS_FOLDER, exist_ok=True)

    for code in unique_codes:
        code_df = soilfull_df[soilfull_df['code'] == code]
        sample_id = code_df['SampleID'].values[0]  # Extraindo o SampleID para o título
        
        # Criar um gráfico para cada código
        plt.figure(figsize=(10, 6))
        plot_color = 'black'  # Cor dos pontos

        # Plotar todos os pontos de theta
        for index, row in code_df.iterrows():
            if row['h'] == 0:
                plt.scatter(0, row['Theta'], color=plot_color, marker='x', s=100, label='φ', clip_on=False)  # Ponto para h=0 como 'X' maior e identificado como φ
            else:
                plt.scatter(row['log_h'], row['Theta'], color=plot_color, alpha=0.6)

        # Plotar os valores de theta1, theta2, theta3 em relação a h1, h2, h3
        adherence_row = adherence_df[adherence_df['code'] == code]
        if not adherence_row.empty:
            # Extraindo parâmetros
            ThetaR = adherence_row['ThetaR'].values[0]
            ThetaS = adherence_row['ThetaS'].values[0]
            alpha = adherence_row['alpha'].values[0]
            n = adherence_row['n'].values[0]

            # Filtrando h maior que 30 para a reta
            h_values_above_30 = np.linspace(30, soilfull_df['h'].max(), 100)
            Theta_values_above_30 = ThetaR + (ThetaS - ThetaR) / (1 + (alpha * h_values_above_30) ** n) ** (1 - 1 / n)
            
            # Plotando a reta apenas para h > 30 em cinza
            plt.plot(np.log10(h_values_above_30), Theta_values_above_30, linestyle='--', color='gray')

            # Adicionar valores de theta1, theta2, theta3 como quadrados sólidos
            plt.scatter(np.log10(adherence_row['h1'].values[0]), adherence_row['theta1'].values[0], color=plot_color, marker='s', s=100)  # Aumentar tamanho
            plt.scatter(np.log10(adherence_row['h2'].values[0]), adherence_row['theta2'].values[0], color=plot_color, marker='s', s=100)  # Aumentar tamanho
            plt.scatter(np.log10(adherence_row['h3'].values[0]), adherence_row['theta3'].values[0], color=plot_color, marker='s', s=100)  # Aumentar tamanho

            # Arredondando valores para a legenda
            error_max = round(adherence_row['ERRORMAX'].values[0], 3)
            total_rmse = round(adherence_row['TOTALRMSE'].values[0], 3)

            # Adicionando informações à legenda com o "X" para φ ao lado do texto
            handles = [
                plt.Line2D([0], [0], color='w', label=f'ERRORMAX: {error_max} cm³/cm³'),  # ERRORMAX
                    plt.Line2D([0], [0], color='w', label=f'RMSE$_{{30-18000}}$: {total_rmse} cm³/cm³'),  # RMSE
                    plt.Line2D([0], [0], marker='s', color='black', label='θ1, θ2, θ3', markerfacecolor='black', markersize=10, linestyle='None'),  # Quadrados sólidos para θ1, θ2, θ3
                    plt.Line2D([0], [0], marker='x', color='black', label='φ', markersize=10, linestyle='None'),
                    plt.Line2D([0], [0], marker='o', color='gray', label='θ', markersize=6, linestyle='None')
            ]
            plt.legend(handles=handles, loc='upper right')

        # Configurações do gráfico
        plt.axhline(0, color='black', lw=0.8)  # Linha horizontal no eixo y=0
        plt.axvline(0, color='black', lw=0.8)  # Linha vertical no eixo x=0
        plt.xlabel("log h (cm)")
        plt.ylabel("Theta (cm³/cm³)")
        plt.title(f"Retention curve - Soil {sample_id}")  # Usando SampleID como título
        plt.grid()
        plt.tight_layout()

        # Ajustar limites dos eixos para garantir que (0,0) coincide corretamente
        plt.xlim(left=0)  # Garantir que o eixo x comece em 0
        plt.ylim(bottom=0)  # Garantir que o eixo y comece em 0

        # Salvar gráfico
        plt.savefig(os.path.join(PLOTS_FOLDER, f'retention_curve_code_{code}.jpg'))
        plt.close()


    def classify_suborder(row):
        diff = row['theta0'] - row['ThetaR']
        if 0 < diff <= 0.20:
            return 1
        elif 0.20 < diff <= 0.40:
            return 2
        elif 0.40 < diff <= 0.60:
            return 3
        elif diff > 0.60:
            return 4

    def map_color(adherence):
        if adherence == 'High':
            return 'green'
        elif adherence == 'Medium':
            return '#FFC300'  # Amarelo mais brilhante
        elif adherence == 'Low':
            return 'red'

    def plot_ternary(df, title, filename):
        scale = 100
        try:
            figure, tax = ternary.figure(scale=scale)
        except AttributeError:
            raise Exception("A função 'figure' não está disponível no módulo 'ternary'. Verifique a instalação.")

        figure.set_size_inches(5, 5)
        
        # Adiciona um grid de 81 triângulos menores com cinza mais escuro
        step = scale / 9
        grid_color = '#a9a9a9'  # Cinza mais escuro

        for i in range(9):
            for j in range(9 - i):
                x1, y1 = i * step, j * step
                x2, y2 = (i + 1) * step, j * step
                x3, y3 = i * step, (j + 1) * step

                # Cinza mais escuro e pontilhado
                tax.line((x1, y1), (x2, y2), color=grid_color, linestyle='--', linewidth=0.5)
                tax.line((x2, y2), (x3, y3), color=grid_color, linestyle='--', linewidth=0.5)
                tax.line((x3, y3), (x1, y1), color=grid_color, linestyle='--', linewidth=0.5)

        # Agora desenhamos o triângulo principal sobre o grid para garantir que ele fique por cima
        tax.boundary(linewidth=1.0)

        fontsize = 12
        offset = 0.2
        tax.left_axis_label("    W15000 (%)", fontsize=fontsize, offset=offset, rotation=-240)
        tax.right_axis_label("W60-W15000 (%)", fontsize=fontsize, offset=offset, rotation=240)
        tax.bottom_axis_label("      A60 (100-W60) (%)", fontsize=fontsize, offset=offset, rotation=180)

        # Adicionando os textos nas laterais do triângulo, a 10 pixels de distância
        # Ajustando as posições dos textos para que estejam nas laterais corretas e a distância adequada
        tax.annotate("Microspace (%)", position=(62, 54), fontsize=10, rotation=60, va='center', ha='center')  # Lado esquerdo
        tax.annotate("Mesospace (%)", position=(-11.6, 54), fontsize=10, rotation=-60, va='center', ha='center')   # Lado direito
        tax.annotate("Macrospace (%)", position=(54, -11.6), fontsize=10, rotation=0, va='center', ha='center')      # Parte inferior

        for index, row in df.iterrows():
            color = map_color(row['Adherence'])
            tax.scatter([[row['A60%'], row['W15000%'], row['W60%-W15000%']]], marker='o', s=35, color=color)  # Tamanho dobrado

        plt.gca().invert_xaxis()
        triangles = [((66.7, 0), (66.7, 33.3)), ((33.3, 0), (33.3, 66.7)), ((66.7, 33.3), (0, 33.3)),
                     ((33.3, 66.7), (0, 66.7)), ((0, 66.7), (66.7, 0)), ((0, 33.3), (33.3, 0))]
        for p1, p2 in triangles:
            tax.line(p1, p2, linewidth=1, color='black')

        titles = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
        midpoints = [(81, 9, 5), (47, 9.5), (13, 9.5), (59, 20), (24, 20), (45, 45), (13, 45), (26, 52), (12, 78)]
        for title_point, point in zip(titles, midpoints):
            tax.annotate(title_point, point)

        plt.title(title, pad=20)

        # Ajustando os ticks para os pontos desejados
        tax.ticks(axis='lbr', ticks=[0, 33.3, 66.7, 100], linewidth=1, offset=0.025)
        tax.get_axes().axis('off')
        tax.clear_matplotlib_ticks()

        legend_elements = [
            plt.Line2D([0], [0], marker='o', color='w', label='Genuine Soil', markerfacecolor='green', markersize=5),
            plt.Line2D([0], [0], marker='o', color='w', label='Adopted Soil', markerfacecolor='#FFC300', markersize=5),
            plt.Line2D([0], [0], marker='o', color='w', label='Rejected Soil', markerfacecolor='red', markersize=5)
        ]
        plt.legend(handles=legend_elements, loc='upper left', title="Adherence", fontsize='small')

        plt.savefig(filename, format='jpg')
        plt.close()

    # Criar um diretório para armazenar os plots
    os.makedirs(PLOTS_FOLDER, exist_ok=True)

    # Plotar gráficos adicionais para cada ordem, salvando apenas a que tiver dados
    for order in range(1, 5):
        filtered_df = df[df.apply(lambda row: classify_suborder(row) == order, axis=1)]
        if not filtered_df.empty:  # Verifica se há dados para a ordem
            plot_ternary(filtered_df, f"Suborder {order}", os.path.join(PLOTS_FOLDER, f'suborder_plot_1soil.jpg'))

    files_to_include = [
        'suborder_plot_1soil.jpg',
        'retention_curve_code_1.jpg'
    ]

    
    df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER, 'YourSoilClassified.xlsx'))
    df.drop(columns=["code", "Adherence", "W60", "W15000", "A60", "W60-W15000"], inplace=True)

    df.rename(columns={"SampleID": "code"}, inplace=True)

    df.to_excel(os.path.join(DOWNLOAD_FOLDER,'YourSoilClassified.xlsx'), index=False)

    # Criando o arquivo ZIP com os arquivos especificados
    with zipfile.ZipFile(os.path.join(DOWNLOAD_FOLDER, 'ternary_plots.zip'), 'w') as zipf:
        for root, _, files in os.walk(PLOTS_FOLDER):
            for file in files:
                if file in files_to_include:  # Inclui apenas os arquivos desejados
                    zipf.write(os.path.join(root, file), file)



@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file part'})

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'No selected file'})

    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Process the file and generate the output files
        process_excel(filepath)

        # Provide download links
        file1 = '/download_xlsx'  # Updated to point to download route
        file2 = '/download_zip'    # Updated to point to download route

        return jsonify({'success': True, 'file1': file1, 'file2': file2})


def process_excel(excel_file):
    df = pd.read_excel(excel_file)
    # Criação do arquivo SoilFull

    if "h(cm)" in df.columns:
        df.rename(columns={"h(cm)": "h"}, inplace=True)
    
    df.to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilFull.xlsx'), index=False)
    
    if 'h' in df.columns and 'theta' in df.columns:
        code1(df)
    else:
        code2(df)

def code1(df):
    df['h'] = pd.to_numeric(df['h'], errors='coerce')
    df = df.dropna(subset=['h'])
    
    # Criação da coluna codeorign com todos os códigos
    df['codeorign'] = df['code']
    
    # Criar um DataFrame para solos descartados antes do sort
    soils_discarded = pd.DataFrame(columns=['codeorign', 'codedis'])
    
    codes_before_filtering = df['code'].unique()  # Captura todos os códigos antes da filtragem
    
    codes_with_h_0_1 = df.groupby('code').filter(lambda x: ((x['h'] >= 0) & (x['h'] <= 1)).any())['code'].unique()
    for code in codes_with_h_0_1:
        closest_theta_value = df[(df['code'] == code) & (df['h'] >= 0) & (df['h'] <= 1)].sort_values(by='h').iloc[0]['theta']
        df = pd.concat([df, pd.DataFrame({'code': [code], 'h': [0], 'theta': [closest_theta_value], 'codeorign': [code]})])

    df_head_0 = df[df['h'] == 0].groupby('code').head(1)
    df_head_30_80 = df[(df['h'] >= 30) & (df['h'] <= 80)].copy()
    df_head_30_80.loc[:, 'h_diff'] = abs(df_head_30_80['h'] - 60)
    df_head_30_80 = df_head_30_80.groupby('code').apply(lambda x: x.loc[x['h_diff'].idxmin()]).reset_index(drop=True)
    df_head_200_500 = df[(df['h'] >= 250) & (df['h'] <= 500)].copy()
    df_head_200_500.loc[:, 'h_diff'] = abs(df_head_200_500['h'] - 330)
    df_head_200_500 = df_head_200_500.groupby('code').apply(lambda x: x.loc[x['h_diff'].idxmin()]).reset_index(drop=True)
    df_head_9000_18000 = df[(df['h'] >= 9000) & (df['h'] <= 18000)].copy()
    df_head_9000_18000.loc[:, 'h_diff'] = abs(df_head_9000_18000['h'] - 15000)
    df_head_9000_18000 = df_head_9000_18000.groupby('code').apply(lambda x: x.loc[x['h_diff'].idxmin()]).reset_index(drop=True)
    
    df_sorted = pd.concat([df_head_0, df_head_30_80, df_head_200_500, df_head_9000_18000])
    grouped = df_sorted.groupby(['code', 'h'])['theta'].first().reset_index()
    grouped['PT'] = ''
    grouped.loc[grouped['h'] == 0, 'PT'] = grouped.loc[grouped['h'] == 0, 'theta']
    
    # Filtra códigos com pelo menos 4 parâmetros
    sample_ids_to_keep = grouped.groupby('code')['h'].count().loc[lambda x: x >= 4].index
    discarded_codes = set(codes_before_filtering) - set(sample_ids_to_keep)  # Códigos descartados
    
    # Preencher o DataFrame de solos descartados
    soils_discarded['codeorign'] = list(codes_before_filtering)
    soils_discarded['codedis'] = [code if code in discarded_codes else None for code in codes_before_filtering]

    # Renomear a coluna codeorign para Sample_ID e garantir que os valores sejam texto
    soils_discarded['Sample_ID'] = soils_discarded['codeorign'].astype(str)
    soils_discarded.drop(columns=['codeorign'], inplace=True)
    
    # Criar a coluna OBS
    soils_discarded['OBS'] = soils_discarded['codedis'].apply(lambda x: "this soil could not be classified because it didn't meet requirements, see about us" if pd.notna(x) else '-')

    grouped = grouped[grouped['code'].isin(sample_ids_to_keep)]
    sample_id_map = {sample_id: i + 1 for i, sample_id in enumerate(grouped['code'].unique())}
    grouped['code'] = grouped['code'].map(sample_id_map)
    
    grouped[['code', 'h', 'theta', 'PT']].to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilSorted.xlsx'), index=False)
    
    # Criar o arquivo SoilComplete.xlsx com valores de h >= 30
    df_complete = df[df['code'].isin(sample_ids_to_keep) & (df['h'] >= 30)][['code', 'h', 'theta']]
    df_complete['Sample_ID'] = df_complete['code']
    df_complete['code'] = df_complete['code'].map(sample_id_map)
    
    df_complete.to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilComplete.xlsx'), index=False)
    
    # Criar o arquivo SoilFull.xlsx com todos os valores de h >= 0
    df_full = df[df['code'].isin(sample_ids_to_keep)][['code', 'h', 'theta']]
    df_full['Sample_ID'] = df_full['code']
    df_full['code'] = df_full['code'].map(sample_id_map)
    
    df_full.to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilFull.xlsx'), index=False)
    
    # Salvar solos descartados
    soils_discarded.to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilsDiscarted.xlsx'), index=False)

    # Calculations for Bounds.xlsx
    calculate_bounds()

    # Generate Parameters.xlsx
    calculate_parameters()

def code2(df):
    codes = []
    h_values = []
    theta_values = []
    
    # Criação da coluna codeorign com todos os códigos
    df['codeorign'] = df['code']  # Captura os códigos originais
    
    codes_before_filtering = df['code'].unique()  # Captura todos os códigos antes da filtragem
    
    # Criar um DataFrame para solos descartados antes do sort
    soils_discarded = pd.DataFrame(columns=['codeorign', 'codedis'])
    
    for index, row in df.iterrows():
        code = row['code']
        for column, value in row.items():
            if column != 'code' and not pd.isna(value):
                codes.append(code)
                h_values.append(column)
                theta_values.append(value)

    new_df = pd.DataFrame({'code': codes, 'h': h_values, 'theta': theta_values})
    new_df['h'] = pd.to_numeric(new_df['h'], errors='coerce')
    new_df = new_df.dropna(subset=['h'])
    
    # Criação da coluna codeorign no new_df
    new_df['codeorign'] = new_df['code']
    
    # Identificar os códigos que já possuem uma linha com h == 0
    codes_with_h_0 = new_df[new_df['h'] == 0]['code'].unique()

    codes_with_h_0_1 = new_df.groupby('code').filter(lambda x: ((x['h'] >= 0) & (x['h'] <= 1)).any())['code'].unique()
    for code in codes_with_h_0_1:
        if code not in codes_with_h_0:  # Verificar se o código já tem h == 0
            closest_theta_value = new_df[(new_df['code'] == code) & (new_df['h'] >= 0) & (new_df['h'] <= 1)].sort_values(by='h').iloc[0]['theta']
            new_df = pd.concat([new_df, pd.DataFrame({'code': [code], 'h': [0], 'theta': [closest_theta_value], 'codeorign': [code]})])

    df_head_0 = new_df[new_df['h'] == 0].groupby('code').head(1)
    df_head_30_80 = new_df[(new_df['h'] >= 30) & (new_df['h'] <= 80)].copy()
    df_head_30_80.loc[:, 'h_diff'] = abs(df_head_30_80['h'] - 60)
    df_head_30_80 = df_head_30_80.groupby('code').apply(lambda x: x.loc[x['h_diff'].idxmin()]).reset_index(drop=True)
    
    df_head_200_500 = new_df[(new_df['h'] >= 250) & (new_df['h'] <= 500)].copy()
    df_head_200_500.loc[:, 'h_diff'] = abs(df_head_200_500['h'] - 330)
    df_head_200_500 = df_head_200_500.groupby('code').apply(lambda x: x.loc[x['h_diff'].idxmin()]).reset_index(drop=True)
    
    df_head_9000_18000 = new_df[(new_df['h'] >= 9000) & (new_df['h'] <= 18000)].copy()
    df_head_9000_18000.loc[:, 'h_diff'] = abs(df_head_9000_18000['h'] - 15000)
    df_head_9000_18000 = df_head_9000_18000.groupby('code').apply(lambda x: x.loc[x['h_diff'].idxmin()]).reset_index(drop=True)

    # Concatenar os DataFrames e criar o final
    df_sorted = pd.concat([df_head_0, df_head_30_80, df_head_200_500, df_head_9000_18000])
    grouped = df_sorted.groupby(['code', 'h'])['theta'].first().reset_index()
    grouped['PT'] = ''
    grouped.loc[grouped['h'] == 0, 'PT'] = grouped.loc[grouped['h'] == 0, 'theta']

    # Filtra códigos com pelo menos 4 parâmetros
    sample_ids_to_keep = grouped.groupby('code')['h'].count().loc[lambda x: x >= 4].index
    discarded_codes = set(codes_before_filtering) - set(sample_ids_to_keep)  # Códigos descartados
    
    # Preencher o DataFrame de solos descartados
    soils_discarded['codeorign'] = list(codes_before_filtering)
    soils_discarded['codedis'] = [code if code in discarded_codes else None for code in codes_before_filtering]

    # Renomear a coluna codeorign para Sample_ID e garantir que os valores sejam texto
    soils_discarded['Sample_ID'] = soils_discarded['codeorign'].astype(str)
    soils_discarded.drop(columns=['codeorign'], inplace=True)
    
    # Criar a coluna OBS
    soils_discarded['OBS'] = soils_discarded['codedis'].apply(lambda x: "this soil could not be classified because it didn't meet requirements, see about us" if pd.notna(x) else '-')

    grouped = grouped[grouped['code'].isin(sample_ids_to_keep)]
    sample_id_map = {sample_id: i + 1 for i, sample_id in enumerate(grouped['code'].unique())}
    grouped['code'] = grouped['code'].map(sample_id_map)

    # Criar o arquivo SoilComplete.xlsx com os códigos filtrados (após a contagem dos parâmetros)
    df_complete = new_df[new_df['code'].isin(sample_ids_to_keep)][['code', 'h', 'theta']]
    df_complete['Sample_ID'] = df_complete['code']
    df_complete['code'] = df_complete['code'].map(sample_id_map)

    df_complete.to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilComplete.xlsx'), index=False)

    # Criar o arquivo SoilFull.xlsx (uma cópia do SoilComplete, sem a filtragem de h > 30)
    df_full = new_df[new_df['code'].isin(sample_ids_to_keep)]  # Garantir que são apenas os códigos não descartados
    df_full[['code', 'h', 'theta', 'codeorign']] = df_full[['code', 'h', 'theta', 'codeorign']]  # Incluir todos os valores de h

    # Reiniciar a enumeração da coluna 'code' para o SoilFull.xlsx
    sample_id_map_full = {sample_id: i + 1 for i, sample_id in enumerate(df_full['code'].unique())}
    df_full['code'] = df_full['code'].map(sample_id_map_full)

    df_full[['code', 'h', 'theta', 'codeorign']].to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilFull.xlsx'), index=False)

    # Criar o arquivo SoilSorted.xlsx (após a filtragem dos códigos)
    grouped[['code', 'h', 'theta', 'PT']].to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilSorted.xlsx'), index=False)

    # Salvar solos descartados
    soils_discarded.to_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilsDiscarted.xlsx'), index=False)

    # Calculations for Bounds.xlsx
    calculate_bounds()

    # Generate Parameters.xlsx
    calculate_parameters()

   

def calculate_bounds():
    df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted.xlsx'))

    # Calculate the alpha and T15000 values
    codes = []
    t15000_values = []
    alpha_values = []
    for code in df['code'].unique():
        df_temp = df[df['code'] == code]
        third_theta = df_temp.iloc[2]['theta']
        fourth_theta = df_temp.iloc[3]['theta']
        first_theta = df_temp.iloc[0]['theta']
        h_third_line = df_temp.iloc[2]['h']
        alpha = (((third_theta - (fourth_theta - 0.01)) / (first_theta - (fourth_theta - 0.01))) ** (-1 / 0.5) - 1) ** (
                    1 - 0.5) / h_third_line
        t15000_temp = df_temp.iloc[3]['theta'] - 0.01
        codes.append(code)
        t15000_values.append(t15000_temp)
        alpha_values.append(alpha)
    df_temp = pd.DataFrame({'code': codes, 'T15000': t15000_values, 'alpha': alpha_values})
    df_temp.to_csv(os.path.join(DOWNLOAD_FOLDER,'InitialParam.txt'), header=False, index=False, sep='\t')

    # Calculate bounds
    med0 = df[df['h'] == 0]['theta'].mean()
    med15000 = df[df['h'] >= 9000]['theta'].mean() - 0.01
    min0 = df[df['h'] == 0]['theta'].min()
    max0 = df[df['h'] == 0]['theta'].max()
    max15000 = df[df['h'] >= 9000]['theta'].max()
    data = {
        'Med0': [med0],
        'Med15000': [med15000],
        'Min0': [min0],
        'Max0': [max0],
        'Max15000': [max15000]
    }
    df_result = pd.DataFrame(data)
    df_result.to_excel(os.path.join(DOWNLOAD_FOLDER,'Bounds.xlsx'), index=False)

    # Importar a biblioteca openpyxl
    import openpyxl
    def insert_row_excel():
        # Carregar o arquivo Excel
        wb = openpyxl.load_workbook(os.path.join(DOWNLOAD_FOLDER,"SoilSorted.xlsx"))
        # Selecionar a primeira planilha
        ws = wb.active
        row = 2  # Começando pela linha 2
        while ws.cell(row=row, column=1).value is not None:  # Enquanto houver dados na coluna A
            if ws.cell(row=row, column=1).value != ws.cell(row=row + 1, column=1).value:
                # Se o valor da célula atual for diferente da célula abaixo
                ws.insert_rows(row + 1)  # Inserir uma linha abaixo da linha atual
                row += 1  # Mover para a próxima linha
            row += 1  # Mover para a próxima linha

        # Salvar as alterações no arquivo Excel
        wb.save(os.path.join(DOWNLOAD_FOLDER,"SoilSorted.xlsx"))

    # Chamar a função para inserir linhas
    insert_row_excel()

    # Inicializar uma lista para armazenar os dados transpostos
    transposed_data = []

    # Loop para transpor os dados para cada grupo de dados na planilha original
    for site_id, group in df.groupby("code"):
        # Transpor os dados do Head
        head = group["h"].head(16).reset_index(drop=True)
        matric_potential_values = list(head.combine_first(pd.Series([0]*16)))

        # Transpor os dados do Theta
        theta = group["theta"].head(16).reset_index(drop=True)
        water_content_values = list(theta.combine_first(pd.Series([0]*16)))

        # Adicionar os dados transpostos à lista
        transposed_data.append([site_id, len(group), *matric_potential_values, *water_content_values, group["PT"].iloc[0]])

    # Salvar os dados em um arquivo de texto
    with open(os.path.join(DOWNLOAD_FOLDER,"ret.txt"), "w") as f:
        for line in transposed_data:
            formatted_line = "\t".join(map(lambda x: '{:.3f}'.format(x) if isinstance(x, float) else str(x), line))
            f.write(formatted_line + "\n")

def calculate_parameters():
    import numpy as np
    from scipy.optimize import curve_fit
    import pandas as pd
    def vg4(par, h, par1):
        return np.abs(par[0]) + (par1 - np.abs(par[0])) / (1 + (par[1] * h) ** par[2]) ** (1 - 1 / par[2])
    def vg4ssq(par, press, theta, par1):
        theta1 = vg4(par, press, par1)
        diff = theta - theta1
        ssq = np.dot(diff, diff)
        return ssq
    def vg4ssq2(press, par0, par2, par3, par1):
        par = [par0, par2, par3]
        return vg4(par, press, par1)
    df_bounds = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'Bounds.xlsx'))
    med15000_value = df_bounds['Med15000'].iloc[0]
    med0_value = df_bounds['Med0'].iloc[0]
    min0_value = df_bounds['Min0'].iloc[0]
    max15000_value = df_bounds['Max15000'].iloc[0]
    max0_value = df_bounds['Max0'].iloc[0]
    df_iniciais = pd.read_csv(os.path.join(DOWNLOAD_FOLDER,'InitialParam.txt'), header=None, sep='\t', names=['code', 'T15000', 'alpha'])
    ret = np.loadtxt(os.path.join(DOWNLOAD_FOLDER,'ret.txt'))
    MAXSAMP = ret.shape[0]
    MAXPNT = 3
    res2 = np.zeros((MAXSAMP, 4))
    rmse2 = np.zeros(MAXSAMP)
    hdata = np.zeros((MAXSAMP, MAXPNT))
    tdata = np.zeros((MAXSAMP, MAXPNT))
    nret = np.zeros(MAXSAMP)
    for i in range(MAXSAMP):
        id, npoints = int(ret[i, 0]), int(ret[i, 1])
        nret[i] = MAXPNT
        if npoints >= 4:
            theta = ret[i, 19:22]
            press = ret[i, 3:6]
            t15000_value = df_iniciais.loc[df_iniciais['code'] == id, 'T15000'].values[0]
            alpha_value = df_iniciais.loc[df_iniciais['code'] == id, 'alpha'].values[0]
            t15000_value = np.clip(t15000_value, 0, max15000_value)
            alpha_value = np.clip(alpha_value, 0.001, 10000)
            par1_value = ret[i, 18]
            par = np.array([t15000_value, alpha_value, 2])
            bounds = ([0, 0.001, 1], [max15000_value, 10000, 10])
            res2a, _ = curve_fit(lambda press, par0, par2, par3: vg4ssq2(press, par0, par2, par3, par1_value),
                                 press, theta, p0=par, bounds=bounds, method='trf', maxfev=100000, absolute_sigma=True)
            ssq2a = vg4ssq(res2a, press, theta, par1_value)
            res2[i, 0] = res2a[0]
            res2[i, 1] = par1_value
            res2[i, 2] = res2a[1]
            res2[i, 3] = res2a[2]
            ssq2 = ssq2a
            rmse2[i] = np.sqrt(ssq2 / nret[i])
        else:
            res2[i, :] = -9.9
            rmse2[i] = -9.9
    with open(os.path.join(DOWNLOAD_FOLDER,'res2.txt'), 'wt+') as fres2:
        for i in range(MAXSAMP):
            fres2.write(f'{i + 1} {res2[i, 0]:.5f} {res2[i, 1]:.5f} {res2[i, 2]:.5f} {res2[i, 3]:.5f} {rmse2[i]:.5f}\n')
    columns = ['ThetaR', 'ThetaS', 'alpha', 'n']
    df_results = pd.DataFrame(res2, columns=columns)
    df_results['m'] = 1 - 1 / df_results['n']
    df_results['RMSE'] = rmse2
    df_results.insert(0, 'code', range(1, len(df_results) + 1))
    output_filename = "Parameters.xlsx"
    df_results.to_excel(os.path.join(DOWNLOAD_FOLDER,output_filename), index=False)
    # Continue processing with the provided code
    import pandas as pd

    # Load the Excel file
    df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted.xlsx'))

    # Drop the 'PT' column
    df = df.drop(columns=['PT'])

    # Define a function to transpose the 'h' and 'theta' data for each code
    def transpose_h_theta(df):
        new_df = pd.DataFrame()
        for code in df['code'].unique():
            code_data = df[df['code'] == code].reset_index(drop=True)
            h_values = code_data['h']
            theta_values = code_data['theta']
            for i, (h, theta) in enumerate(zip(h_values, theta_values)):
                new_df.at[code, f'h{i}'] = h
                new_df.at[code, f'theta{i}'] = theta
                new_df.at[code, 'code'] = code  # Include the 'code' column
        return new_df

    # Apply the function and reorder the columns
    new_df = transpose_h_theta(df)
    new_df = new_df[['code', 'h0', 'theta0', 'h1', 'theta1', 'h2', 'theta2', 'h3', 'theta3']]

    # Save the new Excel file
    new_df.to_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted_Head.xlsx'), index=False)

    import pandas as pd
    import numpy as np

    # Load the Excel files
    df_soil_complete = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilComplete.xlsx'))
    df_parameters = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'Parameters.xlsx'))

    # Merge the dataframes based on the 'code' column
    df_merged = pd.merge(df_soil_complete, df_parameters, on='code')

    # Calculate the new column 'theta_calculated'
    def calculate_theta_calculated(row):
        h = row['h']
        thetaR = row['ThetaR']
        thetaS = row['ThetaS']
        alpha = row['alpha']
        m = row['m']
        n = row['n']
        return thetaR + (thetaS - thetaR) / (1 + (alpha * h) ** n) ** (m)
    df_merged['theta_calculated'] = df_merged.apply(calculate_theta_calculated, axis=1)

    # Calculate the new column 'RMSETOTAL'
    def calculate_RMSETOTAL(row):
        code = row['code']
        theta = row['theta']
        theta_calculated = row['theta_calculated']

        # Select all rows with the same code
        subset = df_merged[df_merged['code'] == code]

        # Calculate RMSE
        num_lines = len(subset)
        if num_lines > 3:
            divisor = num_lines - 3
        else:
            divisor = 3  # Avoid division by zero
        sum_squared_diff = np.sum((subset['theta'] - subset['theta_calculated']) ** 2)
        TOTALRMSE = np.sqrt(sum_squared_diff / divisor)
        return TOTALRMSE
    df_merged['TOTALRMSE'] = df_merged.apply(calculate_RMSETOTAL, axis=1)
    df_merged.to_excel(os.path.join(DOWNLOAD_FOLDER,'RMSE.xlsx'), index=False)

    import pandas as pd
    import numpy as np

    # Load the Excel files
    df_soil = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted.xlsx'))
    df_parameters = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'Parameters.xlsx'))

    # Remove rows where h is equal to 0 from the SoloSorted file
    df_soil = df_soil[df_soil['h'] != 0]

    # Merge the dataframes using 'code' as the key
    df_merged = pd.merge(df_soil, df_parameters, on='code', how='inner')

    # Calculate theta_calculated using the provided formula
    df_merged['theta_calculated'] = df_merged['ThetaR'] + (df_merged['ThetaS'] - df_merged['ThetaR']) / (1 + (df_merged['alpha'] * df_merged['h']) ** df_merged['n']) ** (df_merged['m'])

    # Calculate the differences between theta and theta_calculated
    df_merged['difference'] = abs(df_merged['theta'] - df_merged['theta_calculated'])

    # Group by 'code' and calculate the maximum absolute value of 'difference' for each group
    df_errormax = df_merged.groupby('code')['difference'].max().reset_index()

    # Rename the column to 'ERRORMAX'
    df_errormax.rename(columns={'difference': 'ERRORMAX'}, inplace=True)

    # Merge the dataframes to add the 'ERRORMAX' column to the original dataframe
    df_final = pd.merge(df_merged, df_errormax, on='code', how='inner')

    # Select only the columns 'code', 'theta', 'theta_calculated', and 'ERRORMAX'
    df_final = df_final[['code', 'theta', 'theta_calculated', 'ERRORMAX']]

    # Save the final dataframe to an Excel file
    df_final.to_excel(os.path.join(DOWNLOAD_FOLDER,'ERRORMAX.xlsx'), index=False)
    import pandas as pd
    import numpy as np

    # Load the RMSE and ERROMAX Excel files
    df_rmse = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'RMSE.xlsx'))
    df_errormax = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'ERRORMAX.xlsx'))

    # Create the DataFrame 'df' based on the RMSE DataFrame
    df = df_rmse.copy()

    # Map the first ERROMAX value for each code
    first_errormax = df_errormax.groupby('code')['ERRORMAX'].first().reset_index()

    # Rename the column to avoid conflicts
    first_errormax.rename(columns={'ERRORMAX': 'First_ERRORMAX'}, inplace=True)

    # Merge DataFrames based on the 'code' column
    df = pd.merge(df, first_errormax, on='code')

    # Create the new 'ERRORMAX' column based on the first ERRORMAX value for each code
    df['ERRORMAX'] = df['First_ERRORMAX']

    # Drop the 'First_ERROMAX' column as it is no longer needed
    df.drop(columns=['First_ERRORMAX'], inplace=True)

    # Adding new columns
    df['W60'] = ((1 + (df['alpha'] * 60) ** df['n']) ** -df['m'])
    df['W15000'] = ((1 + (df['alpha'] * 15000) ** df['n']) ** -df['m'])
    df['A60'] = 1 - df['W60']
    df['W60-W15000'] = df['W60'] - df['W15000']
    df['W60%'] = df['W60'] * 100
    df['W15000%'] = df['W15000'] * 100
    df['A60%'] = 100 - df['W60%']
    df['W60%-W15000%'] = df['W60%'] - df['W15000%']

    # Remove 'h' and 'theta' columns
    df.drop(columns=['h', 'theta', 'theta_calculated'], inplace=True)

    # Keep only the first row for each 'code'
    df = df.groupby('code').first().reset_index()

    # Add the 'Adherence' column to the DataFrame
    df['Adherence'] = np.where((df['ERRORMAX'] <= 0.02) & (df['TOTALRMSE'] <= 0.035), 'High',
                                np.where((df['ERRORMAX'] > 0.04) | (df['TOTALRMSE'] > 0.07), 'Low', 'Medium'))

    # Add the 'Classification' column to the DataFrame
    df['Classification'] = np.where(df['Adherence'] == 'High', 'Genuine Soil',
                                     np.where(df['Adherence'] == 'Medium', 'Adopted Soil', 'Rejected Soil'))

    # Load the SoloSorted_Head data
    df_head = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'SoilSorted_Head.xlsx'))

    # Merge DataFrames using the 'code' column as the joining key, with df_head coming first
    df = pd.merge(df_head, df, on='code', how='inner')

    # Save the merged DataFrame to a new Excel file
    df.to_excel(os.path.join(DOWNLOAD_FOLDER,'Adherence.xlsx'), index=False)

     # Additional processing and classification
    import pandas as pd

    # Load the Adherence.xlsx file
    df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER,'Adherence.xlsx'))

    def classify_suborder(row):
        diff = row['theta0'] - row['ThetaR']
        if 0 < diff <= 0.20:
            return 1
        elif 0.20 < diff <= 0.40:
            return 2
        elif 0.40 < diff <= 0.60:
            return 3
        elif diff > 0.60:
            return 4
        else:
            return None

    # Function to classify the main order
    def classify_order(row):
        if 2/3 <= row["A60"] < 1 and 0 < row['W60-W15000'] < 1/3 and 0 < row['W15000'] < 1/3:
            return "A"
        elif 1/3 <= row["A60"] < 2/3 and 0 < row['W60-W15000'] < 1/3 and 0 < row['W15000'] < 1/3:
            return "D"
        elif 1/3 <= row["A60"] < 2/3 and 1/3 <= row['W60-W15000'] < 2/3 and 0 < row['W15000'] < 1/3:
            return "B"
        elif 0 < row["A60"] < 1/3 and 2/3 <= row['W60-W15000'] < 1 and 0 < row['W15000'] < 1/3:
            return "C"
        elif 0 < row["A60"] < 1/3 and 1/3 <= row['W60-W15000'] < 2/3 and 0 < row['W15000'] < 1/3:
            return "E"
        elif 0 < row["A60"] < 1/3 and 1/3 <= row['W60-W15000'] < 2/3 and 1/3 <= row['W15000'] < 2/3:
            return "G"
        elif 0 < row["A60"] < 1/3 and 0 < row['W60-W15000'] < 1/3 and 2/3 <= row['W15000'] < 1:
            return "I"
        elif 0 < row["A60"] < 1/3 and 0 < row['W60-W15000'] < 1/3 and 1/3 <= row['W15000'] < 2/3:
            return "H"
        elif 1/3 <= row["A60"] < 2/3 and 0 < row['W60-W15000'] < 1/3 and 1/3 <= row['W15000'] < 2/3:
            return "F"
        else:
            return "Error"

    # Apply the functions to classify order and subordinate order
    df["Order"] = df.apply(classify_order, axis=1)
    df["Suborder"] = df.apply(classify_suborder, axis=1)

    # Ensure Suborder is an integer
    df["Suborder"] = df["Suborder"].fillna(0).astype(int)

    # Create the 'Family' column
    df["Family"] = df["Order"] + df["Suborder"].astype(str)

    # Dicionário para Suborder
    suborder_dict = {
        1: "Low Effective Porosity",
        2: "Moderate Effective Porosity",
        3: "High Effective Porosity",
        4: "Very High Effective Porosity"
    }

    # Dicionário para Order
    order_dict = {
        "A": "Highly Macrospacious Soil",
        "C": "Highly Mesospacious Soil",
        "I": "Highly Microspacious Soil",
        "D": "Macrospacious Soil",
        "E": "Mesospacious Soil",
        "H": "Microspacious Soil",
        "B": "Macro-Mesospacious Soil",
        "F": "Macro-Microspacious Soil",
        "G": "Meso-Microspacious Soil"
    }

    # Função para criar a nomenclatura da família
    def get_family_nomenclature(row):
        porosity = suborder_dict.get(row['Suborder'], "Unknown Porosity")
        order_name = order_dict.get(row['Order'], "Unknown Soil Type")
        return f"{porosity} - {order_name}"

    # Aplicar a função para criar a nova coluna
    df["Family Nomenclature"] = df.apply(get_family_nomenclature, axis=1)

    # Criar coluna 'a (cm³/cm³)' = theta0 - theta2
    df['a (cm³/cm³)'] = df['theta0'] - df['theta2']

    # Criar coluna 'w (cm³/cm³)' = theta2 - theta3
    df['w (cm³/cm³)'] = df['theta2'] - df['theta3']

    # Criar coluna 'Ksat (cm/d)' = 1931 * (a ^ 1.948)
    df['Ksat (cm/d)'] = 1931 * (df['a (cm³/cm³)'] ** 1.948)


    def classify_a_value(a):
        if a < 0.10:
            return "Low"
        elif 0.10 <= a <= 0.20:
            return "Moderate"
        elif a > 0.20:
            return "High"
        else:
            return "Undefined"

    df['a value'] = df['a (cm³/cm³)'].apply(classify_a_value)

    def classify_w_value(w):
        if w < 0.06:
            return "Low"
        elif 0.06 <= w <= 0.12:
            return "Moderate"
        elif w > 0.12:
            return "High"
        else:
            return "Undefined"

    df['w value'] = df['w (cm³/cm³)'].apply(classify_w_value)

    def classify_ksat_value(ksat):
        if ksat <= 12:
            return "Slow"
        elif 12 < ksat <= 120:
            return "Moderate"
        elif ksat > 120:
            return "Rapid"
        else:
            return "Undefined"

    df['ksat value'] = df['Ksat (cm/d)'].apply(classify_ksat_value)
    
    def classify_hydraulic_class(row):
        a_val = row['a value']
        w_val = row['w value']
        ksat_val = row['ksat value']

        # Primeira parte - restrição
        if (a_val == "Low" and w_val in ["Low", "Moderate", "High"]):
            restriction = "Soil with air restriction"
        elif (a_val in ["High", "Moderate"] and w_val == "Low"):
            restriction = "Soil with water restriction"
        elif ((a_val == "Moderate" and w_val == "High") or
              (a_val == "High" and w_val in ["Moderate", "High"]) or
              (a_val == "Moderate" and w_val == "Moderate")):
            restriction = "Soil without air and water restriction"
        else:
            restriction = "Undefined Restriction"

        # Segunda parte - permeabilidade
        if ksat_val == "Slow":
            permeability = "slow permeability"
        elif ksat_val == "Moderate":
            permeability = "moderate permeability"
        elif ksat_val == "Rapid":
            permeability = "rapid permeability"
        else:
            permeability = "undefined permeability"

        return f"{restriction} and {permeability}"

    # Aplicar a classificação para cada linha
    df['Hydraulic Class'] = df.apply(classify_hydraulic_class, axis=1)

    df.drop(columns=['a value', 'w value', 'ksat value'], inplace=True)

    # Rename the columns as requested
    df.rename(columns={"RMSE": "RMSE_3points", "TOTALRMSE": "RMSE_30cm-15000cm"}, inplace=True)

    # Reorder the columns to have 'Sample_ID' as the first column
    columns_order = ['Sample_ID'] + [col for col in df.columns if col != 'Sample_ID']
    df = df[columns_order]

    # Convert the column SAMPLE_ID to text
    df['Sample_ID'] = df['Sample_ID'].astype(str)

    # Save the DataFrame to an Excel file
    df.to_excel(os.path.join(DOWNLOAD_FOLDER,'YourSoilClassified.xlsx'), index=False)
    
    from matplotlib.backends.backend_pdf import PdfPages
    soilfull_df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilFull.xlsx'))
    adherence_df = pd.read_excel(os.path.join(DOWNLOAD_FOLDER, 'Adherence.xlsx'))

    # Limpeza de memória após o uso
    gc.collect()

    # Convertendo colunas relevantes para log10
    soilfull_df['log_h'] = np.log10(soilfull_df['h'].replace(0, np.nan))  # Transformando h e substituindo 0 por NaN
    soilfull_df['Theta'] = soilfull_df['theta']  # Assumindo que a coluna se chama 'theta'

    unique_codes = soilfull_df['code'].unique()

    # Criar um arquivo PDF para armazenar os plots
    pdf_path = os.path.join(PLOTS_FOLDER, 'retentioncurveplots.pdf')
    
    with PdfPages(pdf_path) as pdf:
        num_plots_per_page = 9
        plots_per_row = 3
        rows = num_plots_per_page // plots_per_row

        for i, code in enumerate(unique_codes):
            if i % num_plots_per_page == 0:  # Inicia uma nova página
                fig, axes = plt.subplots(rows, plots_per_row, figsize=(15, 15))
                axes = axes.flatten()  # Para facilitar o acesso aos eixos

            code_df = soilfull_df[soilfull_df['code'] == code]
            sample_id = adherence_df[adherence_df['code'] == code]['Sample_ID'].values[0]
            
            # Criar um gráfico para cada código
            ax = axes[i % num_plots_per_page]  # Seleciona o eixo correto
            plot_color = 'black'  # Cor dos pontos

            # Plotar todos os pontos de theta
            for index, row in code_df.iterrows():
                if row['h'] == 0:
                    ax.scatter(0, row['Theta'], color=plot_color, marker='x', s=100, label='φ', clip_on=False)  # Ponto para h=0 como "X" grande identificado como φ
                else:
                    ax.scatter(row['log_h'], row['Theta'], color=plot_color, alpha=0.6)

            # Calcular e plotar a reta da equação de Van Genuchten apenas para h > 30
            adherence_row = adherence_df[adherence_df['code'] == code]
            if not adherence_row.empty:
                # Extraindo parâmetros
                ThetaR = adherence_row['ThetaR'].values[0]
                ThetaS = adherence_row['ThetaS'].values[0]
                alpha = adherence_row['alpha'].values[0]
                n = adherence_row['n'].values[0]

                # Filtrando h maior que 30 para a reta
                h_values_above_30 = np.linspace(30, soilfull_df['h'].max(), 100)
                Theta_values_above_30 = ThetaR + (ThetaS - ThetaR) / (1 + (alpha * h_values_above_30) ** n) ** (1 - 1 / n)

                # Adicionar valores de theta1, theta2, theta3 como quadrados sólidos
                ax.scatter(np.log10(adherence_row['h1'].values[0]), adherence_row['theta1'].values[0], color='black', marker='s', s=100)  # Quadrado sólido para θ1
                ax.scatter(np.log10(adherence_row['h2'].values[0]), adherence_row['theta2'].values[0], color='black', marker='s', s=100)  # Quadrado sólido para θ2
                ax.scatter(np.log10(adherence_row['h3'].values[0]), adherence_row['theta3'].values[0], color='black', marker='s', s=100)  # Quadrado sólido para θ3

                # Plotando a reta apenas para h > 30 em cinza
                ax.plot(np.log10(h_values_above_30), Theta_values_above_30, linestyle='--', color='gray')

                # Arredondando valores para a legenda
                error_max = round(adherence_row['ERRORMAX'].values[0], 3)
                total_rmse = round(adherence_row['TOTALRMSE'].values[0], 3)

                # Adicionando informações à legenda com o "X" para φ e quadrados para θ1, θ2, θ3
                handles = [
                    plt.Line2D([0], [0], color='w', label=f'ERRORMAX: {error_max} cm³/cm³'),  # ERRORMAX
                    plt.Line2D([0], [0], color='w', label=f'RRMSE$_{{30-18000}}$: {total_rmse} cm³/cm³'),  # RMSE
                    plt.Line2D([0], [0], marker='s', color='black', label='θ1, θ2, θ3', markerfacecolor='black', markersize=10, linestyle='None'),  # Quadrados sólidos para θ1, θ2, θ3
                    plt.Line2D([0], [0], marker='x', color='black', label='φ', markersize=10, linestyle='None'),
                    plt.Line2D([0], [0], marker='o', color='gray', label='θ', markersize=6, linestyle='None')
                ]
                ax.legend(handles=handles, loc='upper right')

            # Configurações do gráfico
            ax.axhline(0, color='black', lw=0.8)  # Linha horizontal no eixo y=0
            ax.axvline(0, color='black', lw=0.8)  # Linha vertical no eixo x=0
            ax.set_xlabel("log h (cm)")
            ax.set_ylabel("Theta (cm³/cm³)")
            ax.set_title(f"Retention curve - Soil {sample_id}")  # Usando SampleID como título
            ax.grid()
            ax.set_xlim(left=0)  # Garantir que o eixo x comece em 0
            ax.set_ylim(0, 1.0)  # Garantir que o eixo y comece em 0

            if (i + 1) % num_plots_per_page == 0 or i == len(unique_codes) - 1:  # Se a página estiver cheia ou for o último gráfico
                # Remover eixos vazios para evitar gráficos em branco
                for j in range((i % num_plots_per_page) + 1, num_plots_per_page):
                    fig.delaxes(axes[j])  # Remove os eixos não utilizados

                plt.tight_layout()
                pdf.savefig(fig)  # Salva a figura no PDF
                plt.close(fig)    # Fecha a figura para liberar memória
        


    

    # Function to classify the subordinate order
    def classify_suborder(row):
        diff = row['theta0'] - row['ThetaR']
        if 0 < diff <= 0.20:
            return 1
        elif 0.20 < diff <= 0.40:
            return 2
        elif 0.40 < diff <= 0.60:
            return 3
        elif diff > 0.60:
            return 4

    def map_color(adherence):
        if adherence == 'High':
            return 'green'
        elif adherence == 'Medium':
            return '#FFC300'  # Amarelo mais brilhante
        elif adherence == 'Low':
            return 'red'

    def plot_ternary(df, title, filename):
        scale = 100
        try:
            figure, tax = ternary.figure(scale=scale)
        except AttributeError:
            raise Exception("A função 'figure' não está disponível no módulo 'ternary'. Verifique a instalação.")

        figure.set_size_inches(5, 5)
        
        # Adiciona um grid de 81 triângulos menores com cinza mais escuro
        step = scale / 9
        grid_color = '#a9a9a9'  # Cinza mais escuro

        for i in range(9):
            for j in range(9 - i):
                x1, y1 = i * step, j * step
                x2, y2 = (i + 1) * step, j * step
                x3, y3 = i * step, (j + 1) * step

                # Cinza mais escuro e pontilhado
                tax.line((x1, y1), (x2, y2), color=grid_color, linestyle='--', linewidth=0.5)
                tax.line((x2, y2), (x3, y3), color=grid_color, linestyle='--', linewidth=0.5)
                tax.line((x3, y3), (x1, y1), color=grid_color, linestyle='--', linewidth=0.5)

        # Agora desenhamos o triângulo principal sobre o grid para garantir que ele fique por cima
        tax.boundary(linewidth=1.0)

        fontsize = 12
        offset = 0.2
        tax.left_axis_label("    W15000 (%)", fontsize=fontsize, offset=offset, rotation=-240)
        tax.right_axis_label("W60-W15000 (%)", fontsize=fontsize, offset=offset, rotation=240)
        tax.bottom_axis_label("      A60 (100-W60) (%)", fontsize=fontsize, offset=offset, rotation=180)

        # Adicionando os textos nas laterais do triângulo, a 10 pixels de distância
        # Ajustando as posições dos textos para que estejam nas laterais corretas e a distância adequada
        tax.annotate("Microspace (%)", position=(62, 54), fontsize=10, rotation=60, va='center', ha='center')  # Lado esquerdo
        tax.annotate("Mesospace (%)", position=(-11.6, 54), fontsize=10, rotation=-60, va='center', ha='center')   # Lado direito
        tax.annotate("Macrospace (%)", position=(54, -11.6), fontsize=10, rotation=0, va='center', ha='center')      # Parte inferior

        for index, row in df.iterrows():
            color = map_color(row['Adherence'])
            tax.scatter([[row['A60%'], row['W15000%'], row['W60%-W15000%']]], marker='o', s=35, color=color)

        plt.gca().invert_xaxis()
        triangles = [((66.7, 0), (66.7, 33.3)), ((33.3, 0), (33.3, 66.7)), ((66.7, 33.3), (0, 33.3)),
                     ((33.3, 66.7), (0, 66.7)), ((0, 66.7), (66.7, 0)), ((0, 33.3), (33.3, 0))]
        for p1, p2 in triangles:
            tax.line(p1, p2, linewidth=1, color='black')

        titles = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
        midpoints = [(81, 9, 5), (47, 9.5), (13, 9.5), (59, 20), (24, 20), (45, 45), (13, 45), (26, 52), (12, 78)]
        for title_point, point in zip(titles, midpoints):
            tax.annotate(title_point, point)

        plt.title(title, pad=20)

        # Ajustando os ticks para os pontos desejados
        tax.ticks(axis='lbr', ticks=[0, 33.3, 66.7, 100], linewidth=1, offset=0.025)
        tax.get_axes().axis('off')
        tax.clear_matplotlib_ticks()

        legend_elements = [
            plt.Line2D([0], [0], marker='o', color='w', label='Genuine Soil', markerfacecolor='green', markersize=5),
            plt.Line2D([0], [0], marker='o', color='w', label='Adopted Soil', markerfacecolor='#FFC300', markersize=5),
            plt.Line2D([0], [0], marker='o', color='w', label='Rejected Soil', markerfacecolor='red', markersize=5)
        ]
        plt.legend(handles=legend_elements, loc='upper left', title="Adherence", fontsize='small')

        plt.savefig(filename, format='jpg')
        plt.close()

    # Criar um diretório para armazenar os plots
    os.makedirs(PLOTS_FOLDER, exist_ok=True)

    # Plotar o gráfico ternário
    plot_ternary(df, "ALL SUBORDERS", os.path.join(PLOTS_FOLDER, 'allsuborders_plot.jpg'))

    # Plotar gráficos adicionais para cada ordem
    for order in range(1, 5):
        filtered_df = df[df.apply(lambda row: classify_suborder(row) == order, axis=1)]
        plot_ternary(filtered_df, f"Suborder {order}", os.path.join(PLOTS_FOLDER, f'suborder_{order}_plot.jpg'))


    files_to_include = [
        'allsuborders_plot.jpg',
        'suborder_1_plot.jpg',
        'suborder_2_plot.jpg',
        'suborder_3_plot.jpg',
        'suborder_4_plot.jpg',
        'retentioncurveplots.pdf'
    ]

    # Criando o arquivo ZIP com os arquivos especificados
    with zipfile.ZipFile(os.path.join(DOWNLOAD_FOLDER, 'ternary_plots.zip'), 'w') as zipf:
        for root, _, files in os.walk(PLOTS_FOLDER):
            for file in files:
                if file in files_to_include:  # Inclui apenas os arquivos desejados
                    zipf.write(os.path.join(root, file), file)


    your_soil_classified = pd.read_excel(os.path.join(DOWNLOAD_FOLDER, 'YourSoilClassified.xlsx'))
    soils_discarded = pd.read_excel(os.path.join(DOWNLOAD_FOLDER, 'SoilsDiscarted.xlsx'))


# Fazer a união dos DataFrames com uma junção externa, excluindo 'codedis'
    merged_df = pd.merge(your_soil_classified, soils_discarded[['Sample_ID', 'OBS']], 
                     how='outer', on='Sample_ID')

# Reorganizar as colunas para que 'OBS' venha logo após 'Sample_ID'
    column_order = ['Sample_ID', 'OBS'] + [col for col in merged_df.columns if col not in ['Sample_ID', 'OBS']]
    merged_df = merged_df[column_order]

    # Substituir o DataFrame original
    your_soil_classified = merged_df

    columns_to_remove = ["code", "Adherence", "W60", "W15000", "A60", "W60-W15000"]
    your_soil_classified = your_soil_classified.drop(columns=[col for col in columns_to_remove if col in your_soil_classified.columns])

    your_soil_classified = your_soil_classified.rename(columns={"Sample_ID": "code"})


    with pd.ExcelWriter(os.path.join(DOWNLOAD_FOLDER, 'YourSoilClassified.xlsx'), engine='openpyxl') as writer:
        your_soil_classified.to_excel(writer, index=False)

    # Acessar o workbook e a worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']  # Altere para o nome da sua planilha, se necessário

    # Ajustar a largura da coluna e permitir quebra de linha
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Obter a letra da coluna
            for cell in column:
                try:
                     if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width
            for cell in column:
                cell.alignment = cell.alignment.copy(horizontal='center', wrap_text=True)  # Centraliza o texto e ativa a quebra de linha
     


@app.route('/assets/<path:filename>')  # Nova rota para servir arquivos da pasta assets
def assets(filename):
    return send_from_directory(ASSETS_FOLDER, filename)


@app.route('/download_xlsx')
def download_xlsx():
    return send_from_directory(DOWNLOAD_FOLDER, 'YourSoilClassified.xlsx', as_attachment=True)

@app.route('/download_zip')
def download_zip():
    return send_from_directory(DOWNLOAD_FOLDER, 'ternary_plots.zip', as_attachment=True)

@app.route('/ternary_plots/<filename>')
def ternary_plots(filename):
    return send_from_directory(PLOTS_FOLDER, filename)


if __name__ == '__main__':
    app.run(debug=True)
