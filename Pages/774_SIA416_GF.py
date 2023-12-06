import streamlit as st
import pandas as pd
import plotly.express as px
import io
import base64
from PIL import Image
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# ========== FUNCTIONS ==========

def true_if_present(series):
    return pd.Series(True, index=series)

def get_Wohnen(GF_dfsia, GF_Mobimo):
    for floor in GF_Mobimo['Geschossflächen oberirdisch'].unique():
        for raumname in ['Bestand', 'Neubau']:
            # Apply the filters with the current floor and raumname
            filtered_df = GF_dfsia[(GF_dfsia['Ursprungsgeschoss Name'] == floor) & 
                                   (GF_dfsia['Raumname'] == raumname) & 
                                   (GF_dfsia['Nutzer (Zeichenfolge)'] == 'Wohnen')]

            # Sum the values from 'Berechnete Fläche (NRF)'
            sum_of_values = filtered_df['Berechnete Fläche (NRF)'].sum()

            # Update the corresponding column in GF_Mobimo based on raumname
            if raumname == 'Bestand':
                GF_Mobimo.loc[GF_Mobimo['Geschossflächen oberirdisch'] == floor, 'Wohnen Bestand'] = sum_of_values
            elif raumname == 'Neubau':
                GF_Mobimo.loc[GF_Mobimo['Geschossflächen oberirdisch'] == floor, 'Wohnen Neubau'] = sum_of_values

    return GF_Mobimo

def get_Wohnen_Gewerbe(GF_dfsia, GF_Mobimo):
    for floor in GF_Mobimo['Geschossflächen oberirdisch'].unique():
        for category in ['Wohnen', 'Gewerbe']:
            for raumname in ['Bestand', 'Neubau']:
                # Apply the filters with the current floor, category, and raumname
                filtered_df = GF_dfsia[(GF_dfsia['Ursprungsgeschoss Name'] == floor) & 
                                       (GF_dfsia['Raumname'] == raumname) & 
                                       (GF_dfsia['Nutzer (Zeichenfolge)'] == category)]

                # Sum the values from 'Berechnete Fläche (NRF)'
                sum_of_values = filtered_df['Berechnete Fläche (NRF)'].sum()

                # Update the corresponding column in GF_Mobimo based on category and raumname
                col_name = f"{category} {raumname}"  # e.g., 'Wohnen Bestand', 'Gewerbe Neubau'
                GF_Mobimo.loc[GF_Mobimo['Geschossflächen oberirdisch'] == floor, col_name] = sum_of_values

    return GF_Mobimo

def get_Wohnen_Gewerbe_Gemeinschaft(GF_dfsia, GF_Mobimo):
    for floor in GF_Mobimo['Geschossflächen oberirdisch'].unique():
        for category in ['Wohnen', 'Gewerbe', 'Gemeinschaft']:
            for raumname in ['Bestand', 'Neubau']:
                # Apply the filters with the current floor, category, and raumname
                filtered_df = GF_dfsia[(GF_dfsia['Ursprungsgeschoss Name'] == floor) & 
                                       (GF_dfsia['Raumname'] == raumname) & 
                                       (GF_dfsia['Nutzer (Zeichenfolge)'] == category)]

                # Sum the values from 'Berechnete Fläche (NRF)'
                sum_of_values = filtered_df['Berechnete Fläche (NRF)'].sum()

                # Update the corresponding column in GF_Mobimo based on category and raumname
                col_name = f"{category} {raumname}"  # e.g., 'Wohnen Bestand', 'Gemeinschaft Neubau'
                GF_Mobimo.loc[GF_Mobimo['Geschossflächen oberirdisch'] == floor, col_name] = sum_of_values

    return GF_Mobimo

def get_GV_Wohnen_Gewerbe_Gemeinschaft(GF_dfsia, GF_Mobimo):
    for floor in GF_Mobimo['Geschossflächen oberirdisch'].unique():
        for category in ['Wohnen', 'Gewerbe', 'Gemeinschaft']:
            for raumname in ['Bestand', 'Neubau']:
                # Apply the filters with the current floor, category, and raumname
                filtered_df = GF_dfsia[(GF_dfsia['Ursprungsgeschoss Name'] == floor) & 
                                       (GF_dfsia['Raumname'] == raumname) & 
                                       (GF_dfsia['Nutzer (Zeichenfolge)'] == category)]

                # Sum the values from 'Berechnete Fläche (NRF)'
                sum_of_values = filtered_df['Volumen (netto)'].sum()

                # Update the corresponding column in GF_Mobimo based on category and raumname
                col_name = f"{category} {raumname}"  # e.g., 'Wohnen Bestand', 'Gemeinschaft Neubau'
                GF_Mobimo.loc[GF_Mobimo['Geschossflächen oberirdisch'] == floor, col_name] = sum_of_values

    return GF_Mobimo

def get_UG_Wohnen_Gewerbe_Gemeinschaft(df_input, df_output):
    for floor in df_output['Geschossflächen oberirdisch'].unique():
        for category in ['Keller', 'Keller Gewerbe', 'Keller Technik', 'Hobbyräume']:
            for raumname in ['Bestand', 'Neubau']:
                # Apply the filters with the current floor, category, and raumname
                filtered_df = df_input[(df_input['Ursprungsgeschoss Name'] == floor) & 
                                       (df_input['Raumname'] == raumname) & 
                                       (df_input['Nutzer (Zeichenfolge)'] == category)]

                # Sum the values from 'Berechnete Fläche (NRF)'
                sum_of_values = filtered_df['Berechnete Fläche (NRF)'].sum()

                # Update the corresponding column in df_output based on category and raumname
                col_name = f"{category} {raumname}"
                df_output.loc[df_output['Geschossflächen oberirdisch'] == floor, col_name] = sum_of_values

    return df_output

def get_GV_UG_Wohnen_Gewerbe_Gemeinschaft(df_input, df_output):
    for floor in df_output['Geschossflächen oberirdisch'].unique():
        for category in ['Keller', 'Keller Gewerbe', 'Keller Technik', 'Hobbyräume']:
            for raumname in ['Bestand', 'Neubau']:
                # Apply the filters with the current floor, category, and raumname
                filtered_df = df_input[(df_input['Ursprungsgeschoss Name'] == floor) & 
                                       (df_input['Raumname'] == raumname) & 
                                       (df_input['Nutzer (Zeichenfolge)'] == category)]

                # Sum the values from 'Berechnete Fläche (NRF)'
                sum_of_values = filtered_df['Volumen (netto)'].sum()

                # Update the corresponding column in df_output based on category and raumname
                col_name = f"{category} {raumname}"
                df_output.loc[df_output['Geschossflächen oberirdisch'] == floor, col_name] = sum_of_values

    return df_output

def get_UG_Parking(df_input, df_output):
    for floor in df_output['Geschossflächen oberirdisch'].unique():
        for category in ['Parking']:
            for raumname in ['Bestand', 'Neubau']:
                # Apply the filters with the current floor, category, and raumname
                filtered_df = df_input[(df_input['Ursprungsgeschoss Name'] == floor) & 
                                       (df_input['Raumname'] == raumname) & 
                                       (df_input['Nutzer (Zeichenfolge)'] == category)]

                # Sum the values from 'Berechnete Fläche (NRF)'
                sum_of_values = filtered_df['Berechnete Fläche (NRF)'].sum()

                # Update the corresponding column in df_output based on category and raumname
                col_name = f"{category} {raumname}"
                df_output.loc[df_output['Geschossflächen oberirdisch'] == floor, col_name] = sum_of_values

    return df_output

def get_GV_UG_Parking(df_input, df_output):
    for floor in df_output['Geschossflächen oberirdisch'].unique():
        for category in ['Parking']:
            for raumname in ['Bestand', 'Neubau']:
                # Apply the filters with the current floor, category, and raumname
                filtered_df = df_input[(df_input['Ursprungsgeschoss Name'] == floor) & 
                                       (df_input['Raumname'] == raumname) & 
                                       (df_input['Nutzer (Zeichenfolge)'] == category)]

                # Sum the values from 'Berechnete Fläche (NRF)'
                sum_of_values = filtered_df['Volumen (netto)'].sum()

                # Update the corresponding column in df_output based on category and raumname
                col_name = f"{category} {raumname}"
                df_output.loc[df_output['Geschossflächen oberirdisch'] == floor, col_name] = sum_of_values

    return df_output

def add_columns(df, output, input_1, input_2):
    df[output] = df[input_1] + df[input_2] 
    return df

def convert_columns_to_float(GF_Mobimo):
    # Check if the DataFrame is not empty
    if not GF_Mobimo.empty:
        for column in GF_Mobimo.columns:
            if column != 'Geschossflächen oberirdisch':
                # Convert each value individually, handling apostrophes and potential formatting issues
                GF_Mobimo[column] = GF_Mobimo[column].apply(lambda x: pd.to_numeric(x.replace("'", "") if isinstance(x, str) else x, errors='coerce'))
    return GF_Mobimo

def convert_selected_columns_to_float(GF_Mobimo):
    # List of specific columns to convert
    columns_to_convert = ['Berechnete Fläche (NRF)', 'Fläche', 'Volumen (netto)', 'Höhe']

    # Check if the DataFrame is not empty
    if not GF_Mobimo.empty:
        for column in columns_to_convert:
            if column in GF_Mobimo.columns:
                # Convert each value individually, handling apostrophes and potential formatting issues
                GF_Mobimo[column] = GF_Mobimo[column].apply(lambda x: pd.to_numeric(x.replace("'", "") if isinstance(x, str) else x, errors='coerce'))
    return GF_Mobimo

def calculate_totals(df_input1, df_input2, df_output, input, row):
    # Sum values from 'Wohnen Bestand' in GF_Mobimo and AGF_Mobimo
    total_wohnen_bestand = df_input1[input].sum() + df_input2[input].sum()

    # Find the row index in GF_Total where 'Total Geschossfläche' is 'Geschossfläche oberirdisch inkl. AGF'
    row_index = df_output[df_output['Total Geschossfläche'] == row].index

    # Check if there's exactly one row that matches the condition
    if len(row_index) == 1:
        # Set the value in 'Wohnen Bestand' at the specific row
        df_output.at[row_index[0], input] = total_wohnen_bestand
    else:
        # Handle the case where the condition is met by zero or multiple rows
        st.write("Error: The condition is not met by a unique row.")
    return df_output

def sia416_to_excel(df1, df2, df3, df4, df5, df6, df7, df8, df9, df10):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='GF oberirdisch', index=False)
    df2.to_excel(writer, sheet_name='AGF', index=False)
    df3.to_excel(writer, sheet_name='GF unterirdisch exkl. PP', index=False)
    df4.to_excel(writer, sheet_name='GF Parkplätze', index=False)
    df5.to_excel(writer, sheet_name='Total GF', index=False)
    df6.to_excel(writer, sheet_name='GV oberirdisch', index=False)
    df7.to_excel(writer, sheet_name='AGV', index=False)
    df8.to_excel(writer, sheet_name='GV unterirdisch exkl. PP', index=False)
    df9.to_excel(writer, sheet_name='GV Parkplätze', index=False)
    df10.to_excel(writer, sheet_name='Total GV', index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# ========== MAIN APP ==========

st.title("774 SIA 416 Geschossflächen und Gebäudevolumen")

uploaded_file_sia = st.file_uploader("Choose a file", type='xlsx', key="file_uploader_2")

if uploaded_file_sia is not None:

    dfsia = pd.read_excel(uploaded_file_sia, skiprows=1)
    
    with st.expander("Modellierungsanleitung"):
        st.markdown("""**WICHTIGE ANMERKUNGEN:**  \n
                    PARKING:  \n
                    1. Parkingflächen werden nur in den Untergeschosse gerechnet.  \n 
                       Falls Parkingflächen in Obergeschossen geplannt werden, müssen die Geschosse in Python-Code ergänzt werden.  \n
                    2. Sämtliche Parkingflächen werden an Kategorie "Wohnen" zugeweisen.  \n
                       Falls weitere Kategorien, z.B. "Gewerbe" nötig sind, muss Python-Code und Modell angepasst werden. 
                    """)
    
    with st.expander("INPUT: ArchiCAD XLSX-Liste"):
        st.dataframe(dfsia)

    with st.expander("03 GESCHOSSFLÄCHEN (gem. SIA 416)"):
        GF_dfsia = dfsia[dfsia["Raumkategoriename"] == "Geschossfläche"].copy()
        AGF_dfsia = dfsia[dfsia["Raumkategoriename"] == "Aussen-Geschossfläche"].copy()

        columns = [
            'Geschossflächen oberirdisch', 'Wohnen Bestand', 'Wohnen Neubau', 'Wohnen Total',
            'Gewerbe Bestand', 'Gewerbe Neubau', 'Gewerbe Total',
            'Gemeinschaft Bestand', 'Gemeinschaft Neubau', 'Gemeinschaft Total'
        ]
        #========== Get GF (Geschossflächen oberirdisch) table ==========
        GF_Mobimo = pd.DataFrame(columns=columns)        
        unique_storey_names = GF_dfsia['Ursprungsgeschoss Name'].unique()        
        GF_Mobimo['Geschossflächen oberirdisch'] = unique_storey_names
        numeric_columns = columns[1:]  # All columns except the first one
        GF_Mobimo[numeric_columns] = GF_Mobimo[numeric_columns].fillna(0)
        GF_dfsia = convert_selected_columns_to_float(GF_dfsia)      
        GF_Mobimo = get_Wohnen_Gewerbe_Gemeinschaft(GF_dfsia, GF_Mobimo)        
        GF_Mobimo['Wohnen Total'] = GF_Mobimo['Wohnen Bestand'] + GF_Mobimo['Wohnen Neubau'] 
        GF_Mobimo['Gewerbe Total'] = GF_Mobimo['Gewerbe Bestand'] + GF_Mobimo['Gewerbe Neubau'] 
        GF_Mobimo['Gemeinschaft Total'] = GF_Mobimo['Gemeinschaft Bestand'] + GF_Mobimo['Gemeinschaft Neubau'] 
        GF_Mobimo = GF_Mobimo[~GF_Mobimo['Geschossflächen oberirdisch'].isin(['2. UG', '1. UG'])]
        st.header("Geschossflächen oberirdisch")
        st.dataframe(GF_Mobimo)

        #========== Get AGF (Aussengeschossflächen) table ==========
        AGF_dfsia = convert_selected_columns_to_float(AGF_dfsia)  
        AGF_Mobimo = pd.DataFrame(columns=columns)
        AGF_Mobimo['Geschossflächen oberirdisch'] = unique_storey_names    
        AGF_Mobimo = get_Wohnen_Gewerbe_Gemeinschaft(AGF_dfsia, AGF_Mobimo)
        AGF_Mobimo = convert_columns_to_float(AGF_Mobimo)
        AGF_Mobimo[numeric_columns] = AGF_Mobimo[numeric_columns].fillna(0)
        AGF_Mobimo = AGF_Mobimo[~AGF_Mobimo['Geschossflächen oberirdisch'].isin(['2. UG', '1. UG'])]
        AGF_Mobimo['Wohnen Total'] = AGF_Mobimo['Wohnen Bestand'] + AGF_Mobimo['Wohnen Neubau'] 
        AGF_Mobimo['Gewerbe Total'] = AGF_Mobimo['Gewerbe Bestand'] + AGF_Mobimo['Gewerbe Neubau'] 
        AGF_Mobimo['Gemeinschaft Total'] = AGF_Mobimo['Gemeinschaft Bestand'] + AGF_Mobimo['Gemeinschaft Neubau'] 
        AGF_Mobimo.rename(columns={'Geschossflächen oberirdisch': 'Aussengeschossflächen'}, inplace=True)
        st.header("Aussengeschossflächen")
        st.dataframe(AGF_Mobimo)

        #========== Get GF_UG (Geschossflächen unterirdisch exkl. PP) table ==========
        st.header("Geschossflächen unterirdisch exkl. PP")
        values_to_remove = ['EG', "ZG (EG)", '1. OG', '2. OG', '3. OG', '4. OG', '5. OG', '6. OG', '7. OG']     
        GF_UG_Mobimo_dfsia = GF_dfsia[~GF_dfsia['Ursprungsgeschoss Name'].isin(values_to_remove)]
        GF_UG_Mobimo = pd.DataFrame(columns=columns)
        GF_UG_Mobimo['Geschossflächen oberirdisch'] = unique_storey_names        
        GF_UG_Mobimo = get_UG_Wohnen_Gewerbe_Gemeinschaft(GF_UG_Mobimo_dfsia, GF_UG_Mobimo)
        GF_UG_Mobimo = GF_UG_Mobimo[~GF_UG_Mobimo['Geschossflächen oberirdisch'].isin(values_to_remove)]
        GF_UG_Mobimo.rename(columns={'Geschossflächen oberirdisch': 'Geschossflächen unterirdisch exkl. PP'}, inplace=True)
        if 'Keller Bestand' in GF_UG_Mobimo.columns:
            GF_UG_Mobimo['Wohnen Bestand'] = GF_UG_Mobimo['Keller Bestand']
        if 'Keller Neubau' in GF_UG_Mobimo.columns:
            GF_UG_Mobimo['Wohnen Neubau'] = GF_UG_Mobimo['Keller Neubau']
        if 'Keller Gewerbe Bestand' in GF_UG_Mobimo.columns:
            GF_UG_Mobimo['Gewerbe Bestand'] = GF_UG_Mobimo['Keller Gewerbe Bestand']
        if 'Keller Gewerbe Neubau' in GF_UG_Mobimo.columns:
            GF_UG_Mobimo['Gewerbe Neubau'] = GF_UG_Mobimo['Keller Gewerbe Neubau']
        GF_UG_Mobimo[numeric_columns] = GF_UG_Mobimo[numeric_columns].fillna(0)
        GF_UG_Mobimo['Wohnen Bestand'] = GF_UG_Mobimo['Wohnen Bestand'] + GF_UG_Mobimo['Keller Technik Bestand'] + GF_UG_Mobimo['Hobbyräume Bestand']
        GF_UG_Mobimo['Wohnen Neubau'] = GF_UG_Mobimo['Wohnen Neubau'] + GF_UG_Mobimo['Keller Technik Neubau'] + GF_UG_Mobimo['Hobbyräume Neubau']

        # Create a list to hold new rows
        new_rows = []

        for index, row in GF_UG_Mobimo.iterrows():
            if row['Gewerbe Bestand'] > 0 or row['Gewerbe Neubau'] > 0:
                # Create a new row with all values set to 0
                new_row = {col: 0 for col in GF_UG_Mobimo.columns}

                # Copy the value from 'Geschossflächen unterirdisch exkl. PP' and add suffix
                new_row['Geschossflächen unterirdisch exkl. PP'] = row['Geschossflächen unterirdisch exkl. PP'] + " Gewerbe"

                # Copy the value greater than 0
                if row['Gewerbe Bestand'] > 0:
                    new_row['Gewerbe Bestand'] = row['Gewerbe Bestand']
                    # Set the original value to 0
                    GF_UG_Mobimo.at[index, 'Gewerbe Bestand'] = 0
                if row['Gewerbe Neubau'] > 0:
                    new_row['Gewerbe Neubau'] = row['Gewerbe Neubau']
                    # Set the original value to 0
                    GF_UG_Mobimo.at[index, 'Gewerbe Neubau'] = 0

                # Add the new row to the list
                new_rows.append(new_row)

        new_rows_df = pd.DataFrame(new_rows)
        GF_UG_Mobimo = pd.concat([GF_UG_Mobimo, new_rows_df], ignore_index=True)

        GF_UG_Mobimo = add_columns(GF_UG_Mobimo, 'Wohnen Total', 'Wohnen Bestand', 'Wohnen Neubau')
        GF_UG_Mobimo = add_columns(GF_UG_Mobimo, 'Gewerbe Total', 'Gewerbe Bestand', 'Gewerbe Neubau')
        GF_UG_Mobimo = add_columns(GF_UG_Mobimo, 'Gemeinschaft Total', 'Gemeinschaft Bestand', 'Gemeinschaft Neubau')

        st.dataframe(GF_UG_Mobimo)

        #========== Get GF_UG_PP (Geschossflächen Parkplätze) table ==========
        st.header("Geschossflächen Parkplätze")
        GF_UG_PP_Mobimo = pd.DataFrame(columns=columns)
        GF_UG_PP_Mobimo['Geschossflächen oberirdisch'] = unique_storey_names
        GF_UG_PP_Mobimo[numeric_columns] = GF_UG_PP_Mobimo[numeric_columns].fillna(0)        
        GF_UG_PP_Mobimo = get_UG_Parking(GF_UG_Mobimo_dfsia, GF_UG_PP_Mobimo)
        GF_UG_PP_Mobimo = GF_UG_PP_Mobimo[~GF_UG_PP_Mobimo['Geschossflächen oberirdisch'].isin(values_to_remove)]
        if 'Parking Bestand' in GF_UG_PP_Mobimo.columns:
            GF_UG_PP_Mobimo['Wohnen Bestand'] = GF_UG_PP_Mobimo['Parking Bestand']
        if 'Parking Neubau' in GF_UG_PP_Mobimo.columns:
            GF_UG_PP_Mobimo['Wohnen Neubau'] = GF_UG_PP_Mobimo['Parking Neubau']
        
        GF_UG_PP_Mobimo = add_columns(GF_UG_PP_Mobimo, 'Wohnen Total', 'Wohnen Bestand', 'Wohnen Neubau')
        GF_UG_PP_Mobimo = add_columns(GF_UG_PP_Mobimo, 'Gewerbe Total', 'Gewerbe Bestand', 'Gewerbe Neubau')
        GF_UG_PP_Mobimo = add_columns(GF_UG_PP_Mobimo, 'Gemeinschaft Total', 'Gemeinschaft Bestand', 'Gemeinschaft Neubau')
        st.dataframe(GF_UG_PP_Mobimo)

        #========== Get Total GF (Total Geschossfläche) table ==========
        st.header("Total Geschossfläche")
        GF_Total = pd.DataFrame(columns=columns)
        GF_Total = GF_Total.rename(columns={'Geschossflächen oberirdisch': 'Total Geschossfläche'})
        # Create a DataFrame with the new rows
        new_total_rows = pd.DataFrame({
            'Total Geschossfläche': ['Geschossfläche oberirdisch inkl. AGF', 'Geschossfläche unterirdisch']
        })

        GF_Total = pd.concat([new_total_rows, GF_Total], ignore_index=True)

        GF_Total = calculate_totals(GF_Mobimo, AGF_Mobimo, GF_Total, "Wohnen Bestand", 'Geschossfläche oberirdisch inkl. AGF')
        GF_Total = calculate_totals(GF_Mobimo, AGF_Mobimo, GF_Total, "Wohnen Neubau", 'Geschossfläche oberirdisch inkl. AGF')
        GF_Total = calculate_totals(GF_Mobimo, AGF_Mobimo, GF_Total, "Gewerbe Bestand", 'Geschossfläche oberirdisch inkl. AGF')
        GF_Total = calculate_totals(GF_Mobimo, AGF_Mobimo, GF_Total, "Gewerbe Neubau", 'Geschossfläche oberirdisch inkl. AGF')
        GF_Total = calculate_totals(GF_Mobimo, AGF_Mobimo, GF_Total, "Gemeinschaft Bestand", 'Geschossfläche oberirdisch inkl. AGF')
        GF_Total = calculate_totals(GF_Mobimo, AGF_Mobimo, GF_Total, "Gemeinschaft Neubau", 'Geschossfläche oberirdisch inkl. AGF')
        GF_Total = calculate_totals(GF_UG_Mobimo, GF_UG_PP_Mobimo, GF_Total, "Wohnen Bestand", 'Geschossfläche unterirdisch')
        GF_Total = calculate_totals(GF_UG_Mobimo, GF_UG_PP_Mobimo, GF_Total, "Wohnen Neubau", 'Geschossfläche unterirdisch')
        GF_Total = calculate_totals(GF_UG_Mobimo, GF_UG_PP_Mobimo, GF_Total, "Gewerbe Bestand", 'Geschossfläche unterirdisch')
        GF_Total = calculate_totals(GF_UG_Mobimo, GF_UG_PP_Mobimo, GF_Total, "Gewerbe Neubau", 'Geschossfläche unterirdisch')
        GF_Total = calculate_totals(GF_UG_Mobimo, GF_UG_PP_Mobimo, GF_Total, "Gemeinschaft Bestand", 'Geschossfläche unterirdisch')
        GF_Total = calculate_totals(GF_UG_Mobimo, GF_UG_PP_Mobimo, GF_Total, "Gemeinschaft Neubau", 'Geschossfläche unterirdisch')
        GF_Total = add_columns(GF_Total, 'Wohnen Total', 'Wohnen Bestand', 'Wohnen Neubau')
        GF_Total = add_columns(GF_Total, 'Gewerbe Total', 'Gewerbe Bestand', 'Gewerbe Neubau')
        GF_Total = add_columns(GF_Total, 'Gemeinschaft Total', 'Gemeinschaft Bestand', 'Gemeinschaft Neubau')


        st.dataframe(GF_Total)

        # Create the charts
        col1, col2 = st.columns(2)

        with col1:
        
            colors = [ '#A07FDB', '#B5B0D9', '#F2F2F2', '#F2E9BB', '#F2A285', '#300000', '#F26B5E']

            sums = {
                'Wohnen Bestand': GF_Total['Wohnen Bestand'].sum(),
                'Wohnen Neubau': GF_Total['Wohnen Neubau'].sum(),
                'Gewerbe Bestand': GF_Total['Gewerbe Bestand'].sum(),
                'Gewerbe Neubau': GF_Total['Gewerbe Neubau'].sum(),
                'Gemeinschaft Bestand': GF_Total['Gemeinschaft Bestand'].sum(),
                'Gemeinschaft Neubau': GF_Total['Gemeinschaft Neubau'].sum()
            }

            fig = px.pie(
                values=sums.values(),
                names=sums.keys(),
                hole=0.5,  # This creates the donut shape
                title='Total Geschossfläche',
                # color_discrete_sequence=colors
                color_discrete_sequence=px.colors.qualitative.Set3
            )

            st.plotly_chart(fig)

        with col2:

            # Define the 'Neubau' values (above the x-axis)
            neubau_values = [GF_Total['Wohnen Neubau'].sum(), GF_Total['Gewerbe Neubau'].sum(), GF_Total['Gemeinschaft Neubau'].sum()]

            # Define the 'Bestand' values (below the x-axis, as negative values)
            bestand_values = [-GF_Total['Wohnen Bestand'].sum(), -GF_Total['Gewerbe Bestand'].sum(), -GF_Total['Gemeinschaft Bestand'].sum()]

            # Define the categories for the x-axis
            categories = ['Wohnen', 'Gewerbe', 'Gemeinschaft']

            # Create the bar chart with both 'Neubau' and 'Bestand' values
            fig2 = go.Figure()

            # Add 'Neubau' bars (positive values)
            fig2.add_trace(go.Bar(
                x=categories,
                y=neubau_values,
                name='Neubau',
                
            ))

            # Add 'Bestand' bars (negative values)
            fig2.add_trace(go.Bar(
                x=categories,
                y=bestand_values,
                name='Bestand',
                
            ))

            # Update the layout for aesthetics
            fig2.update_layout(
                title='Vergleich von Neubau und Bestand',
                xaxis_title='Kategorie',
                yaxis_title='Fläche',
                barmode='relative',  # This will stack the bars to show 'Neubau' above 'Bestand'
                yaxis=dict(
                    title='Fläche m2',
                    autorange=True  # This will allow negative values to be displayed below the x-axis
                )
            
            )

            # The code to display the figure would be:
            st.plotly_chart(fig2)



    with st.expander("04 GEBÄUDEVOLUMEN  (gem. SIA 416)"):
        #========== Get GV (Gebäudevolumen oberirdisch) table ==========
        GV_Mobimo = pd.DataFrame(columns=columns)        
        unique_storey_names = GF_dfsia['Ursprungsgeschoss Name'].unique()        
        GV_Mobimo['Geschossflächen oberirdisch'] = unique_storey_names
        numeric_columns = columns[1:]  # All columns except the first one
        GV_Mobimo[numeric_columns] = GV_Mobimo[numeric_columns].fillna(0)     
        GV_Mobimo = get_GV_Wohnen_Gewerbe_Gemeinschaft(GF_dfsia, GV_Mobimo)        
        GV_Mobimo['Wohnen Total'] = GV_Mobimo['Wohnen Bestand'] + GV_Mobimo['Wohnen Neubau'] 
        GV_Mobimo['Gewerbe Total'] = GV_Mobimo['Gewerbe Bestand'] + GV_Mobimo['Gewerbe Neubau'] 
        GV_Mobimo['Gemeinschaft Total'] = GV_Mobimo['Gemeinschaft Bestand'] + GV_Mobimo['Gemeinschaft Neubau'] 
        GV_Mobimo = GV_Mobimo[~GV_Mobimo['Geschossflächen oberirdisch'].isin(['2. UG', '1. UG'])]
        st.header("Gebäudevolumen oberirdisch")
        st.dataframe(GV_Mobimo)

        #========== Get AGV (Aussengeschossvolumen) table ==========
        AGV_Mobimo = pd.DataFrame(columns=columns)
        AGV_Mobimo['Geschossflächen oberirdisch'] = unique_storey_names    
        AGV_Mobimo = get_GV_Wohnen_Gewerbe_Gemeinschaft(AGF_dfsia, AGV_Mobimo)
        AGV_Mobimo = convert_columns_to_float(AGV_Mobimo)
        AGV_Mobimo[numeric_columns] = AGV_Mobimo[numeric_columns].fillna(0)
        AGV_Mobimo = AGV_Mobimo[~AGV_Mobimo['Geschossflächen oberirdisch'].isin(['2. UG', '1. UG'])]
        AGV_Mobimo['Wohnen Total'] = AGV_Mobimo['Wohnen Bestand'] + AGV_Mobimo['Wohnen Neubau'] 
        AGV_Mobimo['Gewerbe Total'] = AGV_Mobimo['Gewerbe Bestand'] + AGV_Mobimo['Gewerbe Neubau'] 
        AGV_Mobimo['Gemeinschaft Total'] = AGV_Mobimo['Gemeinschaft Bestand'] + AGV_Mobimo['Gemeinschaft Neubau'] 
        AGV_Mobimo.rename(columns={'Geschossflächen oberirdisch': 'Aussengeschossvolumen'}, inplace=True)
        st.header("Aussengeschossvolumen")
        st.dataframe(AGV_Mobimo)

        #========== Get GV_UG (Gebäudevolumen unterirdisch exkl. PP) table ==========
        st.header("Gebäudevolumen unterirdisch exkl. PP")
        values_to_remove = ['EG', "ZG (EG)", '1. OG', '2. OG', '3. OG', '4. OG', '5. OG', '6. OG', '7. OG']     
        GV_UG_Mobimo_dfsia = GF_dfsia[~GF_dfsia['Ursprungsgeschoss Name'].isin(values_to_remove)]
        
        GV_UG_Mobimo = pd.DataFrame(columns=columns)
        GV_UG_Mobimo['Geschossflächen oberirdisch'] = unique_storey_names    
           
        GV_UG_Mobimo = get_GV_UG_Wohnen_Gewerbe_Gemeinschaft(GV_UG_Mobimo_dfsia, GV_UG_Mobimo)
        
        GV_UG_Mobimo = GV_UG_Mobimo[~GV_UG_Mobimo['Geschossflächen oberirdisch'].isin(values_to_remove)]
        
        GV_UG_Mobimo.rename(columns={'Geschossflächen oberirdisch': 'Geschossflächen unterirdisch exkl. PP'}, inplace=True)
        if 'Keller Bestand' in GV_UG_Mobimo.columns:
            GV_UG_Mobimo['Wohnen Bestand'] = GV_UG_Mobimo['Keller Bestand']
        if 'Keller Neubau' in GV_UG_Mobimo.columns:
            GV_UG_Mobimo['Wohnen Neubau'] = GV_UG_Mobimo['Keller Neubau']
        if 'Keller Gewerbe Bestand' in GV_UG_Mobimo.columns:
            GV_UG_Mobimo['Gewerbe Bestand'] = GV_UG_Mobimo['Keller Gewerbe Bestand']
        if 'Keller Gewerbe Neubau' in GV_UG_Mobimo.columns:
            GF_UG_Mobimo['Gewerbe Neubau'] = GV_UG_Mobimo['Keller Gewerbe Neubau']
        GV_UG_Mobimo[numeric_columns] = GV_UG_Mobimo[numeric_columns].fillna(0)
        GV_UG_Mobimo['Wohnen Bestand'] = GV_UG_Mobimo['Wohnen Bestand'] + GV_UG_Mobimo['Keller Technik Bestand'] + GV_UG_Mobimo['Hobbyräume Bestand']
        GV_UG_Mobimo['Wohnen Neubau'] = GV_UG_Mobimo['Wohnen Neubau'] + GV_UG_Mobimo['Keller Technik Neubau'] + GV_UG_Mobimo['Hobbyräume Neubau']

        # Create a list to hold new rows
        new_rows = []

        for index, row in GV_UG_Mobimo.iterrows():
            if row['Gewerbe Bestand'] > 0 or row['Gewerbe Neubau'] > 0:
                # Create a new row with all values set to 0
                new_row = {col: 0 for col in GV_UG_Mobimo.columns}

                # Copy the value from 'Geschossflächen unterirdisch exkl. PP' and add suffix
                new_row['Geschossflächen unterirdisch exkl. PP'] = row['Geschossflächen unterirdisch exkl. PP'] + " Gewerbe"

                # Copy the value greater than 0
                if row['Gewerbe Bestand'] > 0:
                    new_row['Gewerbe Bestand'] = row['Gewerbe Bestand']
                    # Set the original value to 0
                    GV_UG_Mobimo.at[index, 'Gewerbe Bestand'] = 0
                if row['Gewerbe Neubau'] > 0:
                    new_row['Gewerbe Neubau'] = row['Gewerbe Neubau']
                    # Set the original value to 0
                    GV_UG_Mobimo.at[index, 'Gewerbe Neubau'] = 0

                # Add the new row to the list
                new_rows.append(new_row)

        new_rows_df = pd.DataFrame(new_rows)
        GV_UG_Mobimo = pd.concat([GV_UG_Mobimo, new_rows_df], ignore_index=True)

        GV_UG_Mobimo = add_columns(GV_UG_Mobimo, 'Wohnen Total', 'Wohnen Bestand', 'Wohnen Neubau')
        GV_UG_Mobimo = add_columns(GV_UG_Mobimo, 'Gewerbe Total', 'Gewerbe Bestand', 'Gewerbe Neubau')
        GV_UG_Mobimo = add_columns(GV_UG_Mobimo, 'Gemeinschaft Total', 'Gemeinschaft Bestand', 'Gemeinschaft Neubau')

        st.dataframe(GV_UG_Mobimo)

        #========== Get GV_UG_PP (Gebäudevolumen Parkplätze) table ==========
        st.header("Geschossflächen Parkplätze")
        GV_UG_PP_Mobimo = pd.DataFrame(columns=columns)
        GV_UG_PP_Mobimo['Geschossflächen oberirdisch'] = unique_storey_names
        GV_UG_PP_Mobimo[numeric_columns] = GV_UG_PP_Mobimo[numeric_columns].fillna(0)        
        GV_UG_PP_Mobimo = get_GV_UG_Parking(GV_UG_Mobimo_dfsia, GV_UG_PP_Mobimo)
        GV_UG_PP_Mobimo = GV_UG_PP_Mobimo[~GV_UG_PP_Mobimo['Geschossflächen oberirdisch'].isin(values_to_remove)]
        if 'Parking Bestand' in GV_UG_PP_Mobimo.columns:
            GV_UG_PP_Mobimo['Wohnen Bestand'] = GV_UG_PP_Mobimo['Parking Bestand']
        if 'Parking Neubau' in GF_UG_PP_Mobimo.columns:
            GV_UG_PP_Mobimo['Wohnen Neubau'] = GV_UG_PP_Mobimo['Parking Neubau']
        
        GV_UG_PP_Mobimo = add_columns(GV_UG_PP_Mobimo, 'Wohnen Total', 'Wohnen Bestand', 'Wohnen Neubau')
        GV_UG_PP_Mobimo = add_columns(GV_UG_PP_Mobimo, 'Gewerbe Total', 'Gewerbe Bestand', 'Gewerbe Neubau')
        GV_UG_PP_Mobimo = add_columns(GV_UG_PP_Mobimo, 'Gemeinschaft Total', 'Gemeinschaft Bestand', 'Gemeinschaft Neubau')
        st.dataframe(GV_UG_PP_Mobimo)

        #========== Get Total GV (Total Gebäudevolumen) table ==========
        st.header("Total Gebäudevolumen")
        GV_Total = pd.DataFrame(columns=columns)
        
        GV_Total = GV_Total.rename(columns={'Geschossflächen oberirdisch': 'Total Geschossfläche'})
        # Create a DataFrame with the new rows
        new_total_rows = pd.DataFrame({
            'Total Geschossfläche': ['Geschossfläche oberirdisch inkl. AGF', 'Geschossfläche unterirdisch']
        })

        GV_Total = pd.concat([new_total_rows, GV_Total], ignore_index=True)
        GV_Total = calculate_totals(GV_Mobimo, AGV_Mobimo, GV_Total, "Wohnen Bestand", 'Geschossfläche oberirdisch inkl. AGF')        
        GV_Total = calculate_totals(GV_Mobimo, AGV_Mobimo, GV_Total, "Wohnen Neubau", 'Geschossfläche oberirdisch inkl. AGF')
        GV_Total = calculate_totals(GV_Mobimo, AGV_Mobimo, GV_Total, "Gewerbe Bestand", 'Geschossfläche oberirdisch inkl. AGF')
        GV_Total = calculate_totals(GV_Mobimo, AGV_Mobimo, GV_Total, "Gewerbe Neubau", 'Geschossfläche oberirdisch inkl. AGF')
        GV_Total = calculate_totals(GV_Mobimo, AGV_Mobimo, GV_Total, "Gemeinschaft Bestand", 'Geschossfläche oberirdisch inkl. AGF')
        GV_Total = calculate_totals(GV_Mobimo, AGV_Mobimo, GV_Total, "Gemeinschaft Neubau", 'Geschossfläche oberirdisch inkl. AGF')
        GV_Total = calculate_totals(GV_UG_Mobimo, GV_UG_PP_Mobimo, GV_Total, "Wohnen Bestand", 'Geschossfläche unterirdisch')
        GV_Total = calculate_totals(GV_UG_Mobimo, GV_UG_PP_Mobimo, GV_Total, "Wohnen Neubau", 'Geschossfläche unterirdisch')
        GV_Total = calculate_totals(GV_UG_Mobimo, GV_UG_PP_Mobimo, GV_Total, "Gewerbe Bestand", 'Geschossfläche unterirdisch')
        GV_Total = calculate_totals(GV_UG_Mobimo, GV_UG_PP_Mobimo, GV_Total, "Gewerbe Neubau", 'Geschossfläche unterirdisch')
        GV_Total = calculate_totals(GV_UG_Mobimo, GV_UG_PP_Mobimo, GV_Total, "Gemeinschaft Bestand", 'Geschossfläche unterirdisch')
        GV_Total = calculate_totals(GV_UG_Mobimo, GV_UG_PP_Mobimo, GV_Total, "Gemeinschaft Neubau", 'Geschossfläche unterirdisch')
        GV_Total = add_columns(GV_Total, 'Wohnen Total', 'Wohnen Bestand', 'Wohnen Neubau')
        GV_Total = add_columns(GV_Total, 'Gewerbe Total', 'Gewerbe Bestand', 'Gewerbe Neubau')
        GV_Total = add_columns(GV_Total, 'Gemeinschaft Total', 'Gemeinschaft Bestand', 'Gemeinschaft Neubau')

        GV_Total.rename(columns={'Total Geschossfläche': 'Total Gebäudevolumen'}, inplace=True)
        GV_Total['Total Gebäudevolumen'] = GV_Total['Total Gebäudevolumen'].replace(
            'Geschossfläche oberirdisch inkl. AGF', 'Gebäudevolumen oberirdisch inkl. AGF'
        )
        GV_Total['Total Gebäudevolumen'] = GV_Total['Total Gebäudevolumen'].replace(
            'Geschossfläche unterirdisch', 'Gebäudevolumen unterirdisch'
        )
        st.dataframe(GV_Total)
        
        # Create the charts
        col3, col4 = st.columns(2)

        with col3:
        
            colors = [ '#A07FDB', '#B5B0D9', '#F2F2F2', '#F2E9BB', '#F2A285', '#300000', '#F26B5E']

            sums = {
                'Wohnen Bestand': GV_Total['Wohnen Bestand'].sum(),
                'Wohnen Neubau': GV_Total['Wohnen Neubau'].sum(),
                'Gewerbe Bestand': GV_Total['Gewerbe Bestand'].sum(),
                'Gewerbe Neubau': GV_Total['Gewerbe Neubau'].sum(),
                'Gemeinschaft Bestand': GV_Total['Gemeinschaft Bestand'].sum(),
                'Gemeinschaft Neubau': GV_Total['Gemeinschaft Neubau'].sum()
            }

            fig3 = px.pie(
                values=sums.values(),
                names=sums.keys(),
                hole=0.5,  # This creates the donut shape
                title='Total Gebäudevolumen',
                # color_discrete_sequence=colors
                color_discrete_sequence=px.colors.qualitative.Set3
            )

            st.plotly_chart(fig3)

        with col4:

            # Define the 'Neubau' values (above the x-axis)
            neubau_values = [GV_Total['Wohnen Neubau'].sum(), GV_Total['Gewerbe Neubau'].sum(), GV_Total['Gemeinschaft Neubau'].sum()]

            # Define the 'Bestand' values (below the x-axis, as negative values)
            bestand_values = [-GV_Total['Wohnen Bestand'].sum(), -GV_Total['Gewerbe Bestand'].sum(), -GV_Total['Gemeinschaft Bestand'].sum()]

            # Define the categories for the x-axis
            categories = ['Wohnen', 'Gewerbe', 'Gemeinschaft']

            # Create the bar chart with both 'Neubau' and 'Bestand' values
            fig4 = go.Figure()

            # Add 'Neubau' bars (positive values)
            fig4.add_trace(go.Bar(
                x=categories,
                y=neubau_values,
                name='Neubau',
                
            ))

            # Add 'Bestand' bars (negative values)
            fig4.add_trace(go.Bar(
                x=categories,
                y=bestand_values,
                name='Bestand',
                
            ))

            # Update the layout for aesthetics
            fig4.update_layout(
                title='Vergleich von Neubau und Bestand',
                xaxis_title='Kategorie',
                yaxis_title='Volumen',
                barmode='relative',  # This will stack the bars to show 'Neubau' above 'Bestand'
                yaxis=dict(
                    title='Volumen m3',
                    autorange=True  # This will allow negative values to be displayed below the x-axis
                )
            
            )

            # The code to display the figure would be:
            st.plotly_chart(fig4)

    
    with st.expander("DOWNLOAD DIE BERECHNUNGEN"):
        st.write("Weihnachten kam früh!")

        sia416_data_xlsx = sia416_to_excel(GF_Mobimo, AGF_Mobimo, GF_UG_Mobimo, GF_UG_PP_Mobimo, GF_Total, GV_Mobimo, AGV_Mobimo, GV_UG_Mobimo, GV_UG_PP_Mobimo, GV_Total)
        st.download_button(label='📥 Download XLSX', data=sia416_data_xlsx, file_name='data.xlsx', mime='application/vnd.ms-excel')