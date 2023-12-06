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

import pandas as pd

def sum_area_by_categories(df1):
    df2 = df1.groupby(['Raumkategoriename', 'Nutzer (Zeichenfolge)', 'Ursprungsgeschoss Name'])['Berechnete Fläche (NRF)'].sum().reset_index()
    return df2


def update_df2_from_df1(df1, df2):
    # Iterate over each row in df1
    for index, row in df1.iterrows():
        # Check for the first condition (HNF Bestand)
        if row['Raumkategoriename'] == 'HNF Bestand' and row['Nutzer (Zeichenfolge)'] == 'Wohnen':
            df2_index = df2[df2['Hauptnutzflächen (HNF)'] == row['Ursprungsgeschoss Name']].index
            if not df2_index.empty:
                df2.at[df2_index[0], 'Wohnen Bestand'] = row['Berechnete Fläche (NRF)']

        # Check for the additional condition (HNF Neubau)
        elif row['Raumkategoriename'] == 'HNF Neubau' and row['Nutzer (Zeichenfolge)'] == 'Wohnen':
            df2_index = df2[df2['Hauptnutzflächen (HNF)'] == row['Ursprungsgeschoss Name']].index
            if not df2_index.empty:
                df2.at[df2_index[0], 'Wohnen Neubau'] = row['Berechnete Fläche (NRF)']

        # Check for the additional condition (HNF Neubau)
        elif row['Raumkategoriename'] == 'HNF Neubau' and row['Nutzer (Zeichenfolge)'] == 'Gewerbe':
            df2_index = df2[df2['Hauptnutzflächen (HNF)'] == row['Ursprungsgeschoss Name']].index
            if not df2_index.empty:
                df2.at[df2_index[0], 'Gewerbe Neubau'] = row['Berechnete Fläche (NRF)']

        # Check for the additional condition (HNF Neubau)
        elif row['Raumkategoriename'] == 'HNF Bestand' and row['Nutzer (Zeichenfolge)'] == 'Gewerbe':
            df2_index = df2[df2['Hauptnutzflächen (HNF)'] == row['Ursprungsgeschoss Name']].index
            if not df2_index.empty:
                df2.at[df2_index[0], 'Gewerbe Bestand'] = row['Berechnete Fläche (NRF)']

                # Check for the additional condition (HNF Neubau)
        elif row['Raumkategoriename'] == 'HNF Neubau' and row['Nutzer (Zeichenfolge)'] == 'Gemeinschaft':
            df2_index = df2[df2['Hauptnutzflächen (HNF)'] == row['Ursprungsgeschoss Name']].index
            if not df2_index.empty:
                df2.at[df2_index[0], 'Gemeinschaft Neubau'] = row['Berechnete Fläche (NRF)']

        # Check for the additional condition (HNF Neubau)
        elif row['Raumkategoriename'] == 'HNF Bestand' and row['Nutzer (Zeichenfolge)'] == 'Gemeinschaft':
            df2_index = df2[df2['Hauptnutzflächen (HNF)'] == row['Ursprungsgeschoss Name']].index
            if not df2_index.empty:
                df2.at[df2_index[0], 'Gemeinschaft Bestand'] = row['Berechnete Fläche (NRF)']

    return df2

def sort_df2_by_HNF(df2, order_list):
    # Defining a custom sorting function using the order list
    sorter_index = dict(zip(order_list, range(len(order_list))))

    # Generating a rank column based on the custom sorting function
    df2['Rank'] = df2['Hauptnutzflächen (HNF)'].map(sorter_index)

    # Sorting by the rank column and dropping it afterwards
    df2_sorted = df2.sort_values(['Rank']).drop('Rank', axis=1)

    return df2_sorted


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

    # Find the row index in HNF_Mobimo where 'Total Geschossfläche' is 'Geschossfläche oberirdisch inkl. AGF'
    row_index = df_output[df_output['Total Geschossfläche'] == row].index

    # Check if there's exactly one row that matches the condition
    if len(row_index) == 1:
        # Set the value in 'Wohnen Bestand' at the specific row
        df_output.at[row_index[0], input] = total_wohnen_bestand
    else:
        # Handle the case where the condition is met by zero or multiple rows
        st.write("Error: The condition is not met by a unique row.")
    return df_output

def hnf_to_excel(df1):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='HNF', index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# ========== MAIN APP ==========

st.title("774 SIA 416 Hauptnutzflächen")

uploaded_file_hnf = st.file_uploader("Choose a file", type='xlsx', key="file_uploader_3")

if uploaded_file_hnf is not None:

    dfhnf = pd.read_excel(uploaded_file_hnf, skiprows=1)
    
    with st.expander("INPUT: ArchiCAD XLSX-Liste"):
        st.dataframe(dfhnf)

    with st.expander("05 HAUPTNUTZFLÄCHEN (gem. SIA 416)"):
        
        columns = [
            'Hauptnutzflächen (HNF)', 'Wohnen Bestand', 'Wohnen Neubau', 'Wohnen Total',
            'Gewerbe Bestand', 'Gewerbe Neubau', 'Gewerbe Total',
            'Gemeinschaft Bestand', 'Gemeinschaft Neubau', 'Gemeinschaft Total'
        ]
        #========== Get Hauptnutzflächen (HNF) table ==========
        HNF_Mobimo = pd.DataFrame(columns=columns)        
        unique_storey_names = dfhnf['Ursprungsgeschoss Name'].unique()        
        HNF_Mobimo['Hauptnutzflächen (HNF)'] = unique_storey_names
        numeric_columns = columns[1:]  # All columns except the first one
        HNF_Mobimo[numeric_columns] = HNF_Mobimo[numeric_columns].fillna(0)
        dfhnf = sum_area_by_categories(dfhnf)
        # st.dataframe(dfhnf)
    
        HNF_Mobimo = update_df2_from_df1(dfhnf, HNF_Mobimo)      
        HNF_Mobimo['Wohnen Total'] = HNF_Mobimo['Wohnen Bestand'] + HNF_Mobimo['Wohnen Neubau'] 
        HNF_Mobimo['Gewerbe Total'] = HNF_Mobimo['Gewerbe Bestand'] + HNF_Mobimo['Gewerbe Neubau'] 
        HNF_Mobimo['Gemeinschaft Total'] = HNF_Mobimo['Gemeinschaft Bestand'] + HNF_Mobimo['Gemeinschaft Neubau'] 
        # HNF_Mobimo = HNF_Mobimo[~HNF_Mobimo['Hauptnutzflächen (HNF)'].isin(['2. UG', '1. UG'])]
        st.header("Hauptnutzflächen (HNF)")

        order_list = ['2. UG', '1. UG', 'EG', 'ZG (EG)', '1. OG', '2. OG', '3. OG', '4. OG', '5. OG', '6. OG', '7. OG']
        HNF_Mobimo = sort_df2_by_HNF(HNF_Mobimo, order_list)


        st.dataframe(HNF_Mobimo)

       

        # Create the charts
        col1, col2 = st.columns(2)

        with col1:
        
            colors = [ '#A07FDB', '#B5B0D9', '#F2F2F2', '#F2E9BB', '#F2A285', '#300000', '#F26B5E']

            sums = {
                'Wohnen Bestand': HNF_Mobimo['Wohnen Bestand'].sum(),
                'Wohnen Neubau': HNF_Mobimo['Wohnen Neubau'].sum(),
                'Gewerbe Bestand': HNF_Mobimo['Gewerbe Bestand'].sum(),
                'Gewerbe Neubau': HNF_Mobimo['Gewerbe Neubau'].sum(),
                'Gemeinschaft Bestand': HNF_Mobimo['Gemeinschaft Bestand'].sum(),
                'Gemeinschaft Neubau': HNF_Mobimo['Gemeinschaft Neubau'].sum()
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
            neubau_values = [HNF_Mobimo['Wohnen Neubau'].sum(), HNF_Mobimo['Gewerbe Neubau'].sum(), HNF_Mobimo['Gemeinschaft Neubau'].sum()]

            # Define the 'Bestand' values (below the x-axis, as negative values)
            bestand_values = [-HNF_Mobimo['Wohnen Bestand'].sum(), -HNF_Mobimo['Gewerbe Bestand'].sum(), -HNF_Mobimo['Gemeinschaft Bestand'].sum()]

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




    with st.expander("DOWNLOAD DIE BERECHNUNGEN"):
        st.write("Weihnachten kam früh!")

        sia416_data_xlsx = hnf_to_excel(HNF_Mobimo)
        st.download_button(label='📥 Download XLSX', data=sia416_data_xlsx, file_name='data.xlsx', mime='application/vnd.ms-excel')