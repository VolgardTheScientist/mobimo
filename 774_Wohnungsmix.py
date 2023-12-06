import streamlit as st
import pandas as pd
import plotly.express as px
import io
import base64
from PIL import Image
import numpy as np



# Set the page to wide mode
st.set_page_config(layout="wide")

st.title("774 Wohnungsmix")

# File uploader that only accepts .xlsx files
uploaded_file = st.file_uploader("Choose a file", type='xlsx', key="file_uploader_1")


if uploaded_file is not None:
    # Read Excel file, skipping the first row
    df = pd.read_excel(uploaded_file, skiprows=1)

    # Extract the pattern 'XXXX' from Keller entries
    # keller_df = df[df['Wohnungstyp (Optionen-Set)'] == 'Keller']
    keller_df = df[df['Wohnungstyp (Optionen-Set)'] == 'Keller'].copy()

    keller_df['CodePattern'] = keller_df['Einheitscode-MOBIMO (Zeichenfolge)'].apply(lambda x: x.split('.')[2] if isinstance(x, str) and x.count('.') >= 3 else None)
    # DEBUGGING:
    # st.write(keller_df)

    # Custom aggregation function
    def custom_aggregation(x):
        # Geschoss logic
        geschoss = ', '.join(x['Ursprungsgeschoss Name'].unique())

        # Rest of the logic
        reduit_fläche = x[x['Raumname'] == 'Reduit']['Gemessene Fläche'].sum() if 'Reduit' in x['Raumname'].values else 0
        wohnungstyp = x['Wohnungstyp (Optionen-Set)'].iloc[0] if x['Wohnungstyp (Optionen-Set)'].nunique() == 1 else 'Fehler: Raumtyp im Modell falsch zugewiesen'
        hnf = x['Gemessene Fläche'].sum() - reduit_fläche

        return pd.Series({
            'HNF': hnf,
            'Wohnungsgrösse': wohnungstyp,
            'Reduit-Fläche': reduit_fläche,
            'Geschoss': geschoss            
        })

    # Group by 'Einheitscode-MOBIMO (Zeichenfolge)' and apply the custom aggregation
    wm = df.groupby('Einheitscode-MOBIMO (Zeichenfolge)').apply(custom_aggregation).reset_index()

    # Filter wm for valid Wohnungstyp values
    valid_wohnungstyp = ["1.5 Zimmer Wohnung", "2.5 Zimmer Wohnung", "3.5 Zimmer Wohnung", "4.5 Zimmer Wohnung", "6.5 Zimmer Wohnung", "Loft", "Fehler: Raumtyp im Modell falsch zugewiesen"]
    wm = wm[wm['Wohnungsgrösse'].isin(valid_wohnungstyp)]

    # Rename columns
    wm.rename(columns={'Einheitscode-MOBIMO (Zeichenfolge)': 'Wohnung'}, inplace=True)

    # 'Keller-Fläche' Logic
    wm['Keller-Fläche'] = wm['Wohnung'].apply(
        lambda x: keller_df[keller_df['CodePattern'] == x[-4:]]['Gemessene Fläche'].sum()
        if len(x.split('.')) >= 2 and x[-4:].isdigit() and x[-4:] in keller_df['CodePattern'].values else 0
    )

    # Add 'Gesamtfläche' column to wm
    wm['Gesamtfläche'] = wm['HNF'] + wm['Reduit-Fläche']

    # 'Standort-Kellerraum' Logic
    wm['Standort-Kellerraum'] = wm['Wohnung'].apply(
        lambda x: keller_df[keller_df['CodePattern'] == x[-4:]]['Einheitscode-MOBIMO (Zeichenfolge)'].iloc[0].split('.')[-1]
        if len(x.split('.')) >= 2 and x[-4:].isdigit() and x[-4:] in keller_df['CodePattern'].values else "-"
    )

    # Reordering the columns in wm
    wm = wm[["Wohnung", "Wohnungsgrösse", "Geschoss", "HNF", "Reduit-Fläche", "Standort-Kellerraum", "Keller-Fläche", "Gesamtfläche"]]

    # Create pivot table
    wf = wm.pivot_table(
        index='Wohnungsgrösse', 
        columns='Geschoss', 
        values='HNF', 
        aggfunc='sum', 
        fill_value=0
    )

    # Count the occurrences of each 'Wohnungsgrösse' in wm
    count_series = wm.groupby('Wohnungsgrösse').size()

    # Convert the series to a DataFrame
    count_df = count_series.to_frame('Anzahl Wohnungen')

    # Join the count DataFrame with wf
    wf = wf.join(count_df, on='Wohnungsgrösse')


    # Reset index to make 'Wohnungsgrösse' a column again if needed
    wf.reset_index(inplace=True)

    # Check if the columns exist in wf and then sum them into a new column
    columns_to_sum = ["EG", "EG, ZG (EG)", "EG, ZG (EG), 1. OG"]
    if all(col in wf.columns for col in columns_to_sum):
        wf['EG-1.OG'] = wf[columns_to_sum].sum(axis=1)
        wf.drop(columns=columns_to_sum, inplace=True)
    else:
        st.error("One or more columns to group are missing in 'wf'")

    # # Convert column names to a list
    # column_names = wf.columns.tolist()
 
    # # Display the list of column names
    # st.write(column_names)

    # ========== Quality Gate - check if all room types inside a flat are of the same type ==========
    if "Fehler: Raumtyp im Modell falsch zugewiesen" in wm['Wohnungsgrösse'].values:
        st.error("""Die **Raumtyp** Eigenschaft ist im ArchiCAD-Modell falsch zugewiesen, korrigiere den Fehler und lade die XLSX Liste erneut.""")

    # Check for duplicates in the specific column 'Raumcode-MOBIMO lang (Zeichenfolge)'
    duplicates = df[df['Raumcode-MOBIMO lang (Zeichenfolge)'].duplicated(keep=False)]

    # If there are any duplicates, display a warning
    if not duplicates.empty:
        # Extracting only the duplicate values from the specified column
        duplicate_values = duplicates['Raumcode-MOBIMO lang (Zeichenfolge)'].unique()
        # Displaying the warning with the list of duplicate values
        st.error(f"Duplikaten gefunden: {', '.join(duplicate_values)}")



    # Add a new column 'Total (m2)' that sums values row-wise
    wf['Total (m2)'] = wf.sum(axis=1, numeric_only=True)


    # Add a new row 'Total' that sums values column-wise    
    wf.loc['Total'] = wf.sum()


    # If 'Wohnungsgrösse' is the index and you want to maintain it in the new row
    wf.at['Total', 'Wohnungsgrösse'] = 'Total'

    # Calculate the average HNF and add it as a new column
    wf['Durchschnittliche HNF (m²)'] = (wf['Total (m2)'] / wf['Anzahl Wohnungen']).round(2)

    # Handle division by zero if there are rows with 'Anzahl Wohnungen' as 0
    wf['Durchschnittliche HNF (m²)'].fillna(0, inplace=True)

    # Get the total from the 'Total (m2)' column's 'Total' row
    total_m2 = wf.at['Total', 'Total (m2)']

    # Calculate the percentage for each 'Wohnungsgrösse'
    wf['Durchschnittliche HNF (%)'] = (wf['Total (m2)'] / total_m2 * 100).round(2)

    # Ensure the 'Total' row for the new percentage column is set to 100%
    wf.at['Total', 'Durchschnittliche HNF (%)'] = 100

    # Reordering the columns in wf
    column_order = ["Wohnungsgrösse", "EG-1.OG", "1. OG", "2. OG", "3. OG", 
                    "4. OG", "5. OG", "6. OG", "7. OG", "Anzahl Wohnungen", 
                    "Total (m2)", "Durchschnittliche HNF (m²)", "Durchschnittliche HNF (%)"]

    # Reorder columns and handle any missing columns
    wf = wf.reindex(columns=column_order)

    # Convert values to float before displaying: 
    # List of columns to convert to float
    columns_to_convert = ["EG-1.OG", "1. OG", "2. OG", "3. OG", 
                        "4. OG", "5. OG", "6. OG", "7. OG", "Anzahl Wohnungen", 
                        "Total (m2)", "Durchschnittliche HNF (m²)", "Durchschnittliche HNF (%)"]

    # Convert each column in the list to float
    for col in columns_to_convert:
        if col in wf.columns:
            wf[col] = pd.to_numeric(wf[col], errors='coerce')

    # Fill any NaN values that arise from conversion errors
    wf.fillna(0, inplace=True)

    # Convert 'Wohnungsgrösse' column to string type
    wf['Wohnungsgrösse'] = wf['Wohnungsgrösse'].astype(str)


    # Create expandable containers for DataFrames
    with st.expander("INPUT: ArchiCAD XLSX-Liste"):
        st.dataframe(df)

    with st.expander("Wohnungsmix"):
        st.dataframe(wm)

    with st.expander("HNF mit Reduits ohne Kellerräume"):
        st.dataframe(wf)
        # st.write("Data types in wf DataFrame:", wf.dtypes)

        # Plotly Charts

        def create_donut_chart():
             # Drop the 'Total' row for the chart
            # wf_chart_data = wf.drop('Total', errors='ignore')

            wf_chart_data = wf.copy()  # Create a copy if you want to keep the original df unchanged
            wf_chart_data.drop('Total', errors='ignore', inplace=True)            

            # Create a donut chart using Plotly
            fig = px.pie(wf_chart_data, names='Wohnungsgrösse', values='Total (m2)', hole=0.3,
                color='Wohnungsgrösse', color_discrete_map=color_mapping)

            # Customize labels to be outside
            fig.update_traces(textposition='outside')

            # Format the percentage data with two decimal places
            fig.update_traces(texttemplate='%{label}: %{percent:.2%}')

            # Customize layout to place the legend under the chart
            fig.update_layout(
                title_text='Wohnungsmix',
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=-0.5,
                    xanchor="center",
                    x=0.5
                ),
                uniformtext_minsize=12,
                uniformtext_mode='hide'
            )
            return fig

        def create_bar_chart():
            # Create a bar plot using Plotly

            # Drop the 'Total' row for the chart
            # wf_chart_data = wf.drop('Total', errors='ignore')

            wf_chart_data = wf.copy()  # Create a copy if you want to keep the original df unchanged
            wf_chart_data.drop('Total', errors='ignore', inplace=True)    

            # Create a bar plot using Plotly with the filtered data
            fig2 = px.bar(wf_chart_data, x='Wohnungsgrösse', y='Anzahl Wohnungen', text='Anzahl Wohnungen',
                color='Wohnungsgrösse', color_discrete_map=color_mapping)

            # Customize layout
            fig2.update_layout(
                title_text='Anzahl Wohnungstypen',
                xaxis_title='Wohnungsgrösse',
                yaxis_title='Anzahl Wohnungen',
                uniformtext_minsize=8,
                uniformtext_mode='hide',
                margin=dict(b=170),   # Increase bottom margin
            )
            return fig2


        # Define the color mapping
        color_mapping = {
            "1.5 Zimmer Wohnung": "#CAB4D9",
            "2.5 Zimmer Wohnung": "#B8D9D5",
            "3.5 Zimmer Wohnung": "#F2D06B",
            "4.5 Zimmer Wohnung": "#F2C6A0",
            "6.5 Zimmer Wohnung": "#F28066",
            "Fehler: Raumtyp im Modell falsch zugewiesen": "#FF0000"
            }

        col1, col2 = st.columns(2)

        with col1:
            fig = create_donut_chart()

            # Display the plot in Streamlit
            st.plotly_chart(fig)

        with col2:
            fig2 = create_bar_chart()

            # Display the plot in Streamlit
            st.plotly_chart(fig2)

    # Download data

    # Function to return XLSX file from dataframes
    def to_excel(wm, wf):
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        wm.to_excel(writer, sheet_name='Wohnungsmix', index=False)
        wf.to_excel(writer, sheet_name='HNF', index=False)
        writer.close()
        processed_data = output.getvalue()
        return processed_data

    # Function to convert plotly figure to image and encode
    # Convert Plotly figures to PNG bytes
    def plotly_fig_to_png_bytes(fig):
        img_bytes = fig.to_image(format="png")
        return img_bytes

    # Use the function to convert figures to PNG bytes
    donut_image = plotly_fig_to_png_bytes(fig)
    bar_image = plotly_fig_to_png_bytes(fig2)
 

    data_xlsx = to_excel(wm, wf)
    st.download_button(label='📥 Download XLSX', data=data_xlsx, file_name='data.xlsx', mime='application/vnd.ms-excel')


    st.download_button(
        label="📥 Download Donut Chart",
        data=donut_image,
        file_name="donut_chart.png",
        mime="image/png"
    )


    st.download_button(
        label="📥 Download Bar Chart",
        data=bar_image,
        file_name="bar_chart.png",
        mime="image/png"
    )