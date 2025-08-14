# Import python packages
import streamlit as st
from snowflake.snowpark.context import get_active_session
from io import BytesIO
from datetime import datetime
import pandas as pd
import numpy as np # ### MODIFICATION ### - Ajout de numpy pour gérer les NULLs plus facilement (pd.NA)
import openpyxl


# Write directly to the app
st.set_page_config(layout='wide')
st.title("🛍️ Promotions Input App")

# Get the current credentials
SESSION = get_active_session()

# Constantes de configurations
CONFIG = {
    "DATABASE": "CH_PRD_DB",
    "SCHEMA": "DWH_WRK",
    "STAGE": "INPUT_PROMO_PROJECT",
    "PROMO_TABLE": "CH_PRD_DB.DWH_WRK.BASELINEPREDICTION_REF_PROMOTION",
    "CATEGORY_TABLE": "CH_PRD_DB.DWH_WRK.MAPPING_SEGMENT_CATEGORY"
}



#function to apply styling to the app
#the variable 'style' contains css which is applied to the page using 'st.markdown'
def style():
    """Gestion des styles du front de l'app streamlit."""
    style = """
        <style>
            @import url('http://www.mostardesign.com');
            @import url('https://fonts.cdnfonts.com/css/poppins');

            /* Overall styling base : backgroung image*/
            .stMain{
                background-color: rgba(255, 255, 255, 0.5);
                color : #143F49 !important;
                font-family: FilsonPro, sans-serif;
                margin: 0;
                background-image: url('https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcThL15R22HA6nwDsNdHI7q88gY3-fEyQQrfIiywtN4jQTVd63fs');
                background-size: cover;
                background-position: center;
                background-attachment: fixed;
            }
            .stMain::before {
                content: "";
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(255, 255, 255, 0.98);
            }
            
            /* Change background when screen width is less than 768px (tablets & phones) */
            @media (max-width: 768px) {
                .stMain {
                    background-color: white; /* Darker background */
                    background-image: none; /* Remove the background image */
                }
            
                .stMain::before {
                    background-color: white; /* Adjust overlay color */
                }
            }
            
            /* Change background for very small screens (mobile devices) */
            @media (max-width: 480px) {
                .stMain {
                    background-color: white; /* Even darker */
                }
            
                .stMain::before {
                    background-color: white;
                }
            }

            /* Titles */
            h1 {
                color : #E20E17 !important;
                font-family: "filson-pro", sans-serif;
                font-weight: 600;
                font-style: normal;
                font-size: 40px;
            }

            /* Paragraphs */
            p {
                color : black !important;
                font-family: "filson-pro", sans-serif;
                font-weight: 400;
                font-style: normal;
            }

            /* Expanders */
            [data-testid=stExpander] details{
                border-color: #F8C3C5;
            }
            [data-baseweb="tag"] {
                background-color: #E20E17 !important;
            }

            [data-baseweb="tag"] path{
                color: white !important;
            }
            
            path{
                color: #E20E17 !important;
            }
            [data-testid=stExpanderDetails] details p{
                color: black !important;
                font-weight: lighter !important;
                font-size: 14px;
            }

            /* Radio buttons */
            [data-testid=stRadio] p{
                color: black !important;
                font-weight: bold;
                font-size: 15px;
            }
            [data-testid=stRadio] [role=radiogroup]{
                padding-left: 140px;
            }

            /* Buttons */
            [data-testid=stButton] button{
                background-color: white;
                border: transparent;
            }
            [data-testid=stButton] p{
                color: black !important;
                font-weight: 600;
                font-size: 13px;
            }
            
            [data-testid=stButton] {
                color: #E20E17 !important;
                font-weight: bold;
            }
            [data-testid="stBaseButton-primary"] {
                background-color: #E20E17 !important;
                border: none !important;
            }
            [data-testid="stBaseButton-primary"] p {
                color: white !important;
            }

            
            /* SIDEBAR (here we have no sidebar but the code is here in case one is added in the future) */
            [data-testid=stSidebar] {
                background-color: #E20E17;
                color: white !important;
            }
            [data-testid=stSidebar] h4 {
                color: white;
            }
            [data-testid=stSidebar] h5 {
                color: #FFE5E6;
                font-weight: lighter;
                font-size: smaller;
            }
            [data-testid=stSidebar] p {
                color: white !important;
            }
            [data-testid=stSidebar] a {
                color: white !important;
                font-size: 14px;
            }
            [data-testid=stSidebar] details {
                background-color: rgba(255, 229, 230, 0.3) !important;
                color: white;
            }
            
            [data-testid=stSidebar] details p{
                color: white !important;
                font-weight: normal;
                font-size: 14px ;
            }
            
            [data-testid=stSidebar] details h6{
                color: white !important;
                font-weight: light !important;
                font-size: smaller;
            }
            [data-testid=stSidebar] li {
                font-size: smaller;
            }
            [data-testid=stSidebar] p {
                font-size: smaller;
            }
            [data-testid=stSidebar] .stButton{
                color: #E20E17 !important;
                font-weight: normal !important;
            }
            [data-testid=stSidebar] .stButton p{
                color: #E20E17 !important;
                font-weight: bold !important;
                font-size: 14px;
            }
            
            [data-testid=stSidebar] [data-testid=stExpander] summary p{
                color: white !important;
                font-weight: bold !important;
                font-size: 14px;
            }
            
            
            /* Target the sidebar collapse button */
            [data-testid="stSidebarCollapseButton"] button {
                color: red !important;  /* Change icon color */
                background-color: white !important;  /* Change background color */
                border-radius: 5px; /* Optional: Round edges */
            }
    
            /* Optional: Change color on hover */
            [data-testid="stSidebarCollapseButton"] button:hover {
                background-color: #FFE5E6 !important;
                color: white !important;
            }

            
            /* DARK MODE */
            @media (prefers-color-scheme: dark) {
                .stMain{
                    background-color: rgba(255, 255, 255, 0.5);
                    color : #143F49 !important;
                    font-family: FilsonPro, sans-serif;
                    margin: 0;
                    background-image: url('https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcThL15R22HA6nwDsNdHI7q88gY3-fEyQQrfIiywtN4jQTVd63fs');
                    background-size: cover;
                    background-position: center;
                    background-attachment: fixed; 
                }
                .stMain::before {
                    content: "";
                    position: absolute;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    background-color: rgba(0, 0, 0, 0.98); 
                }
                
                /* Change background when screen width is less than 768px (tablets & phones) */
                @media (max-width: 768px) {
                    .stMain {
                        background-color: black; /* Darker background */
                        background-image: none; /* Remove the background image */
                    }
                
                    .stMain::before {
                        background-color: black; /* Adjust overlay color */
                    }
                }
                
                h1 {
                    color : #E20E17 !important;
                    font-family: "filson-pro", sans-serif;
                    font-weight: 600;
                    font-style: normal;
                    font-size: 25px;
                }
                p {
                    color : white !important;
                    font-family: "filson-pro", sans-serif;
                    font-weight: 400;
                    font-style: normal;
                }
                [data-testid=stBottomBlockContainer] {
                    background-color: black !important;
                    color : black !important;
                    font-family: FilsonPro, sans-serif;
                }
                [data-testid=stBottom] {
                    background-color: black !important;
                    color : black !important;
                    font-family: FilsonPro, sans-serif;
                }
                [data-testid=stBottom] > :nth-child(1) {
                  background-color: transparent !important;
                }
                [data-testid=stChatInput] {
                    background-color: #F8C3C5 !important;
                    border: transparent !important;
                }
                textarea[aria-label="Wie lautet deine Frage?"]:not(:placeholder-shown):focus{
                 color: black;
                }
                [data-testid=stChatMessage][data-role="assistant"] {
                    background-color: rgba(226, 14, 23, 0.15) !important;
                    padding: 1rem 3em 1rem 1rem;
                }
                [data-testid=stChatMessage]:first-of-type {
                    background-color: #143F49 !important;
                    color: white !important;
                    border-radius: 8px !important;
                    padding: 1rem 3em 1rem 1rem;
                }
        
                /* Target user messages */
                [data-testid=stChatMessage]:nth-of-type(even) {
                    background-color: #143F49 !important;
                    color: black !important;
                    padding: 1rem 3em 1rem 1rem;
                }
                /* Target the container that contains the specific header */
                div:has(> #feedback-header) div[data-testid="stVerticalBlockBorderWrapper"] {
                    border: 2px solid red !important;  /* Red border */
                    border-radius: 8px !important;  /* Optional: Rounded edges */
                    padding: 10px !important;  /* Optional: Adjust padding */
                }
                [data-testid=stExpander] details{
                    border-color: #F8C3C5;
                }
                [data-testid=stRadio] p{
                    color: #E20E17 !important;
                    font-weight: bold;
                    font-size: 15px;
                }
                [data-testid=stRadio] [role=radiogroup]{
                    padding-left: 140px;
                }
                [data-testid=stButton] button{
                    background-color: white;
                    border: transparent;
                }
                [data-testid=stButton] p{
                    color: black !important;
                    font-weight: 600;
                    font-size: 13px;
                }
                
                [data-testid=stButton] {
                    color: #E20E17 !important;
                    font-weight: bold;
                }
                [data-testid="stBaseButton-primary"] {
                    background-color: #E20E17 !important;
                    border: none !important;
                }
                [data-testid="stBaseButton-primary"] p {
                    color: white !important;
                }
                
                path{
                    color: #E20E17 !important;
                }
    
                
                /* SIDEBAR */
                [data-testid=stSidebar] {
                    background-color: #E20E17;
                    color: white !important;
                }
                [data-testid=stSidebar] h4 {
                    color: white;
                }
                [data-testid=stSidebar] h5 {
                    color: #FFE5E6;
                    font-weight: lighter;
                    font-size: smaller;
                }
                [data-testid=stSidebar] p {
                    color: white !important;
                }
                [data-testid=stSidebar] a {
                    color: white !important;
                    font-size: 14px;
                }
                [data-testid=stSidebar] details {
                    background-color: rgba(255, 229, 230, 0.3) !important;
                    color: white;
                }
                
                [data-testid=stSidebar] details p{
                    color: white !important;
                    font-weight: normal;
                    font-size: 14px ;
                }
                
                [data-testid=stSidebar] details h6{
                    color: white !important;
                    font-weight: light !important;
                    font-size: smaller;
                }
                [data-testid=stSidebar] li {
                    font-size: smaller;
                }
                [data-testid=stSidebar] p {
                    font-size: smaller;
                }
                [data-testid=stSidebar] .stButton{
                    color: #E20E17 !important;
                    font-weight: normal !important;
                }
                [data-testid=stSidebar] .stButton p{
                    color: #E20E17 !important;
                    font-weight: bold !important;
                    font-size: 14px;
                }
                
                [data-testid=stSidebar] [data-testid=stExpander] summary p{
                    color: white !important;
                    font-weight: bold !important;
                    font-size: 14px;
                }
                
                [data-testid=stExpanderDetails] details p{
                    color: black !important;
                    font-weight: lighter !important;
                    font-size: 14px;
                }
                
                /* Target the sidebar collapse button */
                [data-testid="stSidebarCollapseButton"] button {
                    color: red !important;  /* Change icon color */
                    background-color: white !important;  /* Change background color */
                    border-radius: 5px; /* Optional: Round edges */
                }
        
                /* Optional: Change color on hover */
                [data-testid="stSidebarCollapseButton"] button:hover {
                    background-color: #FFE5E6 !important;
                    color: white !important;
            }
            }
        </style>
    
    """
    st.markdown(style, unsafe_allow_html=True)
    _, logo = st.columns((7,1))
    with logo :
        st.image("https://andros-asia.com/wp-content/uploads/andros-logo.png", width=100)
    return None

def update_table_no_truncate(table, edited_df, numeric_columns=None):
    """
    Updates a Snowflake table with data from a pandas DataFrame without truncating the table.

    Args:
        - table (str): Snowflake table name.
        - edited_df (pandas.DataFrame): DataFrame containing rows to insert.
        - numeric_columns (list, optional): List of numeric column names to be inserted without quotes.
    """
    if numeric_columns is None:
        numeric_columns = []

    # Get the raw connection object from the session.
    conn = SESSION.connection
    cursor = conn.cursor()

    # Get column names
    columns = edited_df.columns.tolist()
    col_list = ', '.join(columns)  # Format as "col1, col2, col3"

    # Use `?` as the placeholder for Snowflake
    placeholders = ', '.join(['?'] * len(columns))
    insert_query = f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})"

    # Convert DataFrame rows to a list of tuples
    values = [
        tuple(None if pd.isna(row[col]) else row[col] for col in columns)
        for _, row in edited_df.iterrows()
    ]

    try:
        # Execute batch insert
        cursor.executemany(insert_query, values)

        # Commit the transaction
        cursor.execute("COMMIT")

    except Exception as e:
        st.write(f"Error executing SQL: {e}")
    
    finally:
        # Close cursor
        cursor.close()

def delete_rows_table(table, ids_to_delete):
    """
    Delete rows from Snowflake table based on a list of rows deleted by the user

    Args:
        rows_to_delete_df (str): Name of the table to update.
    """
    # Get the raw connection object from the session.
    conn = SESSION.connection
        
    # Create a cursor object for executing SQL queries.
    cursor = conn.cursor()
    
    for id in ids_to_delete:
        
        # Assuming the table has an 'ID' column as the unique identifier
        delete_query = f"""
        DELETE FROM {table}
        WHERE PROMOTION_ID = {id};
        """
        cursor.execute(delete_query)

    # Commit changes
    cursor.execute("COMMIT")

    # Close the cursor
    cursor.close()

def update_snowflake_table(table, edited_df, numeric_columns=['']):
    """
    Updates a Snowflake table based on a given DataFrame on basis of promotion_id.

    Args:
        table (str): Name of the table to update.
        edited_df (pd.DataFrame): DataFrame containing the new values.
        numeric_columns (list, optional): List of columns that should not be wrapped in quotes.
    """
    # Get the raw connection object from the session.
    conn = SESSION.connection
        
    # Create a cursor object for executing SQL queries.
    cursor = conn.cursor()

    columns = edited_df.columns

    for _, row in edited_df.iterrows():
        set_clauses = []
        
        for col in columns:
            value = row[col]
            
            # Format value based on data type
            if pd.isna(value):  # Handle NaN values
                formatted_value = "NULL"
            elif col in numeric_columns:  # Numeric columns (no quotes)
                formatted_value = str(value)
            else:  # String columns (wrapped in quotes)
                formatted_value = f"'{value}'"
            
            set_clauses.append(f"{col} = {formatted_value}")

        set_query = ", ".join(set_clauses)
        
        # Assuming the table has an 'ID' column as the unique identifier
        update_query = f"""
        UPDATE {table}
        SET {set_query}
        WHERE PROMOTION_ID = '{row['PROMOTION_ID']}';
        """
        cursor.execute(update_query)

    # Commit changes
    cursor.execute("COMMIT")

    # Close the cursor
    cursor.close()

def save_as_excel_button(df, df_name, key):
        """
        Generate a Streamlit button to download a DataFrame as an Excel file.

        Args:
        - df (DataFrame): DataFrame to be saved.
        - df_name (str): Name of the file (without extension).
        """

        # Round all numbers so they only have two numbers after the comma
        df = df.round(2)

        @st.cache_data # Cache the function to avoid redundant processing.
        def to_excel(df):
            """
            Convert the DataFrame to an Excel file using an in-memory buffer.

            Args:
            - df (DataFrame): DataFrame to convert.

            Returns:
            - Bytes: Excel data in memory.
            """
            # Create an in-memory buffer for storing the Excel data.
            output = BytesIO()
            # Use openpyxl as the engine for writing Excel files.
            writer = pd.ExcelWriter(output, engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name=df_name) # Write data to the first sheet.
            writer.close() # Close the writer to finalize the data.
            # Retrieve the processed Excel data.
            processed_data = output.getvalue()
            return processed_data 
        
        # Generate the Excel data from the DataFrame.
        excel_data = to_excel(df)
        
        # Create a download button in Streamlit for downloading the Excel file.
        _, save_as_excel_col = st.columns([3,1])
        with save_as_excel_col:
            st.download_button(
                label="Download data as Excel file",
                key=key,
                data=excel_data,
                file_name=f"{df_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

def body():
    """
    Main method to define the user interface and behavior of the Promo Input App.
    It creates and displays the different filters, the resulting table which can be modifed and a save button
    """
    style()
    st.write("")
    user_action = st.radio(
            "What do you wish to do ?",
            ["Review or modify existing promotions", "Input a new promotion"],
            horizontal=True,
            index=None,
    )

    st.write("")
    
    if user_action:
        if user_action != "Review or modify existing promotions":
            st.subheader("Choose your values")
            
        # Path to the logs table used for tracking changes
        promo_table = CONFIG['PROMO_TABLE']
        categories_table = CONFIG['CATEGORY_TABLE']
        
        # Load the content of the tabkes into dataframes
        loaded_table = SESSION.table(promo_table)
        promo_df = loaded_table.to_pandas()
        promo_df_all = promo_df.copy()
        categories_table = SESSION.table(categories_table)
        categories_df = categories_table.to_pandas()
        categories_df = categories_df[categories_df["BRAND_NAME"].isin(['ANDROS', 'BONNE MAMAN'])]

        # ### CORRECTION ### - Ligne corrigée
        current_year = datetime.now().year
        years = list(range(2023, current_year+3))
        if user_action ==  "Review or modify existing promotions":
            years += ['All']
        weeks = list(range(1, 52+1))
        promo_types = sorted(list(set(promo_df["PROMOTION_TYPE"])))
        
        if user_action ==  "Review or modify existing promotions":
            promo_types.insert(0, 'Any')
        else: 
            promo_types.insert(0, 'Other')
        duration_types = ['Tag(e)', 'Woche(n)']
        if user_action ==  "Review or modify existing promotions":
            duration_types = ['Any', 'Tag(e)', 'Woche(n)']
            
        regions = set(promo_df['REGION'])
        regions = sorted(list(set(part for region in regions for part in region.split('+'))))
        mechanisms = sorted(list(set(promo_df['MECHANISM'])))
        if user_action ==  "Review or modify existing promotions":
            mechanisms.insert(0, 'Any')
        else: 
            mechanisms.insert(0, 'Other')
        brands = sorted(list(set(promo_df["BRAND_NAME"])))
        #add filter on brands
        
        cent_off_types = set(promo_df['CENT_OFF_BY_PIECE_TYPE'])
        
        ### FILTERS
        ## 1st LAYER
        brand_col, _, global_col = st.columns([3,0.2, 2])
        with brand_col :
            if user_action !=  "Review or modify existing promotions":
                brand_filter = st.selectbox(
                    "Brand",
                    brands
                )
                if not promo_df.empty:
                    promo_df = promo_df[promo_df["BRAND_NAME"] == brand_filter]
                    if brand_filter :
                        categories_df = categories_df[categories_df["BRAND_NAME"] == brand_filter]
            else: 
                brand_filter = st.multiselect(
                    "Brand",
                    brands,
                    []
                )
                if brand_filter :
                    if not promo_df.empty:
                        promo_df = promo_df[promo_df["BRAND_NAME"].isin(brand_filter)]

        # Define the lists of product categories depending on the selected brand(s)
        if user_action ==  "Review or modify existing promotions":
            if len(brand_filter) == 1:
                categories_df = categories_df[categories_df["BRAND_NAME"] == brand_filter[0]]
            else :
                categories_df = categories_df[categories_df["BRAND_NAME"].isin(['ANDROS', 'BONNE MAMAN'])]
        else:
            categories_df = categories_df[categories_df["BRAND_NAME"] == brand_filter]
                
        product_category = list(set(categories_df["KATEGORIE"]))
        product_category = sorted(product_category, key=str.casefold)
        
        product_undercategory = list(set(categories_df["UNTERKATEGORIE"]))
        product_undercategory = list(filter(lambda x: x is not None, product_undercategory)) 
        product_undercategory = sorted(product_undercategory, key=str.casefold)
        
        product_segment = list(set(categories_df["SEGMENT"]) )
        product_segment = list(filter(lambda x: x is not None, product_segment)) 
        product_segment = sorted(product_segment, key=str.casefold)
        
        product_undersegment = list(set(categories_df["UNTERSEGMENT"]))
        product_undersegment = list(filter(lambda x: x is not None, product_undersegment)) 
        product_undersegment = sorted(product_undersegment, key=str.casefold)
        
        categories_df['SKU_ID'] = categories_df.apply(lambda row: f"{row['SKU_BESCHREIBUNG']} ({row['CLIENT_ARTIKEL_ID']})" if row['CLIENT_ARTIKEL_ID'] else row['SKU_BESCHREIBUNG'], axis=1)
        product_id = set(categories_df["CLIENT_ARTIKEL_ID"])
        product_sku_id = list(set(categories_df["SKU_ID"]))
        product_sku_id = list(filter(lambda x: x is not None, product_sku_id)) 
        product_sku_id = sorted(product_sku_id, key=str.casefold)
        
        with global_col :
            if user_action ==  "Review or modify existing promotions":
                is_global = st.selectbox(
                    "All Brands?",
                    ["Any", "No(False)", "Yes(True)"]
                )
            else:
                is_global = st.selectbox(
                    "All Brands?",
                    ["No(False)", "Yes(True)"]
                )
        if is_global == "No(False)":
            is_global = False
            if not promo_df.empty:
                promo_df = promo_df[promo_df["IS_GLOBAL"] == is_global]
        elif is_global == "Yes(True)":
            is_global = True
            if not promo_df.empty:
                promo_df = promo_df[promo_df["IS_GLOBAL"] == is_global]
        
        ## 2nd LAYER
        year_col, _, coop_week_col, _, _ = st.columns([2,0.2, 3, 0.2, 3])
        with year_col :
            year_filter = st.selectbox(
                "Year",
                years
            )
        if year_filter != "All":
            if not promo_df.empty:
                promo_df = promo_df[promo_df["YEAR"] == year_filter]
        
        with coop_week_col :
            week_filter = st.multiselect(
                "Coop Week(s)",
                weeks,
                []
            )
        
        if len(week_filter) > 1 :
            weeks_str = [str(x) for x in week_filter]
            weeks_str_formatted = "+".join(weeks_str)
            if not promo_df.empty:
                promo_df = promo_df[ (promo_df["WEEK"] == weeks_str_formatted) |(promo_df["WEEK"].str.split('+').apply(lambda x: any(week in x for week in weeks_str)))]
        elif len(week_filter) == 1:
            weeks_str = str(week_filter[0])
            if not promo_df.empty:
                promo_df = promo_df[ (promo_df["WEEK"] == weeks_str) |(promo_df["WEEK"].str.split('+').apply(lambda x: any(week in x for week in weeks_str)))]
        
        ## 3rd LAYER
        promo_type_col, duration_type_col, duration_col, mechanism_col = st.columns([2, 2, 2, 3])
        with promo_type_col :
            promo_type_filter = st.selectbox(
                "Promo Type",
                promo_types
            )
        if promo_type_filter == "Other":
            with promo_type_col:
                if not promo_df.empty:
                    promo_type_filter = st.text_input("New Promo Type:")
        if promo_type_filter != "Any":
            if not promo_df.empty:
                promo_df = promo_df[promo_df["PROMOTION_TYPE"] == promo_type_filter]
        
        with duration_type_col :
            duration_type_filter = st.selectbox(
                "Duration",
                duration_types
            )
        max_duration = 1
        if duration_type_filter == 'Woche(n)':
            max_duration = 52
        else: 
            max_duration = 7
        duration_filter_input=""
        if duration_type_filter != "Any":
            with duration_col :
                duration_filter = st.number_input(
                    "",
                    min_value=1,
                    max_value=max_duration
                )
                duration_filter_input = f"{str(duration_filter)} {duration_type_filter[:-3]}"
                if duration_filter > 1 :
                    if duration_type_filter == 'Woche(n)':
                        duration_filter_input += "n"
                    else:
                        duration_filter_input += "e" 
            duration_filter = f"{str(duration_filter)} {duration_type_filter[0]}"
            if not promo_df.empty:
                promo_df = promo_df[promo_df["DURATION"].str[:3] == duration_filter]
            
        with mechanism_col :
            mechanism_filter = st.selectbox(
                "Promo Mechanism",
                mechanisms
            )
        if mechanism_filter == "Other":
            with mechanism_col:
                mechanism_filter = st.text_input("New Promo Mechanism:")
        if mechanism_filter != "Any":
            if not promo_df.empty:
                promo_df = promo_df[promo_df["MECHANISM"] == mechanism_filter]
    
        ## 4th LAYER
        region_filter = st.multiselect(
            "Region(s)",
            regions,
            []
        )
        if region_filter:
            if not promo_df.empty:
                promo_df = promo_df[promo_df["REGION"].str.split('+').apply(lambda x: any(region in x for region in region_filter))]
        
        ## 5th LAYER
        cat_col, undercat_col, seg_col, underseg_col = st.columns([1, 1, 1, 1])
        with cat_col :
            cat_filter = st.multiselect(
                "Product Category(ies)",
                product_category
        )
        if cat_filter:
            if not promo_df.empty:
                promo_df = promo_df[
                        promo_df["PRODUCT_CATEGORY"].apply(lambda x: any(cat in x.split(' && ') for cat in cat_filter) if x and x != "nan" else True)
                ]
            # filter available undercategories depending on chosen categories
            product_undercategory = sorted(list(set(categories_df[categories_df["KATEGORIE"].isin(cat_filter)]
                                        ["UNTERKATEGORIE"]
                                       )))
            
            product_segment = sorted(
                categories_df[categories_df["KATEGORIE"].isin(cat_filter)]["SEGMENT"]
                .dropna()
                .unique()
            )
            product_undersegment = sorted(list(set(categories_df[categories_df["KATEGORIE"].isin(cat_filter)]
                                       ["UNTERSEGMENT"]
                                        .dropna()
                                        .unique()
                                      )))
            product_sku_id = sorted(list(set(categories_df[categories_df["KATEGORIE"].isin(cat_filter)]
                            ["SKU_ID"]
                            .dropna()
                            .unique()
                            )))
        
        else : 
            cat_filter = sorted(list(product_category))
                
        with undercat_col :
            undercat_filter = st.multiselect(
                "Undercategory(ies)",
                product_undercategory
        )
            
        if undercat_filter:
            if not promo_df.empty:
                promo_df = promo_df[
                    promo_df["PRODUCT_UNDERCATEGORY"].apply(lambda x: any(undercat in x.split(' && ') for undercat in undercat_filter) if x and x != "nan" else True)
                ]
            
            
            # filter available undercategories depending on chosen categories
            product_segment = sorted(list(set(categories_df[
                                        (categories_df["UNTERKATEGORIE"].isin(undercat_filter))
                                  ]
                                  ["SEGMENT"]
                                 )))
            product_undersegment = sorted(list(set(categories_df[
                                       (categories_df["UNTERKATEGORIE"].isin(undercat_filter))
                                  ]
                                  ["UNTERSEGMENT"]
                                 )))
            product_sku_id = sorted(list(set(categories_df[ 
                                        (categories_df["UNTERKATEGORIE"].isin(undercat_filter))
                                  ]
                                  ["SKU_ID"]
                                 )))
    
        else : 
            if user_action ==  "Review or modify existing promotions":
                undercat_filter = sorted(list(product_undercategory))
        with seg_col :
            seg_filter = st.multiselect(
                "Segment(s)",
                product_segment
        )
        if seg_filter:
            if not promo_df.empty:
                promo_df = promo_df[
                    promo_df["PRODUCT_SEGMENT"].apply(lambda x: any(seg in x.split(' && ') for seg in seg_filter) if x and x != "nan" else True)
                ]
    
            # filter available undercategories depending on chosen categories
            product_undersegment = sorted(list(set(categories_df[
                                       (categories_df["SEGMENT"].isin(seg_filter))
                                  ]
                                  ["UNTERSEGMENT"]
                                 )))
            product_sku_id = sorted(list(set(categories_df[
                                    (categories_df["SEGMENT"].isin(seg_filter))
                                    ]
                                    ["SKU_ID"]
                                )))
        else : 
            if user_action ==  "Review or modify existing promotions":
                seg_filter = sorted(list(product_segment))
            
        with underseg_col :
            underseg_filter = st.multiselect(
                "Undersegment(s)",
                product_undersegment
        )
        if underseg_filter:
            if not promo_df.empty:
                promo_df = promo_df[
                    promo_df["PRODUCT_UNDERSEGMENT"].apply(lambda x: any(underseg in x.split(' && ') for underseg in underseg_filter) if x and x != "nan" else True)   
                ]
            # filter available undercategories depending on chosen categories
            product_sku_id = sorted(list(set(categories_df[
                                (categories_df["UNTERSEGMENT"].isin(underseg_filter))
                                ]
                                ["SKU_ID"]
                                )))
        else : 
            if user_action ==  "Review or modify existing promotions":
                underseg_filter = sorted(list(product_undersegment))
            
        ## 6th LAYER
        product_id_filter = st.multiselect(
            "Product SKU(s)",
            sorted(product_sku_id)
        ) 
        if product_id_filter:
            product_id_filter_formatted=[]
            for i in range(0, len(product_id_filter)):
                product_id = product_id_filter[i][-10:-1]
                product_sku = product_id_filter[i][0:-12]
                product_id_formatted = str(product_id) + " " + str(product_sku)
                product_id_filter_formatted.append(product_id_formatted)
            promo_df = promo_df[
                promo_df["PRODUCT_ID"].apply(lambda x: any(id in x.split(' && ') for id in product_id_filter_formatted) if x and x != "nan" else True)   
            ]
                    
    
        ## Values Input
        st.write("---")
        # if user_action !=  "Review or modify existing promotions":
        #     promotion_total_cost_col, _, _ = st.columns([1, 0.2, 2])
        #     with promotion_total_cost_col:
        #         promotion_total_cost_filter = st.number_input(
        #             "Cent Off Total",
        #             min_value=0.00,
        #             value=0.0,
        #             step=1.00,
        #             help="Valeur totale de la promotion"
        #         )


                        
                
        if user_action != "Review or modify existing promotions":
            # Première ligne de colonnes: WB et le nouveau mode Cent Off
            col_wb, col_mode = st.columns([2, 3]) # WB prend 1 part, le mode Cent Off prend 2 parts pour l'espace
        
            with col_wb:
                WB_filter = st.number_input(
                    "WB",
                    min_value=0.00,
                    step=1.00,
                    key="wb_filter_input" # Ajout d'une clé unique
                )
            
            with col_mode:
                cent_off_mode = st.radio(
                    "Cent Off By Piece Mode",
                    ("INSERTED", "CALCULATED"), # Les options
                    index=1, # "CALCULATED" est l'option par défaut (index 1 car "INSERTED" est index 0)
                    horizontal=True, # Affiche les options côte à côte
                    key="cent_off_mode_radio" # Ajout d'une clé unique
                )
                # st.write(f"Mode sélectionné : {cent_off_mode}") # Utile pour le débogage
        
        
            # Ajout d'un séparateur visuel (optionnel)
            st.markdown("---") 
        
            # Deuxième ligne de colonnes: Type, Montant, et le nouveau Coût Total
            # J'ai ajusté un peu les ratios pour que l'espace soit mieux géré quand Promotion Total Cost disparaît
            col_type, col_amount, col_total_cost = st.columns([1, 1, 1]) 
        
            with col_type:
                cent_off_type_filter = st.selectbox(
                    "Cent Off By Piece Type",
                    cent_off_types,
                    key="cent_off_type_select" # Ajout d'une clé unique
                )
        
            with col_amount:
                cent_off_filter = st.number_input(
                    f"Cent Off By Piece ({cent_off_type_filter})", # Le f-string est conservé
                    min_value=0.00,
                    step=1.00,
                    key="cent_off_amount_input" # Ajout d'une clé unique
                )
            
            # Initialisation de la variable pour éviter NameError
            promotion_by_piece_total_cost = None 
        
            with col_total_cost:
                # La condition est ici !
                if cent_off_mode == "INSERTED":
                    promotion_by_piece_total_cost = st.number_input(
                        "Promotion By Piece Total Cost",
                        min_value=0.00,
                        step=1.00,
                        help="This represents the total cost per piece for the promotion.",
                        key="promo_by_piece_total_cost_input" # Ajout d'une clé unique
                    )
                else : 
                    promotion_by_piece_total_cost = pd.NA
        
            # Vous pouvez afficher les valeurs pour vérifier (pour le débogage)
            
        else:
            st.write("Section pour 'Review or modify existing promotions'")


        ############
        ###### COLUMNS CONFIGURATION ######
        # Configure column behaviors in the data editor
        column_config = {}
        disabled_cols = []
    
        ###### BRAND ######
        #If input mode, set brand to the chosen one and disable the column
        if user_action !=  "Review or modify existing promotions":
                disabled_cols.append('BRAND_NAME')
                column_config["BRAND_NAME"] = st.column_config.TextColumn(
                    default = brand_filter[0], 
                    disabled = True
                )
        else:
            column_config["BRAND_NAME"] = st.column_config.SelectboxColumn(
                options = brands,
                required = True
            )
    
        ###### IS_GLOBAL ######
        if user_action ==  "Review or modify existing promotions":
            # Configure each column based on filters
            column_config['IS_GLOBAL'] = st.column_config.CheckboxColumn(
                label='ALL BRANDS',
                required=True,
                default=False
            )
        else:
            column_config['IS_GLOBAL'] = st.column_config.CheckboxColumn(
                label='ALL BRANDS',
                default= bool(is_global), 
                disabled=True
            )
    
        ###### YEAR ######
        if user_action !=  "Review or modify existing promotions":
                disabled_cols.append('YEAR')
                column_config["YEAR"] = st.column_config.NumberColumn(
                    default = year_filter, 
                    disabled = True
                )
        else:
            column_config['YEAR'] = st.column_config.NumberColumn(
                required=True,
                min_value= years[0],
                max_value= years[-2],
                help = f"Please chose a year between {years[0]} and {years[-2]}"
            )
    
        ###### WEEK ######
        input_weeks_str_formatted=""
        if len(week_filter) >= 1:
            for i in range(0, len(week_filter)):
                input_weeks_str_formatted += str(week_filter[i]) + "+"
            input_weeks_str_formatted = input_weeks_str_formatted[:-1]
        elif len(week_filter) == 0:
            input_weeks_str_formatted=None
        if user_action !=  "Review or modify existing promotions":
            column_config['WEEK'] = st.column_config.TextColumn(
                required=True,
                default=str(input_weeks_str_formatted),
                max_chars=20,
                validate=r"^\d{1,2}(?:\+\d{1,2})*$",
                disabled= True
            )
        else:
            column_config['WEEK'] = st.column_config.TextColumn(
                required=True,
                help="Please enter the weeks in this format '24+25'",
                default=str(input_weeks_str_formatted),
                max_chars=20,
                validate=r"^\d{1,2}(?:\+\d{1,2})*$",
            )
        
        ###### PROMO_TYPE ######
        if user_action ==  "Review or modify existing promotions":
            promo_types.remove("Any")
            column_config['PROMOTION_TYPE'] = st.column_config.SelectboxColumn(
                        required=True,
                        options= promo_types
                )
        else:
            disabled_cols.append('PROMOTION_TYPE')
            column_config["PROMOTION_TYPE"] = st.column_config.TextColumn(
                        default= promo_type_filter, 
                        disabled=True
            )
    
        ###### DURATION ######
        if user_action ==  "Review or modify existing promotions":
            column_config['DURATION'] = st.column_config.TextColumn(
                        required=True
                )
        else:
            disabled_cols.append('DURATION')
            column_config["DURATION"] = st.column_config.TextColumn(
                        default= duration_filter, 
                        disabled=True
            )
            
        ###### MECHANISM ######
        if user_action ==  "Review or modify existing promotions":
            mechanisms.remove("Any")
            column_config['MECHANISM'] = st.column_config.SelectboxColumn(
                        required=True,
                        options= mechanisms
                )
        else:
            disabled_cols.append('MECHANISM')
            column_config["MECHANISM"] = st.column_config.TextColumn(
                        default= mechanism_filter, 
                        disabled=True
            )
    
        ###### REGION ######
        input_regions_str_formatted=""
        if len(region_filter) >= 1:
            for i in range(0, len(region_filter)):
                input_regions_str_formatted += str(region_filter[i]) + "+"
            input_regions_str_formatted = input_regions_str_formatted[:-1]
        else:
            input_regions_str_formatted=None
        if user_action !=  "Review or modify existing promotions":
            column_config['REGION'] = st.column_config.TextColumn(
                required=True,
                default=str(input_weeks_str_formatted),
                max_chars=20,
                validate=r"^[A-Z]{2,4}(?:\+[A-Z]{2,4})*$",
                disabled= True
            )
        else:
            column_config['REGION'] = st.column_config.TextColumn(
                required=True,
                help="Please enter the regions in this format 'SR+OT'",
                default=str(input_regions_str_formatted),
                max_chars=20,
                validate=r"^[A-Z]{2,4}(?:\+[A-Z]{2,4})*$",
            )
        
        
        ###### PRODUCT_CATEGORY ######
        input_categories_str_formatted=""
        if len(cat_filter) >= 1:
            for i in range(0, len(cat_filter)):
                input_categories_str_formatted += str(cat_filter[i]) + " && "
            input_categories_str_formatted = input_categories_str_formatted[:-4]     
        else:
            input_categories_str_formatted=None
        if user_action !=  "Review or modify existing promotions":
            column_config['PRODUCT_CATEGORY'] = st.column_config.TextColumn(
                required=True,
                default=str(input_categories_str_formatted),
                disabled= True
            )
        else:
            column_config['PRODUCT_CATEGORY'] = st.column_config.TextColumn(
                required=True,
                help="Please enter the regions in this format '{categorie1} && {categorie2}...'",
                default=str(input_categories_str_formatted),
                disabled= True
            )
            
        ###### PRODUCT_UNDERCATEGORY ######
        input_undercategories_str_formatted=""
        if len(undercat_filter) >= 1:
            for i in range(0, len(undercat_filter)):
                input_undercategories_str_formatted += str(undercat_filter[i]) + " && "
            input_undercategories_str_formatted = input_undercategories_str_formatted[:-4]     
        else:
            input_undercategories_str_formatted=None
        if user_action !=  "Review or modify existing promotions":
            column_config['PRODUCT_UNDERCATEGORY'] = st.column_config.TextColumn(
                required=True,
                default=str(input_undercategories_str_formatted),
                disabled= True
            )
        else:
            column_config['PRODUCT_UNDERCATEGORY'] = st.column_config.TextColumn(
                required=True,
                help="Please enter the regions in this format '{undercategorie1} && {undercategorie2}...'",
                default=str(input_undercategories_str_formatted),
                disabled= True
            )
        
        ###### PRODUCT_SEGMENT ######
        input_segment_str_formatted=""
        if len(seg_filter) >= 1:
            for i in range(0, len(seg_filter)):
                input_segment_str_formatted += str(seg_filter[i]) + " && "
            input_segment_str_formatted = input_segment_str_formatted[:-4]     
        else:
            input_segment_str_formatted=None
        if user_action !=  "Review or modify existing promotions":
            column_config['PRODUCT_SEGMENT'] = st.column_config.TextColumn(
                required=True,
                default=str(input_segment_str_formatted),
                disabled= True
            )
        else:
            column_config['PRODUCT_SEGMENT'] = st.column_config.TextColumn(
                required=True,
                help="Please enter the segments in this format '{segment1} && {segment2}...'",
                default=str(input_segment_str_formatted),
                disabled= True
            )
    
        
        ###### PRODUCT_UNDERSEGMENT ######
        input_undersegment_str_formatted=""
        if len(underseg_filter) >= 1:
            for i in range(0, len(underseg_filter)):
                input_undersegment_str_formatted += str(underseg_filter[i]) + " && "
            input_undersegment_str_formatted = input_undersegment_str_formatted[:-4]     
        else:
            input_undersegment_str_formatted=None
        if user_action !=  "Review or modify existing promotions":
            column_config['PRODUCT_UNDERSEGMENT'] = st.column_config.TextColumn(
                required=True,
                default=str(input_undersegment_str_formatted),
                disabled= True
            )
        else:
            column_config['PRODUCT_UNDERSEGMENT'] = st.column_config.TextColumn(
                required=True,
                help="Please enter the segments in this format '{segment1} && {segment2}...'",
                default=str(input_undersegment_str_formatted),
                disabled= True
            )
        
        
        ###### PRODUCT_ID ######
        input_product_id_str_formatted=""
        if len(product_id_filter) >= 1:
            for i in range(0, len(product_id_filter)):
                product_id = product_id_filter[i][-10:-1]
                product_sku = product_id_filter[i][0:-12]
                input_product_id_str_formatted += str(product_id) + " " + str(product_sku) + " && "
            input_product_id_str_formatted = input_product_id_str_formatted[:-4]  
            #st.write(input_product_id_str_formatted)
        else:
            input_product_id_str_formatted=None
        if user_action !=  "Review or modify existing promotions":
            column_config['PRODUCT_ID'] = st.column_config.TextColumn(
                required=True,
                default=str(input_product_id_str_formatted),
                disabled= True
            )
        else:
            column_config['PRODUCT_ID'] = st.column_config.TextColumn(
                required=True,
                help="Please enter the product ids in this format '{productid1 sku1} && {productid2 sku2}...'",
                default=str(input_product_id_str_formatted),
                disabled= True
            )

        ###### WB ######
        if user_action ==  "Review or modify existing promotions":
            column_config['WB'] = st.column_config.NumberColumn(
                required=True,
                default=0.00,
                min_value=0.00,
                step=1.00
            )

        ###### CENT_OFF_BY_PIECE ######
        if user_action ==  "Review or modify existing promotions":
            column_config['CENT_OFF_BY_PIECE'] = st.column_config.NumberColumn(
                required=True,
                default=0.00,
                min_value=0.00,
                step=1.00
            )

        ###### CENT_OFF_BY_PIECE_TYPE ######
        if user_action ==  "Review or modify existing promotions":
            column_config['CENT_OFF_BY_PIECE_TYPE'] = st.column_config.SelectboxColumn(
                required=True,
                default="AMOUNT",
                options=cent_off_types
            )
        
        ###### promotion_total_cost ######
        if user_action == "Review or modify existing promotions":
            column_config['promotion_total_cost'] = st.column_config.NumberColumn(
                required=False, # Peut être NULL pour les anciennes promotions
                default=0.00,
                min_value=0.00,
                step=1.00,
                format="%.2f"
            )
            
    
        if user_action ==  "Review or modify existing promotions":
            st.subheader("Filtered Promotions")
            if promo_df.empty: 
                st.warning("Sorry no promotion matches with the selected criteria. If you want to create such a promotion, please select 'Input a new promotion'.")
            else:
                with st.expander("Learn how to interact with the following table" ):
                    st.write("""
                    In the table below you can input multiple changes as needed.
                    \n\n 
                    You can modify the values in certain columns in case a mistake was made for example. 
                    \n\n
                    Be minduful of the format required in these columns, you can hover on the column names for more info regarding the formats.
                    If you input values in the wrong format it simply won't be saved in the table.
                    \n\n
                    The columns PRODUCT_CATEGORY, PRODUCT_UNDERCATEGORY, PRODUCT_SEGMENT, PRODUCT_UNDERSEGMENT and PRODUCT_ID cannot be modified. Their values are indeed too sensitive and subject to typing or spelling mistakes.
                    If you see a promotion that does not have the correct values in these columns, we suggest you to delete the line.
                    To do so, select the line in the first column on the left of the table and either press "DELETE" or click on the little bin icon at the top right of the table.
                    Then you can choose "Input a new promotion" at the top of this page and redefine the promotion correctly.
                    \n\n
                    Be careful! The following table enables you to create new lines. However we recommend creating new promotions in the 'Input a new promotion' section (select above).
                    Any line created here will not be saved.
                    \n\n
                    After your modifications, please don't forget to save your changes by pressing the 'Save' button.
                    """)
                df_sorted = promo_df.copy()
                def get_min_week(week_str):
                    if pd.isna(week_str):
                        return float('inf')  # ou None si tu veux les laisser à part
                    return min(int(w) for w in week_str.split('+'))
                    
                df_sorted['WEEK_MIN'] = df_sorted['COOP_WEEK'].apply(get_min_week)
                df_sorted = df_sorted.sort_values(['YEAR', 'WEEK_MIN'], ascending=[True, True])
                
                promo_df_edited = st.data_editor(
                        df_sorted, 
                        disabled=disabled_cols, 
                        column_order=('YEAR', 'COOP_WEEK', 'PROMOTION_TYPE', 'IS_GLOBAL', 'DURATION',
                                      'REGION', 'MECHANISM', 'BRAND_NAME', 
                                      'PRODUCT_CATEGORY', 'PRODUCT_UNDERCATEGORY', 'PRODUCT_SEGMENT', 
                                      'PRODUCT_UNDERSEGMENT', 'PRODUCT_ID', 
                                      'WB', 'CENT_OFF_BY_PIECE', 'CENT_OFF_BY_PIECE_TYPE', 'PROMOTION_TOTAL_COST', 'TIMESTAMP', 'USER_ID'), 
                        column_config=column_config,
                        hide_index = True, 
                        num_rows='dynamic',
                        width=10000
                    )
                promo_df_edited = promo_df_edited.drop('WEEK_MIN', axis=1)

                # Compare the edited dataframe and the original one. If there are some differences, save those
                new_rows_df = pd.concat([promo_df_edited, promo_df]).drop_duplicates(keep=False)
                new_rows_df = new_rows_df.drop_duplicates(subset=["PROMOTION_ID"], keep="first")
                og_ids = set(promo_df["PROMOTION_ID"])
                curr_ids = set(promo_df_edited["PROMOTION_ID"])
                deleted_ids = og_ids - curr_ids
                new_rows_df = new_rows_df[new_rows_df["PROMOTION_ID"].isin(curr_ids)]

                #Check if promo_level is defined in all the rows, 
                #if not display a warning message explaining that some lines will not be saved because they were created by adding a line which is forbidden
                new_rows_df = new_rows_df.dropna(subset=["PROMOTION_LEVEL"])
                
                save_as_excel_button(promo_df_edited, "promotions", "dl_promo_df_edited")
                
                save_button_col, save_text_col = st.columns((2, 3))
                if 'was_saved' not in st.session_state:
                    st.session_state.was_saved = False
                with save_button_col:
                    # Button to save the changes made to the dataframe
                    if st.button("Save changes", key="save_test_partner", type="primary"):
                        st.session_state.missing_uptodate = False
                        st.session_state.was_saved = True
                    
                        with st.spinner('Your data is being uploaded'):
                            if not new_rows_df.empty:
                                update_snowflake_table(CONFIG["PROMO_TABLE"], new_rows_df, numeric_columns=["PROMO_ID", "YEAR", "WB", "CENT_OFF_BY_PIECE", "promotion_total_cost"])
                            if len(deleted_ids)>0:
                                delete_rows_table(CONFIG["PROMO_TABLE"], deleted_ids)
                        with save_text_col:
                            st.success('The table has been updated.')
                        
                        
            
            
        else:
            promo_level = ""
            new_promo_df = pd.DataFrame({
                    'YEAR':[year_filter],
                    'WEEK':[input_weeks_str_formatted], 
                    'COOP_WEEK':[input_weeks_str_formatted],
                    'PROMOTION_TYPE':[promo_type_filter],
                    'IS_GLOBAL':[is_global],
                    'DURATION':[duration_filter_input],
                    'REGION':[input_regions_str_formatted],
                    'MECHANISM':[mechanism_filter],
                    'BRAND_NAME':[brand_filter],
                    'PROMOTION_LEVEL' :[promo_level],
                    'PRODUCT_CATEGORY':[input_categories_str_formatted if input_categories_str_formatted != "" else None],
                    'PRODUCT_UNDERCATEGORY':[input_undercategories_str_formatted if input_undercategories_str_formatted != "" else None],
                    'PRODUCT_SEGMENT':[input_segment_str_formatted if input_segment_str_formatted != "" else None],
                    'PRODUCT_UNDERSEGMENT':[input_undersegment_str_formatted if input_undersegment_str_formatted != "" else None],
                    'PRODUCT_ID':[input_product_id_str_formatted if input_product_id_str_formatted != "" else None],
                    'WB':[WB_filter],
                    'CENT_OFF_BY_PIECE':[cent_off_filter],
                    'CENT_OFF_BY_PIECE_TYPE':[cent_off_type_filter],
                    'PROMOTION_BY_PIECE_TOTAL_COST':[promotion_by_piece_total_cost], 
                    'TIMESTAMP' : [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                    'USER_ID' : [st.experimental_user.email],
                    'MODE_CENT_OFF_BY_PIECE': [cent_off_mode]
                })
            level_df = new_promo_df[["PRODUCT_CATEGORY", "PRODUCT_UNDERCATEGORY", "PRODUCT_SEGMENT", "PRODUCT_UNDERSEGMENT", "PRODUCT_ID"]].copy()

            ##### CALCULATE PROMO_LEVEL #####
            promo_level_df = level_df.copy() 
            nan_value = float("NaN")
            promo_level_df.replace("", nan_value, inplace=True)
            promo_level_df.dropna(how='all', axis=1, inplace=True)
            if len(promo_level_df.columns) == 0:
                promo_level_col = ""
            if promo_level_df.empty:
                st.warning("Please chose at least one category, undercategory, segment AND/OR undersegment.")
            else:
                promo_level_col = promo_level_df.columns[-1]
                if promo_level_col == "PRODUCT_ID":
                    promo_level = "PRODUCT"
                else:
                    promo_level = promo_level_col.split('_', 1)[1]
                new_promo_df["PROMOTION_LEVEL"]=promo_level
    
    
                ##### CALCULATE ALL LEVELS #####
                levels = list(level_df.columns)
                last_level_index = levels.index(promo_level_col)
                last_level_value = new_promo_df[promo_level_col][0]
                last_level_values = last_level_value.split(" && ")
                if promo_level_col == "PRODUCT_ID":
                    category_level_col = "CLIENT_ARTIKEL_ID"
                    for i in range (0, len(last_level_values)):
                        last_level_values[i] = last_level_values[i][:9]
                    
                else: 
                    category_level_col = promo_level.replace("D", "T")
                    category_level_col = category_level_col.replace("C", "K")
                    category_level_col = category_level_col.replace("Y", "IE")
    
                for i in range (1, last_level_index+1):
                    curr_level_col = levels[last_level_index - i]
                    curr_cat_col = curr_level_col
                    curr_cat_col = curr_cat_col.split('_', 1)[1]
                    curr_cat_col = curr_cat_col.replace("D", "T")
                    curr_cat_col = curr_cat_col.replace("C", "K")
                    curr_cat_col = curr_cat_col.replace("Y", "IE")
                    
                    curr_level_value = categories_df[(categories_df[category_level_col].isin(last_level_values))][curr_cat_col].reset_index(drop=True)
                    curr_level_value = " && ".join(sorted(set(filter(None, curr_level_value))))

                    new_promo_df[curr_level_col]=curr_level_value if curr_level_value != "" else None
                    
                    product_sku_id = set(categories_df[
                                        (categories_df["UNTERSEGMENT"].isin(underseg_filter))
                                        ]
                                        ["SKU_ID"]
                                        )
                st.write("---")
                #Check if a similar promotion already exists
                similar_promo_exists = promo_df.fillna("").apply(
                    lambda row: all(
                        row[1:16].astype(str).str.strip().str.lower() == 
                        new_promo_df.iloc[0, 0:15].fillna("").astype(str).str.strip().str.lower()
                    ), 
                    axis=1
                ).any()

                if similar_promo_exists :
                    st.warning("""
                        It appears this promotion already exists (see below).
                        \n\n If you want to edit it, please select "Review or modify existing promotions" at the top of the page and 
                        edit the values in the displayed table. 
                    """)
                    st.dataframe(promo_df, hide_index=True)

                else:
                    st.write("""
                            Here is the resulting promotion below. \n\n
                            Please check that all cells are filled correctly. 
                            If there is a mistake, please correct it using the option boxes above. \n\n
                            Once you are sure about your data, you can save it by pressing the button below.
                            """
                            )
                    
                    result_df = st.data_editor(
                                new_promo_df, 
                                disabled=new_promo_df.columns, 
                                column_order=('YEAR', 'COOP_WEEK', 'PROMOTION_TYPE', 'IS_GLOBAL', 'DURATION',
                                              'REGION', 'MECHANISM', 'BRAND_NAME', 
                                              'PRODUCT_CATEGORY', 'PRODUCT_UNDERCATEGORY', 'PRODUCT_SEGMENT', 
                                              'PRODUCT_UNDERSEGMENT', 'PRODUCT_ID', 'WB', 'CENT_OFF_BY_PIECE', 
                                              'CENT_OFF_BY_PIECE_TYPE', 'MODE_CENT_OFF_BY_PIECE', 'PROMOTION_BY_PIECE_TOTAL_COST'), 
                                column_config=column_config,
                                hide_index = True, 
                                width=10000
                    )
                    missing_vals_count = 0
                    mandatory_fileds_str = ""
                    mandatory_fileds = ['COOP_WEEK', 'PROMOTION_TYPE', 'REGION', 'MECHANISM']
                    for i in range (0, len(mandatory_fileds)):
                        if result_df[mandatory_fileds[i]][0] == None or result_df[mandatory_fileds[i]][0] == "":
                            mandatory_fileds_str += f"- {mandatory_fileds[i]} \n\n"
                            missing_vals_count += 1
                    if missing_vals_count != 0 :
                        with st.container(border=True):
                            st.write("Careful! there are some empty values that should be filled, see the list below and make sure to fill them before saving. \n\nOtherwise you won't be able to save your inputs.")
                            st.write("Missing values in columns: ")
                            st.write(mandatory_fileds_str)
                    else:
                        save_button_col, save_text_col = st.columns((2, 3))
    
                        if 'was_saved' not in st.session_state:
                            st.session_state.was_saved = False
                
                        with save_button_col:
                            # Button to save the changes made to the dataframe
                            if st.button("Save changes", key="save_test_partner", type="primary"):
                                st.session_state.missing_uptodate = False
                                st.session_state.was_saved = True
                
                                with st.spinner('Your data is being uploaded'):
                                    update_table_no_truncate(CONFIG["PROMO_TABLE"], new_promo_df, numeric_columns=["PROMO_ID", "YEAR", "WB", "CENT_OFF_BY_PIECE", "promotion_total_cost"])
                                
                                with save_text_col:
                                    st.success('Congratulations! The table has been updated.')
        
        if 'was_saved' not in st.session_state:
            st.session_state.was_saved = False
            
        if st.session_state.was_saved == True:
            st.write("---")
            loaded_table = SESSION.table(promo_table)
            updated_promo_df = loaded_table.to_pandas()
            #TODO
            st.subheader("Here are all of your promotions updated")
            if user_action !=  "Review or modify existing promotions":
                st.caption("""Your new promotion is at the end of the table, you can either scroll to view it or click on 'PROMOTION_ID' to order the column.""")
            
            df_sorted = updated_promo_df.copy()
            def get_min_week(week_str):
                if pd.isna(week_str):
                    return float('inf')
                return min(int(w) for w in week_str.split('+'))
                    
            df_sorted['WEEK_MIN'] = df_sorted['COOP_WEEK'].apply(get_min_week)
            df_sorted = df_sorted.sort_values(['YEAR', 'WEEK_MIN'], ascending=[True, True])
            df_sorted = df_sorted.drop('WEEK_MIN', axis=1)
            column_config = {
                    "IS_GLOBAL": st.column_config.Column(
                        "ALL BRANDS"
                    )
                }
            st.dataframe(df_sorted, hide_index=True, column_config=column_config)
            save_as_excel_button(updated_promo_df, "promotions", "dl_updated_promo_df")
        else:
            st.subheader("All Promotions")
            column_config = {
                    "IS_GLOBAL": st.column_config.Column(
                        "ALL BRANDS"
                    )
                }
            df_sorted = promo_df_all.copy()
            def get_min_week(week_str):
                if pd.isna(week_str):
                    return float('inf')
                return min(int(w) for w in week_str.split('+'))
                
            df_sorted['WEEK_MIN'] = df_sorted['COOP_WEEK'].apply(get_min_week)
            df_sorted = df_sorted.sort_values(['YEAR', 'WEEK_MIN'], ascending=[True, True])
            df_sorted = df_sorted.drop('WEEK_MIN', axis=1)
            st.dataframe(df_sorted, column_config=column_config)

if __name__ == '__main__':
    
    #The main function that initializes the Streamlit application. 
    #It retrieves the user's role, and calls the body of the application
    #to display content based on the user's role.
    
    #Returns:
    #- None
    

    # Call the body function to render the main content of the app, passing in the user's role
    body()


    
def filter_df(self, cols, vals):
        """
        Filters the DataFrame based on matching column values. Returns a new instance of 
        StreamlitDataframe containing the filtered DataFrame.

        Args:
            cols (list): A list of column names to filter by.
            vals (list): A list of values corresponding to each column in 'cols'.

        Returns:
            StreamlitDataframe: A new StreamlitDataframe instance with the filtered DataFrame, 
                                or None if the number of columns and values do not match.
        """

        # Check if the number of columns matches the number of values
        if len(cols) != len(vals): 
            st.warning('Your number of columns if different from your number of values')
            return None
        else: 
            # Filter the DataFrame based on the specified columns and values
            filtered_df = self.df.copy()

            for i_col in range (0, len(cols)):
                if vals[i_col] != 'All':
                    if isinstance(vals[i_col], str):
                        filtered_df = filtered_df[filtered_df[f'{cols[i_col]}'] == vals[i_col]]
                    else:
                        filtered_df = filtered_df[filtered_df[f'{cols[i_col]}'] == vals[i_col]]            
                
            # Return a new StreamlitDataframe instance containing the filtered DataFrame
            return filtered_df
            
def save_as_csv_button(df, df_name):
        """
        This function generates a Streamlit button to allow users to download a DataFrame as a CSV file.

        - The DataFrame (`df`) is first rounded to 2 decimal places for consistency in numerical data.
        - The `convert_df` function, which is decorated with `@st.cache_data`, converts the DataFrame to a CSV format 
        and caches the result to avoid redundant computations during reruns.
        - The CSV data is then passed to the `st.download_button` method, which creates a download button in the Streamlit app.
        - When the button is clicked, the CSV file will be downloaded with the specified file name (`df_name.csv`).
        """
        # Round all numbers so they only have two numbers after the comma
        df = df.round(2)

        @st.cache_data
        def convert_df(df):
            # Cache the DataFrame conversion to CSV format to improve efficiency on reruns
            return df.to_csv().encode("utf-8")
        
        csv = convert_df(df)
        
        st.download_button(
            label="Download data as CSV",
            data=csv,
            file_name=f"{df_name}.csv",
            mime="text/csv",
        )