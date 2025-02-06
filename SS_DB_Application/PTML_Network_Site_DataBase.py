import os
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle, Font, Alignment, Border, Side, PatternFill
from openpyxl.styles import borders
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.drawing.image import Image
import warnings
import sys
from datetime import datetime

#-------------------------------------------------------------
# Set the page title and favicon (logo in browser tab)
st.set_page_config(
    page_title="PTML Network Site DataBase",  # Title of the web page
    page_icon="D:/Advance_Data_Sets/PTML_DB/Huawei.jpg",  # Path to the image file you want to use as favicon
)
#---------------------------------------------------------
# Create two columns
col1, col2 = st.columns([2, 1])  # Adjust column width ratio as needed

# Left column: Markdown with personal details
with col1:
    st.markdown("""
        # [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
        Sr. RF Planning & Optimization Engineer  
        BSc Telecommunications Engineering, School of Engineering  
        MS Data Science, School of Business and Economics  
        **University of Management & Technology**  
        **Mobile:**     +923018412180  
        **Email:**  umersaeed81@hotmail.com  
        **Address:** Dream Gardens, Defence Road, Lahore  
    """)

# Right column: Logo image centered
with col2:
    # Use a placeholder to center the image within the column
    col2.markdown("<div style='display: flex; justify-content: center;'>", unsafe_allow_html=True)
    col2.image("D:/Advance_Data_Sets/PTML_DB/Huawei.jpg", width=100)
    col2.markdown("</div>", unsafe_allow_html=True)
#-----------------------------------------------------------------------------------------------#
# Replace with your given date in YYYY-MM-DD format
given_date_str = "2026-01-03"  # Example date
given_date = datetime.strptime(given_date_str, "%Y-%m-%d")
# Get today's date
today_date = datetime.today()

# Compare dates
if given_date < today_date:
    print("OK.")
    sys.exit(0)  # Exit the script
else:
    print("NoK.")
#-----------------------------------------------------------------------------------------------#     
# Streamlit App
#st.title("PTML Network Site DataBase")
st.markdown("<h1 style='color: maroon;'>PTML Network Site DataBase</h1>", unsafe_allow_html=True)
#-----------------------------------------------------------------------------------------------# 
# # Input File Paths
# input_path = st.text_input("**Input Excel File Path 📂**", "D:/Advance_Data_Sets/PTML_DB/Cells_DB_Mid_Dec_2024.xlsx")

# # Check if the file extension is .xlsx
# if input_path and not input_path.lower().endswith('.xlsx'):
#     st.error("Only .xlsx files are allowed! Please provide a valid file.")

# # Check if the folder exists
# elif not os.path.exists(os.path.dirname(input_path)):
#     st.error(f"The folder path does not exist: {os.path.dirname(input_path)}")

# # Check if the file exists
# elif not os.path.isfile(input_path):
#     st.error(f"The file does not exist: {input_path}")

# else:
#     st.success("Valid file selected.")


#-----------------------------------------------------------------------------------------------#
st.markdown("<h3 style='color: maroon;'>Input and Output File Path</h3>", unsafe_allow_html=True)
#-----------------------------------------------------------------------------------------------# 
# # Input File Path
# input_path = st.text_input("**Input Excel File Path 📂**", "D:/Advance_Data_Sets/PTML_DB/Cells_DB_Mid_Dec_2024.xlsx")

# # Check if the file extension is .xlsx
# if input_path and not input_path.lower().endswith('.xlsx'):
#     st.error("Only .xlsx files are allowed! Please provide a valid file.")

# # Check if the folder exists
# elif not os.path.exists(os.path.dirname(input_path)):
#     st.error(f"The folder path does not exist: {os.path.dirname(input_path)}")

# # Check if the file exists
# elif not os.path.isfile(input_path):
#     st.error(f"The file does not exist: {input_path}")

# else:
#     # Check the number of sheets in the Excel file
#     try:
#         xls = pd.ExcelFile(input_path)
#         num_sheets = len(xls.sheet_names)
        
#         if num_sheets < 3:
#             st.error(f"The Excel file must contain at least 3 sheets! Found only {num_sheets}.")
#         else:
#             st.success(f"Valid file selected with {num_sheets} sheets.")
    
#     except Exception as e:
#         st.error(f"Error reading the Excel file: {e}")

#-----------------------------------------------------------------------------------------------# 
# Input File Path
input_path = st.text_input("**Input Excel File Path 📂**", "D:/Advance_Data_Sets/PTML_DB/Cells_DB_Mid_Dec_2024.xlsx")

# Check if the file extension is .xlsx
if input_path and not input_path.lower().endswith('.xlsx'):
    st.error("Only .xlsx files are allowed! Please provide a valid file.")

# Check if the folder exists
elif not os.path.exists(os.path.dirname(input_path)):
    st.error(f"The folder path does not exist: {os.path.dirname(input_path)}")

# Check if the file exists
elif not os.path.isfile(input_path):
    st.error(f"The file does not exist: {input_path}")

else:
    try:
        # Load the Excel file
        xls = pd.ExcelFile(input_path)
        num_sheets = len(xls.sheet_names)

        # Check if the Excel file has at least 3 sheets
        if num_sheets < 3:
            st.error(f"The Excel file must contain at least 3 sheets! Found only {num_sheets}.")
        else:
            missing_column_sheets = []
            
            # Check if all sheets contain the 'PMO Status' column
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, nrows=0)  # Read only headers
                if 'PMO Status' not in df.columns:
                    missing_column_sheets.append(sheet)

            if missing_column_sheets:
                st.error(f"The following sheets are missing the 'PMO Status' column: {', '.join(missing_column_sheets)}")
            else:
                st.success(f"Valid file selected with {num_sheets} sheets, and all sheets contain the 'PMO Status' column.")

    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")

#-----------------------------------------------------------------------------------------------#                
# Output File Paths

# Get today's date in the required format
today_date = datetime.today().strftime("%d%m%Y")
base_path = f"D:/Advance_Data_Sets/PTML_DB/PTML_Cell_List_{today_date}"
file_extension = ".xlsx"
output_path = f"{base_path}{file_extension}"

# Check if the base_path exists
if not os.path.exists(os.path.dirname(base_path)):
    st.error(f"Error: The specified directory {os.path.dirname(base_path)} does not exist.")
else:
    # Function to generate a unique filename
    def get_unique_filename(base_path, extension):
        counter = 1
        new_filename = f"{base_path}{extension}"
        while os.path.exists(new_filename):
            new_filename = f"{base_path}_{counter}{extension}"
            counter += 1
        return new_filename

    # Get a unique output path if the file already exists
    output_path = get_unique_filename(base_path, file_extension)

    # Allow the user to see and modify the suggested filename
    user_output_path = st.text_input("**Output Excel File Path 📤**", output_path)

    # Verify the folder path and filename entered by the user
    if user_output_path:
        folder = os.path.dirname(user_output_path)
        if not os.path.exists(folder):
            st.error(f"Error: The specified folder '{folder}' does not exist.")
        else:
            if os.path.exists(user_output_path):
                st.warning("The file already exists, a unique name will be generated.")
            else:
                st.success(f"File path is valid: {user_output_path}")

#--------------------------------------------------------------------------------------
# Fixed Column Order
def_columns = {
    "2G": ['Tech region', 'Site ID', 'Site Type', 'Cell ID simple', 'Current Hgt', 'Beam Width', 'Current Azimuth',
           'Current E-Tilt', 'New MSC ID', 'New BSC', 'LAC', 'CGI', 'City Name', 'Province', 'District', 'Tehsil',
           'Sector Name', 'Covered Area', 'BSIC', 'BCCH ARFCN', 'Long', 'Degree', 'Min', 'Sec', 'Latitude',
           'GSM Antenna', 'DCS Antenna', 'DB Antenna', 'TRIB Antenna', 'Total Antenna Count', 'PMO Status'],
    
    "3G": ['Tech region', '2G Site ID', '3G Site ID', 'CL Site Tech', 'Freq. Band', 'Cell ID simple', 'PSC',
           'RNC ID', 'LAC', 'CGI', '3G Site Name', 'Current Hgt', 'Current Azimuth', 'Current E-Tilt', 'City',
           'Province', 'District', 'Tehsil', 'Longitude', 'Latitude', 'Site Type', 'Frequency DOWNLINK',
           'Frequency UPLINK', 'Horizontal BW', 'Vertical BW', 'Antenna Type', 'PMO Status'],
    
    "4G": ['Tech region', '4G Site ID', 'Cell No.', '2G Site ID', '3G Site ID', 'eNodeB ID', '4G spectrum BW',
           'Cell Freq. Band', 'CL Site Tech', 'ECI', 'ECGI', 'TAC', '4G Site Name', 'Current Hgt',
           'Current Azimuth', 'Current E-Tilt', 'Latitude', 'Longitude', 'Site Type', 'City', 'Province',
           'District', 'Tehsil', 'New Antenna Type', 'PMO Status']
}
#-----------------------------------------------------------------------------------------------# 
@st.cache_data
def load_sheets(file_path):
    """Load sheet names from the Excel file."""
    try:
        return pd.ExcelFile(file_path).sheet_names
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return []

# Load available sheets
sheet_names = load_sheets(input_path)

st.markdown("<h3 style='color: maroon;'>Select Sheet Name</h3>", unsafe_allow_html=True)

# Default sheet selection
selected_sheets = {
    "2G": st.selectbox("**Select 2G Sheet Name 📄**", sheet_names, index=sheet_names.index("2G") if "2G" in sheet_names else 0),
    "3G": st.selectbox("**Select 3G Sheet Name 📄**", sheet_names, index=sheet_names.index("3G") if "3G" in sheet_names else 0),
    "4G": st.selectbox("**Select 4G Sheet Name 📄**", sheet_names, index=sheet_names.index("4G") if "4G" in sheet_names else 0)
}



# # Let user modify required columns (without column order input)
# selected_columns = {
#     "2G": st.multiselect("**Select Required Columns for 2G 📝**", def_columns["2G"], default=def_columns["2G"]),
#     "3G": st.multiselect("**Select Required Columns for 3G 📝**", def_columns["3G"], default=def_columns["3G"]),
#     "4G": st.multiselect("**Select Required Columns for 4G 📝**", def_columns["4G"], default=def_columns["4G"]),
# }

st.markdown("<h3 style='color: maroon;'>Select the requied Columns</h3>", unsafe_allow_html=True)

# Create three columns
col1, col2, col3 = st.columns(3)

with col1:
    selected_columns_2G = st.multiselect("**Select Required Columns for 2G 📝**", def_columns["2G"], default=def_columns["2G"])

with col2:
    selected_columns_3G = st.multiselect("**Select Required Columns for 3G 📝**", def_columns["3G"], default=def_columns["3G"])

with col3:
    selected_columns_4G = st.multiselect("**Select Required Columns for 4G 📝**", def_columns["4G"], default=def_columns["4G"])


# Validate PMO Status selection
missing_pmo = []
for tech, selected_cols in {"2G": selected_columns_2G, "3G": selected_columns_3G, "4G": selected_columns_4G}.items():
    if "PMO Status" not in selected_cols:
        missing_pmo.append(tech)

if missing_pmo:
    st.warning(f"⚠️ 'PMO Status' must be selected for {', '.join(missing_pmo)}")

# Store selections in a dictionary
selected_columns = {
    "2G": selected_columns_2G,
    "3G": selected_columns_3G,
    "4G": selected_columns_4G
}

#-----------------------------------------------------------------------
# # Let user select PMO Status values
# pmo_status_options = ["CL", "NCL", "NA", "NCL-relocation", "Planned"]
# selected_pmo_status = st.multiselect("**Select PMO Status values 🛠️**", pmo_status_options, default=pmo_status_options)
#-----------------------------------------------------------------------
# # Let user select PMO Status values using checkboxes
# pmo_status_options = ["CL", "NCL", "NA", "NCL-relocation", "Planned"]
# selected_pmo_status = []

# st.write("**Select PMO Status values 🛠️**")
# for status in pmo_status_options:
#     if st.checkbox(status, value=True):
#         selected_pmo_status.append(status)
#-----------------------------------------------------------------------

st.markdown("<h3 style='color: maroon;'>Select required PMO Status</h3>", unsafe_allow_html=True)

# Let user select PMO Status values using checkboxes
pmo_status_options = ["CL", "NCL", "NA", "NCL-relocation", "Planned"]
selected_pmo_status = []

st.write("**Select PMO Status values 🛠️**")
for status in pmo_status_options:
    default_checked = status in ["CL", "NCL"]
    if st.checkbox(status, value=default_checked):
        selected_pmo_status.append(status)
#-----------------------------------------------------------------------
# Load and filter data
def load_and_filter_data(sheet_name, selected_cols, pmo_status_values):
    """Load data from the selected sheet, enforce selected column order, and filter rows based on PMO Status."""
    try:
        df = pd.read_excel(input_path, sheet_name=sheet_name, dtype={'ECGI': str})
        df = df[[col for col in selected_cols if col in df.columns]]
        df = df[df['PMO Status'].isin(pmo_status_values)].reset_index(drop=True).drop(columns=['PMO Status'])
        return df
    except Exception as e:
        st.error(f"Error processing sheet {sheet_name}: {e}")
        return pd.DataFrame()

# Format Excel file
def format_excel(file_path):
# Set Tab Color (All the Tabs)
    """Apply formatting to the Excel file."""
    wb = openpyxl.load_workbook(file_path)
    colors = ["00B0F0", "0000FF", "ADD8E6", "87CEFA"]
    
    for i, ws in enumerate(wb):
        ws.sheet_properties.tabColor = colors[i % len(colors)]

# Font, Alignment and Border (All the Sheets)
    border = Border(left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin'))
    
    style = NamedStyle(name="styled_cell")
    style.font = Font(name='Calibri Light', size=8)
    style.alignment = Alignment(horizontal='center', vertical='center')
    style.border = border
    wb.add_named_style(style)
    
    wb.calculation.calcMode = 'manual'
    
    for ws in wb:
        for row in ws.iter_rows():
            for cell in row:
                cell.style = "styled_cell"
    
    wb.calculation.calcMode = 'auto'

# WrapText of header and Formatting (All the sheets)    
    for ws in wb:
        for row in ws.iter_rows(min_row=1, max_row=1):
            for cell in row:
                cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
                cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
                cell.font = Font(color="FFFFFF", bold=True, size=11, name='Calibri Light')

# Set Filter on the Header (All the sheet)    
    for ws in wb:
        # Get the first row
        first_row = ws[1]
        # Apply the filter on the first row
        ws.auto_filter.ref = f"A1:{get_column_letter(len(first_row))}1"


# Set Column Width (All the sheets)
    for ws in wb.worksheets:
        # Iterate over all columns in the sheet
        for column in ws.columns:
            # Get the current width of the column
            current_width = ws.column_dimensions[column[0].column_letter].width
            # Get the maximum width of the cells in the column
            length = max(len(str(cell.value)) for cell in column)
            # Set the width of the column to fit the maximum width, if it's greater than the current width
            if length > current_width:
                ws.column_dimensions[column[0].column_letter].width = length
    
# Insert a New Sheet (as First Sheet)    
    ws_title = wb.create_sheet("Title Page", 0)
    # Merge Specific Row and Columns
    ws_title.merge_cells(start_row=12, start_column=5, end_row=18, end_column=17)
    # Fill the Merge Cells
    ws_title.cell(row=12, column=5).value = 'PTML Network Site DataBase'

# Formatting Tital Page Report    
    # Access the first row starting from row 3
    first_row1 = list(ws_title.rows)[11]
    # Iterate through the cells in the first row starting from column E
    for cell in first_row1[4:]:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
        cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type = "solid")
        font = openpyxl.styles.Font(color="FFFFFF",bold=True,size=60,name='Calibri Light')
        cell.font = font

# Hide the gridlines    
    ws_title.sheet_view.showGridLines = False
# Hide the headings
    ws_title.sheet_view.showRowColHeaders = False

# Hyper Link For Title Page
    # loop through all sheets in the workbook and insert the hyperlink to each sheet
    row = 22
    for ws in wb:
        if ws.title != "Title Page":
            hyperlink_cell = ws_title.cell(row=row, column=5)
            hyperlink_cell.value = ws.title
            hyperlink_cell.hyperlink = "#'{}'!A1".format(ws.title)
            hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            #hyperlink_cell.border = border
            hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            # set the height of the cell
            ws_title.row_dimensions[row].height = 15
            # set the colum width of column 5
            ws_title.column_dimensions[get_column_letter(5)].width = 30
            row += 1

# Hyper Link For Sub Pages    
    # Loop through all sheets in the workbook and insert the hyperlink to each sheet
    for i, ws in enumerate(wb.worksheets):
        # Check if the sheet is not the Title Page
        if ws.title != "Title Page":
            # Add hyperlink to cell in the last column+2 of the sheet
            hyperlink_cell = ws.cell(row=2, column=ws.max_column+2)
            hyperlink_cell.value = "Back to Table of Contents"
            hyperlink_cell.hyperlink = "#'{}'!E{}".format("Title Page", 22)
            hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            # Set border on the hyperlink column
            for row in ws.iter_rows(min_row=2, max_row=4, min_col=hyperlink_cell.column, max_col=hyperlink_cell.column):
                for cell in row:
                    cell.border = border
            # Set width of the hyperlink column
            col_letter = openpyxl.utils.get_column_letter(hyperlink_cell.column)
            ws.column_dimensions[col_letter].width = 25

            # Add hyperlink to cell in the last column+2 of the sheet for next sheet
            if i < len(wb.worksheets)-1:
                next_hyperlink_cell = ws.cell(row=3, column=ws.max_column)
                next_hyperlink_cell.value = "Next Sheet"
                next_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i+1].title)
                next_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
                next_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            
            # Add hyperlink to cell in the last column+2 of the sheet for previous sheet
            if i > 0:
                prev_hyperlink_cell = ws.cell(row=4, column=ws.max_column)
                prev_hyperlink_cell.value = "Previous Sheet"
                prev_hyperlink_cell.hyperlink = "#'{}'!A1".format(wb.worksheets[i-1].title)
                prev_hyperlink_cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
                prev_hyperlink_cell.alignment = openpyxl.styles.Alignment(horizontal='center')

## Inset Image on the title page    
    # inset the Huawei logo
    img = Image('D:/Advance_Data_Sets/PTML_DB/Huawei.jpg')
    img.width = 7 * 15
    img.height = 7 * 15
    ws_title.add_image(img,'E3')

    # inset the PTCL logo
    img1 = Image('D:/Advance_Data_Sets/PTML_DB/PTCL.png')
    ws_title.add_image(img1,'M3')

    wb.save(file_path)
    return file_path

st.markdown("<h3 style='color: maroon;'>Export Output</h3>", unsafe_allow_html=True)

# Process and export data
if st.button("Process and Export Data 🚀"):
    with st.spinner("Processing data... ⏳"):
        df_2g = load_and_filter_data(selected_sheets["2G"], selected_columns["2G"], selected_pmo_status)
        df_3g = load_and_filter_data(selected_sheets["3G"], selected_columns["3G"], selected_pmo_status)
        df_4g = load_and_filter_data(selected_sheets["4G"], selected_columns["4G"], selected_pmo_status)

        with pd.ExcelWriter(output_path) as writer:
            df_2g.to_excel(writer, sheet_name="2G_Cells", index=False)
            df_3g.to_excel(writer, sheet_name="3G_Cells", index=False)
            df_4g.to_excel(writer, sheet_name="4G_Cells", index=False)

        #st.success(f"Filtered data saved to {output_path} ✅")

        # Format the file
        formatted_file = format_excel(output_path)
        st.success(f"Formatted file saved: {formatted_file} 🎨")
