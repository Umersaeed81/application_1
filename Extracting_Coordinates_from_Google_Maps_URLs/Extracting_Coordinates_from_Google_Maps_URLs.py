import os
import re
import requests
import pandas as pd
import streamlit as st
from tqdm import tqdm
import openpyxl
from datetime import datetime

#-------------------------------------------------------------
# Set the page title and favicon (logo in browser tab)
st.set_page_config(
    page_title="Extracting Coordinates from Google Maps URLs and Exporting to Excel",
    page_icon="D:/Advance_Data_Sets/PTML_DB/Huawei.jpg",
)

# Layout with two columns
col1, col2 = st.columns([2, 1])
#-------------------------------------------------------------
# Left column: Personal details
with col1:
    st.markdown("""
        # [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
        **Sr. RF Planning & Optimization Engineer**  
        **BSc Telecommunications Engineering, School of Engineering**  
        **MS Data Science, School of Business and Economics**  
        **University of Management & Technology**  
        **Mobile:** +923018412180  
        **Email:** umersaeed81@hotmail.com  
        **Address:** Dream Gardens, Defence Road, Lahore  
    """)

# Right column: Logo image
with col2:
    st.image("D:/Advance_Data_Sets/PTML_DB/Huawei.jpg", width=100)

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

# Streamlit App Title
st.markdown("<h1 style='color: maroon;'>Extracting Coordinates from Google Maps URLs and Exporting to Excel</h1>", unsafe_allow_html=True)

#-----------------------------------------------------------------------------------------------
# Input File Path
st.markdown("<h3 style='color: maroon;'>Input and Output File Path</h3>", unsafe_allow_html=True)

input_path = st.text_input("**Input Excel File Path 📂**", "D:/Advance_Data_Sets/google_map/input_google_map_urls.xlsx")

# Output File Path
today_date = datetime.today().strftime("%d%m%Y")
base_path = f"D:/Advance_Data_Sets/google_map/google_map_lat_long_{today_date}.xlsx"

# Ensure unique filename
counter = 1
output_path = base_path
while os.path.exists(output_path):
    output_path = f"D:/Advance_Data_Sets/google_map/google_map_lat_long_{today_date}_{counter}.xlsx"
    counter += 1

user_output_path = st.text_input("**Output Excel File Path 📤**", output_path)

#-----------------------------------------------------------------------------------------------
# Function to Extract Coordinates
def extract_coordinates(google_maps_url):
    try:
        response = requests.head(google_maps_url, allow_redirects=True, timeout=10)
        resolved_url = response.url  

        # Match '@lat,lng' format
        match_at = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", resolved_url)
        if match_at:
            return match_at.group(1), match_at.group(2)  

        # Match 'q=lat,lng' format
        match_q = re.search(r"q=(-?\d+\.\d+),(-?\d+\.\d+)", resolved_url)
        if match_q:
            return match_q.group(1), match_q.group(2)  

        return None, None
    except Exception as e:
        return None, None

#-----------------------------------------------------------------------------------------------
# Button to Start Web Scraping
if st.button("Click Here to Start Web Scraping 🚀"):
    if not os.path.isfile(input_path):
        st.error(f"The file does not exist: {input_path}")
    else:
        try:
            df = pd.read_excel(input_path)

            if 'URL' not in df.columns:
                st.error("The 'URL' column is missing in the file!")
            else:
                st.success("File is valid, and 'URL' column found.")
                
                total_urls = len(df)
                st.write(f"Total URLs to Process: **{total_urls}**")

                # Initialize progress bar
                progress_bar = st.progress(0)
                progress_text = st.empty()

                for index, url in tqdm(enumerate(df['URL']), desc="Processing URLs", unit="URL", total=total_urls):
                    lat, lng = extract_coordinates(url)  
                    df.at[index, 'Latitude'] = lat  
                    df.at[index, 'Longitude'] = lng  

                    # Update Progress Bar
                    progress_percentage = int(((index + 1) / total_urls) * 100)
                    progress_bar.progress(progress_percentage)
                    progress_text.text(f"Processing {index + 1} out of {total_urls} ({progress_percentage}%)")

                # Save to Excel
                df.to_excel(user_output_path, index=False)
                st.success(f"Data saved successfully to: {user_output_path}")

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
#-----------------------------------------------------------------------------------------------
