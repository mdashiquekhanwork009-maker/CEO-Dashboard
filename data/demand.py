import pandas as pd

# Load source CSV
source_df = pd.read_csv('C:/Data/demand_data.csv')

# Overwrite target CSV with source data
source_df.to_csv('C:/Users/E36250444/OneDrive - JoulestoWatts Business Solutions Pvt Ltd/Desktop/Streamlit Dashboard/data/demand_data.csv', index=False)

print("Target CSV fully replaced with source data!")