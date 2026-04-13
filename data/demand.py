import pandas as pd

# Load the CSV files
source_df = pd.read_csv('C:/Data/demand_data.csv')
target_df = pd.read_csv('C:/Users/E36250444/OneDrive - JoulestoWatts Business Solutions Pvt Ltd/Desktop/Streamlit Dashboard/data/demand_data.csv')

# Set the key column (common column)
key_column = 'ID'

# Set index for easy update
source_df.set_index(key_column, inplace=True)
target_df.set_index(key_column, inplace=True)

# Update target with source values
target_df.update(source_df)

# Reset index back
target_df.reset_index(inplace=True)

# Save updated file
target_df.to_csv('updated_target.csv', index=False)

print("CSV updated successfully!")