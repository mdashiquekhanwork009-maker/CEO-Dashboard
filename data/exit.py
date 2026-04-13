import pandas as pd

try:
    source_path = 'C:/Data/exit.csv'
    target_path = 'C:/Users/E36250444/OneDrive - JoulestoWatts Business Solutions Pvt Ltd/Desktop/Streamlit Dashboard/data/exit.csv'

    source_df = pd.read_csv(source_path)
    source_df.to_csv(target_path, index=False)

except Exception as e:
    with open("error_log.txt", "a") as f:
        f.write(str(e) + "\n")