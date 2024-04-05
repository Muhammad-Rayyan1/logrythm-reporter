import os
import pandas as pd
import matplotlib.pyplot as plt
import argparse

def load_data(file_path):
    df = pd.read_csv(file_path, index_col=False)
    if 'Alarm Date' in df.columns: # remove alarm date cuz ##### will mess things up
        df = df.drop(columns=['Alarm Date'])
    return df

def map_statuses(df):
    status_mapping = {
        'New': 'New',
        'Escalated': 'Open',
        'Working': 'Open',
        'Open': 'Open',
        'Closed: False Alarm': 'Closed',
        'Closed: Reported': 'Closed',
        'Closed: Resolved': 'Closed',
        'Closed: Unresolved': 'Closed',
        'Closed: Monitor': 'Closed'
    }
    df['Alarm Status'] = df['Alarm Status'].map(status_mapping)
    return df

def create_folders_and_process_data(df):
    log_source_entities = df['Log Source Entity'].unique()
    for entity in log_source_entities:
        entity_str = str(entity)
        folder_path = os.path.join(os.getcwd(), entity_str)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"Folder '{entity_str}' created.")
        else:
            print(f"Folder '{entity_str}' already exists.")
        
        filtered_df = df[df['Log Source Entity'] == entity]

        if not filtered_df.empty:
            csv_file_path = os.path.join(folder_path, f"{entity_str}_filtered.xlsx")

            writer = pd.ExcelWriter(csv_file_path, engine='xlsxwriter')
            filtered_df.to_excel(writer, sheet_name='Data', index=False)

            pivot_df = filtered_df.groupby('Alarm Status').size().reset_index(name='Count')
            pivot_df.to_excel(writer, sheet_name='Graph', index=False)

            writer._save()
            print(f"Data and Graph table added to '{entity_str}_filtered.xlsx'.")

            plt.figure(figsize=(10, 6))
            bars = plt.bar(pivot_df['Alarm Status'], pivot_df['Count'])
            plt.xlabel('Alarm Status')
            plt.ylabel('Count')
            plt.title('Alarm Status Distribution')
            plt.xticks(rotation=45, ha='right')
            
            for bar in bars:
                yval = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2, yval + 0.5, int(yval), ha='center', va='bottom')
            
            plt.tight_layout()
            plt.savefig(os.path.join(folder_path, 'alarm_status.png'))
            plt.close()

            print(f"Bar chart saved as 'alarm_status.png' in '{entity_str}' folder.")

            with pd.ExcelWriter(csv_file_path, mode='a', engine='openpyxl') as writer:
                pivot_df_with_alarm_rule = filtered_df.pivot_table(index='Alarm Rule Name', columns='Alarm Status', aggfunc='size', fill_value=0)
                pivot_df_with_alarm_rule.to_excel(writer, sheet_name='total')
                print(f"'total' sheet added to '{entity_str}_filtered.xlsx'.")

        else:
            print(f"No data found for '{entity_str}'.")

def main():
    parser = argparse.ArgumentParser(description='Process CSV data and generate analysis reports.')
    parser.add_argument('file', type=str, help='Path to the CSV file')

    args = parser.parse_args()
    file_path = args.file

    df = load_data(file_path)
    df = map_statuses(df)
    create_folders_and_process_data(df)

if __name__ == "__main__":
    main()
