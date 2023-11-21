import pandas as pd
import os


df = pd.read_csv('Expression.csv', header=None)  # Read the file without a header
df = df.transpose().reset_index(drop=True)
df = df.iloc[1:]
df.columns = ['Gene Expression']
df['Gene Expression'] = df['Gene Expression'].astype(float).apply(lambda x: '{:.4f}'.format(x))
df.to_csv('ExpressionColumn.csv', index=False, header=True, mode='w')  # Specify header=True

print("Process completed: Processing of Expression")

df1 = pd.read_csv('ExpressionColumn.csv')
df2 = pd.read_csv('Columns.csv')
last_rows = df1.tail(377)  
last_column = last_rows.stack().reset_index(drop=True)                  
df2['Gene Expression'] = last_column
df2.to_csv('Organized Data.csv', index=False)  # Replace with the desired output file name and path


print("Process completed: Adding A new expression column to data")

csv_file_path = 'Organized Data.csv'

column_name = 'structure_name'

try:
    df = pd.read_csv(csv_file_path)
    unique_values = df[column_name].unique()

    for value in unique_values:
        folder_name = f"{value}"
        os.makedirs(folder_name, exist_ok=True)
        filtered_df = df[df[column_name] == value]
        file2 = pd.read_csv('DonorInformation.csv')
        ever_tbi_w_loc_value = 'N'
        file2_filtered = file2[file2['ever_tbi_w_loc'] == ever_tbi_w_loc_value]
        merged_data = filtered_df.merge(file2_filtered, on='donor_id', how='inner')
        merged_data = merged_data.sort_values(by=['act_demented', 'sex'])
        merged_data['Gene Expression'] = merged_data['Gene Expression'].astype(float).apply(lambda x: '{:.4f}'.format(x))
        column_to_remove = 'ColumnToRemove'
        if column_to_remove in merged_data:
            merged_data = merged_data.drop(columns=[column_to_remove])
         # Save the modified data to a new Excel file in the folder
        modified_filename = os.path.join(folder_name, f'{value}.xlsx')
        with pd.ExcelWriter(modified_filename, engine='xlsxwriter') as writer:
            merged_data.to_excel(writer, sheet_name='Organized Data', index=False)

            # Create new sheets for dementia (male and females) and non-dementia (male and females)
            dementia_male = merged_data[(merged_data['act_demented'] == 'Dementia') & (merged_data['sex'] == 'M')]
            dementia_female = merged_data[(merged_data['act_demented'] == 'Dementia') & (merged_data['sex'] == 'F')]
            non_dementia_male = merged_data[(merged_data['act_demented'] == 'No Dementia') & (merged_data['sex'] == 'M')]
            non_dementia_female = merged_data[(merged_data['act_demented'] == 'No Dementia') & (merged_data['sex'] == 'F')]

            dementia_male.to_excel(writer, sheet_name='Dementia_Male', index=False)
            dementia_female.to_excel(writer, sheet_name='Dementia_Female', index=False)
            non_dementia_male.to_excel(writer, sheet_name='Non-Dementia_Male', index=False)
            non_dementia_female.to_excel(writer, sheet_name='Non-Dementia_Female', index=False)

except FileNotFoundError:
    print(f"File not found: '{csv_file_path}'")
except Exception as e:
    print(f"An error occurred: {str(e)}")

os.remove('ExpressionColumn.csv')
os.remove('Organized Data.csv')
