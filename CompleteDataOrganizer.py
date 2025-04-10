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

# Dictionary to store data for each region
region_data = {}

try:
    df = pd.read_csv(csv_file_path)
    unique_values = df[column_name].unique()
    
    # Create a new DataFrame for the region comparison
    region_comparison = pd.DataFrame()

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
            
            Alzheimer = merged_data[(merged_data['dsm_iv_clinical_diagnosis'].str.contains('Alzheimer'))]
            other_diagnosis = merged_data[(merged_data['dsm_iv_clinical_diagnosis'] != 'No Dementia') & (~merged_data['dsm_iv_clinical_diagnosis'].str.contains('Alzheimer'))]
            no_dementia = merged_data[(merged_data['dsm_iv_clinical_diagnosis'] == 'No Dementia')]
            
            no_dementia_braak_more_than_3 = merged_data[(merged_data['dsm_iv_clinical_diagnosis'] == 'No Dementia') & (merged_data['braak'] > 3)]
            no_dementia_braak_less_than_3 = merged_data[(merged_data['dsm_iv_clinical_diagnosis'] == 'No Dementia') & (merged_data['braak'] <= 3)]

            dementia_braak_less_than_3 = merged_data[(merged_data['dsm_iv_clinical_diagnosis'] != 'No Dementia') & (merged_data['braak'] <= 3)]
            dementia_braak_more_than_3 = merged_data[(merged_data['dsm_iv_clinical_diagnosis'] != 'No Dementia') & (merged_data['braak'] > 3)]
            
            # Filter for braak values less than or equal to 3
            braak_3_or_less = merged_data[merged_data['braak'] <= 3]

            # Filter for braak values greater than 3
            braak_more_than_3 = merged_data[merged_data['braak'] > 3]


            dementia_male.to_excel(writer, sheet_name='Dementia_Male', index=False)
            dementia_female.to_excel(writer, sheet_name='Dementia_Female', index=False)
            non_dementia_male.to_excel(writer, sheet_name='Non-Dementia_Male', index=False)
            non_dementia_female.to_excel(writer, sheet_name='Non-Dementia_Female', index=False)
            
            Alzheimer.to_excel(writer, sheet_name='Alzheimer Disease', index=False)
            other_diagnosis.to_excel(writer, sheet_name='Other Diagnosis', index=False)
            no_dementia.to_excel(writer, sheet_name='No Dementia', index=False)
            
            no_dementia_braak_more_than_3.to_excel(writer, sheet_name='No Dementia braak more than 3', index=False)
            no_dementia_braak_less_than_3.to_excel(writer, sheet_name='No Dementia braak less than 3', index=False)
            
            dementia_braak_more_than_3.to_excel(writer, sheet_name='Dementia braak more than 3', index=False)
            dementia_braak_less_than_3.to_excel(writer, sheet_name='Dementia braak less than 3', index=False)
        
        # Store the gene expression data for this region
        dementia_data = merged_data[merged_data['act_demented'] == 'Dementia']['Gene Expression'].astype(float).tolist()
        non_dementia_data = merged_data[merged_data['act_demented'] == 'No Dementia']['Gene Expression'].astype(float).tolist()
        
        # Store in the dictionary
        region_data[value] = {
            'Dementia': dementia_data,
            'No Dementia': non_dementia_data
        }
                                                                                                                            
    # Create a region comparison Excel file with horizontal layout
    # Each region is a column, with dementia status as rows
    # Select the first 4 regions (or all if less than 4)
    selected_regions = list(unique_values)[:4] if len(unique_values) >= 4 else list(unique_values)
    
    # Define region colors (one for each region)
    region_colors = ['#E6F2FF', '#FFE6E6', '#E6FFE6', '#FFE6FF']
    if len(selected_regions) > len(region_colors):
        # Add more colors if needed
        region_colors.extend(['#FFFFCC', '#CCFFFF', '#FFCCCC', '#CCFFCC'])
    
    # Create the Excel file
    with pd.ExcelWriter('Region_Comparison.xlsx', engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Region Comparison')
        
        # Define cell formats
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#D9D9D9',
            'border': 1
        })
        
        data_format = workbook.add_format({
            'align': 'center',
            'border': 1
        })
        
        no_dementia_header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#CCFFCC',
            'border': 1
        })
        
        dementia_header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#FFCCCC',
            'border': 1
        })
        
        # Write the headers
        worksheet.write(0, 0, 'Region', header_format)
        worksheet.write(1, 0, 'No Dementia', no_dementia_header_format)  # Switched order
        worksheet.write(2, 0, 'Dementia', dementia_header_format)        # Switched order
        
        # Set column width for first column
        worksheet.set_column(0, 0, 15)
        
        # Number of data points per region
        data_points = 29
        
        # For each region, create a column section
        current_col = 1  # Start from column B (index 1)
        
        for region_idx, region in enumerate(selected_regions):
            # Get the color for this region
            region_color = region_colors[region_idx % len(region_colors)]
            
            # Create a format for this region's header
            region_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': region_color,
                'border': 1
            })
            
            # Create a format for this region's data cells (same color but not bold)
            region_data_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': region_color,
                'border': 1
            })
            
            # Calculate the start and end columns for this region
            start_col = current_col
            end_col = start_col + data_points - 1
            
            # Write region name and merge cells for region header
            worksheet.merge_range(0, start_col, 0, end_col, region, region_format)
            
            # Set column width
            worksheet.set_column(start_col, end_col, 12)
            
            # Get dementia and non-dementia values
            dementia_values = region_data[region]['Dementia']
            non_dementia_values = region_data[region]['No Dementia']
            
            # Write non-dementia values in row 1 (switched order)
            for data_idx in range(data_points):
                value = non_dementia_values[data_idx] if data_idx < len(non_dementia_values) else (non_dementia_values[-1] if non_dementia_values else 0.0)
                worksheet.write(1, start_col + data_idx, value, region_data_format)  # Using region-specific formatting
            
            # Write dementia values in row 2 (switched order)
            for data_idx in range(data_points):
                value = dementia_values[data_idx] if data_idx < len(dementia_values) else (dementia_values[-1] if dementia_values else 0.0)
                worksheet.write(2, start_col + data_idx, value, region_data_format)  # Using region-specific formatting
            
            # Update the current column for the next region
            current_col = end_col + 1
        
        # Create an additional sheet with vertical layout
        vert_worksheet = workbook.add_worksheet('Vertical Layout')
        
        # Define cell formats
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#D9D9D9',
            'border': 1
        })
        
        # Define diagnosis type formats
        no_dementia_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#CCFFCC',
            'border': 1
        })
        
        alzheimer_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#FFCCCC',
            'border': 1
        })
        
        other_diagnosis_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#FFFFCC',
            'border': 1
        })
        
        # Set column width for data columns
        column_width = 15
        
        # Calculate how many columns we need (number of regions * 3)
        total_cols = len(selected_regions) * 3
        
        # Set up all column widths at once
        for col in range(total_cols):
            vert_worksheet.set_column(col, col, column_width)
        
        # Current column tracker
        current_col = 0
        
        # For each region
        for region_idx, region in enumerate(selected_regions):
            # Get the color for this region
            region_color = region_colors[region_idx % len(region_colors)]
            
            # Create a format for this region's header
            region_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': region_color,
                'border': 1
            })
            
            # Create data formats with matching region color
            region_data_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': region_color,
                'border': 1
            })
            
            # Get the starting column for this region
            start_col = current_col
            end_col = start_col + 2  # 3 columns per region (one for each diagnosis type)
            
            # Write region name in the first row and merge 3 cells horizontally
            vert_worksheet.merge_range(0, start_col, 0, end_col, region, region_format)
            
            # Write diagnosis types in the second row
            vert_worksheet.write(1, start_col, 'No Dementia', no_dementia_format)
            vert_worksheet.write(1, start_col + 1, 'Alzheimer Disease', alzheimer_format)
            vert_worksheet.write(1, start_col + 2, 'Other Diagnosis', other_diagnosis_format)
            
            # Get the data for this region
            filtered_df = df[df[column_name] == region]
            file2 = pd.read_csv('DonorInformation.csv')
            ever_tbi_w_loc_value = 'N'
            file2_filtered = file2[file2['ever_tbi_w_loc'] == ever_tbi_w_loc_value]
            merged_data_for_region = filtered_df.merge(file2_filtered, on='donor_id', how='inner')
            merged_data_for_region = merged_data_for_region.sort_values(by=['act_demented', 'sex'])
            merged_data_for_region['Gene Expression'] = merged_data_for_region['Gene Expression'].astype(float).apply(lambda x: '{:.4f}'.format(x))
            
            # Get expression values for different diagnosis types
            no_dementia_values = merged_data_for_region[merged_data_for_region['dsm_iv_clinical_diagnosis'] == 'No Dementia']['Gene Expression'].tolist()
            alzheimer_values = merged_data_for_region[merged_data_for_region['dsm_iv_clinical_diagnosis'].str.contains('Alzheimer')]['Gene Expression'].tolist()
            other_diagnosis_values = merged_data_for_region[(merged_data_for_region['dsm_iv_clinical_diagnosis'] != 'No Dementia') & 
                                                           (~merged_data_for_region['dsm_iv_clinical_diagnosis'].str.contains('Alzheimer'))]['Gene Expression'].tolist()
            
            # Write all values for No Dementia
            for i, value in enumerate(no_dementia_values):
                if i < 30:  # Limit to 30 rows
                    vert_worksheet.write(i + 2, start_col, value, region_data_format)
            
            # Write all values for Alzheimer Disease
            for i, value in enumerate(alzheimer_values):
                if i < 30:  # Limit to 30 rows
                    vert_worksheet.write(i + 2, start_col + 1, value, region_data_format)
            
            # Write all values for Other Diagnosis
            for i, value in enumerate(other_diagnosis_values):
                if i < 30:  # Limit to 30 rows
                    vert_worksheet.write(i + 2, start_col + 2, value, region_data_format)
            
            # Update current column for next region
            current_col = end_col + 1
            
            # Update the current column for the next region
            current_col = end_col + 1
            
    print("Region comparison Excel file created successfully!")
    
    print("Region comparison Excel file created successfully!")

except FileNotFoundError:
    print(f"File not found: '{csv_file_path}'")
except Exception as e:
    print(f"An error occurred: {str(e)}")

os.remove('ExpressionColumn.csv')
os.remove('Organized Data.csv')
