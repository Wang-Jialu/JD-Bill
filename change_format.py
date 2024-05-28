import os
import pandas as pd
import zipfile
import os

def unzip_files(source_dir, dest_dir):
    # Ensure the destination directory exists
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)

    # Loop through all files in the source directory
    for item in os.listdir(source_dir):
        if item.endswith('.zip'):
            file_path = os.path.join(source_dir, item)
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(dest_dir)
            print(f'Unzipped {item} into {dest_dir}')

# Define source and destination directories
source_directory = r"path\to\source\folder"
folder_path = r"path\to\raw_data\folder"

# Call the function to unzip files
unzip_files(source_directory, folder_path)


bill_template_file = r"path\to\bill_template.xlsx"

bill_template = pd.read_excel(bill_template_file, sheet_name=None, header=None)

# Create a dictionary mapping column names from the second row of template (key) 
# to their corresponding column names from the first row (value)
outbound_template = dict(zip(bill_template['Outbound'].iloc[1], bill_template['Outbound'].iloc[0]))
storage_template = dict(zip(bill_template['Storage'].iloc[1], bill_template['Storage'].iloc[0]))
inbound_template = dict(zip(bill_template['Inbound'].iloc[1], bill_template['Inbound'].iloc[0]))


def process_form(df,template):
    """
    Iterate over the items in the template dictionary, 
    checking if each column specified in the template exists in the input DataFrame. 
    If a column exists, add it to the output DataFrame with the mapped column name.
    """

    output = pd.DataFrame()
    for column_name, mapped_column_name in template.items():
        if column_name in df.columns:
            output[mapped_column_name] =df[column_name]
    return output



file_paths = [os.path.join(folder_path, file) for file in os.listdir(folder_path)]

summary_df = pd.DataFrame(columns=['客户', '计费项目', '仓库编号', '结算币种含税金额'])


for file_path in file_paths:

    event_data = pd.read_excel(file_path, sheet_name=None)
    # Identify the customer associated with that file, as each file represents data for a single customer
    first_customer_name = event_data[next(iter(event_data))]['付款对象'].iloc[0]
    output_file = f"{first_customer_name}.xlsx"
    output_data = {}


    for sheet_name, df in event_data.items():
        # In each sheet, '计费项目' remains consistent.'
        first_category = df['计费项目'].iloc[0]
        # Ensure uniformity in column names
        df = df.rename(columns={'仓库编码': '仓库编号'})
        df['结算币种含税金额 '] = df['结算币种含税金额'].astype(float)
        grouped_data = df.groupby('仓库编号')['结算币种含税金额'].sum().reset_index()

        grouped_data['客户'] = first_customer_name
        grouped_data['计费项目'] = first_category
        summary_df = pd.concat([summary_df, grouped_data], ignore_index=True)
    

        if '出库' in sheet_name:
            output_data[sheet_name] = process_form(df,outbound_template)
        elif '在库' in sheet_name:
            output_data[sheet_name] = process_form(df,storage_template)
        elif '入库'in sheet_name:
            output_data[sheet_name] = process_form(df,inbound_template)
        else:
            # If none of the above conditions are met, the original DataFrame df is retained.
            print(f"处理文件失败：{file_path}，工作表名称：{sheet_name}")
            output_data[sheet_name] = df

    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, output_df in output_data.items():
            output_df.to_excel(writer, sheet_name=sheet_name, index=False)


summary_df.to_excel('summary.xlsx', index=False)