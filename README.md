# JD-Bill
### JD-Bill is a collection of scripts and Jupyter notebooks designed to help the company’s finance department efficiently manage and process customer billing. This repository includes the following files:


### 1. [Monthly Carrier Summary and Customer Data Splitting](https://github.com/Wang-Jialu/JD-Bill/blob/main/merge_split_for_accountants.ipynb)
This Python script processes spreadsheets from various last-mile carriers, each stored in separate folders. It first merges the Excel files within each folder to create a monthly summary for each carrier. For FedEx-O shipments, it calculates the receivable amount using a rate card and maps shipping zip codes to warehouse locations based on a zip code-to-warehouse mapping.

The script also splits the merged summary file by customer names, saving each customer's data into individual folders named after the customers, consolidating their last-mile products into a single file. To improve performance, it uses multithreading for reading and processing multiple files concurrently.

Key functions include listing files in a folder, counting sheets in an Excel file, cleaning text, converting waybill numbers to strings, and performing fee calculations for FedEx-O shipments. The script processes all folders in the current directory by default but can be adapted to handle a specific folder if needed.

### 2. [Enhanced Revenue and Cost Analysis](https://github.com/Wang-Jialu/JD-Bill/blob/main/merge_split_accounts_receivable.ipynb)
This Python script is similar to the previous version but includes key enhancements. It still processes spreadsheets from various last-mile carriers, merging data and calculating fees. However, it now differentiates between self-operated and non-self-operated warehouses by mapping account numbers to warehouse locations for UPS carrier. Additionally, it summarizes revenue and costs for each carrier and generates a consolidated report, providing a detailed analysis by customer and warehouse location.

### 3. [Customer Warehouse Billing Data Formatting](https://github.com/Wang-Jialu/JD-Bill/blob/main/change_format.py)
This script handles customer warehouse billing data, which includes different types of operations such as inbound, storage, and outbound. It unzips files containing this data, formats it according to predefined templates, and generates separate Excel files for each customer. Additionally, it creates a summary file of all charges per customer.

### 4. [Customer-Specific Billing Data Reformatting](https://github.com/Wang-Jialu/JD-Bill/blob/main/specific_customer_format.ipynb)
This code reformats a customer’s last mile billing data according to a specific template. It reads the event data and a template from Excel files, maps the columns from the event data to the template, and creates a new DataFrame with the required format. The reformatted data is then saved to an Excel file.

### 5. [AP and AR Reconciliation and Tracking Comparison](https://github.com/Wang-Jialu/JD-Bill/blob/main/common_ap_ar.ipynb)
This code is designed to help reconcile AP (Accounts Payable) and AR (Accounts Receivable) data by comparing tracking numbers from two folders: one for cost data and one for receivables.

It extracts and formats tracking numbers from Excel and CSV files, then identifies common tracking numbers and those unique to the receivables. The results are summarized in an Excel file, with separate sheets for common and unique data, and categorized based on file content.

### Integration and Efficiency in Billing Processes
These tools collectively enhance the efficiency and accuracy of the billing process, enabling the finance department to handle customer invoices quickly and accurately. Each script and notebook is designed to perform specific tasks that contribute to a streamlined workflow, reducing manual effort and minimizing errors.

### Using Google Colab
Formatting: [Colab Notebook Link](https://colab.research.google.com/drive/19GiDYlWBpogcqy3ZcJeYPl8Y_e_EDpXa#scrollTo=8Yn-M4kk0lr5)

