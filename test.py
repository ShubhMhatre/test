import streamlit as st
import pandas as pd
import csv
import json

st.title("Excel File Processor")

secondary = st.file_uploader("Upload Secondary Sales File", type="xlsx")
pcc = st.file_uploader("Upload PCC Growth Bonus File", type="xlsx")

def csv_to_json(csv_file_path, json_file_path):
    # Read CSV file and convert it to a list of dictionaries
    with open(csv_file_path, 'r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        data = list(csv_reader)
     
    # Write the data to a JSON file
    with open(json_file_path, 'w') as json_file:
        json.dump(data, json_file, indent=4)





if secondary and pcc:
    
    file_path = 'Secondary Sales.xlsx'
    df = pd.read_excel(secondary)
    
    # Filter for Fevicol Products and the relevant fiscal years (2024 and 2025)
    fevicol_products = ['FEVICOL HEATX', 'FEVICOL SH', 'FEVICOL MARINE', 'FEVICOL PROBOND', 
                        'FEVICOL HI-PER', 'FEVICOL HI-PER STAR', 'FEVICOL SR 998', 'FEVICOL SPEEDX']
    
    df_filtered = df[df['BI_Product'].isin(fevicol_products) & df['Fiscal Year'].isin([2024, 2025])]
    
    # Pivot the data to create columns for each month's sales volumes
    df_pivot = pd.pivot_table(df_filtered, 
                              index=['Group Code', 'Fiscal Year'], 
                              columns='Month', 
                              values='Sec Sales Vol Kgs', 
                              aggfunc='sum', 
                              fill_value=0).reset_index()
    
    # Separate data for 2025 and 2024 for comparison
    df_2025 = df_pivot[df_pivot['Fiscal Year'] == 2025].copy()
    df_2024 = df_pivot[df_pivot['Fiscal Year'] == 2024].copy()
    
    # Merge the 2025 and 2024 data to compare Q2 volumes (Jul, Aug, Sep)
    merged = pd.merge(df_2025, df_2024, on='Group Code', how='left', suffixes=('_2025', '_2024'))
    
    # Calculate the Q2 total volumes for 2025 and 2024
    merged['Total_Volume_Q3_2025'] = merged[['Oct_2025', 'Nov_2025', 'Dec_2025']].sum(axis=1)
    merged['Total_Volume_Q3_2024'] = merged[['Oct_2024', 'Nov_2024', 'Dec_2024']].sum(axis=1).fillna(0)
    
    # Function to determine the slab, points per kg, next slab, etc.
    def calculate_slab_and_points(volume):
        if volume >= 5500:
            slab = '5500+'
            points_per_kg = 6
            next_slab = 5500
            next_slab_points_per_kg = 6
        elif volume >= 3300:
            slab = '3300+'
            points_per_kg = 5
            next_slab = 5500
            next_slab_points_per_kg = 6
        elif volume >= 2000:
            slab = '2000+'
            points_per_kg = 4
            next_slab = 3300
            next_slab_points_per_kg = 5
        elif volume >= 1000:
            slab = '1000+'
            points_per_kg = 3
            next_slab = 2000
            next_slab_points_per_kg = 4
        elif volume >= 450:
            slab = '450+'
            points_per_kg = 2
            next_slab = 1000
            next_slab_points_per_kg = 3
        else:
            slab = 'None'
            points_per_kg = 0
            next_slab = 450
            next_slab_points_per_kg = 2
    
        balance_to_next_slab = max(next_slab - volume, 0)
        
        return slab, points_per_kg, next_slab, next_slab_points_per_kg, balance_to_next_slab
    
    # Apply the function to calculate slab, points, and next slab details
    slab_results = merged['Total_Volume_Q3_2025'].apply(calculate_slab_and_points)
    merged[['Current_Slab', 'Current_Points_Per_Kg', 'Next_Slab', 'Next_Slab_Points_Per_Kg', 'Balance_to_Next_Slab']] = pd.DataFrame(slab_results.tolist(), index=merged.index)
    
    # Calculate current and next slab Khazana points
    merged['Current_Khazana_Points'] = merged['Total_Volume_Q3_2025'] * merged['Current_Points_Per_Kg']
    merged['Next_Slab_Khazana_Points'] = merged['Next_Slab'] * merged['Next_Slab_Points_Per_Kg']
    
    # Growth requirements based on last year's volume
    def calculate_growth_required(volume_2024):
        if volume_2024 >= 5500:
            return 0
        elif volume_2024 >= 3300:
            return 0.05
        elif volume_2024 >= 2000:
            return 0.08
        elif volume_2024 >= 1000:
            return 0.15
        elif volume_2024 >= 450:
            return 0.20
        else:
            return 0.20
    
    # Apply the growth requirement
    merged['Growth_Required'] = merged['Total_Volume_Q3_2024'].apply(calculate_growth_required)
    
    # Calculate target for growth bonus and balance for growth bonus
    merged['Target_for_Growth_Bonus'] = merged['Total_Volume_Q3_2024'] * (1 + merged['Growth_Required'])
    merged['Balance_for_Growth_Bonus'] = merged['Target_for_Growth_Bonus'] - merged['Total_Volume_Q3_2025']
    
    # Calculate growth Khazana points
    merged['Growth_Khazana_Points'] = merged['Target_for_Growth_Bonus'] * 2
    
    # Select and rename the final columns for output
    output_columns = {
        'Group Code': 'Group Code',
        'Oct_2025': 'Fevicol Oct Volume',
        'Nov_2025': 'Fevicol Nov Volume',
        'Dec_2025': 'Fevicol Dec Volume',
        'Total_Volume_Q3_2025': 'Fevicol Total Volume Q3 TY',
        'Total_Volume_Q3_2024': 'Fevicol Total Volume Q3 LY',
        'Current_Slab': 'Fevicol Current Slab',
        'Current_Points_Per_Kg': 'Fevicol Current Points per kg',
        'Next_Slab': 'Fevicol Next Slab',
        'Next_Slab_Points_Per_Kg': 'Fevicol Next Slab Points per kg',
        'Balance_to_Next_Slab': 'Fevicol Balance to Next Slab',
        'Current_Khazana_Points': 'Fevicol Current Khazana Points',
        'Next_Slab_Khazana_Points': 'Fevicol Next Slab Khazana Points',
        'Growth_Required': 'Fevicol Growth Required',
        'Target_for_Growth_Bonus': 'Target for growth bonus',
        'Balance_for_Growth_Bonus': 'Balance for growth bonus',
        'Growth_Khazana_Points': 'Growth Khazana Points'
    }
    
    final_output = merged[list(output_columns.keys())].rename(columns=output_columns)
    
    # Save the result to an Excel file
    # final_output.to_excel('fevicol.xlsx', index=False)
    fevicol_df = final_output
    
    
    
    ############################################ Premiumization Details
    # Load the data from the Excel file
    file_path = 'Secondary Sales.xlsx'
    df = pd.read_excel(secondary)
    
    # Define the products and months of interest
    products_of_interest = ['FEVICOL HI-PER', 'FEVICOL HI-PER STAR']
    months_of_interest = ["Oct", "Nov"]
    fiscal_year_of_interest = 2025
    
    # Filter the data for the products and fiscal year of interest
    filtered_df = df[
        (df["BI_Product"].isin(products_of_interest)) & 
        (df["Fiscal Year"] == fiscal_year_of_interest) & 
        (df["Month"].isin(months_of_interest))
    ]
    
    # Pivot table to get monthly sales volumes for each dealer
    pivot_df = filtered_df.pivot_table(
        # index=["Group Code", "Dealer Name"],
        index=["Group Code"],
        columns="Month",
        values="Sec Sales Vol Kgs",
        aggfunc="sum",
        fill_value=0
    ).reset_index()
    
    # Rename columns for clarity
    # pivot_df.columns = ["Group Code", "Dealer Name", "Hiper Jul", "Hiper Sep", "Hiper Aug"]
    pivot_df.columns = ["Group Code", "Hiper Oct", "Hiper Nov"]
    
    # Calculate the total volume
    pivot_df["Hiper Total"] = pivot_df["Hiper Oct"] + pivot_df["Hiper Nov"]
    # Define the slab conditions
    def get_slab_pm(volume):
        if volume >= 300:
            return "300+", 2, 300, 2
        else:
            return "None", 0, 300, 2
    
    # Apply the slab conditions
    pivot_df[["Hiper Current Slab", "Hiper Khazana Points per KG", "Hiper Next Slab", "Hiper Next Slab Khazana Points"]] = pivot_df.apply(
        lambda row: pd.Series(get_slab_pm(row["Hiper Total"])), axis=1
    )
    
    # Calculate the balance for Sep
    pivot_df["Hiper Balance for Q3"] = pivot_df["Hiper Next Slab"] - pivot_df["Hiper Total"]
    pivot_df["Hiper Balance for Q3"] = pivot_df["Hiper Balance for Q3"].apply(lambda x: max(x, 0))
    
    # Calculate Khaza points
    pivot_df["Hiper Current Khaza Points"] = pivot_df["Hiper Total"] * pivot_df["Hiper Khazana Points per KG"]
    pivot_df["Hiper Next Slab Khazana Points"] = pivot_df["Hiper Next Slab"] * pivot_df["Hiper Next Slab Khazana Points"]
    
    # Write the result to a new Excel file
    output_file_path = "hiper.xlsx"
    # pivot_df.to_excel(output_file_path, index=False)
    hiper_df = pivot_df
    
    print(f"Processed data has been saved to {output_file_path}")
    
    
    
    
    ############################################ Masterlok Details
    # Load the data from the Excel file
    file_path = 'Secondary Sales.xlsx'
    df = pd.read_excel(secondary)
    
    # Define the products and months of interest
    products_of_interest = ["MASTERLOK", "MASTERLOK XTRA", "BULBOND", "BULBOND XTRA"]
    months_of_interest = ["Oct", "Nov"]
    fiscal_year_of_interest = 2025
    
    # Filter the data for the products and fiscal year of interest
    filtered_df = df[
        (df["BI_Product"].isin(products_of_interest)) & 
        (df["Fiscal Year"] == fiscal_year_of_interest) & 
        (df["Month"].isin(months_of_interest))
    ]
    
    # Pivot table to get monthly sales volumes for each dealer
    pivot_df = filtered_df.pivot_table(
        # index=["Group Code", "Dealer Name"],
        index=["Group Code"],
        columns="Month",
        values="Sec Sales Vol Kgs",
        aggfunc="sum",
        fill_value=0
    ).reset_index()
    
    # Rename columns for clarity
    pivot_df.columns = ["Group Code", "Masterlok Oct", "Masterlok Nov"]
    
    # Calculate the total volume
    pivot_df["Masterlok Total"] = pivot_df["Masterlok Oct"] + pivot_df["Masterlok Nov"]
    
    # Define the slab conditions
    def get_slab_ml(volume):
        if volume >= 1000:
            return "1000+", 12, 1000, 12
        elif volume >= 500:
            return "500+", 8, 1000, 12
        elif volume >= 300:
            return "300+", 6, 500, 8
        elif volume >= 100:
            return "100+", 4, 300, 6
        else:
            return "None", 0, 100, 4
    
    # Apply the slab conditions
    pivot_df[["Masterlok Current Slab", "Masterlok Khazana Points per KG", "Masterlok Next Slab", "Masterlok Next Slab Khazana Points"]] = pivot_df.apply(
        lambda row: pd.Series(get_slab_ml(row["Masterlok Total"])), axis=1
    )
    
    # Calculate the balance for Sep
    pivot_df["Masterlok Balance for Q3"] = pivot_df["Masterlok Next Slab"] - pivot_df["Masterlok Total"]
    pivot_df["Masterlok Balance for Q3"] = pivot_df["Masterlok Balance for Q3"].apply(lambda x: max(x, 0))
    
    # Calculate Khaza points
    pivot_df["Masterlok Current Khaza Points"] = pivot_df["Masterlok Total"] * pivot_df["Masterlok Khazana Points per KG"]
    pivot_df["Masterlok Next Slab Khazana Points"] = pivot_df["Masterlok Next Slab"] * pivot_df["Masterlok Next Slab Khazana Points"]
    
    # Write the result to a new Excel file
    output_file_path = "Masterlok.xlsx"
    # pivot_df.to_excel(output_file_path, index=False)
    masterlok_df = pivot_df
    print(f"Processed data has been saved to {output_file_path}")
    
    ####################################### PCC details
    # Load the data from the Excel file
    file_path = 'PCC Data.xlsx'
    df = pd.read_excel(pcc)
    
    # Convert the columns to numeric, coercing errors to NaN
    df["Bal to Growth Bonus Tgt1"] = pd.to_numeric(df["Bal to Growth Bonus Tgt1"], errors='coerce')
    df["Bal to Growth Bonus Tgt2"] = pd.to_numeric(df["Bal to Growth Bonus Tgt2"], errors='coerce')
    
    
    # Add the new columns
    df["Expected Slab 1 Vol balance"] = df["Bal to Growth Bonus Tgt1"] / 4
    df["Expected Slab 2 Vol balance"] = df["Bal to Growth Bonus Tgt2"] / 4
    
    # Save the updated dataframe to a new Excel file
    output_file_path = "Updated_pccdata.xlsx"
    # df.to_excel(output_file_path, index=False)
    updated_pccdata_df = df
    
    print(f"Updated data has been saved to {output_file_path}")
    
    ###########################################################3#Merge all files and make json
    # Load the data from the Excel files
    # fevicol_df = pd.read_excel("fevicol.xlsx")
    # hiper_df = pd.read_excel("hiper.xlsx")
    # msg_df = pd.read_excel("msg.xlsx")
    # masterlok_df = pd.read_excel("Masterlok.xlsx")
    # updated_pccdata_df = pd.read_excel("Updated_pccdata.xlsx")
    
    # Merge the dataframes on "Group Code" using outer join
    merged_df = fevicol_df.merge(hiper_df, on=["Group Code"], how="outer") \
                          .merge(masterlok_df, on=["Group Code"], how="outer") \
                          .merge(updated_pccdata_df, on=["Group Code"], how="left")
    # Ensure that we have only one "Dealer Name" and "Group Code" in the final dataframe
    # Drop duplicate "Dealer Name" columns
    # dealer_name_cols = [col for col in merged_df.columns if col.startswith('Dealer Name')]
    # merged_df['Dealer Name'] = merged_df[dealer_name_cols].bfill(axis=1).iloc[:, 0]
    # merged_df = merged_df.drop(columns=dealer_name_cols[1:])
    
    
    secondary_sales_df = pd.read_excel(secondary)
    grouped_sales_df = secondary_sales_df.groupby('Group Code')['Dealer Name'].apply(lambda x: ', '.join(x.unique())).reset_index()
    merged_df = pd.merge(merged_df, grouped_sales_df, on='Group Code', how='left')
    
    
    merged_df.to_excel("merged.xlsx", index=False)
    # Convert the merged DataFrame to JSON
    
    merged_df = pd.read_excel("merged.xlsx")
    merged_json = merged_df.to_json(orient="records")
    
    # Save the JSON to a file
    with open("merged_data.json", "w") as json_file:
        json_file.write(merged_json)

    # Read the files into dataframes
    # df1 = pd.read_excel(uploaded_file1)
    # df2 = pd.read_excel(uploaded_file2)

    # # Perform processing (replace this with your own logic)
    # result = df1.merge(df2, on='YourCommonColumn')  # Example operation

    # # Output to Excel file
    # result.to_excel("output.xlsx", index=False)

    # Download link for output file
    st.download_button(label="Download Processed File", data=merged_df.to_excel(index=False, engine='openpyxl'), file_name="merged.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
