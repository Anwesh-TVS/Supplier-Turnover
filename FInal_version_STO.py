def supplier_turnover():

    try:  
        import pandas as pd
        import numpy as np
        from datetime import datetime
        import os
        import warnings
        from hdbcli import dbapi 
        warnings.filterwarnings("ignore")

        df = pd.read_excel("./Turnover OCT'23.XLSX")
        # Filter by plant 

        plant = ['HOS','MYSR','HP','H001','H002','SPW1','SPWH','JMWH','HEX1']
        df = df[df['PLANT'].isin(plant)]
        df['VENDOR'] = df['VENDOR'].dropna(axis = 0)
        #print(len(df))
        #print("Initial dataframe ", df["PLANT"].unique())
        # Filter by Vendor
        df['VENDOR'] = df['VENDOR'].str[-5:]
        df.dropna(inplace = True)
        #print("After dropping null values ", df["PLANT"].unique())

        df['VENDOR'].astype(str)

        df["MATERIAL"].astype(str)
        df["PLANT"].astype(str)

        # Filter rows where 'Last_5_Digits' number starts with '2'

        filtered_df = df[df['VENDOR'].str.startswith('2')]
        #print(len(filtered_df))


        #print("After vendor startswith 2 ", filtered_df["PLANT"].unique())
        ######Part Series with NE,NB,N7 in 5 Series PO to be excluded.

        # Filter by PO-NUMBER
        exclude_part_series = ['NE', 'NB', 'N7']

        # Cast 'PO-NUMBER' column to a string
        #filtered_df['PO-NUMBER'] = filtered_df['PO-NUMBER'].astype(str)

        # Create conditions for filtering
        '''condition1 = ~filtered_df['PO-NUMBER'].astype(str).str.startswith('5')
        condition2 = ~filtered_df['MATERIAL'].str.startswith(tuple(exclude_part_series))
        filtered_df = filtered_df[condition1 & condition2]
        print(len(filtered_df))'''

        #filtered_df = filtered_df.loc[~filtered_df['PO-NUMBER'].astype(str).str.startswith('5') & ~filtered_df['MATERIAL'].str.startswith(tuple(exclude_part_series))]

        filtered_df['PLANT'] = filtered_df['PLANT'].str.strip()

        exclude_plant = ["HOS","SPWH","SPW1"]

        #filtered_df = filtered_df.loc[~((filtered_df['PO-NUMBER'].astype(str).str.startswith('5')) & (filtered_df['MATERIAL'].str.startswith(tuple(exclude_part_series))) & (filtered_df['PLANT'].isin(exclude_plant)))]

        filtered_df = filtered_df.loc[~((filtered_df['MATERIAL'].str.startswith(tuple(exclude_part_series))) & (filtered_df['PLANT'].isin(exclude_plant)))]



        #print(len(filtered_df))
        #print("After PO and material filter",filtered_df["PLANT"].unique())

        #Filter by Raw  Material
        excl_vendor = ["20330" , "45755" , "20480", "45585","21695"]
        # Remove leading and trailing whitespaces
        filtered_df['VENDOR'] = filtered_df['VENDOR'].str.strip()

        filtered_df = filtered_df.loc[~filtered_df['VENDOR'].isin(excl_vendor)]
        #print(len(filtered_df))
        #BMW – Part Series NE,NB,N7-Other than HOS ,SPW1 ,SPWH Plant & If the parts goes to MYSR ,HP to be considered .
        #########################################################################################################################

        #exclude_part_series_bmw = ['NE', 'NB', 'N7']
        # Create conditions for filtering
        #condition11 = ~final_df['MATERIAL'].str.startswith(tuple(exclude_part_series_bmw))
        #condition12 = ~final_df['PLANT'].isin(['HOS', 'SPW1', 'SPWH'])
        #condition13 = (final_df['PLANT'] == 'MYSR') | (final_df['PLANT'] == 'HP')

        # Combine conditions with the & operator
        #final_df = final_df[condition11 & condition12 & condition13]
        #print(len(final_df))

        ### Last filter table

        # Combine conditions with the & operator


        #Vendor Code : Other Then 2 seies can be deleted 

        filter_df2 = filtered_df[filtered_df["VENDOR"].str.startswith("2")]

        #print(filter_df2["PLANT"].unique())
        #To be excluded :
        #PO Number : Starting with 121,123,122 Series to be deleted

        #PO Number : Starting with 5500 & 94 Series to be deleted
        removed_PO = ["121", "131", "122", "123", "5500", "94"]


        #filter_df2 = filter_df2.loc[(filter_df2["PO-NUMBER"] != removed_PO)]


        condition3 = ~filter_df2['PO-NUMBER'].astype(str).str.startswith(tuple(map(str, removed_PO)))

        # Apply condition using loc
        filter_df2 = filter_df2.loc[condition3, :]

        removed_material = []
        #Removed material

        material_removed = ["WM4862","WW1362","F200070","WW4862"]

        filter_df2  = filter_df2[~filter_df2['MATERIAL'].isin(material_removed)]

        #filter_df2 = filter_df2[~filter_df2["PO-NUMBER"].astype(str).str.startswith(tuple(map(str, removed_PO)))]
        #filter_df2["PLANT"].unique()

        
        conn =dbapi.connect(
                    address='10.121.3.243',
                    port=30015,
                    user='HPR_DEVELOP',
                    password='HPRdevelop123*',
                    encrypt=True,
                    sslValidateCertificate=False
                )
        cursor=conn.cursor()

        cursor.execute('''select distinct a.matnr,b.maktx material_Desc,extwg,case when extwg ='0001' then 'Proprietory'
                                 when extwg ='0002' then 'Non-Proprietory'
                                 when extwg ='0003' then 'MIW-Common'
                                 when extwg ='0004' then 'MIW-Excl.spares'
                                 when extwg ='0005' then 'Imported'
                                 when extwg ='0006' then 'Tools'
                                 when extwg ='1001' then '3 Wheeler'
                                 when extwg ='1002' then 'Casting & Machining'
                                 when extwg ='1003' then 'Forging & Machining'
                                 when extwg ='1004' then 'IDM' 
                                 when extwg ='1005' then 'Plastic & Rubber'
                                 when extwg ='1006' then 'Press & fabrication'
                                 when extwg ='1007' then 'Prop-Electrical'
                                 when extwg ='1008' then 'Prop-Mechanical'
                                 when extwg ='1009' then 'Raw Material'
                                 when extwg ='1010' then 'Sticker'
                                 when extwg ='1011' then 'Paint'
                                 when extwg ='1012' then 'EV'
                                 else NULL end as Category
                                from "SAPPRD_HANA_BASE_EDITION"."MARA"  as A
                                join "SAPPRD_HANA_BASE_EDITION"."MAKT" as b
                                on a.matnr = b.matnr
                                where  extwg <>'' ''')
        category=pd.DataFrame(cursor.fetchall())
        category.columns = [desc[0] for desc in cursor.description]

        final_df  = filter_df2.merge(category,how = 'inner',left_on='MATERIAL',right_on='MATNR')
        final_df['Turn_Over(in Cr)'] = round (final_df['VALUE']/1e7,2)

        current_date = datetime.now()
        current_year = current_date.year
        current_month = current_date.month

        # Convert 'YEAR MONTH' column to extract the year and month as separate columns
        final_df['YEAR'] = final_df['YEAR MONTH'].astype(str).str[:4].astype(int)
        final_df['MONTH'] = final_df['YEAR MONTH'].astype(str).str[4:].astype(int)
        plant_wise = final_df.groupby('PLANT')["Turn_Over(in Cr)"].sum().reset_index()

        # Assuming 'category_totals' is your DataFrame
        total_sum_plant = plant_wise['Turn_Over(in Cr)'].sum()
                                                                                                                                                                    
        # Add a row for total sum
        total_row_plant = {'PLANT': 'TOTAL', 'Turn_Over(in Cr)': total_sum_plant}
        plant_wise = pd.concat([plant_wise, pd.DataFrame([total_row_plant])], ignore_index=True)



        filtered_data_cat = final_df.copy()
        # Filter the DataFrame for the current year and months up to the current month
        filtered_data_cat = filtered_data_cat[(filtered_data_cat['YEAR'] == current_year) & (filtered_data_cat['MONTH'] <= current_month)]

        #material_bmw = ['K2200040','K3212850','K6200140','KE161500','M1200170','N3200240','N8200350','N8220500','N9141300','N9170360','N9216940','N9216950','NF210650','R2010770']

        vendor_is_exclude = ["20670", "20037"]

        # Exclude rows based on the 'VENDOR' column
        vendor_deduction = filtered_data_cat[filtered_data_cat['VENDOR'].isin(vendor_is_exclude)]

        #category_totals = filtered_data_cat[~filtered_data_cat['MATERIAL'].isin(material_bmw)]

        # Exclude rows based on conditions in 'PLANT' and 'CATEGORY' columns
        deducted_value = filtered_data_cat.loc[((filtered_data_cat["PLANT"] == "JMWH") & (filtered_data_cat["CATEGORY"] == "Casting & Machining"))]

        # Group by 'CATEGORY' and sum the 'Turn_Over(in Cr)' for vendor_deduction
        vendor_deduction = vendor_deduction.groupby("CATEGORY")["Turn_Over(in Cr)"].sum().reset_index()

        # Group by 'CATEGORY' and sum the 'Turn_Over(in Cr)' for deducted_value
        deducted_value = deducted_value.groupby("CATEGORY")["Turn_Over(in Cr)"].sum().reset_index()

        # Group by 'CATEGORY' and sum the 'Turn_Over(in Cr)' for category_totals
        category_totals = filtered_data_cat.groupby('CATEGORY')['Turn_Over(in Cr)'].sum().reset_index()

        # Deduct values for 'Casting & Machining' category from 'Press & fabrication' category in category_totals
        category_totals.loc[category_totals['CATEGORY'] == 'Press & fabrication', 'Turn_Over(in Cr)'] -= deducted_value.loc[deducted_value['CATEGORY'] == 'Casting & Machining', 'Turn_Over(in Cr)'].values[0]

        # Deduct values for 'Prop-Mechanical' category from 'Press & fabrication' category in category_totals
        category_totals.loc[category_totals['CATEGORY'] == 'Press & fabrication', 'Turn_Over(in Cr)'] -= vendor_deduction.loc[vendor_deduction['CATEGORY'] == 'Prop-Mechanical', 'Turn_Over(in Cr)'].values[0]


        total_sum = category_totals['Turn_Over(in Cr)'].sum()

        # Add a row for total sum
        total_row = {'CATEGORY': 'TOTAL', 'Turn_Over(in Cr)': total_sum}
        category_totals = pd.concat([category_totals, pd.DataFrame([total_row])], ignore_index=True)

        supplier_summary  = final_df.groupby(['VENDOR NAME','CATEGORY'])['Turn_Over(in Cr)'].sum()
        supplier_summary = pd.DataFrame(supplier_summary)
        supplier_summary = supplier_summary.reset_index()

        filter_data_month = final_df.copy()

        # Convert 'YEAR MONTH' column to extract the year and month as separate columns
        #final_df['YEAR'] = final_df['YEAR MONTH'].astype(str).str[:4].astype(int)
        #final_df['MONTH'] = final_df['YEAR MONTH'].astype(str).str[4:].astype(int)

        # Calculate the current year and month
        current_date = datetime.now()
        current_year = current_date.year
        current_month = current_date.strftime("%b").upper() + str(current_date.year)[2:]

        # Filter the DataFrame for the current year and months up to the current month
        filtered_data = filter_data_month[(filter_data_month['YEAR'] == current_year) & ((filter_data_month['YEAR'] != current_year) | (filter_data_month['MONTH'] <= current_date.month))]

        # Create a new column with only the month name
        filtered_data['Month Name'] = (
            filtered_data['MONTH']
            .apply(lambda x: datetime.strptime(str(x), "%m").strftime("%b"))
            .str.upper() + ' ' + filtered_data['YEAR'].astype(str).str[2:]
        )


        # Define a custom order for month names
        month_order = [
            'APR ' + str(current_year)[2:],
            'MAY ' + str(current_year)[2:],
            'JUN ' + str(current_year)[2:],
            'JUL ' + str(current_year)[2:],
            'AUG ' + str(current_year)[2:],
            'SEP ' + str(current_year)[2:],
            'OCT ' + str(current_year)[2:],
            'NOV ' + str(current_year)[2:],
            'DEC ' + str(current_year)[2:],
            'JAN ' + str(current_year + 1)[2:],
            'FEB ' + str(current_year + 1)[2:],
            'MAR ' + str(current_year + 1)[2:]
        ]



        # Group by 'Month Name' and 'VENDOR NAME','VENDOR' and calculate the total turnover
        result = filtered_data.groupby(['Month Name', 'VENDOR','VENDOR NAME'])['Turn_Over(in Cr)'].sum().reset_index()

        # Pivot the table to have months as columns with the custom order
        pivot_result = result.pivot(index=(['VENDOR','VENDOR NAME']), columns='Month Name', values='Turn_Over(in Cr)').reindex(columns=month_order)

        pivot_result
        # Calculate the total for each vendor
        pivot_result['Turn_Over(in Cr)'] = pivot_result.sum(axis=1)


        # Fill NaN values with 0
        pivot_result.fillna(0, inplace=True)

        pivot_result = pd.DataFrame(pivot_result)
        # Display the result

        pivot_result = pivot_result.reset_index()
        filter_data_vendor = final_df.copy()

        vendor_group = filter_data_vendor.groupby(['VENDOR','VENDOR NAME','PLANT'])['Turn_Over(in Cr)'].sum().reset_index()
        plant_group = vendor_group.pivot(index =['VENDOR','VENDOR NAME'],columns = ['PLANT'],values ='Turn_Over(in Cr)')
        plant_group.fillna(0,inplace = True)
        plant_group['Total (in Cr)'] = plant_group.sum(axis=1)
        plant_group = plant_group.reset_index()
        excluded = filter_df2.merge(category, how='left', right_on='MATNR', left_on='MATERIAL')
        excluded = excluded[excluded["CATEGORY"].isna()]

        excluded_po  = [20000885 ,10001937,20000548,40001868,10003774,71000161,10004627,40001079]

        #bmw_excluded =['K2200040','K3212850','K6200140','KE161500','M1200170','N3200240','N8200350','N8220500','N9141300','N9170360','N9216940','N9216950','NF210650','R2010770']

        #excluded = excluded[~excluded['MATERIAL'].isin(material_bmw) & ~excluded["PO-NUMBER"].isin(excluded_po)]

        #excluded = excluded[~excluded['MATERIAL'].isin(material_bmw)]

        #excluded = excluded[~excluded["PO-NUMBER"].isin(excluded_po)]
        excluded = pd.DataFrame(excluded)

        # Define the file name and path
        file_name = 'supplier_turnover.xlsx'  # Desired file name
        file_path = os.path.join(os.getcwd(), file_name)  # Full file path

        # Create a dictionary with DataFrames and their sheet names
        dataframes = {
            'supplier_summary': supplier_summary,
            'pivot_result': pivot_result,
            'plant_wise_value': plant_wise,
            'plant_group': plant_group,
            'category_totals': category_totals,
            'excluded' :  excluded,
            'Raw_Data':final_df
        }

    # Save each DataFrame in separate sheets
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            for sheet_name, df in dataframes.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)

                # Set column widths based on conditions
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                worksheet.set_column('A:A', 30)

                if sheet_name in ['supplier_summary','plant_wise_value', 'plant_group', 'category_totals', 'excluded','final_df']:
                    worksheet.set_column(1, 1, 20)  # Set the width of the 2nd column to 20

                if sheet_name == 'supplier_summary':
                    worksheet.set_column(2, 2, 20)  # Set the width of the 3rd column to 20

                if sheet_name == 'pivot_result':
                    worksheet.set_column(14, 14, 20)  # Set the width of the 14th column to 20
                
                if sheet_name == 'plant_group':
                    worksheet.set_column(10, 10, 20)  # Set the width of the 11th column to 20

                # Set header formatting
                header_format = workbook.add_format({'bold': True, 'bg_color': '#FFA500'})
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)

    except Exception as e:
        print(f"An error occurred: {e}")
    
supplier_turnover()

