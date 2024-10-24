from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd
import numpy as np
from openpyxl.styles import Alignment

input_file = r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\6.BHS.xlsx'
input_sheet = 'Sheet1'
output_path = r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Python output\2148Analysis_BHS24oct.xlsx'

df = pd.read_excel(input_file,input_sheet)
df=df[df['Q2'] == 1]









# Writing Error index in dataframe

# Error_n = ['Count1: Zone Quota ']
# df_name= pd.DataFrame(Error_n) 
# df_name.to_excel(output_path, index=False, header=False)


# Step 1: Create a list of topic names
topics = ['Table1 :Combined table of Counts for total awareness and Total usage','Table2:Shifting from previous to current regular brand','Table 3 : Count:Age wise brands','Table4 : Count:SEC wise brands']


df_name = pd.DataFrame(topics, columns=['Logic list'])
df_name.to_excel(output_path, index=False)




def inside_append_dataframe_with_blank_rows(file_path, dataframe, blank_rows=2):
    """
    Appends a DataFrame to an existing Excel file with a specified number of blank rows in between.

    Parameters:
    - file_path: str, path to the Excel file
    - dataframe: pd.DataFrame, DataFrame to append
    - blank_rows: int, number of blank rows between appended DataFrames (default is 5)
    """
    
    with pd.ExcelWriter(
            file_path,
            engine='openpyxl',
            mode='a',
            if_sheet_exists='overlay') as writer:
        reader = pd.read_excel(file_path)
        dataframe.to_excel(
            writer,
            startrow=reader.shape[0] + blank_rows,
            index=True,
            header=True)

   
        
    workbook = load_workbook(file_path)
    worksheet = workbook.active

    # Center-align the text in the appended DataFrame
    for row in worksheet.iter_rows(min_row=startrow + 1, max_row=startrow + len(dataframe) + 3, min_col=1, max_col=len(dataframe.columns) + 1):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # Save the workbook
    workbook.save(file_path)     







def outside_append_dataframe_with_blank_rows(file_path, dataframe, blank_rows=5):
    """
    Appends a DataFrame to an existing Excel file with a specified number of blank rows in between.

    Parameters:
    - file_path: str, path to the Excel file
    - dataframe: pd.DataFrame, DataFrame to append
    - blank_rows: int, number of blank rows between appended DataFrames (default is 5)
    """
    
    with pd.ExcelWriter(
            file_path,
            engine='openpyxl',
            mode='a',
            if_sheet_exists='overlay') as writer:
        reader = pd.read_excel(file_path)
        dataframe.to_excel(
            writer,
            startrow=reader.shape[0] + blank_rows,
            index=True,
            header=True)











# Count1: Tom  


# Error_n = ['Count1 :Tom']
# df_name= pd.DataFrame(Error_n) 
# existing_df = pd.read_excel(output_path)
# startrow = existing_df.shape[0] + 4
# with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
#     df_name.to_excel(writer, startrow=startrow, index=False, header=False)




df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)


df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','MOUB')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))





df_s = df_global['Q14'].value_counts()


Valuecount1=pd.DataFrame(df_s)
# total_sum = Valuecount1.sum()
# total_sum.name = 'Total_sum'
# Valuecount1=Valuecount1._append(total_sum)




Valuecount1.index = Valuecount1.index.to_series().replace(dict_format)








# Count 2 = Spont brand count






df_global=df.copy()

df_global.reset_index(drop=True, inplace=True)



df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','BrandcolumnQ14')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount2=pd.DataFrame(sum_of_columns)







# Count 3 = Unaided awareness









Valuecount3 = pd.concat([Valuecount1, Valuecount2], axis=1)
Valuecount3.columns = ['TOM', 'Spont']
Valuecount3['Unaided awareness']=Valuecount3.sum(min_count=1,axis=1)






# Count 4 = Aided Awareness brand count






df_global=df.copy()

df_global.reset_index(drop=True, inplace=True)


df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','Brand15')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount4=pd.DataFrame(sum_of_columns)






# Count 5 = Total Awareness





df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)



df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','Brandtotalaware')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount5=pd.DataFrame(sum_of_columns)








# Count 6 = Ever smoked brand count







df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)



df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','Brandeversmoke(Q16)')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount6=pd.DataFrame(sum_of_columns)












# Count 7 = Last 2 year smoked brand count



df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)



df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','Brand(Q17)')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount7=pd.DataFrame(sum_of_columns)





# Count 8 = Last 1 year smoked brand count



df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)



df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','Q18')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount8=pd.DataFrame(sum_of_columns)















# Count 9 = Last 1 month smoked brand count



df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)



df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','Q19')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount9=pd.DataFrame(sum_of_columns)










# Count 10 = Last 2 week smoked brand count



df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)



df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','Q20.1')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount10=pd.DataFrame(sum_of_columns)











# Count 11 = Last 1 week smoked brand count



df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)



df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','Q20')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))
df_global.rename(columns=dict_format, inplace=True)

renamed_columns = list(dict_format.values())

# Adding a new column 'Row_Sum' that contains the sum of each row across the renamed columns
sum_of_columns= df_global[renamed_columns].sum()




Valuecount11=pd.DataFrame(sum_of_columns)








# Count 12 = MOUB




df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)


df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','MOUB')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))





df_s = df_global['Q21'].value_counts()


Valuecount12=pd.DataFrame(df_s)
# total_sum = Valuecount1.sum()
# total_sum.name = 'Total_sum'
# Valuecount1=Valuecount1._append(total_sum)




Valuecount12.index = Valuecount12.index.to_series().replace(dict_format)






# Concat


Error_n = [' Table1 :Combined table of Counts for total awareness and Total usage']
df_name= pd.DataFrame(Error_n) 
existing_df = pd.read_excel(output_path)
startrow = existing_df.shape[0] + 4
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df_name.to_excel(writer, startrow=startrow, index=False, header=False)




resultant = pd.concat([Valuecount3,Valuecount4,Valuecount5,Valuecount6,Valuecount7,Valuecount8,Valuecount9,Valuecount10,Valuecount11,Valuecount12], axis=1)
resultant.columns = ['TOM', 'Spont', 'Unaided','Aided awareness','Total Awareness','Ever smoked','Last 2 year smoked','Last 1 year smoked', 'Last 1 month smoked', 'Last 2 weeks smoked','Last 1 weeks smoked','Regularly smoking' ]

total_sum = resultant.sum()
total_sum.name = 'Total_sum'
resultant=resultant._append(total_sum)





inside_append_dataframe_with_blank_rows(output_path, resultant)





















# Count = Shifting from previous to current regular brand 

Error_n = ['Table2:Shifting from previous to current regular brand  ']
df_name= pd.DataFrame(Error_n) 
existing_df = pd.read_excel(output_path)
startrow = existing_df.shape[0] + 4
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df_name.to_excel(writer, startrow=startrow, index=False, header=False)





df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','MOUB')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))




df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)

df_global['Q21']=df_global['Q21'].replace(dict_format)

df_global['Q24'] = df_global['Q24'].replace(-1, np.nan)
df_global['Q24']=df_global['Q24'].replace(dict_format)



shift_matrix = pd.crosstab(df_global['Q24'], df_global['Q21'])



inside_append_dataframe_with_blank_rows(output_path, shift_matrix)








# Count:Age wise brands 




Error_n = ['Table 3 : Count:Age wise brands ']

df_name= pd.DataFrame(Error_n) 
existing_df = pd.read_excel(output_path)
startrow = existing_df.shape[0] + 4
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df_name.to_excel(writer, startrow=startrow, index=False, header=False)




df_global=df.copy()
df_global.reset_index(drop=True, inplace=True)
df_global['Q6'] = pd.to_numeric(df_global['Q6'], errors='coerce')
def categorize_scores(score):
    
    if 21<=score <= 25:
        return '21-25 Yrs'
    elif 26 <= score <= 30:
        return '26-30 Yrs'
    elif 31<=score <= 35:
        return '31-35 Yrs'
    elif 36<=score<=40:
        return '36-40 Yrs'
    elif 41<=score<=45:
        return '41-45 Yrs'
    elif 46<=score<=50:
        return '46-50 Yrs'
    else:
        return 'Other'
    

df_global['Age'] = df_global['Q6'].apply(categorize_scores) 






df_dictformation = pd.read_excel(r'D:\Input\Cigarette\Brand Health Study(October 2024)\Input\Data input\BHS 22-Oct-24_DataMap.xlsx','MOUB')



# Join two columns into a dictionary
# Assuming 'Column1' and 'Column2' are the names of the columns you want to join
dict_format = dict(zip(df_dictformation['Column'], df_dictformation['Brand']))




df_global['Q21'] = df_global['Q21'].replace(dict_format)

df_modified = df_global.loc[:,['Q21','Age']]

crosstab_result = pd.crosstab(df_modified['Age'], df_modified['Q21'])



# dynamic_headers = [f'brands{i}' for i in range(1, len(crosstab_result.columns) + 1)]

count1=pd.DataFrame(crosstab_result )

inside_append_dataframe_with_blank_rows(output_path, count1)







