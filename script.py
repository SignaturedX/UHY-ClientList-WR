import pandas as pd

print("Client List script successfully launched.")
print('Running: Pandas', pd.__version__)
print("-----------------------------------------------")

input_sheet = pd.read_excel('input/read/List1.xlsx')
output_sheet = pd.read_excel('input/compare/List2.xlsx')

inumber_series = pd.Series(input_sheet.get('Client number', default="Client number ERROR: No column named 'Client number' found. (INPUT)"))
iname_series = pd.Series(input_sheet.get('Client name', default="Client name ERROR: No column named 'Client name' found. (INPUT)"))
ipartner_series = pd.Series(input_sheet.get('Partner', default="Partner ERROR: No column named 'Partner' found. (INPUT)"))

onumber_series = pd.Series(output_sheet.get('Client number', default="Client number ERROR: No column named 'Client number' found. (OUTPUT)"))
oname_series = pd.Series(output_sheet.get('Client name', default="Client name ERROR: No column named 'Client name' found. (OUTPUT)"))
opartner_series = pd.Series(output_sheet.get('Partner', default="Partner ERROR: No column named 'Partner' found. (OUTPUT)"))

updated_partners = pd.DataFrame(columns=['Client number', 'Client name', 'Partner'])

print("Loaded sheet data: Successfully.")
print("Commencing data transfer...")
for i in range(len(ipartner_series.array)):
    if iname_series.array[i] == oname_series.array[i] and inumber_series.array[i] == onumber_series.array[i]:
        updated_partners = pd.concat([onumber_series.where(onumber_series == onumber_series.array[i]),
                                      oname_series.where(oname_series == oname_series.array[i]),
                                      ipartner_series.where(ipartner_series == ipartner_series.array[i])], axis=1, ignore_index=True)
        print("iterated")

print("Data transfer complete.")
print(updated_partners)

writer = pd.ExcelWriter('output/updated_partners.xlsx')

updated_partners.to_excel(writer, 'Sheet1')
writer.save()