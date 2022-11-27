import pandas as pd
import sys
from colorama import Fore, init
init(autoreset=True)

print(Fore.GREEN + "Client List script successfully launched.")
print('Running: Pandas', pd.__version__, 'Preferred (1.5.2) | ', 'Python', sys.version)
print("-----------------------------------------------")

input_sheet = pd.read_excel('input/read/List1.xlsx')
output_sheet = pd.read_excel('input/compare/List2.xlsx')

inumber_series = pd.Series(
    input_sheet.get('Client number', default="Client number ERROR: No column named 'Client number' found. (INPUT)"))
iname_series = pd.Series(
    input_sheet.get('Client name', default="Client name ERROR: No column named 'Client name' found. (INPUT)"))
ipartner_series = pd.Series(
    input_sheet.get('Partner', default="Partner ERROR: No column named 'Partner' found. (INPUT)"))

onumber_series = pd.Series(
    output_sheet.get('Client number', default="Client number ERROR: No column named 'Client number' found. (OUTPUT)"))
oname_series = pd.Series(
    output_sheet.get('Client name', default="Client name ERROR: No column named 'Client name' found. (OUTPUT)"))
opartner_series = pd.Series(
    output_sheet.get('Partner', default="Partner ERROR: No column named 'Partner' found. (OUTPUT)"))

updated_partners = ["Empty"] * len(ipartner_series.array)
confirmed_cnames = ["Empty"] * len(iname_series.array)
confirmed_cnumbers = ["Empty"] * len(inumber_series.array)

# TODO: add save functionality

def confirmation(mismatch_type: str, case_num: int, input_series: pd.Series, output_series: pd.Series):

    if mismatch_type == "client_name":
        print(Fore.RED + "---------- CLIENT NAME MISMATCH (CASE " + str(case_num + 1) + ") ----------")
        print("Confirm that the client name is still correct; if confirmed, the input client name will be used.")
        print("Enter 'y' to confirm, or 'n' to change it.")
        choice = input("Client name: " + input_series[case_num] + " | " + output_series[case_num] + " (y/n): ")

        if choice == "y":
            print("Confirmed.")
            confirmed_cnames[case_num] = input_series[case_num]
        elif choice == "n":
            print("Enter the correct client name.")
            confirmed_cnames[case_num] = input("Updated client name: ")
            print("Client name mismatch updated. (CASE " + str(case_num + 1) + ")")

    elif mismatch_type == "client_number":

        print(Fore.RED + "---------- CLIENT NUMBER MISMATCH (CASE " + str(case_num + 1) + ") ----------")
        print("Confirm that the client number is still correct; if confirmed, the input client number will be used.")
        print("Enter 'y' to confirm, or 'n' to change it.")
        choice = input("Client number: " + str(input_series[case_num]) + " | " + str(output_series[case_num]) + " (y/n): ")

        if choice == "y":
            print("Confirmed.")
            confirmed_cnumbers[case_num] = input_series[case_num]
        elif choice == "n":
            print("Enter the correct client number.")
            confirmed_cnumbers[case_num] = input("Updated client number: ")
            print("Client number mismatch updated. (CASE " + str(case_num + 1) + ")")

    elif mismatch_type == "empty_partner":
        print(Fore.RED + "---------- EMPTY PARTNER (CASE " + str(case_num + 1) + ") ----------")
        print("Missing a partner entry in the READ file")
        print("Enter the current partner assigned to client: " + iname_series[case_num] + " #" + inumber_series[case_num])
        updated_partners[case_num] = input("Updated partner: ")
        print("Partner entry updated. (CASE " + str(case_num + 1) + ")")


print("Loaded sheet data: Successfully.")
print("Commencing data transfer...")
for i in range(len(ipartner_series.array)):
    if iname_series.array[i] == oname_series.array[i] and inumber_series.array[i] == onumber_series.array[i]:  # all data matches -automatically proceed

        updated_partners[i] = ipartner_series.array[i]
        confirmed_cnames[i] = oname_series.array[i]
        confirmed_cnumbers[i] = onumber_series.array[i]
        print("CASE " + str(i + 1) + ": " + Fore.YELLOW + "Data matches: Successful transfer.")

    elif iname_series.array[i] != oname_series.array[i] and inumber_series.array[i] == onumber_series.array[i]:  # name mismatch - ask for confirmation
        confirmation("client_name", i, iname_series, oname_series)
        confirmed_cnumbers[i] = onumber_series.array[i]
        updated_partners[i] = ipartner_series.array[i]

    elif iname_series.array[i] == oname_series.array[i] and inumber_series.array[i] != onumber_series.array[i]:  # number mismatch - ask for confirmation
        confirmation("client_number", i, inumber_series, onumber_series)
        confirmed_cnames[i] = oname_series.array[i]
        updated_partners[i] = ipartner_series.array[i]

    elif iname_series.array[i] != oname_series.array[i] and inumber_series.array[i] != onumber_series.array[i]:  # both mismatch - ask for confirmation
        confirmation("client_name", i, iname_series, oname_series)
        confirmation("client_number", i, inumber_series, onumber_series)
        updated_partners[i] = ipartner_series.array[i]

    elif iname_series.array[i] == oname_series.array[i] and inumber_series.array[i] == onumber_series.array[i] and ipartner_series.array[i] == "" or "NaN":  # partner mismatch - ask for confirmation
        confirmation("empty_partner", i, iname_series, onumber_series)
        confirmed_cnames[i] = oname_series.array[i]
        confirmed_cnumbers[i] = onumber_series.array[i]

    # updated_partners = pd.concat([onumber_series.where(onumber_series == onumber_series.array[i]),
    #                               oname_series.where(oname_series == oname_series.array[i]),
    #                               ipartner_series.where(ipartner_series == ipartner_series.array[i])], axis=1, ignore_index=True)

print(Fore.GREEN + "Data transfer complete: All checks successful.")
print(updated_partners)

new_sheet = pd.DataFrame(
    {'Client number': confirmed_cnumbers, 'Client name': confirmed_cnames, 'Partner': updated_partners})
print(new_sheet)

# outputs sheet for copy and paste
writer = pd.ExcelWriter('output/new_sheet.xlsx')

new_sheet.to_excel(writer, 'Sheet1')
writer.save()
