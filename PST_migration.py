import pandas as pd
import time
import MailController as mc
import os


def something_went_wrong(message):
    print(" ¯\_(ツ)_/¯ Something went wrong, closing!" + message)
    time.sleep(3)
    exit()


def user_inputs():
    config_file = os.getcwd().replace("\\", "\\\\")
    with open(config_file + "\\config.txt") as config:
        for line in config:
            try:
                if "path_of_input_file" in line:
                    path_of_input_file = line.split("<")[1].replace(">\n", "")

                if "input_worksheet" in line:
                    input_worksheet = line.split("<")[1].replace(">\n", "")

                if "path_of_output_files" in line:
                    path_of_output_files = line.split("<")[1].replace(">\n", "")

                if "ticket_number" in line:
                    ticket_number = line.split("<")[1].replace(">\n", "").replace(">", "")
            except Exception:
                something_went_wrong(" [!] inputs are probably wrong.")

    check = False

    while not check:
        print("Make sure all the settings are set properly:\n")
        print("1. Path of input files: {0} \n2. Worksheet: "
              "{1} \n3. Path of output files: ""{2} \n4. Ticket number: {3}"
              .format(path_of_input_file, input_worksheet, path_of_output_files, ticket_number))

        graph_seperator()

        a = input('Type "yes/y" if you want continue. Type "no/n" if you want to change settings: ').lower()

        graph_seperator()
        result_a = ["yes", "y"]
        result_b = ["no", "n"]

        if a in result_a:
            path_of_input_file = path_of_input_file.replace("\\", "\\\\")
            path_of_output_files = path_of_output_files.replace("\\", "\\\\")
            ticket_number = str(ticket_number)

            check = True

        elif a in result_b:
            path_of_input_file = input("[!] Add path to the starting file --> ")
            path_of_input_file.replace("\\", "\\\\")
            input_worksheet = input("[!] Add name of the worksheet --> ")

            path_of_output_files = input("[!] Where should be new files saved? --> ")
            path_of_output_files.replace("\\", "\\\\")

            ticket_number = input("[!] What's ticket number? --> ")
        else:
            print("\n[!] Wrong input!")

    return path_of_input_file, path_of_output_files, input_worksheet, ticket_number


def get_data(path, worksheet):
    try:
        data_parsed = pd.read_excel(path, sheet_name=worksheet)
        return data_parsed
    except Exception:
        something_went_wrong("[!] Path or worksheet names might be wrong. If not try to close the excel file and"
                             "run the program afterwards.")


def graph_seperator():
    print("\n" + 50 * "-")


def dict_creator(data):
    data_in_list = {}
    for part in data:
        case = {part: []}
        data_in_list.update(case)

    return data_in_list


def get_countries(data):
    countries = []

    for country in data:
        if country not in countries:
            countries.append(country)

    country_list = dict_creator(countries)

    return country_list


def item_creator(data):
    headers = dict_creator(data.columns)
    header_lists = []
    all_items = []

    for header in headers:
        header_lists.append(data[header].tolist())

    item_length = len(data)
    for count in range(item_length):
        new_item = []
        for item in header_lists:
            new_item.append(item[count])
        all_items.append(new_item)

    return all_items


def record_parser(items):
    headers = dict_creator(data.columns)
    record = {}
    all_records = []

    for item in items:
        counter = 0
        for key in headers:
            case = {key: str(item[counter])}
            record.update(case)
            counter += 1

        all_records.append(record.copy())

    return all_records


def country_selector(records):
    records_by_country = get_countries(data["Country"].tolist())

    for record in records:
        for key in records_by_country:
            if str(key) in str(record.get("Country")):
                records_by_country[key].append(record)

    return records_by_country


def create_xlsx(data, output_path):
    file_paths = []

    try:
        for key in data:
            country = data[key]
            create_file = pd.DataFrame(country)
            writer = pd.ExcelWriter(output_path + "\\" + key + ".xlsx", engine="xlsxwriter")
            create_file.to_excel(writer, sheet_name=key, index=False)
            writer.save()

            file_paths.append(output_path + "\\\\" + key + ".xlsx")

    except:
        something_went_wrong("")

    return file_paths


def send_mail_contacts(files, ticket, template):
    contact_file = os.getcwd().replace("\\", "\\\\")
    sent = []
    not_sent = []
    all_full_contacts = []
    country_exist = []
    for file in files:
        with open(contact_file + "\\contacts.txt") as contacts:
            for line in contacts:
                country_exist.append(line.split(":")[0].lower())
                all_full_contacts.append(line)

        if file.split("\\")[-1].replace(".xlsx", "").lower() in country_exist:
            for line in all_full_contacts:
                if file.split("\\")[-1].replace(".xlsx", "").lower() == line.split(":")[0].lower():
                    address = line.split("<")[1].replace(">\n", "")
            try:
                sent.append(mc.MailSender(address, ticket, file, template).sent_mail())
            except Exception:
                something_went_wrong(" [!] Mail could not been sent. Make sure you turned Outlook on.")

        else:
            not_sent.append(file.split("\\")[-1].replace(".xlsx", ""))

    print("\n          ------ Success! ------")
    for mail in sent:
        print(mail)

    print("\n      ----- Has not been sent ------")
    for mail in not_sent:
        print(mail + ".xlsx" + " ---> has not been sent due to missing contact")


ahoj = '''
          PST_migrations!!

░░░░░░▄▀▒▒▒▒░░░░▒▒▒▒▒▒▒▒▒▒▒▒▒█    TURN OUTLOOK ON!!
░░░░░█▒▒▒▒░░░░▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒▒█    TURN OUTLOOK ON!!
░░░░█▒▒▄▀▀▀▀▀▄▄▒▒▒▒▒▒▒▒▒▄▄▀▀▀▀▀▀▄   TURN OUTLOOK ON!!
░░▄▀▒▒▒▄█████▄▒█▒▒▒▒▒▒▒█▒▄█████▄▒█   TURN OUTLOOK ON!!
░█▒▒▒▒▐██▄████▌▒█▒▒▒▒▒█▒▐██▄████▌▒█   TURN OUTLOOK ON!!
▀▒▒▒▒▒▒▀█████▀▒▒█▒░▄▒▄█▒▒▀█████▀▒▒▒█   TURN OUTLOOK ON!!
▒▒▐▒▒▒░░░░▒▒▒▒▒█▒░▒▒▀▒▒█▒▒▒▒▒▒▒▒▒▒▒▒█   TURN OUTLOOK ON!!
▒▌▒▒▒░░░▒▒▒▒▒▄▀▒░▒▄█▄█▄▒▀▄▒▒▒▒▒▒▒▒▒▒▒▌   TURN OUTLOOK ON!!
▒▌▒▒▒▒░▒▒▒▒▒▒▀▄▒▒█▌▌▌▌▌█▄▀▒▒▒▒▒▒▒▒▒▒▒▐   TURN OUTLOOK ON!!
▒▐▒▒▒▒▒▒▒▒▒▒▒▒▒▌▒▒▀███▀▒▌▒▒▒▒▒▒▒▒▒▒▒▒▌  TURN OUTLOOK ON!!
▀▀▄▒▒▒▒▒▒▒▒▒▒▒▌▒▒▒▒▒▒▒▒▒▐▒▒▒▒▒▒▒▒▒▒▒█  TURN OUTLOOK ON!!
▀▄▒▀▄▒▒▒▒▒▒▒▒▐▒▒▒▒▒▒▒▒▒▄▄▄▄▒▒▒▒▒▒▄▄▀  TURN OUTLOOK ON!!
▒▒▀▄▒▀▄▀▀▀▄▀▀▀▀▄▄▄▄▄▄▄▀░░░░▀▀▀▀▀▀    TURN OUTLOOK ON!!
▒▒▒▒▀▄▐▒▒▒▒▒▒▒▒▒▒▒▒▒▐ 
'''

print(ahoj)
while True:

    input_file, output_files, worksheet, ticket = user_inputs()
    data = get_data(input_file, worksheet)
    items = item_creator(data)
    all_records = record_parser(items)
    complete_data = country_selector(all_records)

    write_data = create_xlsx(complete_data, output_files)
    template_decider = input('\nWhich template do you want to use?\nFor "File Server" press "FS",\n'
                             'for "Awaiting – Discovered" pres "A",\n'
                             'for "Owner unclear - User disagree" press "O",\n'
                             'for "Not Enabled Users" press "N",\n'
                             'for "Failed Items" press "FI"\n'
                             '\033[1m[+] Your choice: \033[0m')

    if template_decider.upper() not in ("FS", "A", "N", "O", "FI"):
        something_went_wrong("[!] If the program failed in this stage you can start panicking and screaming because "
                             "idk how this could happen.")

    send_mail_contacts(write_data, ticket, template_decider)

    graph_seperator()

    close = input('Press "C" to end the program: ')
    if close.upper() == "C":
        print("closing...")
        time.sleep(4)
        break
