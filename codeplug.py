import sys
import os
import csv
from openpyxl import load_workbook


class HD1CodePlugSystem:

    def __init__(self,
                 name, talkgroup_sheet_name, tx, rx, template, radio_id
                ):
        self._name = name
        self._talkgroup_sheet_name = talkgroup_sheet_name
        self._tx = tx
        self._rx = rx
        self._template = template
        self._radio_id = radio_id

        self._talkgroups = {}

    def add_talkgroup(self, talkgroup):
        self._talkgroups[talkgroup._talkgroup] = talkgroup

    def talkgroups(self):
        return self._talkgroups


class HD1CodePlugTemplate:
    
    def __init__(self,
                 name, data
                ):
        self._name = name
        self._data = data


class HD1CodePlugTalkGroup:
    
    def __init__(self,
                 talkgroup, slot, name, short_name
                ):
        self._talkgroup = talkgroup
        self._slot = slot
        self._name = name
        self._short_name = short_name


class HD1CodePlugPriorityContact:

    def __init__(self,
                 number,
                 call_type,
                 contact_alias,
                 call_id
                ):

        self._number = number
        self._call_type = call_type
        self._contact_alias = contact_alias
        self._call_id = call_id

    def __str__(self):
        return "{0}, {1}, {2}, {3}".format(self._number, self._call_type, self._contact_alias, self._call_id)

    def populate_fields(self):
        return [self._number, self._call_type, self._call_type, "", "", "", self._call_id]


class HD1CodePlugChannelInformation:

    def __init__(self, 
                number,
                system,
                tg,
                template
                ):

        self._number = number
        self._system = system
        self._talkgroup = tg
        self._template = template

    def populate_fields(self):
        working = self._template._data.copy()

        self._replace(working, "$NUMBER", self._number)
        self._replace(working, "$TX", self._system._tx)
        self._replace(working, "$RX", self._system._rx)
        self._replace(working, "$ALIAS", self._talkgroup._short_name)
        self._replace(working, "$SLOT", "Slot{0}".format(self._talkgroup._slot))
        self._replace(working, "$CONTACT", "Priority Contacts: TG {0}".format(self._talkgroup._talkgroup))

        self._replace(working, "$RADIOID", self._system._radio_id)

        return working
    
    def _replace(self, working, name, value):
        index = 0
        for cell in working:
            if cell == name:
                working[index] = value
                break
            index = index + 1


BASE_INFO_SHEET = "HD1 Base Info"

PRIORITY_CONTACTS_SHEET = "HD1 Priority Contacts"
PRIORITY_CONTACTS_HEADER = "No.,Call Type,Contacts Alias,City,Province,Country,Call ID"

CHANNEL_INFORMATION_SHEET = "HD1 Channel Information"
CHANNEL_INFORMATION_HEADER = "No.,Channel Type,Channel Alias,Rx Frequency,Tx Frequency,Tx Power,TOT,VOX,VOX Level,Scan Add/Step,Channel Work Alone,Default to Talkaround,Band Width,Dec QT/DQT,Enc QT/DQT,Tx Authority,Relay,Work Mode,Slot,ID Setting,Color Code,Encryption,Encryption Type,Encryption Key,Promiscuous,Tx Authority,Kill Code,WakeUp Code,Contacts,Rx Group Lists,Group Lists 1,Group Lists 2,Group Lists 3,Group Lists 4,Group Lists 5,Group Lists 6,Group Lists 7,Group Lists 8,Group Lists 9,Group Lists 10,Group Lists 11,Group Lists 12,Group Lists 13,Group Lists 14,Group Lists 15,Group Lists 16,Group Lists 17,Group Lists 18,Group Lists 19,Group Lists 20,Group Lists 21,Group Lists 22,Group Lists 23,Group Lists 24,Group Lists 25,Group Lists 26,Group Lists 27,Group Lists 28,Group Lists 29,Group Lists 30,Group Lists 31,Group Lists 32,Group Lists 33,GPS,Send GPS Info,Receive GPS Info,GPS Timing Report,GPS Timing Report TX Contacts"

ADDRESS_BOOK_CONTACTS_SHEET = "HD1 Address Book Contacts"

class HD1CodePlugSpreadsheet:


    def __init__(self, 
                 spreadsheet_filename,
                 config_sheet = BASE_INFO_SHEET):
        
        print("Loading workbook: ", spreadsheet_filename)

        self._spreadsheet_filename = spreadsheet_filename
        self._config_sheet = config_sheet

        self._workbook = load_workbook(self._spreadsheet_filename, data_only=True)

        self._radio_id = "RADIOID"

        self._systems = {}
        self._templates = {}

        self._priority_contacts = []
        self._channels = []

        self._config = self.load_config_sheet()

    def load_config_sheet(self):

        if self._config_sheet not in self._workbook:
            raise Exception("Config sheet '{0}' missing or mispelled".format(self._config_sheet))

        print("Loading configuration from: ", self._config_sheet)

        self._load_base_info()

    def _load_base_info(self):
        base_info = self._workbook[self._config_sheet]
        self._load_radio_id(base_info)
        self._load_systems(base_info)
        self._load_templates(base_info)

    def _find_start_row(self, sheet, text):
        line = 1
        for row in sheet:
            if row[0].value == text:
                return line
            line = line + 1

        return -1

    def _load_radio_id(self, base_info):
        line_count = self._find_start_row(base_info, 'Radio ID')
        if line_count == -1:
            raise Exception ("No Radio ID defined in config") 
        line_count = line_count + 1  

        self._radio_id = name = base_info["A{0}".format(line_count)].value

    def _load_systems(self, base_info):

        line_count = self._find_start_row(base_info, 'System')
        if line_count == -1:
            raise Exception ("No Systems defined in config") 
        line_count = line_count + 1  

        process = True
        while process is True:
            name = base_info["A{0}".format(line_count)].value
            if name is not None and name != '':
                include = base_info["B{0}".format(line_count)].value
                if include == 'Y':
                    talkgroup_sheet_name = base_info["C{0}".format(line_count)].value
                    tx = base_info["D{0}".format(line_count)].value
                    rx = base_info["E{0}".format(line_count)].value
                    template = base_info["F{0}".format(line_count)].value

                    self._systems[name] = HD1CodePlugSystem(name, talkgroup_sheet_name, tx, rx, template, self._radio_id)

            else:
                process = False  
            line_count = line_count + 1     

    def _load_templates(self, base_info):

        line_count = self._find_start_row(base_info, 'Templates')
        if line_count == -1:
            raise Exception ("No Templates defined in config") 
        line_count = line_count + 1  

        process = True
        while process is True:
            name = base_info["A{0}".format(line_count)].value
            if name is not None and name != '':
                end_column = base_info["B{0}".format(line_count)].value

                start_cell = "C{0}".format(line_count)
                end_cell = "{0}{1}".format(end_column, line_count)

                items = base_info["{0}".format(start_cell):"{0}".format(end_cell)]
                # data = [str(cell.value) for cell in items[0]]
                data = [cell.value for cell in items[0]]


                self._templates[name] = HD1CodePlugTemplate(name, data)

            else:
                process = False  

            line_count = line_count + 1     

    def load_talkgroups(self):

        for talkgroup_system in self._systems.values():
            print("Loading talkgroups for:", talkgroup_system._talkgroup_sheet_name)
            talkgroup_ws = self._workbook[talkgroup_system._talkgroup_sheet_name]
            first = True
            for row in talkgroup_ws:
                if first is False:
                    tg = HD1CodePlugTalkGroup(row[0].value, row[1].value, row[2].value, row[3].value)
                    talkgroup_system.add_talkgroup(tg)
                else:
                    first = False

    def create_priority_contacts(self):

        print("Generating Priority Contacts")

        talkgroup_ids = []

        for talkgroup_system in self._systems.values():
            print("\tProcessing talkgroups for:", talkgroup_system._talkgroup_sheet_name)
            for tg in talkgroup_system.talkgroups():
                talkgroup_ids.append(tg)

        talkgroup_ids.sort()

        deduped_ids = list(dict.fromkeys(talkgroup_ids))

        count = 1
        for id in deduped_ids:
            pc = HD1CodePlugPriorityContact(count, "Group Call", "TG {0}".format(id), id)
            self._priority_contacts.append(pc)
            count = count + 1

        self._write_priority_contacts_to_worksheet()

    def _create_new_worksheet(self, worksheet_name, idx):

        if worksheet_name in self._workbook.sheetnames:
            idx = self._workbook.sheetnames.index(worksheet_name)
            self._workbook.remove(self._workbook[worksheet_name])

        ws =  self._workbook.create_sheet(worksheet_name, idx)

        return ws

    def _write_priority_contacts_to_worksheet(self):

        priority_contacts_ws = self._create_new_worksheet(PRIORITY_CONTACTS_SHEET, 1)

        fields = PRIORITY_CONTACTS_HEADER.split(",")
        column = 1
        for field in fields:
            priority_contacts_ws.cell(1, column).value = field
            column = column + 1

        row = 2        
        for contact in self._priority_contacts:

            fields = contact.populate_fields()

            column = 1
            for field in fields:
                priority_contacts_ws.cell(row, column).value = field
                column = column + 1

            row = row + 1

    def create_channel_information(self):

        print("Generating Channel Informatio")

        count = 1
        for talkgroup_system in self._systems.values():
            print("\tProcessing talkgroups for:", talkgroup_system._talkgroup_sheet_name)
   
            print("\t\tLoading template: ", talkgroup_system._template)
            template = self._templates[talkgroup_system._template]

            for tg in talkgroup_system._talkgroups.values():
                ci = HD1CodePlugChannelInformation(count, talkgroup_system, tg, template)
                self._channels.append(ci)
                count = count + 1

        self._write_channel_info_to_worksheet()

    def _write_channel_info_to_worksheet(self):

        channel_info_ws =  self._create_new_worksheet(CHANNEL_INFORMATION_SHEET, 2)

        fields = CHANNEL_INFORMATION_HEADER.split(",")
        column = 1
        for field in fields:
            channel_info_ws.cell(1, column).value = field
            column = column + 1

        row = 2
        for channel in self._channels:

            fields = channel.populate_fields()

            column = 1
            for field in fields:
                channel_info_ws.cell(row, column).value = field
                column = column + 1

            row = row + 1

    def create_xlsx(self):
        codeplug.load_talkgroups()
        codeplug.create_priority_contacts()
        codeplug.create_channel_information()
        #codeplug.save(self.__spreadsheet_filename)
        codeplug.save("working2.xlsx")

    def save(self, filename):
        print("Saving workbook '{0}".format(filename))

        self._workbook.save(filename)

    def create_csvs(self):

        self._export_sheet_csv(PRIORITY_CONTACTS_SHEET)
        self._export_sheet_csv(CHANNEL_INFORMATION_SHEET)
        self._export_sheet_csv(ADDRESS_BOOK_CONTACTS_SHEET)

    def _export_sheet_csv(self, sheet_name):

        print("Writing {0} to csv".format(sheet_name))

        ws = self._workbook[sheet_name]
        if not ws:
            print("Worksheet '{0}' does not existing, skipping csv".format(sheet_name))

        filepath = sheet_name+".csv"

        if os.path.exists(filepath):
            os.remove(filepath)

        with open(filepath, 'w', newline="") as f:
            c = csv.writer(f)
            for r in ws.rows:
                c.writerow([cell.value for cell in r])


if __name__ == '__main__':

    codeplug = HD1CodePlugSpreadsheet(sys.argv[1])
    
    if sys.argv[2] == 'xlsx':
        codeplug.create_xlsx()
    elif sys.argv[2] == 'csv':
        codeplug.create_csvs()
    else:
        print("Unknown command line option xlsx or csv only")