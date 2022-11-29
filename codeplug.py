from openpyxl import load_workbook
import sys

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

    def to_row(self):
        working = self._template._data.copy()

        self._replace(working, "$NUMBER", str(self._number))
        self._replace(working, "$TX", str(self._system._tx))
        self._replace(working, "$RX", str(self._system._rx))
        self._replace(working, "$ALIAS", str(self._talkgroup._short_name))
        self._replace(working, "$SLOT", "Slot{0}".format(self._talkgroup._slot))
        self._replace(working, "$CONTACT", "Priority Contacts: TG {0}".format(self._talkgroup._talkgroup))

        self._replace(working, "$RADIOID", self._system._radio_id)

        return ", ".join(working)
    
    def _replace(self, working, name, value):
        index = 0
        for cell in working:
            if cell == name:
                working[index] = value
                break
            index = index + 1


class HD1CodePlugSpreadsheet:

    def __init__(self, 
                 spreadsheet_filename,
                 config_sheet = "HD1 Base Info"):
        
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

        self._load_talkgroups()

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
                data = [str(cell.value) for cell in items[0]]

                self._templates[name] = HD1CodePlugTemplate(name, data)

            else:
                process = False  

            line_count = line_count + 1     

    def _load_talkgroups(self):

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

        print("\tGenerating Priority Contacts")

        talkgroup_ids = []

        for talkgroup_system in self._systems.values():
            print("\t\tProcessing talkgroups for:", talkgroup_system._talkgroup_sheet_name)
            for tg in talkgroup_system.talkgroups():
                talkgroup_ids.append(tg)

        talkgroup_ids.sort()

        deduped_ids = list(dict.fromkeys(talkgroup_ids))

        count = 1
        for id in deduped_ids:
            pc = HD1CodePlugPriorityContact(count, "Group Call", "TG {0}".format(id), id)
            self._priority_contacts.append(pc)
            count = count + 1

    def create_channel_information(self):

        print("\tGenerating Channel Informatio")

        count = 1
        for talkgroup_system in self._systems.values():
            print("\t\tProcessing talkgroups for:", talkgroup_system._talkgroup_sheet_name)
   
            print("\t\t\tLoading template: ", talkgroup_system._template)
            template = self._templates[talkgroup_system._template]

            for tg in talkgroup_system._talkgroups.values():
                ci = HD1CodePlugChannelInformation(count, talkgroup_system, tg, template)
                self._channels.append(ci)
                count = count + 1

        for channel in self._channels:
            print(channel.to_row())        


if __name__ == '__main__':

    codeplug = HD1CodePlugSpreadsheet(sys.argv[1])
    
    codeplug.create_priority_contacts()

    codeplug.create_channel_information()

