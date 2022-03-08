# This is a sample Python script.
import wx
from wx import adv as wx_adv
from ctypes import *

# EEPROM eNameplate data structure to be written/read back.
#   - size: Allocation of buffer in Bytes
#   - setValue: Element buffer to be written during Write process, Populated from Form controls before write.
#   - getValue: Element buffer read into during Read process
dict_eeprom_data = {
    "NamePlateRevision": {"size": 2, "setValue": b'1', "getValue": ""},         # Hard-coded in this programmer
    "PartNumber": {"size": 15, "setValue": "", "getValue": ""},                 # 872000xxx-3yy-z
    "BoardSerial": {"size": 11, "setValue": "", "getValue": ""},                # User PCBA Serial
    "LRUSerial": {"size": 11, "setValue": "", "getValue": ""},                  # User LRU Serial - Optional
    "AssemblyManfDate": {"size": 6, "setValue": "", "getValue": ""},            # YYMMDD  - Not used
    "AssemblyMFGSite": {"size": 2, "setValue": "", "getValue": ""},             # BN/MX/UT  - Not used
    "AssemblyModLevel": {"size": 2, "setValue": "", "getValue": ""},            # Number  - Not used
    "AssemblyModDate": {"size": 6, "setValue": "", "getValue": ""},             # YYMMDD  - Not used
    "AssemblyModSite": {"size": 2, "setValue": "", "getValue": ""},             # BN/MX/UT  - Not used
    "MACcount": {"size": 3, "setValue": b'0', "getValue": ""},                  # Number 0-255
    "MACEntry": [{"size": 6, "setValue": "", "getValue": ""}],                  # Array, 6Byte Hex
    "CRC": {"size": 4, "setValue": "", "getValue": ""},                         # CRC-32 TODO: Implement CRC
}

sFileNameOut = create_string_buffer(b"./out.bin")
sFileNameIn = create_string_buffer(b"./in.bin")


# Labeled Choice box: Choice selection for user.
class UiDropBox(wx.BoxSizer):
    def __init__(self, panel, label, options):
        super(UiDropBox, self).__init__(wx.VERTICAL)
        self.label = wx.StaticText(panel, label=label)
        self.Add(self.label, flag=wx.TOP | wx.EXPAND, border=8)
        self.choice_control = wx.Choice(panel, choices=options)
        self.Add(self.choice_control, flag=wx.BOTTOM | wx.EXPAND, border=8)

    def get_value(self):
        return self.choice_control.GetStringSelection()

    def set_value(self, key_string):
        return self.choice_control.SetStringSelection(key_string)

    def set_selection(self, version_list):
        self.choice_control.Set(version_list)


# Labeled Text entry box: Data entry for user
class UiDate(wx.BoxSizer):
    def __init__(self, panel, label):
        super(UiDate, self).__init__(wx.VERTICAL)
        self.Add(wx.StaticText(panel, label=label), flag=wx.TOP | wx.EXPAND, border=8)
        self.date_control = wx_adv.DatePickerCtrl(panel)
        self.Add(self.date_control, flag=wx.BOTTOM | wx.EXPAND, border=8)

    def set_value(self, date_entry):
        set_date = wx.DateTime()
        year = int(date_entry[0:1].decode('utf-8')) + 2000
        month = int(date_entry[2:3].decode('utf-8'))
        day = int(date_entry[4:5].decode('utf-8'))
        set_date.Set(day, month, year=year)
        self.date_control.SetValue(set_date)

    def get_value(self):
        return self.date_control.GetValue()

    def get_value_bytes(self):
        # Get String representation of data in ISO format (YYYY-MM-DD) and reformat to form (YYMMDD)
        string_date = self.date_control.GetValue().FormatISODate().replace('-', '')[2:]
        return string_date.encode('utf-8')


# Labeled Text entry box: Data entry for user
class UiEntryBox(wx.BoxSizer):
    def __init__(self, panel, label, value):
        self.label_string = label
        super(UiEntryBox, self).__init__(wx.VERTICAL)
        self.label = wx.StaticText(panel, label=self.label_string)
        self.Add(self.label, flag=wx.TOP | wx.EXPAND, border=8)
        self.text_control = wx.TextCtrl(panel, value=value)
        self.Add(self.text_control, flag=wx.BOTTOM | wx.EXPAND, border=8)

    def set_visible(self, b_visible):
        if b_visible:
            self.Show(True)
            self.text_control.Show(True)
            self.label.Show(True)
        else:
            self.Hide(True)
            self.text_control.Hide()
            self.label.Hide()

    def set_value(self, string_entry):
        self.text_control.SetValue(string_entry)

    def get_value(self):
        return self.text_control.GetValue()


# Labeled Read-only text: Report values to user.
class UiReportBox(wx.BoxSizer):
    def __init__(self, panel, label, value):
        self.label_string = label
        super(UiReportBox, self).__init__(wx.VERTICAL)
        self.label = wx.StaticText(panel, label=self.label_string)
        self.Add(self.label, flag=wx.TOP | wx.EXPAND, border=8)
        self.static_text_control = wx.TextCtrl(panel, value=value, style=wx.TE_READONLY)
        self.Add(self.static_text_control, flag=wx.BOTTOM | wx.EXPAND, border=8)

    def set_value(self, string_entry):
        self.static_text_control.SetValue(string_entry)

    def get_value(self):
        return self.static_text_control.GetValue()

    def set_visible(self, b_visible):
        if b_visible:
            self.Show(True)
            self.static_text_control.Show(True)
            self.label.Show(True)
        else:
            self.Hide(True)
            self.static_text_control.Hide()
            self.label.Hide()

    def validate(self, cmp_string):
        validate_pass = self.static_text_control.GetValue() == cmp_string

        if validate_pass:
            self.static_text_control.SetBackgroundColour(wx.GREEN)
        else:
            self.static_text_control.SetBackgroundColour(wx.RED)

        self.static_text_control.Refresh()
        return validate_pass


class UiFrame(wx.Frame):
    def __init__(self, *args, **kw):
        super(UiFrame, self).__init__(*args, **kw)

        self.SetWindowStyle(wx.DEFAULT_FRAME_STYLE & ~(wx.MINIMIZE_BOX | wx.MAXIMIZE_BOX))

        self.SetSize(240, 400)
        self.SetSizeHints(240, 400, 240, 400)

        self.selected_lru = None
        self.selected_variant = None

        pnl = wx.Panel(self)
        sizer = wx.GridBagSizer(hgap=20, vgap=20)
        row_count = 0

        # LRU Serial Entry elements
        self.part_number = UiEntryBox(pnl, label="*Part Number:", value="827SPB004-300-C")
        self.board_serial = UiEntryBox(pnl, label="*Board Serial Entry:", value="MyBoard01")
        self.lru_serial = UiEntryBox(pnl, label="LRU Serial Entry:", value="")
        sizer.Add(self.part_number, pos=(row_count, 0), span=(1, 2), flag=wx.EXPAND)
        row_count += 1
        sizer.Add(self.board_serial, pos=(row_count, 0), span=(1, 2), flag=wx.EXPAND)
        row_count += 1
        sizer.Add(self.lru_serial, pos=(row_count, 0), span=(1, 2), flag=wx.EXPAND)
        row_count += 1

        # MAC Entry elements - generated IP address
        self.lru_mac = UiEntryBox(pnl, label="MAC Address Entry:", value="")
        #TODO Uncomment when IP generation method confirmed
        #self.lru_ip = UiReportBox(pnl, label="Generated IP Address:", value="-")
        sizer.Add(self.lru_mac, pos=(row_count, 0), span=(1, 2), flag=wx.EXPAND)
        #sizer.Add(self.lru_ip, pos=(row_count, 2), span=(1, 1), flag=wx.EXPAND)
        row_count += 1

        # Read/Write buttons
        self.button_read = wx.Button(pnl, label='Read')
        self.button_write = wx.Button(pnl, label='Write')
        #self.button_verify = wx.Button(pnl, label='Verify')
        sizer.Add(self.button_read, pos=(row_count, 0), flag=wx.EXPAND)
        sizer.Add(self.button_write, pos=(row_count, 1), flag=wx.EXPAND)
        #sizer.Add(self.button_verify, pos=(row_count, 2), flag=wx.EXPAND)
        self.Bind(wx.EVT_BUTTON, self.OnWrite, self.button_write)
        self.Bind(wx.EVT_BUTTON, self.OnRead, self.button_read)
        #self.Bind(wx.EVT_BUTTON, self.OnVerify, self.button_verify)
        row_count += 1

        pnl.SetSizer(sizer)


        self.dialog = wx.MessageDialog(pnl, "", "", style=wx.YES_NO)
        self.alert = wx.MessageDialog(pnl, "", "", style=wx.OK)


        # and a status bar
        self.CreateStatusBar()
        self.SetStatusText("Disconnected")




    def parse_mac_entry(self, mac_string):
        # standardise string by removing delimiters. May be entered without.
        mac_string = mac_string.replace('-', '')
        mac_string = mac_string.replace('.', '')

        # Length check, 12 characters (6 Hex bytes expected)
        if len(mac_string) != 12:
            self.SetStatusText("Invalid MAC Length")
            return False

        # Content check, Provided characters are valid HEX
        try:
            return_bytes = bytes.fromhex(mac_string)
        except ValueError:
            self.SetStatusText("Invalid MAC Format")
            return False

        # Reformat MAC string with delimiters '-', Standardise appearance of MAC regardless of form entered.
        mac_array = [mac_string[i:i + 2] for i in range(0, len(mac_string), 2)]
        mac_string = '-'.join(mac_array)
        self.lru_mac.set_value(mac_string)
        self.SetStatusText("MAC Good")

        self.generate_ip()

        return return_bytes

    def generate_ip(self):
        #TODO unlock function when IP greneration method confirmed.
        return
        char_array = self.lru_mac.get_value().split('-')
        int_array = [str(int("0x"+char_value, 16)) for char_value in char_array]
        self.lru_ip.set_value('172.' + '.'.join(int_array[-3:]))

    def parse_read_mac(self):
        mac_bytes = dict_eeprom_data["MACEntry"][0]["getValue"]
        mac_string = str(mac_bytes)
        mac_string_array = mac_string.split('\'')[1]
        mac_string_array2 = mac_string_array.split("\\x")
        self.lru_mac.set_value('-'.join(mac_string_array2[1:]))
        self.generate_ip()

    def OnVerify(self, event):
        func_read()

    def set_alert(self, title, message):
        self.alert.SetTitle(title)
        self.alert.SetMessage(message)
        self.alert.ShowModal()

    def set_yes_no_alert(self, title, message, proceed, stop):
        self.dialog.SetTitle(title)
        self.dialog.SetMessage(message)
        self.dialog.SetYesNoLabels(proceed, stop)
        return self.dialog.ShowModal() == wx.ID_YES

    def OnRead(self, event):
        if not self.set_yes_no_alert("Load from EEPROM", "Replace current form with data read from EEPROM?",
                                     "Replace", "Cancel"):
            return

        if func_connect_programmer():
            func_read()
        else:
            self.SetStatusText("No Programmer detected")

        if dict_eeprom_data["NamePlateRevision"]["getValue"] == b'\xFF\xFF':
            self.clear_form()
            self.set_alert("Read Error", "No data in eNameplate")
            return

        ui_object.set_alert("Read Success", "EEPROM Read Successfully")

        self.part_number.set_value(dict_eeprom_data["PartNumber"]["getValue"].decode("utf8"))
        self.lru_serial.set_value(dict_eeprom_data["LRUSerial"]["getValue"].decode("utf8"))
        self.board_serial.set_value(dict_eeprom_data["BoardSerial"]["getValue"].decode("utf8"))

        if dict_eeprom_data["MACcount"]["getValue"] != b'0':
            self.parse_read_mac()

    def clear_form(self):
        self.part_number.set_value("")
        self.board_serial.set_value("")
        self.lru_serial.set_value("")
        self.lru_mac.set_value("")

    def validate_entry(self, title, element, max_length, required=False):
        entry_string = element.get_value()
        if required and len(entry_string) == 0:
            self.SetStatusText("No " + title + " entered")

        if len(entry_string) > max_length:
            if self.set_yes_no_alert(title + " too long", "Entered " + title +
                                                          " too long (Max " + str(max_length) + " characters). \n"
                                                          + title + " MUST be be truncated", "Continue", "Cancel"):
                entry_string = entry_string[:max_length]
            else:
                self.SetStatusText(title + " too long")
                return False
        return entry_string

    def OnWrite(self, event):
        # Load Part number string from Part number entry control. Truncate to allowable size.
        pn_string = self.validate_entry("Part Number", self.part_number, dict_eeprom_data["PartNumber"]["size"],
                                        required=True)
        if pn_string is not False:
            dict_eeprom_data["PartNumber"]["setValue"] = pn_string.encode("utf-8")
        else:
            return

        # Load "BoardSerial" element from Board Serial entry control. Truncate to allowable size.
        board_serial_string = self.validate_entry("Board Serial", self.board_serial,
                                                  dict_eeprom_data["BoardSerial"]["size"], required=True)
        if board_serial_string is not False:
            dict_eeprom_data["BoardSerial"]["setValue"] = board_serial_string.encode("utf-8")
        else:
            return

        # Load "AssemblySerial" element from LRU Serial entry control. Truncate to allowable size.
        assm_serial_string = self.validate_entry("Lru Serial", self.lru_serial, dict_eeprom_data["LRUSerial"]["size"],
                                                 required=False)
        if assm_serial_string is not False:
            dict_eeprom_data["LRUSerial"]["setValue"] = assm_serial_string.encode("utf-8")
        else:
            dict_eeprom_data["LRUSerial"]["setValue"] = b''

        # Load MAC value string if provided and validate.
        mac_string = self.lru_mac.get_value()
        if len(mac_string) > 0:
            dict_eeprom_data["MACcount"]["setValue"] = b'1'
            mac_value = self.parse_mac_entry(mac_string)
            if not mac_value:
                return
            dict_eeprom_data["MACEntry"][0]["setValue"] = mac_value
        # If No MAC provided assume none loaded in this product.
        else:
            dict_eeprom_data["MACcount"]["setValue"] = b'0'

        if func_connect_programmer():
            func_program()
        else:
            self.set_alert("Write Error", "No Programmer detected")

    # Public setter function allowing external functions to set user messages to Status text bar
    def setStatusText(self, message):
        self.SetStatusText(message)


def func_load_block(buffer, block_num):
    origin = 0x20 + (0x40 * block_num)
    for buffer_index in range(0, len(buffer)):
        lib.E2_WriteBuffer(1, origin + buffer_index, buffer.raw[buffer_index])


def func_load_line(buffer, line_num):
    origin = (0x10 * line_num)
    for buffer_index in range(0, len(buffer)):
        lib.E2_WriteBuffer(1, origin + buffer_index, buffer.raw[buffer_index])


def func_connect_programmer():
    programmer_mode = 1  # E2M_COM = 1
    programmer_file_type = 9  # E2_BIN = 9
    programmer_speed = 1  # SP_MEDIUM = 1
    sComPort = create_string_buffer(10)
    ui_object.setStatusText("Connecting to Programmer")

    num_ports = c_byte(lib.E2_GetNumPorts(programmer_mode)).value
    for port_index in range(0, num_ports):
        try:
            lib.E2_GetPort(port_index, sComPort, programmer_mode)  # Get Com-port of connected device.
            # Device name hardcoded for 24C32 (4KB flash) to avoid known issue with EEPROM programmer device flashing only
            # Tail 4KB of binary in corrupt pattern if provided larger file size.
            lib.E2_InitProgrammer(sComPort, b'24C32', 1)
            if lib.E2_TestComms() == 0:
                ui_object.setStatusText("Programmer Connected")
                return True
        except:
            continue

    ui_object.setStatusText("Disconnected")
    return False


def eeprom_write_buffer(eeprom_item, buffer_index):
    eeprom_buffer = eeprom_item["setValue"]
    item_length = len(eeprom_buffer)
    item_length_allocation = eeprom_item["size"]

    for buffer_offset in range(0, item_length_allocation):
        if buffer_offset < item_length:
            lib.E2_WriteBuffer(1, buffer_index + buffer_offset, eeprom_buffer[buffer_offset])
        else:
            lib.E2_WriteBuffer(1, buffer_index + buffer_offset, 0)


def func_verify():
    pass


def func_read():

    lib.E2_PROG_Open()

    if lib.E2_ReadEEPROM(sFileNameIn, 9) != 0:
        ui_object.setStatusText("EEPROM Read Failed")
        lib.E2_PROG_Close()
        return False

    ui_object.setStatusText("EEPROM Read Completed")
    lib.E2_PROG_Close()

    file = open("./in.bin", "rb")
    buffer_index = 0

    # Parse Control data from EEPROM image.
    for eeprom_key in dict_eeprom_data:
        if eeprom_key == "MACEntry":
            for mac_index in range(0, mac_count):
                eeprom_item = dict_eeprom_data[eeprom_key][mac_index]
                eeprom_item["getValue"] = file.read(eeprom_item["size"]).split(b'\x10')[0]
                buffer_index += eeprom_item["size"]
        else:

            eeprom_item = dict_eeprom_data[eeprom_key]
            eeprom_item["getValue"] = file.read(eeprom_item["size"]).split(b'\x00')[0]
            buffer_index += eeprom_item["size"]

        if (eeprom_key == "MACcount"):
            if eeprom_item["getValue"] != b'\xff\xff\xff':
                mac_count = int(eeprom_item["getValue"])
            else:
                mac_count = 0

    file.close()
    return


def func_program():
    buffer_size = lib.E2_GetBufferSize(1)
    lib.E2_ClearBuffer(2)
    lib.E2_PROG_Open()
    lib.E2_SetResetPolarity(0)
    buffer_index = 0
    mac_count = 0

    # Parse Control data into EEPROM image.
    for eeprom_key in dict_eeprom_data:
        if eeprom_key == "MACEntry":
            for mac_index in range(0, mac_count):
                eeprom_item = dict_eeprom_data[eeprom_key][mac_index]
                eeprom_write_buffer(eeprom_item, buffer_index)
                buffer_index += eeprom_item["size"]
        else:
            eeprom_item = dict_eeprom_data[eeprom_key]
            eeprom_write_buffer(eeprom_item, buffer_index)
            buffer_index += eeprom_item["size"]

        if eeprom_key == "MACcount":
            mac_count = int(eeprom_item["setValue"])

    lib.E2_SaveBuffer(1, sFileNameOut, 9)

    ui_object.setStatusText("Starting program cycle")

    # Program Cycle
    if lib.E2_ProgramEEPROM(sFileNameOut) != 0:
        ui_object.set_alert("Write Error", "EEPROM Program Failed")
        lib.E2_PROG_Close()
        return False
    ui_object.setStatusText("EEPROM Program Completed")

    if lib.E2_ReadEEPROM(sFileNameIn, 9) != 0:
        ui_object.set_alert("Write Error", "EEPROM Read Failed")
        lib.E2_PROG_Close()
        return False
    ui_object.setStatusText("EEPROM Read Completed")

    # Verification Cycle
    if lib.E2_VerifyEEPROM(sFileNameOut) != 0:
        ui_object.set_alert("Write Error", "EEPROM Verification Failed")
        lib.E2_PROG_Close()
        return False
    ui_object.setStatusText("EEPROM Verification Completed")
    ui_object.set_alert("Write Success", "EEPROM Written Successfully")

    lib.E2_PROG_Close()
    return True


lib = windll.LoadLibrary("./E2ISP.DLL")
app = wx.App()
ui_object = UiFrame(None, title="RISE eNameplate Programmer")
ui_object.Show()
app.MainLoop()





