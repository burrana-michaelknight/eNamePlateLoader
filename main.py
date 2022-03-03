# This is a sample Python script.
import wx
from wx import adv as wx_adv
from ctypes import *
from time import sleep

dict_location = {
    "Mexicali": {"code": "MX"},
    "Utah": {"code": "UT"},
    "Brisbane": {"code": "BN"},
}

dict_lru = {
    "872SPB001": {"id": "872SPB001-3", "base_serial": "", "has_mac": True,
                  "variants": {
                      "300": {"code": "00", "versions": {"RevB": "B", "RevC": "C"}}}
                  },
    "872SPB002": {"id": "872SPB002-3", "base_serial": "", "has_mac": False,
                  "variants": {
                      "300": {"code": "00", "versions": {"RevA": "A", "RevB": "B"}}}
                  },
    "872SPB004": {"id": "872SPB004-3", "base_serial": "", "has_mac": False,
                  "variants": {
                      "300": {"code": "00", "versions": {"RevA": "A", "RevB": "B", "RevC": "C"}}}
                  },
}

dict_eeprom_data = {
    "NamePlateRevision": {"size": 2, "setValue": b'1', "getValue": ""},      # Hard-coded in this programmer
    "AssemblyId": {"size": 11, "setValue": "", "getValue": ""},              # 872xxx001
    "AssemblyVariant": {"size": 2, "setValue": "", "getValue": ""},         # A/B/C
    "AssemblyRevision": {"size": 2, "setValue": "", "getValue": ""},        # Number
    "AssemblySerial": {"size": 11, "setValue": "", "getValue": ""},         # User Serial
    "AssemblyManfDate": {"size": 6, "setValue": "", "getValue": ""},        # YYMMDD
    "AssemblyMFGSite": {"size": 2, "setValue": "", "getValue": ""},         # BN/MX/UT
    "AssemblyModLevel": {"size": 2, "setValue": "", "getValue": ""},        # Number
    "AssemblyModDate": {"size": 6, "setValue": "", "getValue": ""},        # YYMMDD
    "AssemblyModSite": {"size": 2, "setValue": "", "getValue": ""},         # BN/MX/UT
    "MACcount": {"size": 1, "setValue": b'0', "getValue": ""},               # Number 0-255
    "MACEntry": [{"size": 6, "setValue": "", "getValue": ""}],              # Array, 6Byte Hex
    "CRC": {"size": 4, "setValue": "", "getValue": ""},                     # CRC-32
    "ConfigVer": {"size": 1, "setValue": "", "getValue": ""},               # Number
    "ConfigLen": {"size": 2, "setValue": "", "getValue": ""}                # Number
}

# Labeled Choice box: Choice selection for user.
class UiDropBox(wx.BoxSizer):
    def __init__(self, panel, label, options):
        super(UiDropBox, self).__init__(wx.VERTICAL)
        self.label = wx.StaticText(panel, label=label)
        self.Add(self.label, flag=wx.TOP | wx.EXPAND, border=8)
        #self.value = options[0]
        self.choice_control = wx.Choice(panel, choices=options)
        self.Add(self.choice_control, flag=wx.BOTTOM | wx.EXPAND, border=8)

    def get_value(self):
        return self.choice_control.StringSelection

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

        self.selected_lru = None
        self.selected_variant = None

        pnl = wx.Panel(self)
        sizer = wx.GridBagSizer(hgap=20, vgap=20)
        row_count = 0

        # Hardware Selection
        self.lru_type_selection = UiDropBox(pnl, "LRU Type Selection:", list(dict_lru.keys()))
        self.lru_variant_selection = UiDropBox(pnl, "LRU Variant Selection:", [])
        self.lru_revision_selection = UiDropBox(pnl, "LRU Revision Selection:", [])

        self.Bind(wx.EVT_CHOICE, self.OnSelectLru, self.lru_type_selection.choice_control)
        self.Bind(wx.EVT_CHOICE, self.OnSelectVarient, self.lru_variant_selection.choice_control)
        self.Bind(wx.EVT_CHOICE, self.OnSelectVersion, self.lru_revision_selection.choice_control)

        sizer.Add(self.lru_type_selection, pos=(row_count, 0), flag=wx.EXPAND)
        sizer.Add(self.lru_variant_selection, pos=(row_count, 1), flag=wx.EXPAND)
        sizer.Add(self.lru_revision_selection, pos=(row_count, 2), flag=wx.EXPAND)
        row_count += 1

        # LRU Serial Entry elements
        self.lru_hw_id = UiReportBox(pnl, label="Hardware PN:", value="-")
        self.lru_serial = UiEntryBox(pnl, label="LRU Serial Entry:", value="")
        self.lru_board_serial = UiEntryBox(pnl, label="Board Serial Entry:", value="")

        sizer.Add(self.lru_hw_id, pos=(row_count, 0), span=(1, 1), flag=wx.EXPAND)
        sizer.Add(self.lru_serial, pos=(row_count, 1), span=(1, 1), flag=wx.EXPAND)
        sizer.Add(self.lru_board_serial, pos=(row_count, 2), span=(1, 1), flag=wx.EXPAND)
        row_count += 1

        # Manufacturer details
        self.manufacture_date = UiDate(pnl, label="Manufacture Date:")
        self.manufacture_loc = UiDropBox(pnl, "Manufactured in", list(dict_location.keys()))
        sizer.Add(self.manufacture_date, pos=(row_count, 0), span=(1, 1), flag=wx.EXPAND)
        sizer.Add(self.manufacture_loc, pos=(row_count, 1), flag=wx.EXPAND)
        row_count += 1

        # MAC Entry elements
        self.lru_mac = UiEntryBox(pnl, label="MAC Address Entry:", value="")
        sizer.Add(self.lru_mac, pos=(row_count, 0), span=(1, 2), flag=wx.EXPAND)
        #row_count += 1

        # Report generated IP address
        self.lru_ip = UiReportBox(pnl, label="Generated IP Address:", value="-")
        sizer.Add(self.lru_ip, pos=(row_count, 2), span=(1, 2), flag=wx.EXPAND)
        row_count += 1

        # Read/Write buttons
        self.button_write = wx.Button(pnl, label='Write')
        self.button_read = wx.Button(pnl, label='Read')
        self.button_verify = wx.Button(pnl, label='Verify')
        sizer.Add(self.button_write, pos=(row_count, 0), flag=wx.EXPAND)
        sizer.Add(self.button_read, pos=(row_count, 1), flag=wx.EXPAND)
        sizer.Add(self.button_verify, pos=(row_count, 2), flag=wx.EXPAND)
        self.Bind(wx.EVT_BUTTON, self.OnWrite, self.button_write)
        self.Bind(wx.EVT_BUTTON, self.OnRead, self.button_read)
        self.Bind(wx.EVT_BUTTON, self.OnVerify, self.button_verify)
        row_count += 1

        pnl.SetSizer(sizer)

        self.dialog = wx.MessageDialog(pnl, "", "", style=wx.YES_NO)

        # and a status bar
        self.CreateStatusBar()
        self.SetStatusText("Disconnected")
        self.SetSize(self.GetBestSize())

    def validate_all(self, event):
        validate_pass = True
        validate_pass &= self.lru_read_serial.validate(self.lru_serial.get_value())
        validate_pass &= self.lru_read_hw_id.validate(self.lru_hw_id.get_value())
        validate_pass &= self.lru_read_mac.validate(self.lru_mac.get_value())
        if validate_pass:
            self.SetStatusText("Validation Passed")
        else:
            self.SetStatusText("Validation Failed")

    def makeMenuBar(self):
        fileMenu = wx.Menu()

        fileMenu.AppendSeparator()
        helloItem = fileMenu.Append(-1, "&Hello...\tCtrl-H",
                                    "Help string shown in status bar for this menu item")
        exitItem = fileMenu.Append(wx.ID_EXIT)

        helpMenu = wx.Menu()
        aboutItem = helpMenu.Append(wx.ID_ABOUT)

        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, "&File")
        menuBar.Append(helpMenu, "&Help")

        self.SetMenuBar(menuBar)

        self.Bind(wx.EVT_MENU, self.OnHello, helloItem)
        self.Bind(wx.EVT_MENU, self.OnExit,  exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

    def SetIpVisibile(self, bSet):
        self.lru_mac.set_visible(bSet)
        #self.lru_read_mac.set_visible(bSet)
        self.lru_ip.set_visible(bSet)

    def OnSelectLru(self, event):
        new_lru = dict_lru[self.lru_type_selection.get_value()]
        self.selected_lru = new_lru

        self.UpdateVarientSelect()
        self.UpdateRevisionSelect()

        # Update form appearence to Hide or expose Mac/IP fields as required.
        self.SetIpVisibile(new_lru["has_mac"])

        self.hw_id_update()

    def OnSelectVarient(self, event):
        new_varient = self.selected_lru["variants"][self.lru_variant_selection.get_value()]
        self.selected_variant = new_varient

        self.UpdateRevisionSelect()
        self.hw_id_update()

    def OnSelectVersion(self, event):
        self.hw_id_update()

    def UpdateVarientSelect(self):
        self.lru_variant_selection.set_selection(list(self.selected_lru["variants"].keys()))
        self.lru_variant_selection.choice_control.Select(len(self.lru_variant_selection.choice_control.GetItems())-1)
        self.selected_variant = self.selected_lru["variants"][self.lru_variant_selection.get_value()]

    def UpdateRevisionSelect(self):
        self.lru_revision_selection.set_selection(list(self.selected_variant["versions"].keys()))
        self.lru_revision_selection.choice_control.Select(len(self.lru_revision_selection.choice_control.GetItems())-1)

    def hw_id_update(self):
        base_string = "!-@ #"
        base_string = base_string.replace("!", self.lru_type_selection.get_value())
        base_string = base_string.replace("@", self.lru_variant_selection.get_value())
        base_string = base_string.replace('#', self.lru_revision_selection.get_value())
        try:
            self.lru_hw_id.set_value(base_string)

        except KeyError:
            self.SetStatusText("Invalid LRU")

    def parse_mac_entry(self):
        mac_string = self.lru_mac.get_value()

        # standardise string by removing delimiters. May be entered without.
        mac_string = mac_string.replace('-', '')
        mac_string = mac_string.replace('.', '')

        # Length check
        if len(mac_string) != 12:
            self.SetStatusText("Invalid MAC Length")
            return False

        # Content check
        try:
            return_bytes = bytes.fromhex(mac_string)
        except ValueError:
            self.SetStatusText("Invalid MAC Format")
            return False

        mac_array = [mac_string[i:i + 2] for i in range(0, len(mac_string), 2)]

        mac_string = '-'.join(mac_array)
        self.lru_mac.set_value(mac_string)
        self.SetStatusText("MAC Good")

        ip_array = [str(int("0x"+mac_array[3+index], 16)) for index in range(0, 3)]
        self.lru_ip.set_value('172.'+'.'.join(ip_array))

        return return_bytes

    def OnVerify(self, event):
        func_read()

    def OnRead(self, event):
        self.dialog.SetTitle("Load from EEPROM")
        self.dialog.SetMessage("Replace current form with data read from EEPROM?")
        self.dialog.SetYesNoLabels("Replace", "Cancel")
        if self.dialog.ShowModal() != wx.ID_YES:
            return

        func_read()

        # Parse EEPROM_data dict into form controls
        # Find LRU from code
        lru_read_string = dict_eeprom_data["AssemblyId"]["getValue"].decode("utf8")
        lru_key_list = [lru_key for lru_key in dict_lru if dict_lru[lru_key]["id"] == lru_read_string]
        if (len(lru_key_list) == 0) or not self.lru_type_selection.set_value(lru_key_list[0]):
            self.setStatusText("LRU Type not recognised")
            return
        self.OnSelectLru(None)

        # Find variant from code
        var_read_string = dict_eeprom_data["AssemblyVariant"]["getValue"].decode("utf8")
        var_key_list = [var_key for var_key in self.selected_lru['variants']
                        if self.selected_lru['variants'][var_key]["code"] == var_read_string]
        if (len(var_read_string) == 0) or not self.lru_variant_selection.set_value(var_key_list[0]):
            self.setStatusText("LRU Variant not recognised")
            return
        self.OnSelectVarient(None)

        # Find Rev from code
        rev_read_string = dict_eeprom_data["AssemblyRevision"]["getValue"].decode("utf8")
        var_key_list = [rev_key for rev_key in self.selected_variant["versions"] if self.selected_variant["versions"][rev_key] == rev_read_string]
        if (len(var_key_list) == 0) or not self.lru_revision_selection.set_value(var_key_list[0]):
            self.setStatusText("LRU Revision not recognised")
            return
        self.OnSelectVersion(None)

        self.lru_serial.set_value(dict_eeprom_data["AssemblySerial"]["getValue"].decode("utf8"))




    def OnWrite(self, event):

        # Load "Assembly ID" element from LRU selection control
        if self.selected_lru is None:
            self.SetStatusText("No LRU selected")
            return
        else:
            hw_lru_string = self.selected_lru["id"].split("%")[0][:dict_eeprom_data["AssemblyId"]["size"]]
            dict_eeprom_data["AssemblyId"]["setValue"] = self.selected_lru["id"].split("%")[0].encode("utf-8")

        # Load "AssemblyVarient" element from LRU selection control
        hw_variant_key = self.lru_variant_selection.get_value()
        hw_varient_string = self.selected_lru["variants"][hw_variant_key]["code"][:dict_eeprom_data["AssemblyVariant"]["size"]]
        dict_eeprom_data["AssemblyVariant"]["setValue"] = hw_varient_string.encode("utf-8")

        # Load "AssemblyRevision" element from LRU Serial entry control
        hw_rev_key = self.lru_revision_selection.get_value()
        assem_revision_string = self.selected_variant["versions"][hw_rev_key][:dict_eeprom_data["AssemblyRevision"]["size"]]
        dict_eeprom_data["AssemblyRevision"]["setValue"] = assem_revision_string.encode("utf-8")

        # Load "AssemblySerial" element from LRU Serial entry control
        assem_serial_string = self.lru_serial.get_value()[:dict_eeprom_data["AssemblySerial"]["size"]]
        if len(assem_serial_string) == 0:
            self.SetStatusText("No LRU Serial entered")
            return
        else:
            dict_eeprom_data["AssemblySerial"]["setValue"] = assem_serial_string.encode("utf-8")

        # Load "AssemblyManfDate" element from LRU Manf Date control
        assem_serial_bytes = self.manufacture_date.get_value_bytes()[:dict_eeprom_data["AssemblyManfDate"]["size"]]
        dict_eeprom_data["AssemblyManfDate"]["setValue"] = assem_serial_bytes

        # Load "AssemblyMFGSite" element from Assm Site select control
        assem_loc_key = self.manufacture_loc.get_value()
        dict_eeprom_data["AssemblyMFGSite"]["setValue"] = \
            dict_location[assem_loc_key]["code"].encode("utf-8")[:dict_eeprom_data["AssemblyMFGSite"]["size"]]

        mac_value = None
        if self.selected_lru['has_mac']:
            dict_eeprom_data["MACcount"]["setValue"] = b'1'
            mac_value = self.parse_mac_entry()
            if not mac_value:
                return
            dict_eeprom_data["MACEntry"][0]["setValue"] = mac_value
        else:
            dict_eeprom_data["MACcount"]["setValue"] = b'0'

        if func_connect_programmer():
            func_program()
        else:
            self.SetStatusText("No Programmer detected")

    def OnExit(self, event):
        """Close the frame, terminating the application."""
        self.Close(True)

    def OnHello(self, event):
        """Say hello to the user."""
        wx.MessageBox("Hello World")

    def OnAbout(self, event):
        """Display an About Dialog"""
        wx.MessageBox("This is a wxPython Hello World sample",
                      "About Hello World 2",
                      wx.OK|wx.ICON_INFORMATION)

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
    programmer_file_type = 11  # E2_BIN = 9
    programmer_speed = 1  # SP_MEDIUM = 1
    sComPort = create_string_buffer(10)
    sDevName = create_string_buffer(30)

    if c_byte(lib.E2_GetNumPorts(programmer_mode)).value == 0:
        return False
    else:
        lib.E2_GetPort(0, sComPort, programmer_mode)  # Get Com-port of connected device.
        lib.E2_GetDevice(0, 17, sDevName)  # Get DevName of connected device.
        return lib.E2_InitProgrammer(b'COM6', b'24C32', 1) == 0


def eeprom_write_buffer(eeprom_item, buffer_index):
    eeprom_buffer = eeprom_item["setValue"]
    item_length = len(eeprom_buffer)
    item_length_allocation = eeprom_item["size"]

    for buffer_offset in range(0, item_length_allocation):
        if buffer_offset < item_length:
            lib.E2_WriteBuffer(1, buffer_index + buffer_offset, eeprom_buffer[buffer_offset])
        else:
            lib.E2_WriteBuffer(1, buffer_index + buffer_offset, 0)


def func_read():
    #TODO: Read EEPROM into file.

    sFileNameIn = create_string_buffer(b"./out.bin")

    file = open("./out.bin", "rb")
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
            eeprom_item["getValue"] = file.read(eeprom_item["size"]).split(b'\x10')[0]
            buffer_index += eeprom_item["size"]

        if eeprom_key == "MACcount":
            mac_count = int(eeprom_item["getValue"])
    return


def func_program():
    sFileNameOut = create_string_buffer(b"./out.bin")
    sFileNameIn = create_string_buffer(b"./in.bin")

    buffer_size = lib.E2_GetBufferSize(1)
    lib.E2_ClearBuffer(1)
    lib.E2_PROG_Open()
    lib.E2_SetResetPolarity(0)
    buffer_index = 0
    mac_count = 0

    # Parse Control data into EEPROM image.
    for eeprom_key in dict_eeprom_data:
        if eeprom_key == "MACEntry":
            for mac_index in range(0, mac_count):
                eeprom_item = dict_eeprom_data[eeprom_key][mac_index]
                eeprom_write_buffer(eeprom_item,buffer_index)
                buffer_index += eeprom_item["size"]
        else:
            eeprom_item = dict_eeprom_data[eeprom_key]
            eeprom_write_buffer(eeprom_item, buffer_index)
            buffer_index += eeprom_item["size"]

        if eeprom_key == "MACcount":
            mac_count = int(eeprom_item["setValue"])

    lib.E2_SaveBuffer(1, sFileNameOut, 9)

    try:
        ui_object.setStatusText("Start program cycle")

        # Erase Cycle
        if lib.E2_EraseAllDevice() != 0:
            ui_object.setStatusText("EEPROM Erasure Failed")
            lib.E2_PROG_Close()
            return False
        ui_object.setStatusText("EEPROM Erasure Completed")

        # Program Cycle
        if lib.E2_EraseAllDevice(lib.E2_ProgramEEPROM(sFileNameOut)) != 0:
            ui_object.setStatusText("EEPROM Program Failed")
            lib.E2_PROG_Close()
            return False
        ui_object.setStatusText("EEPROM Program Completed")

    except OSError:
        pass

    lib.E2_PROG_Close()
    return True

lib = windll.LoadLibrary("./E2ISP.DLL")
app = wx.App()
ui_object = UiFrame(None, title="RISE eNameplate Programmer")
ui_object.Show()
app.MainLoop()





