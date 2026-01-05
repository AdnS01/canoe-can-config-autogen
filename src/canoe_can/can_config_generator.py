import os
import re
import openpyxl
from pathlib import Path


class CAN_CANoeConfig:

    def __init__(self):
        self.filePathDBC = ""
        self.filePathExcel = ""
        self.output_root = ""
        self.repo_root = Path(__file__).resolve().parents[2]
        self.find_input_paths()
        
        self.Signals = {}
        self.Messages = {}
        self.Nodes = []
        self.CSIC = {}
        self.LSBsignal = 0
        self.MSBsignal = 0


    def find_input_paths(self):
        input_root = self.repo_root / Path("inputs") / Path("can")
        dbc_files = sorted((input_root / "dbc").glob("*.dbc"))
        excel_files = sorted((input_root / "excel").glob("*.xlsx"))
        if not dbc_files or not excel_files:
            raise FileNotFoundError(
                f"Expected .dbc and .xlsx under {input_root}/dbc and {input_root}/excel"
            )
        self.filePathDBC = dbc_files[0]
        self.filePathExcel = excel_files[0]

        self.output_root = self.repo_root / Path("outputs") / Path("can")
        self.output_root.mkdir(parents=True, exist_ok=True)

        print("\n[CAN] CANoe Configuration Auto-Generation")


    # Printing All Messages with their parameters
    def Print_Messages(self):
        # Parse the DBC file
        self.DBCparser()
        # Extract all signals with their parameters
        for key, data in self.Messages.items():
            print("Message  : ", key, "\n",
                  "  Sender 1         :   ", data['sender'][0], "\n",
                  "  Sender 2         :   ", data['sender'][1], "\n",
                  "  ID               :   ", hex(data['id']), "\n",
                  "  Cyclicity        :   ", data['cyclicity'], "\n\n")


    # Printing All Messages with their parameters
    def Print_Messages_Signals(self):
        # Parse the DBC file
        self.DBCparser()
        # Extract all signals with their parameters
        for key, data in self.Messages.items():
            print("Message  : ", key, "\n",
                  "-----------------------\n",
                  "  Sender 1         :   ", data['sender'][0], "\n",
                  "  Sender 2         :   ", data['sender'][1], "\n",
                  "  ID               :   ", hex(data['id']), "\n",
                  "  Cyclicity        :   ", data['cyclicity'], "\n",
                  "-----------------------\n",
                  "  Signals        :   ")
            sorted_signals = sorted(data['signals'].items(), key=lambda item: item[1]['LSB'])
            for keySG, dataSG in sorted_signals:
                print(
                    "\t\t", keySG, " :\n",
                    "\t\t\tLSB    :   ", dataSG['LSB'], "\n",
                    "\t\t\tMSB    :   ", dataSG['MSB'], "\n",
                    "\t\t\tLength    :   ", dataSG['length'], "\n",
                    "\t\t\tCheckSumInit    :   ", dataSG['CheckSumInit'], "\n",
                    "\t\t\tInit    :   ", dataSG['init'], "\n",
                    "\t\t\tMin    :   ", dataSG['min'], "\n",
                    "\t\t\tMax    :   ", dataSG['max'], "\n",
                    "\t\t\tVal_Desc : ", dataSG['desc'], "\n",
                    "\t\t-----------------------")


    # Check if the file exists by checking if the path exists or not
    def Check_file_exists(self):
        if os.path.exists(self.filePathDBC):
            print('The file exists\n')
        else:
            print('The specified file does NOT exist\n')


    def Byte_Position(self, startBit, Length):
        LSBbyte = 0
        MSBbyte = 0
        if 0 <= startBit <= 7:
            LSBbyte = 0
            MSBbyte = 7
        elif 8 <= startBit <= 15:
            LSBbyte = 8
            MSBbyte = 15
        elif 16 <= startBit <= 23:
            LSBbyte = 16
            MSBbyte = 23
        elif 24 <= startBit <= 31:
            LSBbyte = 24
            MSBbyte = 31
        elif 32 <= startBit <= 39:
            LSBbyte = 32
            MSBbyte = 39
        elif 40 <= startBit <= 47:
            LSBbyte = 40
            MSBbyte = 47
        elif 48 <= startBit <= 55:
            LSBbyte = 48
            MSBbyte = 55
        elif 56 <= startBit <= 63:
            LSBbyte = 56
            MSBbyte = 63

        self.LSBsignal = (MSBbyte - startBit) + LSBbyte
        self.MSBsignal = self.LSBsignal + Length - 1


    # Extract CheckSumInit Const From Excel File
    def Extract_CSIC(self):
        # Open Excel file
        EXfile = openpyxl.load_workbook(self.filePathExcel)
        # Select the SIGNALS sheet
        Sheet = EXfile['SIGNALS']
        # Extract CheckSumInit Const
        for row in Sheet:
            # Signal name : row[0] && CheckSumInit Const : row[32]
            self.CSIC[row[0].value] = row[32].value


    # Extract all messages with their parameters
    def DBCparser(self):
        Message_Name = ""
        # Strings to search
        BU = re.compile("^BU_:")
        BO = re.compile("^BO_ ")
        SG = re.compile("^ SG_ ")
        MT = re.compile("^BO_TX_BU_ ")
        Cy = re.compile('^BA_ "Periode" BO_ ')
        BA = re.compile('^BA_ "GenSigStartValue" SG_ ')
        VAL = re.compile("^VAL_ ")
        # Extract CheckSum Const && Byte position
        self.Extract_CSIC()
        # Parse the DBC File
        with open(self.filePathDBC, 'r') as DBCfile:
            for line in DBCfile:
                if BU.match(line):
                    self.Nodes = line.split()
                    self.Nodes.pop(0)
                    self.Nodes.remove("BCM")
                # Extract frames with their parameters
                elif BO.match(line):
                    BO_id = re.search(r'BO_\s*(\d+)', line)
                    Id = int(BO_id.group(1))
                    BO_Message = re.search(r'\b\d+\s+(.+?):', line)
                    Message_Name = BO_Message.group(1).strip()
                    BO_Sender = re.search(r'\:\s+\d+\s+(\w+)', line)
                    Sender = BO_Sender.group(1).strip()
                    self.Messages[Message_Name] = {'sender': [], 'signals': {}, 'cyclicity': 0, 'id': Id}
                    self.Messages[Message_Name]['sender'].append(Sender)
                    self.Messages[Message_Name]['sender'].append(None)
                elif SG.match(line):
                    # Extract signal name
                    SG_name = re.search(r'\b\s+(.+?)\s+', line)
                    tempName = SG_name.group(1).strip()
                    # check if the var is a chekcksum var
                    if re.search("CHKSUM", tempName):
                        tempCSIC = self.CSIC[tempName]
                    else:
                        tempCSIC = None
                    # Extract Parameters startbit and length
                    SG_B_L = re.search(r'\:\s+(\d+)\|(\d+)\@', line)
                    tempstartbit = int(SG_B_L.group(1))
                    tempLength = int(SG_B_L.group(2))
                    # Extract Parameters facter and Offset
                    SG_F_O = re.search(r'\((.+?)\,(.+?)\)', line)
                    tempFact = float(SG_F_O.group(1))
                    tempOffset = float(SG_F_O.group(2))
                    # Extract Parameters Max and Min and (type for EnVar Syntax)
                    SG_MinMax_int = re.search(r'\[([+-]?\d+)\|([+-]?\d+)\]', line)
                    SG_MinMax_str = re.search(r'\[(.+?)\|(.+?)\]', line)
                    if SG_MinMax_int:
                        tempMin = SG_MinMax_int.group(1)
                        tempMax = SG_MinMax_int.group(2)
                        tempType = 0
                    else:
                        tempMin = float(SG_MinMax_str.group(1))
                        tempMax = float(SG_MinMax_str.group(2))
                        tempType = 1
                    # Add LSB to Signal
                    self.Byte_Position(tempstartbit, tempLength)
                    # Fill the Signals Dictionary
                    self.Messages[Message_Name]['signals'][tempName] = {'LSB': self.LSBsignal, 'MSB': self.MSBsignal, 'startbit': tempstartbit, 'length': tempLength, 'fact': tempFact, 'offset': tempOffset, 'min': tempMin, 'max': tempMax,
                                         'type': tempType, 'init': 0.0, 'CheckSumInit': tempCSIC, 'desc': "none"}
                # Extract Transmitters
                elif MT.match(line):
                    MT_id = re.search(r'BO_TX_BU_\s*(\d+)\s*:', line)
                    Trans = re.search(r'\:\s*(.+?)\,(.+?)\;', line)
                    for key, data in self.Messages.items():
                        if self.Messages[key]['id'] == int(MT_id.group(1).strip()):
                            data['sender'][0] = Trans.group(1).strip()
                            data['sender'][1] = Trans.group(2).strip()
                # Extract Cyclicity parameter
                elif Cy.match(line):
                    BA_id = re.search(r'BO_\s*(\d+)', line)
                    BA_cy = re.search(r'BO_\s*\d+\s*(\d+)', line)
                    for key, data in self.Messages.items():
                        if self.Messages[key]['id'] == int(BA_id.group(1)):
                            self.Messages[key]['cyclicity'] = int(BA_cy.group(1))
                # Extract Init parameter and Calculate its physical value
                elif BA.match(line):
                    BA_id = re.search(r'SG_\s+(\d+)\s+', line)
                    BA_name = re.search(r'\d+\s+(.+?)\s+\d+;', line)
                    tempBAName = BA_name.group(1).strip()
                    BA_Init = re.search(r'\s+(\d+);', line)
                    tempInit = float(BA_Init.group(1).strip())
                    for key, data in self.Messages.items():
                        if self.Messages[key]['id'] == int(BA_id.group(1)):
                            signal_data = self.Messages[key]['signals'][tempBAName]
                            offset = signal_data['offset']
                            fact = signal_data['fact']
                            signal_data['init'] = (tempInit * fact) + offset
                # Extract Desc parameter for Signals that have Value Table
                elif VAL.match(line):
                    BA_id = re.search(r'VAL_\s+(\d+)\s+', line)
                    VAL_name = re.search(r'\b\d+\s+(.+?)\s+\b', line)
                    tempValName = VAL_name.group(1).strip()
                    VAL_Desc = re.search(r'\d\s+(.*)', line)
                    tempDesc = VAL_Desc.group(1).strip()
                    if BA_id:
                        for key, data in self.Messages.items():
                            if self.Messages[key]['id'] == int(BA_id.group(1)):
                                self.Messages[key]['signals'][tempValName]['desc'] = tempDesc
            # Stop Parsing
            DBCfile.close()


    # Create Environment Variables File Declaration && Environment Variables File Value Descriptions
    def Generate_EnVars(self):
        print("\n[CAN] Generating Environment Variables")
        envars_root = self.output_root / Path("envars")
        envars_root.mkdir(parents=True, exist_ok=True)
        EnVarDecPath = envars_root / 'EV_Dec.txt'
        EnVarValDescPath = envars_root / 'EV_ValDesc.txt'
        # Create Environment Variables File Declaration
        self.Generate_EnVars_Dec(EnVarDecPath)
        # Create Environment Variables File Value Descriptions
        self.Generate_EnVars_ValDesc(EnVarValDescPath)


    # Create Environment Variables File Declaration from Message Signals
    def Generate_EnVars_Dec(self, fileName):
        # Extract all messages with their parameters
        self.DBCparser()
        i = 0
        # Create the Environment Variables
        with open(fileName, 'w') as EnVarDec:
            for key, data in self.Messages.items():
                sorted_signals = sorted(data['signals'].items(), key=lambda item: item[1]['startbit'])
                for keySG, dataSG in sorted_signals:
                    i = i + 1
                    EnVarLine = 'EV_ ENV_' + str(hex(data['id'])) + '_' + keySG + ' : ' + str(dataSG['type']) + ' [' + str(dataSG['min']) + '|' + str(
                        dataSG['max']) + '] "" ' + str(dataSG['init']) + ' ' + str(i) + ' DUMMY_NODE_VECTOR0 Vector__XXX; \n'
                    EnVarDec.write(EnVarLine)
                    if re.search("CHKSUM", keySG):
                        i = i + 1
                        EnVarDec.write('EV_ ENV_' + str(hex(data['id'])) + '_CRC_FAULT : 0 [0|1] "" 0 ' + str(i) + ' DUMMY_NODE_VECTOR0 Vector__XXX; \n')
            # Stop Creating
            EnVarDec.close()
        print("[CAN] Environment Variables declaration file created")



    # Create Environment Variables File Value Descriptions from Message Signals
    def Generate_EnVars_ValDesc(self, fileName):
        # Extract all messages with their parameters
        self.DBCparser()
        # Create the Value Descriptions of the Environment Variables
        with open(fileName, 'w') as EnVarValDesc:
            for key, data in self.Messages.items():
                sorted_signals = sorted(data['signals'].items(), key=lambda item: item[1]['startbit'])
                for keySG, dataSG in sorted_signals:
                    if dataSG['desc'] != "none":
                        EnVarLine = "VAL_ ENV_" + str(hex(data['id'])) + '_' + dataSG['desc'] + " \n"
                        EnVarValDesc.write(EnVarLine)
            # Stop Creating
            EnVarValDesc.close()
        print("[CAN] Environment Variables value descriptions created")


    # Generate the CAPL code without E2E algorithm 
    def Generate_CAPL(self):
        capl_root = self.output_root / Path("capl")
        capl_root.mkdir(parents=True, exist_ok=True)
        # Extract all messages with their parameters
        self.DBCparser()
        print("\n[CAN] Generating CAPL code")
        # Create CAPL Code File for each virtual node
        for node in self.Nodes:
            fileName = node + ".can"
            filePath = capl_root / Path(fileName)
            with open(filePath, 'w') as File:
                # Add file name to a table
                line = '///////////////////// Code CAPL for : ' + node + r' \\\\\\\\\\\\\\\\\\\\\\' + '\n\n'
                File.write(line)
                # Variables {}
                File.write('variables\n')
                File.write('{\n')
                for key, data in self.Messages.items():
                    if node == data['sender'][0] or node == data['sender'][1]:
                        File.write('\n')
                        line = '\tmessage ' + str(hex(data['id'])) + ' ' + '_m' + key + ';\n'
                        File.write(line)
                        if data['cyclicity'] != 0:
                            line = '\tmsTimer _t' + key + ';\n'
                            File.write(line)
                            line = '\tint _' + key + 'CycleTime = ' + str(data['cyclicity']) + ';\n'
                            File.write(line)
                File.write('\n}\n\n\n')
                # On start {}
                File.write('On start \n')
                File.write('{\n')
                for key, data in self.Messages.items():
                    if node == data['sender'][0] or node == data['sender'][1]:
                        if data['cyclicity'] != 0:
                            File.write('\n')
                            line = '\tsetTimer(_t' + key + ', _' + key + 'CycleTime);\n'
                            File.write(line)
                        for SG in data['signals']:
                            line = '\t_m' + key + '.' + SG + '.phys = getValue(ENV_' + str(hex(data['id'])) + '_' + SG + ');\n'
                            File.write(line)
                File.write('\n}\n\n\n')
                # On Timer ... {}
                for key, data in self.Messages.items():
                    if (node == data['sender'][0] or node == data['sender'][1]) and (data['cyclicity'] != 0):
                        line = 'On Timer _t' + key + '\n'
                        File.write(line)
                        File.write('{\n')
                        line = '\toutput(_m' + key + ');\n'
                        File.write(line)
                        line = '\tsetTimer(_t' + key + ', _' + key + 'CycleTime);\n'
                        File.write(line)
                        File.write('}\n\n\n')
                # On EnvVar ... {}
                for key, data in self.Messages.items():
                    if node == data['sender'][0] or node == data['sender'][1]:
                        for SG in data['signals']:
                            line = 'On EnvVar ENV_' + str(hex(data['id'])) + '_' + SG + '\n'
                            File.write(line)
                            File.write('{\n')
                            line = '\t_m' + key + '.' + SG + '.phys = getValue(this);\n'
                            File.write(line)
                            File.write('}\n\n')
                # Stop Generating
                File.close()
        print("[CAN] CAPL code generated for all virtual CAN nodes")



    # Generate the CAPL code with E2E algorithm
    def Generate_CAPL_E2E(self):
        capl_e2e_root = self.output_root / Path("capl_e2e")
        capl_e2e_root.mkdir(parents=True, exist_ok=True)
        # Extract all messages with their parameters
        self.DBCparser()
        print("\n[CAN] Generating CAPL code (E2E enabled)")
        # Create CAPL Code File for each virtual node
        for node in self.Nodes:
            fileName = node + ".can"
            filePath = capl_e2e_root / Path(fileName)
            with open(filePath, 'w') as File:
                # Add file name to a table
                File.write('///////////////////// Code CAPL for : ' + node + r' \\\\\\\\\\\\\\\\\\\\\\' + '\n\n')
                # -----------------------------------------------------------------------------------------------
                # Variables {}
                File.write('variables\n')
                File.write('{\n')
                for key, data in self.Messages.items():
                    if node == data['sender'][0] or node == data['sender'][1]:
                        File.write('\n')
                        line = '\tmessage ' + str(hex(data['id'])) + ' ' + '_m' + key + ';\n'
                        File.write(line)
                        if data['cyclicity'] != 0:
                            line = '\tmsTimer _t' + key + ';\n'
                            File.write(line)
                            line = '\tint _' + key + 'CycleTime = ' + str(data['cyclicity']) + ';\n'
                            File.write(line)
                            # If CHKSUM exist, declare it as a global variable
                            for keySG, dataSG in data['signals'].items():
                                if re.search("CHKSUM", keySG):
                                    File.write('\tint temp_' + str(hex(data['id'])) + ' = 0;\n')
                                    File.write('\tbyte ' + keySG + ';\n')
                                elif re.search("PROCESS", keySG):
                                    File.write('\tbyte ' + keySG + ';\n')
                File.write('\n}\n\n\n')
                # --------------------------------------------------------------------------------------------
                # On start {}
                File.write('On start \n')
                File.write('{\n')
                for key, data in self.Messages.items():
                    if node == data['sender'][0] or node == data['sender'][1]:
                        if data['cyclicity'] != 0:
                            File.write('\n')
                            line = '\tsetTimer(_t' + key + ', _' + key + 'CycleTime);\n'
                            File.write(line)
                        sorted_signals = sorted(data['signals'].items(), key=lambda item: item[1]['startbit'])
                        for keySG, dataSG in sorted_signals:
                            File.write('\t_m' + key + '.' + keySG + '.phys = getValue(ENV_' + str(hex(data['id'])) + '_' + keySG + ');\n')
                File.write('\n}\n\n\n')
                # --------------------------------------------------------------------------------------------
                # On Timer ... {}
                for key, data in self.Messages.items():
                    if (node == data['sender'][0] or node == data['sender'][1]) and (data['cyclicity'] != 0):
                        File.write("\nOn Timer _t" + key + "\n{\n")
                        sorted_signals = sorted(data['signals'].items(), key=lambda item: item[1]['LSB'])
                        for keySG1, dataSG1 in sorted_signals:
                            # E2E Mode
                            if dataSG1['CheckSumInit'] is not None:
                                # Variables Declaration
                                CheckSumString = ""
                                CounterString = ""
                                File.write('\tint sum;\n\tint mod;\n\tint i;\n\n')
                                File.write("\tbyte sig[16];\n\n")
                                for keySG2, dataSG2 in sorted_signals:
                                    if dataSG2['length'] <= 8:
                                        File.write("\tbyte " + keySG2 + ";\n")
                                    elif 8 < dataSG2['length'] <= 16:
                                        File.write("\tword " + keySG2 + ";\n")
                                    else:
                                        File.write("\tdword " + keySG2 + ";\n")
                                File.write('\n\tsum = 0x00;\n\tmod = 0x00;\n\ti = 0;\n\n')
                                File.write('\tfor (i=0; i < 16; i++){\n\t\tsig[i] = 0;\n\t}\n\n')
                                # --------------------------------------------------------------------------------
                                for keySG2, dataSG2 in sorted_signals:
                                    if re.search("CHKSUM", keySG2):
                                        File.write("")
                                    elif re.search("PROCESS", keySG2):
                                        File.write("")
                                    else:
                                        File.write("\t" + str(keySG2) + " = _m" + key + "." + str(keySG2) + ";\n")
                                # --------------------------------------------------------------------------------
                                i = 0
                                File.write("\n")
                                File.write("\tsig[" + str(i) + "] = ")
                                bitNibble = 4
                                length = len(sorted_signals)
                                LastMSB = -1
                                for keySG2, dataSG2 in sorted_signals:
                                    # ---------------------------------------------------------
                                    length = length - 1
                                    data_length = dataSG2['length']
                                    # ---------------------------------------------------------
                                    X = dataSG2['LSB'] - LastMSB - 1
                                    # ---------------------------------------------------------
                                    if re.search("CHKSUM", keySG2):
                                        CheckSumString = keySG2
                                        tab = [""] * 1  # Initialize list with 1 element
                                        tab[0] = dataSG1['CheckSumInit']
                                    elif re.search("PROCESS", keySG2):
                                        CounterString = keySG2
                                        tab = [""] * 1  # Initialize list with 1 element
                                        tab[0] = "(" + CounterString + ")"
                                    else:
                                        if dataSG2['length'] <= 4:
                                            tab = [""] * 1  # Initialize list with 1 element
                                            tab[0] = "(" + keySG2 + ")"
                                        elif 4 < dataSG2['length'] <= 8:
                                            tab = [""] * 2  # Initialize list with 2 elements
                                            tab[0] = "((" + keySG2 + " & 0xF0) >> 4)"
                                            tab[1] = "(" + keySG2 + " & 0x0F)"
                                        elif 8 < dataSG2['length'] <= 12:
                                            tab = [""] * 3  # Initialize list with 3 elements
                                            tab[0] = "((" + keySG2 + " & 0xF00) >> 8)"
                                            tab[1] = "((" + keySG2 + " & 0x0F0) >> 4)"
                                            tab[2] = "(" + keySG2 + " & 0x00F)"
                                        elif 12 < dataSG2['length'] <= 16:
                                            tab = [""] * 4  # Initialize list with 4 elements
                                            tab[0] = "((" + keySG2 + " & 0xF000) >> 12)"
                                            tab[1] = "((" + keySG2 + " & 0x0F00) >> 8)"
                                            tab[2] = "((" + keySG2 + " & 0x00F0) >> 4)"
                                            tab[3] = "(" + keySG2 + " & 0x000F)"
                                        else:
                                            tab = [""] * 6  # Initialize list with 6 elements
                                            tab[0] = "((" + keySG2 + " & 0xF00000) >> 20)"
                                            tab[1] = "((" + keySG2 + " & 0x0F0000) >> 16)"
                                            tab[2] = "((" + keySG2 + " & 0x00F000) >> 12)"
                                            tab[3] = "((" + keySG2 + " & 0x000F00) >> 8)"
                                            tab[4] = "((" + keySG2 + " & 0x0000F0) >> 4)"
                                            tab[5] = "(" + keySG2 + " & 0x00000F)"
                                    # ------------------------------------------------------
                                    for signal in tab:
                                        while X != 0:
                                            if X > bitNibble:
                                                X = X - bitNibble
                                                File.write("0;\n")
                                                i = i + 1
                                                if i != 16:
                                                    File.write("\tsig[" + str(i) + "] = ")
                                                bitNibble = 4
                                            elif X == bitNibble:
                                                File.write("0;\n")
                                                i = i + 1
                                                if i != 16:
                                                    File.write("\tsig[" + str(i) + "] = ")
                                                bitNibble = 4
                                                X = 0
                                            elif X < bitNibble:
                                                bitNibble = bitNibble - X
                                                X = 0
                                        while True:
                                            if data_length % 4 == 0:
                                                temp = 4
                                            else:
                                                temp = data_length % 4
                                            data_length = data_length - temp
                                            if temp > bitNibble:
                                                temp = temp - bitNibble
                                                File.write(signal + " >> " + str(temp) + ";\n")
                                                i = i + 1
                                                if i != 16:
                                                    File.write("\tsig[" + str(i) + "] = ")
                                                data_length = data_length + temp
                                                bitNibble = 4
                                                if length == 0 and data_length == 0:
                                                    File.write("0;\n")
                                            elif temp == bitNibble:
                                                File.write(signal + ";\n")
                                                i = i + 1
                                                if i != 16:
                                                    File.write("\tsig[" + str(i) + "] = ")
                                                    if length == 0 and data_length == 0:
                                                        File.write("0;\n")
                                                bitNibble = 4
                                                temp = 0
                                            elif temp < bitNibble:
                                                bitNibble = bitNibble - temp
                                                File.write(signal + " << " + str(bitNibble) + " | ")
                                                if length == 0 and data_length == 0:
                                                    File.write("0;\n")
                                                temp = 0
                                            if temp == 0:
                                                break

                                    LastMSB = dataSG2['MSB']
                                # --------------------------------------------------------------------
                                File.write("\n\tif (temp_" + str(hex(data['id'])) + " == 0){")
                                # Begin of the ALgo
                                File.write("\n\t\tfor (i=0; i < 16; i++){")
                                File.write('\n\t\t\tsum = sum + sig[i];\n\t\t}\n')
                                File.write('\t\t//write("Sum of nibbles and CheckSumInit : %x", sum);\n')
                                File.write('\t\tmod = sum % 16;\n')
                                File.write('\t\t//write("Modulo of %x : %x", sum, mod);\n')
                                File.write('\t\tmod = ~mod;\n')
                                File.write('\t\t//write("Temp Value of CHKSUM : %x", mod);\n')
                                File.write("\t\t" + CheckSumString + " = mod & 0x0F;\n")
                                File.write('\t\t//write("CHKSUM Value : %x", ' + CheckSumString + ');')
                                File.write("\n\t}")
                                # The end of Algo
                                File.write("\n\telse {")
                                File.write("\n\t\t" + CheckSumString + " = " + CheckSumString + ";")
                                File.write("\n\t\t" + CounterString + " = 0x0F - " + CheckSumString + ";")
                                File.write("\n\t}\n")
                                # fill checksum and counter data byte
                                File.write("\n\t_m" + key + "." + CheckSumString + " = " + CheckSumString + ";")
                                File.write("\n\t_m" + key + "." + CounterString + " = " + CounterString + ";\n")
                                # Incrementation of global counter var
                                File.write("\n\t" + CounterString + " ++;")
                                File.write("\n\tif (" + CounterString + " > 0x0F){")
                                File.write("\n\t\t" + CounterString + " = 0;\n\t}\n")
                        # Normal script
                        File.write('\n\toutput(_m' + key + ');\n')
                        File.write('\tsetTimer(_t' + key + ', _' + key + 'CycleTime);\n')
                        File.write('}\n\n\n')
                # --------------------------------------------------------------------------------------------
                # On EnvVar ... {}
                for key, data in self.Messages.items():
                    if node == data['sender'][0] or node == data['sender'][1]:
                        sorted_signals = sorted(data['signals'].items(), key=lambda item: item[1]['startbit'])
                        for keySG, dataSG in sorted_signals:
                            File.write('On EnvVar ENV_' + str(hex(data['id'])) + '_' + keySG + '\n')
                            File.write('{\n')
                            File.write('\t_m' + key + '.' + keySG + '.phys = getValue(this);\n')
                            File.write('}\n\n')
                            if re.search("CHKSUM", keySG):
                                File.write('On EnvVar ENV_' + str(hex(data['id'])) + '_' + 'CRC_FAULT\n')
                                File.write('{\n')
                                File.write('\ttemp_' + str(hex(data['id'])) + ' = getValue(this);\n')
                                File.write('}\n\n')
                # Stop Generating
                File.close()
        print("[CAN] CAPL code generated for all virtual CAN nodes")



    def Generate_XVP(self, CANname,  img_1, img_2):
        xvp_root = self.output_root / Path("xvp_config")
        xvp_root.mkdir(parents=True, exist_ok=True)
        img_path = self.repo_root / Path("inputs") / Path("can") / Path("img")
        img_Path_1 = img_path / Path(img_1)
        img_Path_2 = img_path / Path(img_2)
        if not img_Path_1 or not img_Path_2:
            raise FileNotFoundError(
                f"Expected img under {img_path}"
            )
        # Call DB parser
        self.DBCparser()
        print("\n[CAN] Generating XML/XVP configuration")
        # For each virtual node
        for node in self.Nodes:
            # Create XSV file
            fileName = node + ".xvp"
            filePath = xvp_root / Path(fileName)
            with open(filePath, 'w') as File:
                # Text shared with all panels : Header
                File.write('<?xml version="1.0"?>\n')
                File.write('<Panel Type="Vector.CANalyzer.Panels.PanelSerializer, Vector.CANalyzer.Panels.Serializer, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null">\n')
                File.write('\t<Object Type="Vector.CANalyzer.Panels.Runtime.Panel, Vector.CANalyzer.Panels.Common, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="Panel" Children="Controls" ControlName="' + node + '_' + CANname + '_CAN_Panel">\n')
                # For each Message
                j = 0
                temp = 0
                Y_panel_temp = 0
                yloc_Groupe_BOX = 140
                xloc_Groupe_BOX = 180
                for key, data in self.Messages.items():
                    if node == data['sender'][0] or node == data['sender'][1]:
                        # Create GroupBoxControl
                        File.write('\t\t<Object Type="Vector.CANalyzer.Panels.Design.GroupBoxControl, Vector.CANalyzer.Panels.CommonControls, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="GroupBoxControl' + str(j) + '" Children="Controls" ControlName="Group Box 13">\n')
                        i = 0
                        Y = 0
                        j = j + 1
                        CRC = False
                        # increment Y location for each msg
                        yloc_Groupe_BOX = yloc_Groupe_BOX + (temp*30) + 45
                        temp = 0
                        # Check the location
                        if yloc_Groupe_BOX >= 2500:
                            if Y_panel_temp < yloc_Groupe_BOX:
                                Y_panel_temp = yloc_Groupe_BOX
                            yloc_Groupe_BOX = 185
                            xloc_Groupe_BOX = xloc_Groupe_BOX + 840
                            temp = 0
                        # create Y size for each msg
                        ysize_Groupe_BOX = 40
                        for keySG, dataSG in data['signals'].items():
                            # Check if CHKSUM exist to add a crc led control
                            if re.search("CHKSUM", keySG):
                                CRC = True
                            # Calculate X & Y size for each signal
                            i = i + 1
                            if i % 2 != 0:
                                # print signal in left side
                                X = 24
                                Y = Y + 30
                                # increment Y size for each sig
                                ysize_Groupe_BOX = ysize_Groupe_BOX + 30
                                temp = temp + 1
                            else:
                                # print signal in right side
                                X = 439
                            # print Static Text Control
                            File.write('\t\t\t<Object Type="Vector.CANalyzer.Panels.Design.StaticTextControl, Vector.CANalyzer.Panels.CommonControls, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="StaticTextControl' + str(i) + '" Children="Controls" ControlName="Static Text">\n')
                            File.write('\t\t\t\t<Property Name="Name">StaticTextControl' + str(i) + '</Property>\n')
                            File.write('\t\t\t\t<Property Name="Size">180, 21</Property>\n')
                            File.write('\t\t\t\t<Property Name="Location">' + str(X) + ', ' + str(Y) + '</Property>\n')
                            File.write('\t\t\t\t<Property Name="Font">Microsoft YaHei, 8.25pt, style=Bold</Property>\n')
                            File.write('\t\t\t\t<Property Name="Text">' + keySG + '</Property>\n')
                            File.write('\t\t\t\t<Property Name="ForeColor">64, 64, 64</Property>\n')
                            File.write('\t\t\t</Object>\n')
                            # Calculate X size for next
                            if i % 2 != 0:
                                X = 209
                            else:
                                X = 624
                            # if Val tab exist
                            if dataSG['desc'] != "none":
                                # Create ComboBoxControl
                                File.write('\t\t\t<Object Type="Vector.CANalyzer.Panels.Design.ComboBoxControl, Vector.CANalyzer.Panels.CommonControls, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="ComboBoxControl' + str(i) + '" Children="Controls" ControlName="Combo Box">\n')
                                File.write('\t\t\t\t<Property Name="Name">ComboBoxControl' + str(i) + '</Property>\n')
                                File.write('\t\t\t\t<Property Name="Size">180, 21</Property>\n')
                                File.write('\t\t\t\t<Property Name="Location">' + str(X) + ', ' + str(Y) + '</Property>\n')
                                File.write('\t\t\t\t<Property Name="DropDownWidth">180</Property>\n')
                                File.write('\t\t\t\t<Property Name="BackColor">Window</Property>\n')
                                File.write('\t\t\t\t<Property Name="ForeColor">WindowText</Property>\n')
                                File.write('\t\t\t\t<Property Name="TextBackColor">138, 185, 241</Property>\n')
                                File.write('\t\t\t\t<Property Name="DisplayLabel">Hide</Property>\n')
                                File.write('\t\t\t\t<Property Name="DescriptionText">' + keySG + ':</Property>\n')
                                File.write('\t\t\t\t<Property Name="UsedValueTable">PhysicalValue</Property>\n')
                                File.write('\t\t\t\t<Property Name="DescriptionSize">0, 0</Property>\n')
                                File.write('\t\t\t\t<Property Name="SymbolConfiguration">8;1;' + CANname + '_CAN;;;ENV_' + str(hex(data['id'])) + '_' + keySG + ';1;;;-1;;;;;;0</Property>\n')
                                File.write('\t\t\t\t<Property Name="TabIndex">' + str(i) + '</Property>\n')
                                File.write('\t\t\t</Object>\n')
                            else:
                                # Create TextBoxControl
                                File.write('\t\t\t<Object Type="Vector.CANalyzer.Panels.Design.TextBoxControl, Vector.CANalyzer.Panels.CommonControls, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="TextBoxControl' + str(i) + '" Children="Controls" ControlName="Input/Output Box">\n')
                                File.write('\t\t\t\t<Property Name="Name">TextBoxControl' + str(i) + '</Property>\n')
                                File.write('\t\t\t\t<Property Name="Size">180, 20</Property>\n')
                                File.write('\t\t\t\t<Property Name="Location">' + str(X) + ', ' + str(Y) + '</Property>\n')
                                File.write('\t\t\t\t<Property Name="AlarmUpperTextColor">WindowText</Property>\n')
                                File.write('\t\t\t\t<Property Name="AlarmLowerBkgColor">Salmon</Property>\n')
                                File.write('\t\t\t\t<Property Name="TextBackColor">Gray</Property>\n')
                                File.write('\t\t\t\t<Property Name="ValueDisplay">Decimal</Property>\n')
                                File.write('\t\t\t\t<Property Name="AlarmUpperBkgColor">IndianRed</Property>\n')
                                File.write('\t\t\t\t<Property Name="AlarmLowerTextColor">WindowText</Property>\n')
                                File.write('\t\t\t\t<Property Name="AlarmGeneralSettings">1;0;0;100</Property>\n')
                                File.write('\t\t\t\t<Property Name="DisplayLabel">Hide</Property>\n')
                                File.write('\t\t\t\t<Property Name="DescriptionSize">0, 0</Property>\n')
                                File.write('\t\t\t\t<Property Name="SymbolConfiguration">8;1;' + CANname + '_CAN;;;ENV_' + str(hex(data['id'])) + '_' + keySG + ';1;;;-1;;;;;;0</Property>\n')
                                File.write('\t\t\t\t<Property Name="TabIndex">' + str(i) + '</Property>\n')
                                File.write('\t\t\t</Object>\n')
                        # CRC
                        if CRC:
                            # Calculate X & Y size
                            i = i + 1
                            if i % 2 != 0:
                                Y = Y + 30
                                ysize_Groupe_BOX = ysize_Groupe_BOX + 30
                                temp = temp + 1
                            # print Static Text Control
                            File.write('\t\t\t<Object Type="Vector.CANalyzer.Panels.Design.StaticTextControl, Vector.CANalyzer.Panels.CommonControls, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="StaticTextControl' + str(i) + '" Children="Controls" ControlName="Static Text ' + str(i) + '">\n')
                            File.write('\t\t\t\t<Property Name="Name">StaticTextControl' + str(i) + '</Property>\n')
                            File.write('\t\t\t\t<Property Name="Size">139, 13</Property>\n')
                            Y = Y + 5
                            X = 624
                            File.write('\t\t\t\t<Property Name="Location">' + str(X) + ', ' + str(Y) + '</Property>\n')
                            File.write('\t\t\t\t<Property Name="Text">Press to induce CRC Failure</Property>\n')
                            File.write('\t\t\t</Object>\n')
                            # print Led Control
                            File.write('\t\t\t<Object Type="Vector.CANalyzer.Panels.Design.LEDControl, Vector.CANalyzer.Panels.CommonControls, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="LEDControl' + str(i) + '" ControlName="LED Control ' + str(i) + '">\n')
                            File.write('\t\t\t\t<Property Name="Name">LEDControl' + str(i) + '</Property>\n')
                            File.write('\t\t\t\t<Property Name="Size">24, 24</Property>\n')
                            Y = Y - 5
                            X = X + 150
                            File.write('\t\t\t\t<Property Name="Location">' + str(X) + ', ' + str(Y) + '</Property>\n')
                            File.write('\t\t\t\t<Property Name="SwitchValuesVTXml">&lt;Version&gt;2&lt;/Version&gt;&lt;Count&gt;2&lt;/Count&gt;&lt;RxValue&gt;0&lt;/RxValue&gt;&lt;LowerUpper&gt;Lower&lt;/LowerUpper&gt;&lt;TxValue&gt;0&lt;/TxValue&gt;&lt;Color&gt;Gray&lt;/Color&gt;&lt;RxValue&gt;1&lt;/RxValue&gt;&lt;LowerUpper&gt;Lower&lt;/LowerUpper&gt;&lt;TxValue&gt;1&lt;/TxValue&gt;&lt;Color&gt;Red&lt;/Color&gt;</Property>\n')
                            File.write('\t\t\t\t<Property Name="SymbolConfiguration">8;1;' + CANname + '_CAN;;;ENV_' + str(hex(data['id'])) + '_' + 'CRC_FAULT;1;;;-1;;;;;;0</Property>\n')
                            File.write('\t\t\t\t<Property Name="TabIndex">' + str(i) + '</Property>\n')
                            File.write('\t\t\t</Object>\n')
                        # GroupBoxControl properties
                        File.write('\t\t\t<Property Name="Name">GroupBoxControl' + str(j) + '</Property>\n')
                        File.write('\t\t\t<Property Name="Size">828, ' + str(int(ysize_Groupe_BOX)) + '</Property>\n')
                        File.write('\t\t\t<Property Name="Location">' + str(int(xloc_Groupe_BOX)) + ', ' + str(int(yloc_Groupe_BOX)) + '</Property>\n')
                        File.write('\t\t\t<Property Name="Font">Verdana, 9pt, style=Bold</Property>\n')
                        File.write('\t\t\t<Property Name="Text">' + key + ':</Property>\n')
                        File.write('\t\t\t<Property Name="TabIndex">' + str(j) + '</Property>\n')
                        File.write('\t\t</Object>\n')
                # calculate Y size of panel
                if Y_panel_temp != 0:
                    yloc_Groupe_BOX = Y_panel_temp
                yloc_Groupe_BOX = yloc_Groupe_BOX + (temp * 30) + 100
                Y_img2 = yloc_Groupe_BOX - 50
                # image 1 (OEM Logo)
                X_panel = xloc_Groupe_BOX + 1000
                X_img1 = (X_panel/2) - 210
                File.write('\t\t<Object Type="Vector.CANalyzer.Panels.Design.PictureBoxControl, Vector.CANalyzer.Panels.CommonControls, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="PictureBoxControl1" ControlName="Picture Box 1">\n')
                File.write('\t\t\t<Property Name="Name">PictureBoxControl1</Property>\n')
                File.write('\t\t\t<Property Name="Size">420, 124</Property>\n')
                File.write('\t\t\t<Property Name="Location">' + str(int(X_img1)) + ', 29</Property>\n')  # need adaptation !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                File.write('\t\t\t<Property Name="ImageFile">' + str(img_Path_1) + '</Property>\n')  # Image Path
                File.write('\t\t\t<Property Name="IsProportional">False</Property>\n')
                File.write('\t\t</Object>\n')
                # image 2 (LEAR Logo)
                File.write('\t\t<Object Type="Vector.CANalyzer.Panels.Design.PictureBoxControl, Vector.CANalyzer.Panels.CommonControls, Version=15.3.89.0, Culture=neutral, PublicKeyToken=null" Name="PictureBoxControl2" ControlName="Picture Box 2">\n')
                File.write('\t\t\t<Property Name="Name">PictureBoxControl2</Property>\n')
                File.write('\t\t\t<Property Name="Size">91, 33</Property>\n')
                File.write('\t\t\t<Property Name="Location">25, ' + str(int(Y_img2)) + '</Property>\n')
                File.write('\t\t\t<Property Name="ImageFile">' + str(img_Path_2) + '</Property>\n')  # Image Path
                File.write('\t\t\t<Property Name="IsProportional">False</Property>\n')
                File.write('\t\t</Object>\n')
                # Text shared with all panels :  : Footer
                Y_panel = yloc_Groupe_BOX
                File.write('\t\t<Property Name="Name">Panel</Property>\n')
                File.write('\t\t<Property Name="Size">' + str(int(X_panel)) + ', ' + str(int(Y_panel)) + '</Property>\n')
                File.write('\t\t<Property Name="BackColor">138, 185, 241</Property>\n') # Panel color
                File.write('\t</Object>\n')
                File.write('</Panel>\n')
                # Stop Generating
                File.close()
        print("[CAN] XML/XVP files generated successfully")
