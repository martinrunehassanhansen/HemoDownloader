##  HemoDownloader 1.2
##  A GUI UTILITY FOR DOWNLOADING DATA FROM HEMOCUE® HBA1C 501 DEVICES
##  Copyright © 2018-2019 Martin Rune Hassan Hansen <martinrunehassanhansen@ph.au.dk>

##  This program is free software: you can redistribute it and/or modify
##  it under the terms of the GNU General Public License as published by
##  the Free Software Foundation, either version 3 of the License, or
##  (at your option) any later version.

##  This program is distributed in the hope that it will be useful,
##  but WITHOUT ANY WARRANTY; without even the implied warranty of
##  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
##  GNU General Public License for more details.

##  You should have received a copy of the GNU General Public License
##  along with this program.  If not, see <https://www.gnu.org/licenses/>.



##  VERSION 1.0.1, 2018-09-22
##  Change compared with version 1.0: The program now gracefully handles HbA1c values listed as "< 4 %".

##  VERSION 1.0.2b, 2018-09-24
##  Change compared with version 1.0.1: The program now gracefully handles HbA1c values listed as "> 14 %".

##  VERSION 1.1, 2019-07-02
##  Changes compared with version 1.0.2b:
##      * Added legal notice
##      * Added help
##      * Added license information
##      * Added 'About' information
##      * Added debug mode
##      * Changed default timeout
##      * Changed headings in exported files
##      * Added more tests of data integrity

##  VERSION 1.2, 2021-04-13
##  Change compared with version 1.1:
##  Bugfix:   Because of a difference in formatting, previous versions of HemoDownloader could not parse data from
##  HemoCue HbA1c 501 devices running software software revision 2014-08-02. This has now been fixed.

__version__ = 1.2


##  IMPORT NECESSARY MODULES
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import filedialog
from os import path
import re, csv, datetime

modulesNotFound = False
try:
    import serial, serial.tools.list_ports
except ModuleNotFoundError:
    modulesNotFound = True
try:
    import xlsxwriter
except ModuleNotFoundError:
    modulesNotFound = True
try:
    import xlwt
except ModuleNotFoundError:
    modulesNotFound = True


##  MAIN WINDOW
class settingsWindow(ttk.Frame):
    
    ##  INITIALIZE THE WINDOW.
    def __init__(self, master):
        #   Define geometry of window.
        ttk.Frame.__init__(self, master=master)
        self.master.title('HemoDownloader 1.2 - Transfer data from HemoCue® HbA1c 501')
        self.master.resizable(False, False)

        #   Create menubar
        self.menubar = tk.Menu(root)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Exit", command=root.destroy)
        self.menubar.add_cascade(label="File", menu=self.filemenu)
        self.editmenu = tk.Menu(self.menubar, tearoff=0)
        self.helpmenu = tk.Menu(self.menubar, tearoff=0)
        self.helpmenu.add_command(label="Instructions for use", command=lambda: self.showHelpBox("Instructions for use"))
        self.helpmenu.add_command(label="About HemoDownloader", command=lambda: self.showHelpBox("About HemoDownloader"))
        self.helpmenu.add_command(label="License information", command=lambda: self.showHelpBox("License information"))
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)
        root.config(menu=self.menubar)
        
        #   Create widgets.
        self.firstTimeExecuted = True
        self.initialDir = r'\\'
        self.outputFilePathLabel = ttk.Label(root, text="File to save:")
        self.outputFilePathLabel.grid(row = 2, column=0, sticky='W', pady=15, padx=10)
        
        self.outputFilename = ''
        self.outputFilenameLabel = ttk.Entry(root,textvariable=self.outputFilename, width=50, state='readonly')
        self.outputFilenameLabel.grid(row=2, column=1, sticky='W', padx=10)

        #   Define list of filetypes for save-as menu.
        self.tabularFiletypes = [('CSV','*.csv'),('Tab Separated Values','*.tsv'),('Excel Workbook','*.xlsx'),('Excel 97-2003 Workbook','*.xls'),('All files','*.*')]

        #   More widgets
        self.saveAsButton = ttk.Button(root,text='Save As...', command=self.save_as_filename)
        self.saveAsButton.grid(row=2, column=2, sticky='W', padx=10)

        self.comportLabel = ttk.Label(root, text="Serial port connection to use:")
        self.comportLabel.grid(row = 3, column=0, sticky='W', padx=10)
        
        self.comportLongName = tk.StringVar()
        self.comportMenu = ttk.OptionMenu(root, self.comportLongName, *[], command=self.select_comport)
        self.comportMenu.grid(row=3, column=1, sticky='W', padx=10)
        self.update_serial_port_list()
        self.registerComportConnectionErrorboxState(False)

        self.timeoutLabel = ttk.Label(root, text="Connection timeout:")
        self.timeoutLabel.grid(row = 4, column=0, sticky='W', padx=10, pady=15)
        
        self.timeoutString = tk.StringVar()
        self.timeoutStringChoices = ['dummy','30 seconds','1 minute','2 minutes','3 minutes','5 minutes','10 minutes']
        self.timeoutMenu = ttk.OptionMenu(root, self.timeoutString, *self.timeoutStringChoices, command=self.set_timeoutSeconds)
        self.timeoutMenu.grid(row=4, column=1, sticky='W', padx=10)
        self.timeoutString.set(self.timeoutStringChoices[5])
        self.set_timeoutSeconds()

        #  Display warning text stating that the software is not approved for medical use.
        self.warningText = """This software is intended for use in epidemiological studies only.
        It is *not* approved as a medical device and must not be used as such.
        That means the software must not be used on human beings for diagnosis, prevention,
        monitoring, prediction, prognosis, treatment or alleviation of disease or any other medical purposes.\n
        By clicking 'RECEIVE DATA' below you are agreeing not to use the software as a medical device.
        """
        self.warningContent = tk.StringVar()
        self.warningContent.set(self.warningText)
        self.warningTextBox = ttk.Label(root, justify='center', textvariable=self.warningContent)
        self.warningTextBox.grid(row=6, column=0,columnspan=3, pady=0)

        #  Button that has to be clicked to start transmission.
        self.startButton = ttk.Button(self.master, text='RECEIVE DATA', command=self.recordData)
        self.startButton.grid(row=7, column=0,columnspan=3, pady=(5,15))
        
    ##  REGISTER WHETHER THERE IS AN OPEN ERROR BOX REPORTING THAT THE SERIAL CONNECTION COULD NOT BE OPENED.
    def registerComportConnectionErrorboxState(self, state):
        self.comportConnectionErrorboxOpen = state

    ##  SELECT WHICH COMPORT TO USE
    def select_comport(self, longComportName):
        try:
            self.comportLongName.set(longComportName)
            self.comportShortName = self.comportsDict[self.comportLongName.get()]
        except KeyError:
            self.after(2500, self.update_serial_port_list)
            
    ##  GET AN ALPHABETICALLY SORTED LIST OF AVAILABLE SERIAL PORTS.
    def list_serial_ports(self):
        try:
            comports = list(serial.tools.list_ports.comports(include_links=False))
            if len(comports) != 0:
                self.comportsDict = {}
                for comport in comports:
                    self.comportsDict.update({comport[1]:comport[0]})
            else:
                self.comportsDict = {'[NO SERIAL PORTS AVAILABLE]':'-'}
        except:
            self.comportsDict = {'[NO SERIAL PORTS AVAILABLE]':'-'}
        self.comportsLongNames = sorted(self.comportsDict)

    ##  UPDATE THE LIST OF SERIAL PORTS EVERY 2.5 SECONDS AND SHOW AN ERROR BOX IF THE SELECTED SERIAL PORT BECOMES UNAVAILABLE.
    def update_serial_port_list(self):
        self.list_serial_ports()
        previouslySelectedComport = self.comportLongName.get()     
        menu = self.comportMenu["menu"]
        menu.delete(0,"end")
        for string in self.comportsLongNames:
            menu.add_command(label=string, command=lambda value=string: self.select_comport(value))
        if previouslySelectedComport != '' and previouslySelectedComport != '[NO SERIAL PORTS AVAILABLE]':
            previousComportAmongCurrentPorts = self.check_connection_comport(previouslySelectedComport)
            if not previousComportAmongCurrentPorts:
                if not self.comportConnectionErrorboxOpen:
                    self.serial_port_connection_lost_error(previouslySelectedComport)
            else:
                self.select_comport(previouslySelectedComport)
        else:
            self.select_comport(self.comportsLongNames[0])
        self.after(2500, self.update_serial_port_list)

    ##  CHECK IF THE SELECTED SERIAL PORT HAS BECOME UNAVAILABLE.
    def check_connection_comport(self, longComportName):
        comports = list(serial.tools.list_ports.comports(include_links=False))
        comportFound = False
        for comport in comports:
            if comport[1] == longComportName:
                comportFound = True
        self.select_comport(self.comportsLongNames[0])
        return comportFound

    ##  ERROR BOX: THE SELECTED SERIAL PORT HAS BECOME UNAVAILABLE
    def serial_port_connection_lost_error(self, longComportName):
        warningMessage = 'The serial port "' + longComportName + '" has been disconnected.\r\nPlease check the connection or use another serial port.'
        tk.messagebox.showwarning("Connection lost",warningMessage, command=self.registerComportConnectionErrorboxState(True))
        self.registerComportConnectionErrorboxState(False)

    ##  ERROR BOX: NO SERIAL PORT AVAILABLE.
    def serial_port_missing_error(self):
        tk.messagebox.showerror("No connection","No serial port availabe for data capture.\r\nPlease connect to a serial port and try again.")

    ##  ERROR BOX: THE SELECTED SERIAL PORT IS CONNECTED, BUT COULD NOT BE OPENED.
    def serial_port_could_not_open_error(self, longComportName):
        errorMessage = 'The serial port "' + longComportName + '" could not be opened.\r\nPlease close any other programs that might be using this serial port, and try again.'
        tk.messagebox.showerror("Serial port error",errorMessage)
 
    ##  ERROR BOX: NO FILENAME SPECIFIED
    def output_filename_missing_error(self):
        tk.messagebox.showerror("Missing filename","Please enter a filename.")

    ##  DIALOG FOR SPECIFYING FILENAME FOR OUTPUT DATA.
    def save_as_filename(self):

        #   If no filename has yet been defined (i.e., this is the first time the dialog is opened), seed the function.
        if len(self.outputFilename) == 0:
            if self.firstTimeExecuted == True:
                self.initialDir = r'\\'
            initialFile = ''
            possibleFiletypes = self.tabularFiletypes

        #   If a filename also already been defined (i.e., this is not the first time that the dialog has been opened), take the old information into account when opening the dialog.
        else:
            #   The initial folder should be the same as the folder that the user navigated to the last time (unless it no longer exists).
            self.initialDir = path.dirname(self.outputFilename)
            if not path.isdir(self.initialDir):
                self.initialDir = r'\\'

            #   Re-use the old filename
            initialFile = path.basename(self.outputFilename)

            #   Re-arrange the list of filetypes so that the previously selected filetype is the first item (meaning it will be pre-selected in the drop-down menu).
            possibleFiletypes = self.tabularFiletypes
            oldExtension = list(path.splitext(initialFile))[1]
            extensionIndex = -1
            for index in range(len(possibleFiletypes)):
                if possibleFiletypes[index][1] == '*' + oldExtension:
                    extensionIndex = index
            if extensionIndex == -1:
                for index in range(len(possibleFiletypes)):
                    if possibleFiletypes[index][1] == '*.*':
                        extensionIndex = index
            firstFiletype = possibleFiletypes.pop(extensionIndex)
            possibleFiletypes.insert(0, firstFiletype)

        #   Set the default extension.
        if possibleFiletypes[0][1] != '*.*':
            defaultExtension = possibleFiletypes[0][1]
        else:
            defaultExtension = ''

        #   Open the dialog, parse the filename that was input and display it on the settings screen.
        f = filedialog.asksaveasfilename(initialdir = self.initialDir, initialfile = initialFile, title = "Save as", filetypes = possibleFiletypes, defaultextension=defaultExtension)
        if f is not None and f != '':
            self.setOutputFilename(f)

    ##  SET THE OUTPUT FILENAME AND DISPLAY IT IN THE APPROPRIATE BOX.
    def setOutputFilename(self, outputFilename):
        self.outputFilename = outputFilename
        self.outputFilenameLabel.config(state='NORMAL')
        self.outputFilenameLabel.delete(0,'end')
        self.outputFilenameLabel.insert(0,self.outputFilename)
        self.outputFilenameLabel.config(state='readonly')
        
    ##  SET THE CONNECTION TIMEOUT IN SECONDS.
    def set_timeoutSeconds(self, dummy='dummy'):
        timeoutNumber = int(self.timeoutString.get().split(' ')[0])
        if timeoutNumber == 30:
            self.timeoutSeconds = timeoutNumber
        else:
            self.timeoutSeconds = timeoutNumber*60
            
    ##  Complain if necessary modules are missing.
    def modules_not_found_error(self):
        tk.messagebox.showerror("Missing modules","One or more of the following necessary modules were not found:\n\npyserial\nXlsxWriter\nxlwt\n\nPlease close the program and install all of these modules before restarting.\nThe easiest way to perform the installation is to execute the script 'src/setup.py'.")
            
    ##  OPEN THE WINDOW FOR RECORDING/PARSING DATA FROM DEVICE AND WRITING IT TO A FILE.
    def recordData(self):

        #   Prepare variables
        comportShortName = self.comportShortName
        comportLongName = self.comportLongName.get()
        timeoutSeconds = self.timeoutSeconds
        outputFilename = self.outputFilename
        tabularFiletypes = self.tabularFiletypes

        #   Complain if necessary modules are missing.
        if modulesNotFound == True:
            self.modules_not_found_error()
            
        #   Complain if no serial port is available.
        elif comportShortName == '-':
            self.serial_port_missing_error()

        #   Complain if no filename has been entered.
        elif outputFilename == '':
            self.output_filename_missing_error()
        else:
            #   Try to open the serial port.
            try:
                serReceiver = serial.Serial(port=comportShortName, baudrate=9600, bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, timeout=None)

            #   Complain if the serial port cannot be opened.
            except serial.serialutil.SerialException:
                comportConnected = self.check_connection_comport(comportLongName)
                if comportConnected:
                    self.serial_port_could_not_open_error(comportLongName)
                else:
                    self.serial_port_connection_lost_error(comportLongName)
                return

            #   Set custom buffer size to avoid losing data during transmission. 200kB should always be sufficient.
            serReceiver.set_buffer_size(rx_size = 204800, tx_size = 204800)
            
            #   During data download, a custom dialog box is displayed if there is a problem with the serial connection.
            #   We make sure that this is the only error message that can be displayed during data download.
            self.registerComportConnectionErrorboxState(True)

            #   Open the window for downloading and processing data.
            s = dataProcessingWindow(self, 'MyTest', comportShortName, comportLongName, timeoutSeconds, serReceiver, outputFilename, tabularFiletypes)

            #   Wait for the window to close.
            self.master.wait_window(s)

            #   When the window has closed, re-able the settings window's ability to display error messages about the serial connection.
            self.registerComportConnectionErrorboxState(False)

            #   Clear the output filename to avoid overwriting data by mistake.
            self.firstTimeExecuted = False
            self.initialDir = path.dirname(self.outputFilename)
            if not path.isdir(self.initialDir):
                self.initialDir = r'\\'
            self.setOutputFilename('')

    ##  OPEN HELP WINDOW
    def showHelpBox(self,helpType):
        s = helpWindow(self, 'MyTest', helpType)
        self.master.wait_window(s)

##  WINDOW FOR RECORDING/PARSING DATA AND WRITING FILE.
class dataProcessingWindow(simpledialog.Dialog):

    ## INITIALIZE THE WINDOW.
    def __init__(self, parent, name, comportShortName, comportLongName, timeoutSeconds, serReceiver, outputFilename, tabularFiletypes):
        tk.Toplevel.__init__(self, master=parent)
        self.name = name
        self.length = 400

        #   Grab important variables.
        self.comportShortName = comportShortName
        self.comportLongName = comportLongName
        self.timeoutSeconds = timeoutSeconds
        self.serReceiver = serReceiver
        self.outputFilename = outputFilename
        self.tabularFiletypes = tabularFiletypes

        #   Create window and widgets.
        self.create_window()
        self.create_widgets()

    #   Create window.
    def create_window(self):
        #   Set focus on the new window.
        self.focus_set()

        #   Make sure that focus stays on the new window.
        self.grab_set()

        #   Only show one window in the taskbar.
        self.transient(self.master)

        #   Configure layout and position over parent window.
        self.title(u'Download data')
        self.resizable(False, False)
        dx = (self.master.master.winfo_width() >> 1) - (self.length >> 1)
        dy = (self.master.master.winfo_height() >> 1) -75
        self.geometry(u'+{x}+{y}'.format(x = self.master.winfo_rootx() + dx,
                                         y = self.master.winfo_rooty() + dy))
        
        # No matter how the window is destroyed, we will shut down cleanly using the function self.close()
        self.protocol(u'WM_DELETE_WINDOW', self.cancelTransfer)

        #   If the user presses the ESCAPE button, the process will cancel.
        self.bind(u'<Escape>', self.cancelTransfer)

    #   Create widgets.
    def create_widgets(self):
        self.bytesReceivedString = tk.StringVar()
        self.countdownString = tk.StringVar()
        self.secondsPassed = 0
        self.timeUntilTimeout = self.timeoutSeconds
        self.progressbarValue = tk.IntVar()
        self.timeoutTimeUnit = tk.StringVar()
        self.maximum = self.timeoutSeconds
        self.bytesReceived = 0
        self.binaryBuffer = b''
    
        ttk.Label(self, textvariable=self.bytesReceivedString).pack(anchor='w', padx=10, pady=10)
        self.progress = ttk.Progressbar(self, maximum=self.maximum, orient='horizontal',
                                        length=self.length, variable=self.progressbarValue, mode='determinate')
        self.progress.pack(padx=10)
        ttk.Label(self, textvariable=self.countdownString).pack(side='left', padx=(10,0), pady=10)
        ttk.Button(self, text='Cancel', command=self.cancelTransfer).pack(anchor='e', padx=(0,10), pady=10)
        #
        self.waitOneSecond()

    ## ENTER A ONE-SECOND LOOP, RECORDING DATA. AT THE SAME TIME, DISPLAY A PROGRESSBAR OF THE TIME TO CONNECTION TIMEOUT.
    def waitOneSecond(self):
        try:
            #   Attempt to read data from the serial port.
            self.getSerialData()

            #   Update the progress bar. Ignore single bytes of FF (hexadecimal, control character sent by HbA1c 501 device when cable is connected).
            if self.binaryStringReceived == b'' or self.binaryStringReceived == b'\xff':
                self.bytesReceivedString.set('Number of bytes received: ' + str(self.bytesReceived))
                self.secondsPassed += 1
                self.timeUntilTimeout = self.timeoutSeconds-self.secondsPassed
                if self.binaryBuffer == b'':
                    self.progressbarValue.set(self.timeUntilTimeout)
                    countdownString = 'Waiting. Timeout in '
                    minutesLeft = self.timeUntilTimeout // 60
                    secondsLeft = self.timeUntilTimeout - 60*minutesLeft
                    if minutesLeft > 0:
                        countdownString += str(minutesLeft) + ' minute'
                        if minutesLeft > 1:
                            countdownString += 's'
                        countdownString += ' and '
                    countdownString += str(secondsLeft) + ' second'
                    if secondsLeft > 1 or secondsLeft == 0:
                        countdownString += 's'
                    countdownString += '.'
                    self.countdownString.set(countdownString)
            else:
                self.bytesReceivedString.set('Number of bytes received: ' + str(self.bytesReceived))
                self.countdownString.set('Transmission in progress. Countdown halted.')
                if self.binaryBuffer == self.binaryStringReceived:
                    self.progress.configure(mode='indeterminate')
                    self.progress.start()
                self.secondsPassed = self.timeoutSeconds - 6

            #   If the countdown is not complete, take another round in the loop after 1 second.
            if self.timeUntilTimeout > 0:
                self.after(1000, self.waitOneSecond)

            #   When the connection times out, close the progress window.
            else:
                self.close()

                #   If no data was received, give the user an error message.
                if self.binaryBuffer == b'':
                    self.connectionTimedOutError()

                #   If data was recived, process it.
                else:
                    self.saveHbA1cData()

        #   In case of serial port communication errors, close the progress window and display an error message.
        except serial.serialutil.SerialException:
            self.close()
            self.serialPortCommError()
            
    ##  FUNCTION TO READ DATA FROM SERIAL PORT
    def getSerialData(self):
        #   Bugfixing mode: If you set "self.bugfixing" to 'True' instead of 'False' and run the program with a python terminal open,
        #   the program will request a path to a binary dump file containing HbA1c data. The bytestring in this file will then
        #   be treated as if it had been received over serial. Furthermore, the program will print each row of data as it is
        #   parsing it.
        self.bugfixing = False

        if self.bugfixing == False:
            self.binaryStringReceived = self.serReceiver.read(self.serReceiver.in_waiting)

        if self.bugfixing == True:
            if self.binaryBuffer == b'':
                bugfixingDataSource = ''
                while bugfixingDataSource == '' or not path.isfile(bugfixingDataSource):
                    bugfixingDataSource = input("\nBUGFIXING MODE: Please write path to binary file containing dump of HbA1c data.\n")
                    bugfixingDataSource = bugfixingDataSource.strip('\"')
                with open(bugfixingDataSource,'rb') as f:
                    self.binaryStringReceived = f.read()
            else:
                self.binaryStringReceived = b''

        if self.binaryStringReceived != b'\xff':
            self.bytesReceived += len(self.binaryStringReceived)
            self.binaryBuffer += self.binaryStringReceived
                
    ##  ERROR BOX: SERIAL CONNECTION ERROR.
    def serialPortCommError(self):
        warningMessage = ('Data transmission was interrupted due to a problem with the serial port "' + self.comportLongName + '".\r\n\r\n')
        if self.binaryBuffer != b'':
            warningMessage += 'Do you want to save a copy of the raw binary data so you can process it manually?'
            saveBinaryDataAnswer = tk.messagebox.askyesno("Error",warningMessage,icon='error',default='yes')
            if saveBinaryDataAnswer == True:
                self.saveBinaryData()
        else:
            warningMessage += 'No data was received before the interruption.'
            errorBox = tk.messagebox.showerror("Error",warningMessage)

    ##  FUNCTION EXECUTED IF THE USER CANCELS DATA RECORDING.
    def cancelTransfer(self, dummy='dummy'):
        self.close()
        if self.binaryBuffer != b'':
            title = 'Cancelled'
            question = ('Operation cancelled by user.\r\n\r\nSave a copy of the raw binary data that had already been received?')
            saveBinaryDataAnswer = tk.messagebox.askyesno(title,question,default='yes')
            if saveBinaryDataAnswer == True:
                self.saveBinaryData()

    ##  ERROR BOX: CONNECTION TIMED OUT
    def connectionTimedOutError(self):
        warningMessage = 'Connection timed out before any data was received.'
        tk.messagebox.showerror("Timeout",warningMessage)
                
    ##  CLOSE THE WINDOW CONTAINING THE PROGRESS BAR, AND CLOSE THE SERIAL CONNECTION.
    def close(self, event=None):
        self.serReceiver.close()
        self.master.focus_set()  # put focus back to the parent window
        self.destroy()  # destroy progress window

    ##  CHECK THE INTEGRITY OF THE BINARY DATA RECORDED.
    def checkDataIntegrity(self):        
        #   Get the string representation of the bytes.
        receivedData = str(self.binaryBuffer)

        #   Remove the leading 'b'.
        try:
            receivedData = receivedData[2:len(receivedData)-1]
        except IndexError:
            self.setUnknownDataStructure()
            return

        #   Split the string into list items, based on the occurrrence of the string 'Data No.:'.
        self.receivedDataRows = receivedData.split(r'Data No.: ')

        #   Now split based on line shifts.
        rowCounter = 0
        for row in self.receivedDataRows:
            self.receivedDataRows[rowCounter] = str(self.receivedDataRows[rowCounter]).split(r'\r\n')
            rowCounter += 1

        #   If the data is from a HemoCue® HbA1c 501 device, it *must* contain the string 'HEMOCUE HbA1c 501' or 'HEMOCUE HbA1c501' (depending on whether it printed a single or multiple pieces of data).
        try:
            if ('HEMOCUE HbA1c 501' in self.receivedDataRows[0][1]) or ('HEMOCUE HbA1c501' in self.receivedDataRows[0][1]):
                self.dataType = 'HbA1c'
            else:
                self.setUnknownDataStructure()
                return
        except IndexError:
            self.setUnknownDataStructure()
            return

        #   If transmission is complete, it will end with binary string b'\x1bd\x03'
        lastRow = self.receivedDataRows[len(self.receivedDataRows)-1]
        lastItem = lastRow[len(lastRow)-1]
        if lastItem == r'\x1bd\x03':
            self.transmissionCompleted = True
        else:
            self.transmissionCompleted = False

        #   Detect if transmission was briefly interrupted and then resumed before the interruption could be detected, resulting in data that cannot be parsed (or results in invalid characters).
        try:
            self.parseHbA1cData()
            for row in self.parsedHbA1cData[1:]:
                for cell in row:
                    if (r'\x' in cell) and (r'\\x' not in cell):
                        self.transmissionCompleted = False
                self.cellCounter = -3
                for cell in row[-2:]:
                    self.cellCounter += 1
                    try:
                        float(cell)
                    except ValueError:
                        if not ((self.cellCounter == -2 and cell == '<4') or (self.cellCounter == -2 and cell == '>14') or (self.cellCounter == -1 and cell == '')):
                            self.transmissionCompleted = False
                try:
                    int(row[0])
                except ValueError:
                    self.transmissionCompleted = False
                    
        except IndexError:
            self.transmissionCompleted = False

        #   Detect if the number of observations is consistent with the values listed in "Data no."
        #   If there is an inconsistency, it is because there was some loss of data during transmission, but by chance it resulted in data that could still be parsed.
        if self.transmissionCompleted == True:
            try:
                lowestDataNo = int(self.parsedHbA1cData[1][0])
                highestDataNo = int(self.parsedHbA1cData[len(self.parsedHbA1cData)-1][0])
                if not (len(self.parsedHbA1cData) - 1 == highestDataNo - lowestDataNo + 1):
                    self.transmissionCompleted = False
                if self.bugfixing == True:
                    print('\n\nNumber of observations in binary data received =',len(self.parsedHbA1cData) - 1)
                    print('Expected number of observations, based on values listed in "Data no." =',highestDataNo - lowestDataNo + 1)
            except ValueError:
                self.transmissionCompleted = False
                print('Error in number of observations.')

    ##  IF THE BINARY DATA RECEIVED HAS AN UNKNOWN STRUCTURE, RECORD THIS.
    def setUnknownDataStructure(self):
        self.dataType = 'unknown'
        self.transmissionCompleted = None
        self.receivedDataRows = None

    ##  PARSE THE BINARY DATA SO THAT IT BECOMES A WELL-DEFINED LIST OF LISTS (CORRESPONDING TO A MATRIX DATASET OF ROWS AND COLUMNS).
    def parseHbA1cData(self):
        self.parsedHbA1cData = []
        for dataRow in self.receivedDataRows[1:]:
            if self.bugfixing == True:
                print('\n\nNow parsing this row of data:\n',dataRow)

            #   There is an inconsistency between software revision 2014-08-02 of the HemoCue HbA1c and the version of the
            #   software that Hemodownloader was developed for. Software revision 2014-08-02 has two less line breaks in
            #   the printed output. This is fixed to avoid problems in the code below.
            if dataRow[3] == '':
                dataRow.pop(3)
            if dataRow[4] == '':
                dataRow.pop(4)

            #   Remove any barcodes
            dataRowNoBarcodes = []
            for cell in dataRow:
                if not cell.startswith(r'\x1dh0\x1dw\x03\x1b$'):
                    cell = cell.replace(r'\x00','')
                    dataRowNoBarcodes.append(cell)
            dataRow = dataRowNoBarcodes
                
            dataID = dataRow[0]
            date = dataRow[1]
            time = dataRow[2].replace('Time:','').replace('Time :','')

            inputDateTimeFormat = ''
            if r'[Y/M/D]' in date:
                inputDateTimeFormat += '%y/%m/%d'
                date = date.replace(r'[Y/M/D]','')
            elif r'[M/D/Y]' in date:
                inputDateTimeFormat += '%m/%d/%y'
                date = date.replace(r'[M/D/Y]','')
            elif r'[D/M/Y]' in date:
                inputDateTimeFormat += '%d/%m/%y'
                date = date.replace(r'[D/M/Y]','')

            if ('AM' in time) or ('PM' in time):
                inputDateTimeFormat += ' %p %I:%M'
            else:
                inputDateTimeFormat += ' %H:%M'

            dateTimeString = re.sub(' +',' ', date + time) 
          
            try:
                dateTime = datetime.datetime.strptime(dateTimeString, inputDateTimeFormat)
                dateTimeString = dateTime.strftime("%Y-%m-%dT%H:%M")
                hbA1cValues = dataRow[3].replace('HbA1c','').replace('mmol/mol','').replace(' ','').split('%')
                hba1cPercent = hbA1cValues[0]
                hba1cMmol = hbA1cValues[1]
                operatorID = dataRow[5]
                patientID = dataRow[7]
                self.parsedHbA1cData.append([dataID,dateTimeString,operatorID,patientID,hba1cPercent,hba1cMmol])
            except:
                print(dataRow)
                print('This causes an error. Press enter to continue.')
                dummy_var = input()
                self.transmissionCompleted = False

        try:
            self.parsedHbA1cData.sort(key=lambda number: number[0])
        except:
            self.transmissionCompleted = False
        self.parsedHbA1cData.insert(0,['data_id','datetime','operator_id','patient_id','hba1c_percent','hba1c_mmol_per_mol'])

    ##  ERROR BOX: THE DATA HAS UNKNOWN STRUCTURE.
    def unknownDataStructureError(self):
        title = 'Unknown data format'
        question = ('Cannot decode the data received. Either an error happened during transmission, ' +
                    'or the data is not from a HemoCue® HbA1c device.\r\n\r\n' +
                    'Do you want to save a copy of the raw binary data so you can process it manually?')
        saveBinaryDataAnswer = tk.messagebox.askyesno(title,question,default='yes',icon='error')
        if saveBinaryDataAnswer == True:
            self.saveBinaryData()

    ##  ERROR BOX: DATA TRANSMISSION WAS INCOMPLETE.
    def incompleteTransmissionError(self):
        title = 'Transmission incomplete'
        question = ('Data transmission was interrupted, probably because the HemoCue® HbA1c 501 device lost power, ' +
                    'or because of a problem with the serial cable.\r\n\r\n' +
                    'Do you want to save a copy of the raw binary data so you can process it manually?')
        saveBinaryDataAnswer = tk.messagebox.askyesno(title,question,default='yes',icon='error')
        if saveBinaryDataAnswer == True:
            self.saveBinaryData()

    ##  WRITE DATA (AS EITHER TABULAR OR BINARY DATA).
    def outputFileWriter(self,mode='tabular'):

        #   Tabular data
        if mode == 'tabular':

            #   Excel format
            if list(path.splitext(self.outputFilename))[1] == '.xlsx':
                #   Find number of columns.
                numberOfColumns = 0
                for row in self.parsedHbA1cData:
                    if len(row) > numberOfColumns:
                        numberOfColumns = len(row)

                #   Create a dictionary of column numbers and the minimum column width in characters.
                columnNumberOfChars = {}
                for columnNumber in range(numberOfColumns):
                    columnNumberOfChars.update({columnNumber:1})

                #   Open workbook
                workbook = xlsxwriter.Workbook(self.outputFilename)
                worksheet = workbook.add_worksheet()
                rowNumber = 0
                colNumber = 0

                #   Loop over the HbA1c data and write it to the cells of the worksheet. When a cell with contents wider than the minimum column width is encountered, update the mimimum width to accomodate the content.
                for row in self.parsedHbA1cData:
                    for cellContent in row:
                        worksheet.write(rowNumber, colNumber, cellContent)
                        if len(cellContent) > columnNumberOfChars[colNumber]:
                            columnNumberOfChars[colNumber] = len(cellContent)
                        colNumber += 1
                    colNumber = 0
                    rowNumber += 1

                #   Write the column widths to the file.
                for columnNumber in range(len(columnNumberOfChars)):
                    worksheet.set_column(columnNumber, columnNumber, columnNumberOfChars[columnNumber]+2)

                #   Close the Excel file.                                    
                workbook.close()

            #   Excel 97-2003 format
            elif list(path.splitext(self.outputFilename))[1] == '.xls':
                #   Find number of columns.
                numberOfColumns = 0
                for row in self.parsedHbA1cData:
                    if len(row) > numberOfColumns:
                        numberOfColumns = len(row)

                #   Create a dictionary of column numbers and the minimum column width in characters.
                columnNumberOfChars = {}
                for columnNumber in range(numberOfColumns):
                    columnNumberOfChars.update({columnNumber:1})
                    
                #   Open workbook.
                oldWorkbook = xlwt.Workbook()
                sheet1 = oldWorkbook.add_sheet("Sheet1")

                #   Loop over the HbA1c data and write it to the cells of the worksheet. When a cell with contents wider than the minimum column width is encountered, update the mimimum width to accomodate the content.
                for rowNumber in range(len(self.parsedHbA1cData)):
                    row = sheet1.row(rowNumber)
                    for columnNumber in range(len(self.parsedHbA1cData[rowNumber])):
                        row.write(columnNumber,self.parsedHbA1cData[rowNumber][columnNumber])
                        if len(self.parsedHbA1cData[rowNumber][columnNumber]) > columnNumberOfChars[columnNumber]:
                            columnNumberOfChars[columnNumber] = len(self.parsedHbA1cData[rowNumber][columnNumber])

                #   Write the column widths to the file.
                for columnNumber in range(len(columnNumberOfChars)):
                    sheet1.col(columnNumber).width = (columnNumberOfChars[columnNumber]+2) * 256

                #   Close the Excel file.
                oldWorkbook.save(self.outputFilename)

            #   Tab-Separated Values
            elif list(path.splitext(self.outputFilename))[1] == '.tsv':
                with open(self.outputFilename, 'w', newline='') as f:
                    writer = csv.writer(f, delimiter='\t')
                    writer.writerows(self.parsedHbA1cData)

            #   Comma-Separated Values (if the user specifies an unknown file extension, data is also saved as CSV).
            else:
                with open(self.outputFilename, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerows(self.parsedHbA1cData)

        #   Raw binary data
        else:
            with open(self.outputFilename,'bw') as f:
                f.write(self.binaryBuffer)

    ##  DIALOG BOX: DATA SUCCESSFULLY WRITTEN TO DISK.
    def fileSuccessfullyWritten(self):
        myMessage = 'Successfully saved data in the file "' + self.outputFilename + '"'
        tk.messagebox.showinfo('Data saved',myMessage)

    ##  DIALOG BOX [USED IF WE CANNOT WRITE TO THE FILE THAT THE USER SPECIFIED]: ASK THE USER IF THEY WANT TO SAVE TO A DIFFERENT FILE.
    def doesUserStillWantToSaveData(self):
        title = 'Error'
        question = ('Could not save to the file "' + self.outputFilename +
                    '".\r\n\r\n' +
                    'Do you want to save to a different file?')
        self.userWantsToSaveData = tk.messagebox.askyesno(title,question,default='yes',icon='error')

    ##  DIALOG FOR SPECIFYING A NEW FILENAME, IN CASE WE CAN'T WRITE TO THE ONE SPECIFIED BY THE USER.
    def defineNewOutputFilename(self, mode='tabular', firstTime=False):
        #   Define the initial directory for the dialog box.
        initialDir = path.dirname(self.outputFilename)
        if not path.isdir(initialDir):
            initialDir = '\\'

        #   Get list of the filetypes that the user can choose (depending on whether we're saving tabular or binary data), and 
        if mode == 'tabular':
            possibleFiletypes = self.tabularFiletypes
        else:
            possibleFiletypes = [("Binary files", "*.bin"),('All files','*.*')]

        #   Do the following, unless it is the first time asking for a filename for a binary file:
        if not (mode=='binary' and firstTime==True):
            #   Re-use the old filename.
                initialFile = path.basename(self.outputFilename)

            #   Re-arrange the list of filetypes so that the previously selected filetype is the first item (meaning it will be pre-selected in the drop-down menu).
                oldExtension = list(path.splitext(initialFile))[1]
                extensionIndex = -1
                for index in range(len(possibleFiletypes)):
                    if possibleFiletypes[index][1] == '*' + oldExtension:
                        extensionIndex = index
                if extensionIndex == -1:
                    for index in range(len(possibleFiletypes)):
                        if possibleFiletypes[index][1] == '*.*':
                            extensionIndex = index
                firstFiletype = possibleFiletypes.pop(extensionIndex)
                possibleFiletypes.insert(0, firstFiletype)
        #   If we are asking for a binary file name for the first time, do this instead:
        else:
            initialFile = ''

        #   Set the default extension.
        if possibleFiletypes[0][1] != '*.*':
            defaultExtension = possibleFiletypes[0][1]
        else:
            defaultExtension = ''

        #   Open the dialog, parse the filename that was input and change output filename as appropriate.
        f = filedialog.asksaveasfilename(initialdir = initialDir, initialfile = initialFile, title = "Save as", filetypes = possibleFiletypes, defaultextension=defaultExtension)
        if f is None or f == '':
            self.userWantsToSaveData = False
        else:
            self.outputFilename = f
            
    ##  WRAPPER FUNCTION TO TEST THE INTEGRETY OF THE HBA1C DATA, PARSE IT AND WRITE IT TO DISK.
    def saveHbA1cData(self):
        self.checkDataIntegrity()
        if self.dataType == 'unknown':
            self.unknownDataStructureError()
        elif self.transmissionCompleted == False:
            self.incompleteTransmissionError()
        else:
            self.userWantsToSaveData = True
            self.saveOrAskForFilenameLoop(mode='tabular')

    ##  WRITE RAW BINARY DATA TO DISK.
    def saveBinaryData(self):
        self.userWantsToSaveData = True
        self.defineNewOutputFilename(mode='binary', firstTime=True)
        self.saveOrAskForFilenameLoop(mode='binary')

    ##  ENDLESS LOOP ASKING FOR A NEW FILENAME UNTIL IT IS EITHER POSSIBLE TO WRITE TO DISK, OR THE USER GIVES UP.
    def saveOrAskForFilenameLoop(self, mode='tabular'):
        while self.userWantsToSaveData == True:
            try:
                self.outputFileWriter(mode=mode)
                self.fileSuccessfullyWritten()
                self.userWantsToSaveData = False
            except (PermissionError, FileNotFoundError):
                self.doesUserStillWantToSaveData()
                if self.userWantsToSaveData == True:
                    self.defineNewOutputFilename(mode=mode, firstTime=False)

##  WINDOW FOR DISPLAYING HELP INFORMATION (INSTRUCTIONS FOR USE, ABOUT OR LICENSE INFO).
class helpWindow(simpledialog.Dialog):

    ## INITIALIZE THE WINDOW.
    def __init__(self, parent, name, helpType):
        tk.Toplevel.__init__(self, master=parent)
        self.name = name

        #   Grab important variable.
        self.helpType = helpType

        #   Create window and widgets.
        self.create_window()
        self.create_widgets()

    #   Create window.
    def create_window(self):
        #   Set focus on the new window.
        self.focus_set()

        #   Make sure that focus stays on the new window.
        self.grab_set()

        #   Only show one window in the taskbar.
        self.transient(self.master)

        #   Configure layout and position over parent window.
        self.helpBoxTitle = "HemoDownloader 1.2 - " + self.helpType
        self.title(self.helpBoxTitle)
        self.resizable(False, False)
        self.geometry(u'+{x}+{y}'.format(x = self.master.winfo_rootx(),
                                         y = self.master.winfo_rooty()))
        
        # Load text to be inserted into the box.
        if self.helpType == "Instructions for use":      
            self.text_to_be_inserted =  root.help_text
        elif self.helpType == "About HemoDownloader":      
            self.text_to_be_inserted =  root.about_text
        elif self.helpType == "License information":      
            self.text_to_be_inserted = root.license_text

        # Calculate height of text box.
        self.textbox_height = self.text_to_be_inserted.count('\n') + 1
        if self.textbox_height > 35:
            self.textbox_height = 35

        # Draw text box ith contents.
        self.textField = tk.Text(self,wrap=tk.WORD, height=self.textbox_height, width=80, padx=15, pady=15)
        self.textField.insert(tk.END, self.text_to_be_inserted)
        self.textField.grid(row = 0, column=0, sticky='W')

        # Configure scroll bar
        self.scrollbar = tk.Scrollbar(self, command=self.textField.yview, orient=tk.VERTICAL)
        self.scrollbar.config(command=self.textField.yview)
        self.textField.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.grid(row = 0, column=1, sticky='NS')
        
        #   If the user presses the ESCAPE button, the window will close.
        self.bind(u'<Escape>', self.closeHelpWindow)

    #   Creae a button for closing the window.
    def create_widgets(self):
        ttk.Button(self, text='Close', command=self.closeHelpWindow).grid(row=1, column=0,columnspan=2, pady=0)

    ##  Function for closing the help box.
    def closeHelpWindow(self, dummy='dummy'):
        self.destroy()

##  DEFINE MAIN WINDOW
root = tk.Tk()
feedback = settingsWindow(root)


##   HELP TEXT
root.help_text ="""HemoDownloader 1.2
A GUI utility for downloading data from HemoCue® HbA1c 501 devices

This program allows biochemical data to be exported from the device HemoCue® HbA1c 501 and saved in a structured database for further processing in statistical packages.

PLEASE NOTE THAT HEMODOWNLOADER IS INTENDED FOR USE IN EPIDEMIOLOGICAL STUDIES ONLY. THE SOFTWARE IS *NOT* APPROVED AS A MEDICAL DEVICE AND MUST NOT BE USED AS SUCH. THAT MEANS THE SOFTWARE MUST NOT BE USED ON HUMAN BEINGS FOR DIAGNOSIS, PREVENTION, MONITORING, PREDICTION, PROGNOSIS, TREATMENT OR ALLEVIATION OF DISEASE OR ANY OTHER MEDICAL PURPOSES.


How to transfer all results in device memory:
 1. Click 'RECEIVE DATA'.
 2. Connect HemoCue® HbA1c 501 device to computer using RS232 null modem cable.
 3. Turn on the device and wait for it to go into stand-by mode.
 4. Press 'MODE' button for 3 seconds.
 5. The device will display 'Set up' and 'Data'.
 6. Press 'DOWN' (▼) or 'UP' (▲) button to select 'Data'.
 7. Press 'MODE' button to display the results in memory.
 8. Press 'PRINTER' button.
 9. Using 'DOWN' (▼) or 'UP' (▲) button, select 'All'.
10. Press 'MODE' button.
11. Wait for data to be transferred to computer.


How to transfer individual results:
 1. Click 'RECEIVE DATA'.
 2. Connect HemoCue® HbA1c 501 device to computer using RS232 null modem cable.
 3. Turn on the device and wait for it to go into stand-by mode.
 4. Press 'MODE' button for 3 seconds.
 5. The device will display 'Set up' and 'Data'.
 6. Press 'DOWN' (▼) or 'UP' (▲) button to select 'Data'.
 7. Press 'MODE' button to display the results in memory.
 8. Scroll through test results using 'DOWN' (▼) and 'UP' (▲) buttons.
 9. When the desired record is shown, press 'PRINTER' button.
10. Using 'DOWN' (▼) or 'UP' (▲) button, select 'Current'.
11. Press 'MODE' button.
12. Wait for data to be transferred to computer."""

##   ABOUT TEXT
root.about_text = """HemoDownloader 1.2
A GUI utility for downloading data from HemoCue® HbA1c 501 devices

Copyright © 2018-2021 Martin Rune Hassan Hansen

To show software license, select 'Help' > 'License information'.


Contact information

email address: martinrunehassanhansen@ph.au.dk

Physical addresses:

  Section for Environment, Work and Health
  Att: Martin Rune Hassan Hansen
  Department of Public Health
  Aarhus University
  Bartholins Allé 2
  DK-8000 Aarhus C
  Denmark

  and 

  The National Research Centre for the Working Environment
  Att: Martin Rune Hassan Hansen
  Lersø Parkallé 105
  DK-2100 København
  Denmark"""

##   LICENSE TEXT
root.license_text = """********************************************************************************
*** HemoDownloader 1.2                                                       ***
********************************************************************************
Copyright © 2018-2021 Martin Rune Hassan Hansen

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.

PLEASE NOTE THAT HEMODOWNLOADER IS INTENDED FOR USE IN EPIDEMIOLOGICAL STUDIES ONLY. THE SOFTWARE IS *NOT* APPROVED AS A MEDICAL DEVICE AND MUST NOT BE USED AS SUCH. THAT MEANS THE SOFTWARE MUST NOT BE USED ON HUMAN BEINGS FOR DIAGNOSIS, PREVENTION, MONITORING, PREDICTION, PROGNOSIS, TREATMENT OR ALLEVIATION OF DISEASE OR ANY OTHER MEDICAL PURPOSES.

HemoDownloader uses the following third-party python modules that are covered by
their own licenses listed below:
* pyserial
* XlsxWriter
* xlwt

The compiled versions of HemoDownloader 1.2 were created using the program PyInstaller that is distributed under a modified GPL license, also listed below.

HemoCue is a registered trademark of HemoCue AB (Ängelholm, Sweden).

********************************************************************************
*** GNU GENERAL PUBLIC LICENSE version 3                                     ***
********************************************************************************

GNU GENERAL PUBLIC LICENSE

Version 3, 29 June 2007

Copyright © 2007 Free Software Foundation, Inc. <https://fsf.org/>

Everyone is permitted to copy and distribute verbatim copies of this license document, but changing it is not allowed.

Preamble

The GNU General Public License is a free, copyleft license for software and other kinds of works.

The licenses for most software and other practical works are designed to take away your freedom to share and change the works. By contrast, the GNU General Public License is intended to guarantee your freedom to share and change all versions of a program--to make sure it remains free software for all its users. We, the Free Software Foundation, use the GNU General Public License for most of our software; it applies also to any other work released this way by its authors. You can apply it to your programs, too.

When we speak of free software, we are referring to freedom, not price. Our General Public Licenses are designed to make sure that you have the freedom to distribute copies of free software (and charge for them if you wish), that you receive source code or can get it if you want it, that you can change the software or use pieces of it in new free programs, and that you know you can do these things.

To protect your rights, we need to prevent others from denying you these rights or asking you to surrender the rights. Therefore, you have certain responsibilities if you distribute copies of the software, or if you modify it: responsibilities to respect the freedom of others.

For example, if you distribute copies of such a program, whether gratis or for a fee, you must pass on to the recipients the same freedoms that you received. You must make sure that they, too, receive or can get the source code. And you must show them these terms so they know their rights.

Developers that use the GNU GPL protect your rights with two steps: (1) assert copyright on the software, and (2) offer you this License giving you legal permission to copy, distribute and/or modify it.

For the developers' and authors' protection, the GPL clearly explains that there is no warranty for this free software. For both users' and authors' sake, the GPL requires that modified versions be marked as changed, so that their problems will not be attributed erroneously to authors of previous versions.

Some devices are designed to deny users access to install or run modified versions of the software inside them, although the manufacturer can do so. This is fundamentally incompatible with the aim of protecting users' freedom to change the software. The systematic pattern of such abuse occurs in the area of products for individuals to use, which is precisely where it is most unacceptable. Therefore, we have designed this version of the GPL to prohibit the practice for those products. If such problems arise substantially in other domains, we stand ready to extend this provision to those domains in future versions of the GPL, as needed to protect the freedom of users.

Finally, every program is threatened constantly by software patents. States should not allow patents to restrict development and use of software on general-purpose computers, but in those that do, we wish to avoid the special danger that patents applied to a free program could make it effectively proprietary. To prevent this, the GPL assures that patents cannot be used to render the program non-free.

The precise terms and conditions for copying, distribution and modification follow.

TERMS AND CONDITIONS

0. Definitions.

“This License” refers to version 3 of the GNU General Public License.

“Copyright” also means copyright-like laws that apply to other kinds of works, such as semiconductor masks.

“The Program” refers to any copyrightable work licensed under this License. Each licensee is addressed as “you”. “Licensees” and “recipients” may be individuals or organizations.

To “modify” a work means to copy from or adapt all or part of the work in a fashion requiring copyright permission, other than the making of an exact copy. The resulting work is called a “modified version” of the earlier work or a work “based on” the earlier work.

A “covered work” means either the unmodified Program or a work based on the Program.

To “propagate” a work means to do anything with it that, without permission, would make you directly or secondarily liable for infringement under applicable copyright law, except executing it on a computer or modifying a private copy. Propagation includes copying, distribution (with or without modification), making available to the public, and in some countries other activities as well.

To “convey” a work means any kind of propagation that enables other parties to make or receive copies. Mere interaction with a user through a computer network, with no transfer of a copy, is not conveying.

An interactive user interface displays “Appropriate Legal Notices” to the extent that it includes a convenient and prominently visible feature that (1) displays an appropriate copyright notice, and (2) tells the user that there is no warranty for the work (except to the extent that warranties are provided), that licensees may convey the work under this License, and how to view a copy of this License. If the interface presents a list of user commands or options, such as a menu, a prominent item in the list meets this criterion.

1. Source Code.

The “source code” for a work means the preferred form of the work for making modifications to it. “Object code” means any non-source form of a work.

A “Standard Interface” means an interface that either is an official standard defined by a recognized standards body, or, in the case of interfaces specified for a particular programming language, one that is widely used among developers working in that language.

The “System Libraries” of an executable work include anything, other than the work as a whole, that (a) is included in the normal form of packaging a Major Component, but which is not part of that Major Component, and (b) serves only to enable use of the work with that Major Component, or to implement a Standard Interface for which an implementation is available to the public in source code form. A “Major Component”, in this context, means a major essential component (kernel, window system, and so on) of the specific operating system (if any) on which the executable work runs, or a compiler used to produce the work, or an object code interpreter used to run it.

The “Corresponding Source” for a work in object code form means all the source code needed to generate, install, and (for an executable work) run the object code and to modify the work, including scripts to control those activities. However, it does not include the work's System Libraries, or general-purpose tools or generally available free programs which are used unmodified in performing those activities but which are not part of the work. For example, Corresponding Source includes interface definition files associated with source files for the work, and the source code for shared libraries and dynamically linked subprograms that the work is specifically designed to require, such as by intimate data communication or control flow between those subprograms and other parts of the work.

The Corresponding Source need not include anything that users can regenerate automatically from other parts of the Corresponding Source.

The Corresponding Source for a work in source code form is that same work.

2. Basic Permissions.

All rights granted under this License are granted for the term of copyright on the Program, and are irrevocable provided the stated conditions are met. This License explicitly affirms your unlimited permission to run the unmodified Program. The output from running a covered work is covered by this License only if the output, given its content, constitutes a covered work. This License acknowledges your rights of fair use or other equivalent, as provided by copyright law.

You may make, run and propagate covered works that you do not convey, without conditions so long as your license otherwise remains in force. You may convey covered works to others for the sole purpose of having them make modifications exclusively for you, or provide you with facilities for running those works, provided that you comply with the terms of this License in conveying all material for which you do not control copyright. Those thus making or running the covered works for you must do so exclusively on your behalf, under your direction and control, on terms that prohibit them from making any copies of your copyrighted material outside their relationship with you.

Conveying under any other circumstances is permitted solely under the conditions stated below. Sublicensing is not allowed; section 10 makes it unnecessary.

3. Protecting Users' Legal Rights From Anti-Circumvention Law.

No covered work shall be deemed part of an effective technological measure under any applicable law fulfilling obligations under article 11 of the WIPO copyright treaty adopted on 20 December 1996, or similar laws prohibiting or restricting circumvention of such measures.

When you convey a covered work, you waive any legal power to forbid circumvention of technological measures to the extent such circumvention is effected by exercising rights under this License with respect to the covered work, and you disclaim any intention to limit operation or modification of the work as a means of enforcing, against the work's users, your or third parties' legal rights to forbid circumvention of technological measures.

4. Conveying Verbatim Copies.

You may convey verbatim copies of the Program's source code as you receive it, in any medium, provided that you conspicuously and appropriately publish on each copy an appropriate copyright notice; keep intact all notices stating that this License and any non-permissive terms added in accord with section 7 apply to the code; keep intact all notices of the absence of any warranty; and give all recipients a copy of this License along with the Program.

You may charge any price or no price for each copy that you convey, and you may offer support or warranty protection for a fee.

5. Conveying Modified Source Versions.

You may convey a work based on the Program, or the modifications to produce it from the Program, in the form of source code under the terms of section 4, provided that you also meet all of these conditions:

a) The work must carry prominent notices stating that you modified it, and giving a relevant date.
b) The work must carry prominent notices stating that it is released under this License and any conditions added under section 7. This requirement modifies the requirement in section 4 to “keep intact all notices”.
c) You must license the entire work, as a whole, under this License to anyone who comes into possession of a copy. This License will therefore apply, along with any applicable section 7 additional terms, to the whole of the work, and all its parts, regardless of how they are packaged. This License gives no permission to license the work in any other way, but it does not invalidate such permission if you have separately received it.
d) If the work has interactive user interfaces, each must display Appropriate Legal Notices; however, if the Program has interactive interfaces that do not display Appropriate Legal Notices, your work need not make them do so.
A compilation of a covered work with other separate and independent works, which are not by their nature extensions of the covered work, and which are not combined with it such as to form a larger program, in or on a volume of a storage or distribution medium, is called an “aggregate” if the compilation and its resulting copyright are not used to limit the access or legal rights of the compilation's users beyond what the individual works permit. Inclusion of a covered work in an aggregate does not cause this License to apply to the other parts of the aggregate.

6. Conveying Non-Source Forms.

You may convey a covered work in object code form under the terms of sections 4 and 5, provided that you also convey the machine-readable Corresponding Source under the terms of this License, in one of these ways:

a) Convey the object code in, or embodied in, a physical product (including a physical distribution medium), accompanied by the Corresponding Source fixed on a durable physical medium customarily used for software interchange.
b) Convey the object code in, or embodied in, a physical product (including a physical distribution medium), accompanied by a written offer, valid for at least three years and valid for as long as you offer spare parts or customer support for that product model, to give anyone who possesses the object code either (1) a copy of the Corresponding Source for all the software in the product that is covered by this License, on a durable physical medium customarily used for software interchange, for a price no more than your reasonable cost of physically performing this conveying of source, or (2) access to copy the Corresponding Source from a network server at no charge.
c) Convey individual copies of the object code with a copy of the written offer to provide the Corresponding Source. This alternative is allowed only occasionally and noncommercially, and only if you received the object code with such an offer, in accord with subsection 6b.
d) Convey the object code by offering access from a designated place (gratis or for a charge), and offer equivalent access to the Corresponding Source in the same way through the same place at no further charge. You need not require recipients to copy the Corresponding Source along with the object code. If the place to copy the object code is a network server, the Corresponding Source may be on a different server (operated by you or a third party) that supports equivalent copying facilities, provided you maintain clear directions next to the object code saying where to find the Corresponding Source. Regardless of what server hosts the Corresponding Source, you remain obligated to ensure that it is available for as long as needed to satisfy these requirements.
e) Convey the object code using peer-to-peer transmission, provided you inform other peers where the object code and Corresponding Source of the work are being offered to the general public at no charge under subsection 6d.
A separable portion of the object code, whose source code is excluded from the Corresponding Source as a System Library, need not be included in conveying the object code work.

A “User Product” is either (1) a “consumer product”, which means any tangible personal property which is normally used for personal, family, or household purposes, or (2) anything designed or sold for incorporation into a dwelling. In determining whether a product is a consumer product, doubtful cases shall be resolved in favor of coverage. For a particular product received by a particular user, “normally used” refers to a typical or common use of that class of product, regardless of the status of the particular user or of the way in which the particular user actually uses, or expects or is expected to use, the product. A product is a consumer product regardless of whether the product has substantial commercial, industrial or non-consumer uses, unless such uses represent the only significant mode of use of the product.

“Installation Information” for a User Product means any methods, procedures, authorization keys, or other information required to install and execute modified versions of a covered work in that User Product from a modified version of its Corresponding Source. The information must suffice to ensure that the continued functioning of the modified object code is in no case prevented or interfered with solely because modification has been made.

If you convey an object code work under this section in, or with, or specifically for use in, a User Product, and the conveying occurs as part of a transaction in which the right of possession and use of the User Product is transferred to the recipient in perpetuity or for a fixed term (regardless of how the transaction is characterized), the Corresponding Source conveyed under this section must be accompanied by the Installation Information. But this requirement does not apply if neither you nor any third party retains the ability to install modified object code on the User Product (for example, the work has been installed in ROM).

The requirement to provide Installation Information does not include a requirement to continue to provide support service, warranty, or updates for a work that has been modified or installed by the recipient, or for the User Product in which it has been modified or installed. Access to a network may be denied when the modification itself materially and adversely affects the operation of the network or violates the rules and protocols for communication across the network.

Corresponding Source conveyed, and Installation Information provided, in accord with this section must be in a format that is publicly documented (and with an implementation available to the public in source code form), and must require no special password or key for unpacking, reading or copying.

7. Additional Terms.

“Additional permissions” are terms that supplement the terms of this License by making exceptions from one or more of its conditions. Additional permissions that are applicable to the entire Program shall be treated as though they were included in this License, to the extent that they are valid under applicable law. If additional permissions apply only to part of the Program, that part may be used separately under those permissions, but the entire Program remains governed by this License without regard to the additional permissions.

When you convey a copy of a covered work, you may at your option remove any additional permissions from that copy, or from any part of it. (Additional permissions may be written to require their own removal in certain cases when you modify the work.) You may place additional permissions on material, added by you to a covered work, for which you have or can give appropriate copyright permission.

Notwithstanding any other provision of this License, for material you add to a covered work, you may (if authorized by the copyright holders of that material) supplement the terms of this License with terms:

a) Disclaiming warranty or limiting liability differently from the terms of sections 15 and 16 of this License; or
b) Requiring preservation of specified reasonable legal notices or author attributions in that material or in the Appropriate Legal Notices displayed by works containing it; or
c) Prohibiting misrepresentation of the origin of that material, or requiring that modified versions of such material be marked in reasonable ways as different from the original version; or
d) Limiting the use for publicity purposes of names of licensors or authors of the material; or
e) Declining to grant rights under trademark law for use of some trade names, trademarks, or service marks; or
f) Requiring indemnification of licensors and authors of that material by anyone who conveys the material (or modified versions of it) with contractual assumptions of liability to the recipient, for any liability that these contractual assumptions directly impose on those licensors and authors.
All other non-permissive additional terms are considered “further restrictions” within the meaning of section 10. If the Program as you received it, or any part of it, contains a notice stating that it is governed by this License along with a term that is a further restriction, you may remove that term. If a license document contains a further restriction but permits relicensing or conveying under this License, you may add to a covered work material governed by the terms of that license document, provided that the further restriction does not survive such relicensing or conveying.

If you add terms to a covered work in accord with this section, you must place, in the relevant source files, a statement of the additional terms that apply to those files, or a notice indicating where to find the applicable terms.

Additional terms, permissive or non-permissive, may be stated in the form of a separately written license, or stated as exceptions; the above requirements apply either way.

8. Termination.

You may not propagate or modify a covered work except as expressly provided under this License. Any attempt otherwise to propagate or modify it is void, and will automatically terminate your rights under this License (including any patent licenses granted under the third paragraph of section 11).

However, if you cease all violation of this License, then your license from a particular copyright holder is reinstated (a) provisionally, unless and until the copyright holder explicitly and finally terminates your license, and (b) permanently, if the copyright holder fails to notify you of the violation by some reasonable means prior to 60 days after the cessation.

Moreover, your license from a particular copyright holder is reinstated permanently if the copyright holder notifies you of the violation by some reasonable means, this is the first time you have received notice of violation of this License (for any work) from that copyright holder, and you cure the violation prior to 30 days after your receipt of the notice.

Termination of your rights under this section does not terminate the licenses of parties who have received copies or rights from you under this License. If your rights have been terminated and not permanently reinstated, you do not qualify to receive new licenses for the same material under section 10.

9. Acceptance Not Required for Having Copies.

You are not required to accept this License in order to receive or run a copy of the Program. Ancillary propagation of a covered work occurring solely as a consequence of using peer-to-peer transmission to receive a copy likewise does not require acceptance. However, nothing other than this License grants you permission to propagate or modify any covered work. These actions infringe copyright if you do not accept this License. Therefore, by modifying or propagating a covered work, you indicate your acceptance of this License to do so.

10. Automatic Licensing of Downstream Recipients.

Each time you convey a covered work, the recipient automatically receives a license from the original licensors, to run, modify and propagate that work, subject to this License. You are not responsible for enforcing compliance by third parties with this License.

An “entity transaction” is a transaction transferring control of an organization, or substantially all assets of one, or subdividing an organization, or merging organizations. If propagation of a covered work results from an entity transaction, each party to that transaction who receives a copy of the work also receives whatever licenses to the work the party's predecessor in interest had or could give under the previous paragraph, plus a right to possession of the Corresponding Source of the work from the predecessor in interest, if the predecessor has it or can get it with reasonable efforts.

You may not impose any further restrictions on the exercise of the rights granted or affirmed under this License. For example, you may not impose a license fee, royalty, or other charge for exercise of rights granted under this License, and you may not initiate litigation (including a cross-claim or counterclaim in a lawsuit) alleging that any patent claim is infringed by making, using, selling, offering for sale, or importing the Program or any portion of it.

11. Patents.

A “contributor” is a copyright holder who authorizes use under this License of the Program or a work on which the Program is based. The work thus licensed is called the contributor's “contributor version”.

A contributor's “essential patent claims” are all patent claims owned or controlled by the contributor, whether already acquired or hereafter acquired, that would be infringed by some manner, permitted by this License, of making, using, or selling its contributor version, but do not include claims that would be infringed only as a consequence of further modification of the contributor version. For purposes of this definition, “control” includes the right to grant patent sublicenses in a manner consistent with the requirements of this License.

Each contributor grants you a non-exclusive, worldwide, royalty-free patent license under the contributor's essential patent claims, to make, use, sell, offer for sale, import and otherwise run, modify and propagate the contents of its contributor version.

In the following three paragraphs, a “patent license” is any express agreement or commitment, however denominated, not to enforce a patent (such as an express permission to practice a patent or covenant not to sue for patent infringement). To “grant” such a patent license to a party means to make such an agreement or commitment not to enforce a patent against the party.

If you convey a covered work, knowingly relying on a patent license, and the Corresponding Source of the work is not available for anyone to copy, free of charge and under the terms of this License, through a publicly available network server or other readily accessible means, then you must either (1) cause the Corresponding Source to be so available, or (2) arrange to deprive yourself of the benefit of the patent license for this particular work, or (3) arrange, in a manner consistent with the requirements of this License, to extend the patent license to downstream recipients. “Knowingly relying” means you have actual knowledge that, but for the patent license, your conveying the covered work in a country, or your recipient's use of the covered work in a country, would infringe one or more identifiable patents in that country that you have reason to believe are valid.

If, pursuant to or in connection with a single transaction or arrangement, you convey, or propagate by procuring conveyance of, a covered work, and grant a patent license to some of the parties receiving the covered work authorizing them to use, propagate, modify or convey a specific copy of the covered work, then the patent license you grant is automatically extended to all recipients of the covered work and works based on it.

A patent license is “discriminatory” if it does not include within the scope of its coverage, prohibits the exercise of, or is conditioned on the non-exercise of one or more of the rights that are specifically granted under this License. You may not convey a covered work if you are a party to an arrangement with a third party that is in the business of distributing software, under which you make payment to the third party based on the extent of your activity of conveying the work, and under which the third party grants, to any of the parties who would receive the covered work from you, a discriminatory patent license (a) in connection with copies of the covered work conveyed by you (or copies made from those copies), or (b) primarily for and in connection with specific products or compilations that contain the covered work, unless you entered into that arrangement, or that patent license was granted, prior to 28 March 2007.

Nothing in this License shall be construed as excluding or limiting any implied license or other defenses to infringement that may otherwise be available to you under applicable patent law.

12. No Surrender of Others' Freedom.

If conditions are imposed on you (whether by court order, agreement or otherwise) that contradict the conditions of this License, they do not excuse you from the conditions of this License. If you cannot convey a covered work so as to satisfy simultaneously your obligations under this License and any other pertinent obligations, then as a consequence you may not convey it at all. For example, if you agree to terms that obligate you to collect a royalty for further conveying from those to whom you convey the Program, the only way you could satisfy both those terms and this License would be to refrain entirely from conveying the Program.

13. Use with the GNU Affero General Public License.

Notwithstanding any other provision of this License, you have permission to link or combine any covered work with a work licensed under version 3 of the GNU Affero General Public License into a single combined work, and to convey the resulting work. The terms of this License will continue to apply to the part which is the covered work, but the special requirements of the GNU Affero General Public License, section 13, concerning interaction through a network will apply to the combination as such.

14. Revised Versions of this License.

The Free Software Foundation may publish revised and/or new versions of the GNU General Public License from time to time. Such new versions will be similar in spirit to the present version, but may differ in detail to address new problems or concerns.

Each version is given a distinguishing version number. If the Program specifies that a certain numbered version of the GNU General Public License “or any later version” applies to it, you have the option of following the terms and conditions either of that numbered version or of any later version published by the Free Software Foundation. If the Program does not specify a version number of the GNU General Public License, you may choose any version ever published by the Free Software Foundation.

If the Program specifies that a proxy can decide which future versions of the GNU General Public License can be used, that proxy's public statement of acceptance of a version permanently authorizes you to choose that version for the Program.

Later license versions may give you additional or different permissions. However, no additional obligations are imposed on any author or copyright holder as a result of your choosing to follow a later version.

15. Disclaimer of Warranty.

THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY APPLICABLE LAW. EXCEPT WHEN OTHERWISE STATED IN WRITING THE COPYRIGHT HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM “AS IS” WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM IS WITH YOU. SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF ALL NECESSARY SERVICING, REPAIR OR CORRECTION.

16. Limitation of Liability.

IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MODIFIES AND/OR CONVEYS THE PROGRAM AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES, INCLUDING ANY GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE USE OR INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED TO LOSS OF DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY YOU OR THIRD PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER PROGRAMS), EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.

17. Interpretation of Sections 15 and 16.

If the disclaimer of warranty and limitation of liability provided above cannot be given local legal effect according to their terms, reviewing courts shall apply local law that most closely approximates an absolute waiver of all civil liability in connection with the Program, unless a warranty or assumption of liability accompanies a copy of the Program in return for a fee.

END OF TERMS AND CONDITIONS


********************************************************************************
*** PYSERIAL LICENSE                                                         ***
********************************************************************************

Copyright (c) 2001-2016 Chris Liechti <cliechti@gmx.net>
All Rights Reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

  * Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.

  * Redistributions in binary form must reproduce the above
    copyright notice, this list of conditions and the following
    disclaimer in the documentation and/or other materials provided
    with the distribution.

  * Neither the name of the copyright holder nor the names of its
    contributors may be used to endorse or promote products derived
    from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

---------------------------------------------------------------------------
Note:
Individual files contain the following tag instead of the full license text.

    SPDX-License-Identifier:    BSD-3-Clause

This enables machine processing of license information based on the SPDX
License Identifiers that are here available: http://spdx.org/licenses/


********************************************************************************
*** XSLXWRITER LICENSE                                                       ***
********************************************************************************

Copyright (c) 2013, John McNamara <jmcnamara@cpan.org>
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

The views and conclusions contained in the software and documentation are those
of the authors and should not be interpreted as representing official policies,
either expressed or implied, of the FreeBSD Project.


********************************************************************************
*** XLWT LICENSE                                                             ***
********************************************************************************

xlwt has various licenses that apply to the different parts of it, they are
listed below:

The license for the work John Machin has done since xlwt was created::

    Portions copyright (c) 2007, Stephen John Machin, Lingfo Pty Ltd
    All rights reserved.

    Redistribution and use in source and binary forms, with or without
    modification, are permitted provided that the following conditions are met:

    1. Redistributions of source code must retain the above copyright notice,
    this list of conditions and the following disclaimer.

    2. Redistributions in binary form must reproduce the above copyright notice,
    this list of conditions and the following disclaimer in the documentation
    and/or other materials provided with the distribution.

    3. None of the names of Stephen John Machin, Lingfo Pty Ltd and any
    contributors may be used to endorse or promote products derived from this
    software without specific prior written permission.

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
    AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
    PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR
    CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
    EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
    PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
    OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
    WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
    OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
    ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

The licensing for the unit tests added as part of the work for Python 3
compatibility is as follows::

  Author:  mozman --<mozman@gmx.at>
  Purpose: test_mini
  Created: 03.12.2010
  Copyright (C) 2010, Manfred Moitzi
  License: BSD licence

The license for pyExcelerator, from which xlwt was forked::

      Copyright (C) 2005 Roman V. Kiseliov
      All rights reserved.

      Redistribution and use in source and binary forms, with or without
      modification, are permitted provided that the following conditions
      are met:

      1. Redistributions of source code must retain the above copyright
         notice, this list of conditions and the following disclaimer.

      2. Redistributions in binary form must reproduce the above copyright
         notice, this list of conditions and the following disclaimer in
         the documentation and/or other materials provided with the
         distribution.

      3. All advertising materials mentioning features or use of this
         software must display the following acknowledgment:
         "This product includes software developed by
          Roman V. Kiseliov <roman@kiseliov.ru>."

      4. Redistributions of any form whatsoever must retain the following
         acknowledgment:
         "This product includes software developed by
          Roman V. Kiseliov <roman@kiseliov.ru>."

      THIS SOFTWARE IS PROVIDED BY Roman V. Kiseliov ``AS IS'' AND ANY
      EXPRESSED OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
      IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
      PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL Roman V. Kiseliov OR
      ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
      SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
      NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
      LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
      HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
      STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
      ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
      OF THE POSSIBILITY OF SUCH DAMAGE.

    Roman V. Kiseliov
    Russia
    Kursk
    Libknecht St., 4

    +7(0712)56-09-83

    <roman@kiseliov.ru>

Portions of xlwt.Utils are based on pyXLWriter which is licensed as follows::

 Copyright (c) 2004 Evgeny Filatov <fufff@users.sourceforge.net>
 Copyright (c) 2002-2004 John McNamara (Perl Spreadsheet::WriteExcel)

 This library is free software; you can redistribute it and/or modify it
 under the terms of the GNU Lesser General Public License as published by
 the Free Software Foundation; either version 2.1 of the License, or
 (at your option) any later version.

 This library is distributed in the hope that it will be useful, but
 WITHOUT ANY WARRANTY; without even the implied warranty of
 MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser
 General Public License for more details:

 https://www.gnu.org/licenses/lgpl.html

pyXLWriter also makes reference to the PERL Spreadsheet::WriteExcel as follows::

  This module was written/ported from PERL Spreadsheet::WriteExcel module
  The author of the PERL Spreadsheet::WriteExcel module is John McNamara
  <jmcnamara@cpan.org>

********************************************************************************
*** PYINSTALLER LICENSE                                                      ***
********************************************************************************

================================
 The PyInstaller licensing terms
================================
 

Copyright (c) 2010-2019, PyInstaller Development Team
Copyright (c) 2005-2009, Giovanni Bajo
Based on previous work under copyright (c) 2002 McMillan Enterprises, Inc.


PyInstaller is licensed under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2 of the License,
or (at your option) any later version.


Bootloader Exception
--------------------

In addition to the permissions in the GNU General Public License, the
authors give you unlimited permission to link or embed compiled bootloader
and related files into combinations with other programs, and to distribute
those combinations without any restriction coming from the use of those
files. (The General Public License restrictions do apply in other respects;
for example, they cover modification of the files, and distribution when
not linked into a combine executable.)
 
 
Bootloader and Related Files
----------------------------

Bootloader and related files are files which are embedded within the
final executable. This includes files in directories:

./bootloader/
./PyInstaller/loader

 
About the PyInstaller Development Team
--------------------------------------

The PyInstaller Development Team is the set of contributors
to the PyInstaller project. A full list with details is kept
in the documentation directory, in the file
``doc/CREDITS.rst``.

The core team that coordinates development on GitHub can be found here:
https://github.com/pyinstaller/pyinstaller.  As of 2015, it consists of:

* Hartmut Goebel
* Martin Zibricky
* David Vierra
* David Cortesi


Our Copyright Policy
--------------------

PyInstaller uses a shared copyright model. Each contributor maintains copyright
over their contributions to PyInstaller. But, it is important to note that these
contributions are typically only changes to the repositories. Thus,
the PyInstaller source code, in its entirety is not the copyright of any single
person or institution.  Instead, it is the collective copyright of the entire
PyInstaller Development Team.  If individual contributors want to maintain
a record of what changes/contributions they have specific copyright on, they
should indicate their copyright in the commit message of the change, when they
commit the change to the PyInstaller repository.

With this in mind, the following banner should be used in any source code file
to indicate the copyright and license terms:


#-----------------------------------------------------------------------------
# Copyright (c) 2005-20l5, PyInstaller Development Team.
#
# Distributed under the terms of the GNU General Public License with exception
# for distributing bootloader.
#
# The full license is in the file COPYING.txt, distributed with this software.
#-----------------------------------------------------------------------------



GNU General Public License
--------------------------

https://gnu.org/licenses/gpl-2.0.html


		    GNU GENERAL PUBLIC LICENSE
		       Version 2, June 1991

 Copyright (C) 1989, 1991 Free Software Foundation, Inc.
                 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA
 Everyone is permitted to copy and distribute verbatim copies
 of this license document, but changing it is not allowed.

			    Preamble

  The licenses for most software are designed to take away your
freedom to share and change it.  By contrast, the GNU General Public
License is intended to guarantee your freedom to share and change free
software--to make sure the software is free for all its users.  This
General Public License applies to most of the Free Software
Foundation's software and to any other program whose authors commit to
using it.  (Some other Free Software Foundation software is covered by
the GNU Library General Public License instead.)  You can apply it to
your programs, too.

  When we speak of free software, we are referring to freedom, not
price.  Our General Public Licenses are designed to make sure that you
have the freedom to distribute copies of free software (and charge for
this service if you wish), that you receive source code or can get it
if you want it, that you can change the software or use pieces of it
in new free programs; and that you know you can do these things.

  To protect your rights, we need to make restrictions that forbid
anyone to deny you these rights or to ask you to surrender the rights.
These restrictions translate to certain responsibilities for you if you
distribute copies of the software, or if you modify it.

  For example, if you distribute copies of such a program, whether
gratis or for a fee, you must give the recipients all the rights that
you have.  You must make sure that they, too, receive or can get the
source code.  And you must show them these terms so they know their
rights.

  We protect your rights with two steps: (1) copyright the software, and
(2) offer you this license which gives you legal permission to copy,
distribute and/or modify the software.

  Also, for each author's protection and ours, we want to make certain
that everyone understands that there is no warranty for this free
software.  If the software is modified by someone else and passed on, we
want its recipients to know that what they have is not the original, so
that any problems introduced by others will not reflect on the original
authors' reputations.

  Finally, any free program is threatened constantly by software
patents.  We wish to avoid the danger that redistributors of a free
program will individually obtain patent licenses, in effect making the
program proprietary.  To prevent this, we have made it clear that any
patent must be licensed for everyone's free use or not licensed at all.

  The precise terms and conditions for copying, distribution and
modification follow.

		    GNU GENERAL PUBLIC LICENSE
   TERMS AND CONDITIONS FOR COPYING, DISTRIBUTION AND MODIFICATION

  0. This License applies to any program or other work which contains
a notice placed by the copyright holder saying it may be distributed
under the terms of this General Public License.  The "Program", below,
refers to any such program or work, and a "work based on the Program"
means either the Program or any derivative work under copyright law:
that is to say, a work containing the Program or a portion of it,
either verbatim or with modifications and/or translated into another
language.  (Hereinafter, translation is included without limitation in
the term "modification".)  Each licensee is addressed as "you".

Activities other than copying, distribution and modification are not
covered by this License; they are outside its scope.  The act of
running the Program is not restricted, and the output from the Program
is covered only if its contents constitute a work based on the
Program (independent of having been made by running the Program).
Whether that is true depends on what the Program does.

  1. You may copy and distribute verbatim copies of the Program's
source code as you receive it, in any medium, provided that you
conspicuously and appropriately publish on each copy an appropriate
copyright notice and disclaimer of warranty; keep intact all the
notices that refer to this License and to the absence of any warranty;
and give any other recipients of the Program a copy of this License
along with the Program.

You may charge a fee for the physical act of transferring a copy, and
you may at your option offer warranty protection in exchange for a fee.

  2. You may modify your copy or copies of the Program or any portion
of it, thus forming a work based on the Program, and copy and
distribute such modifications or work under the terms of Section 1
above, provided that you also meet all of these conditions:

    a) You must cause the modified files to carry prominent notices
    stating that you changed the files and the date of any change.

    b) You must cause any work that you distribute or publish, that in
    whole or in part contains or is derived from the Program or any
    part thereof, to be licensed as a whole at no charge to all third
    parties under the terms of this License.

    c) If the modified program normally reads commands interactively
    when run, you must cause it, when started running for such
    interactive use in the most ordinary way, to print or display an
    announcement including an appropriate copyright notice and a
    notice that there is no warranty (or else, saying that you provide
    a warranty) and that users may redistribute the program under
    these conditions, and telling the user how to view a copy of this
    License.  (Exception: if the Program itself is interactive but
    does not normally print such an announcement, your work based on
    the Program is not required to print an announcement.)

These requirements apply to the modified work as a whole.  If
identifiable sections of that work are not derived from the Program,
and can be reasonably considered independent and separate works in
themselves, then this License, and its terms, do not apply to those
sections when you distribute them as separate works.  But when you
distribute the same sections as part of a whole which is a work based
on the Program, the distribution of the whole must be on the terms of
this License, whose permissions for other licensees extend to the
entire whole, and thus to each and every part regardless of who wrote it.

Thus, it is not the intent of this section to claim rights or contest
your rights to work written entirely by you; rather, the intent is to
exercise the right to control the distribution of derivative or
collective works based on the Program.

In addition, mere aggregation of another work not based on the Program
with the Program (or with a work based on the Program) on a volume of
a storage or distribution medium does not bring the other work under
the scope of this License.

  3. You may copy and distribute the Program (or a work based on it,
under Section 2) in object code or executable form under the terms of
Sections 1 and 2 above provided that you also do one of the following:

    a) Accompany it with the complete corresponding machine-readable
    source code, which must be distributed under the terms of Sections
    1 and 2 above on a medium customarily used for software interchange; or,

    b) Accompany it with a written offer, valid for at least three
    years, to give any third party, for a charge no more than your
    cost of physically performing source distribution, a complete
    machine-readable copy of the corresponding source code, to be
    distributed under the terms of Sections 1 and 2 above on a medium
    customarily used for software interchange; or,

    c) Accompany it with the information you received as to the offer
    to distribute corresponding source code.  (This alternative is
    allowed only for noncommercial distribution and only if you
    received the program in object code or executable form with such
    an offer, in accord with Subsection b above.)

The source code for a work means the preferred form of the work for
making modifications to it.  For an executable work, complete source
code means all the source code for all modules it contains, plus any
associated interface definition files, plus the scripts used to
control compilation and installation of the executable.  However, as a
special exception, the source code distributed need not include
anything that is normally distributed (in either source or binary
form) with the major components (compiler, kernel, and so on) of the
operating system on which the executable runs, unless that component
itself accompanies the executable.

If distribution of executable or object code is made by offering
access to copy from a designated place, then offering equivalent
access to copy the source code from the same place counts as
distribution of the source code, even though third parties are not
compelled to copy the source along with the object code.

  4. You may not copy, modify, sublicense, or distribute the Program
except as expressly provided under this License.  Any attempt
otherwise to copy, modify, sublicense or distribute the Program is
void, and will automatically terminate your rights under this License.
However, parties who have received copies, or rights, from you under
this License will not have their licenses terminated so long as such
parties remain in full compliance.

  5. You are not required to accept this License, since you have not
signed it.  However, nothing else grants you permission to modify or
distribute the Program or its derivative works.  These actions are
prohibited by law if you do not accept this License.  Therefore, by
modifying or distributing the Program (or any work based on the
Program), you indicate your acceptance of this License to do so, and
all its terms and conditions for copying, distributing or modifying
the Program or works based on it.

  6. Each time you redistribute the Program (or any work based on the
Program), the recipient automatically receives a license from the
original licensor to copy, distribute or modify the Program subject to
these terms and conditions.  You may not impose any further
restrictions on the recipients' exercise of the rights granted herein.
You are not responsible for enforcing compliance by third parties to
this License.

  7. If, as a consequence of a court judgment or allegation of patent
infringement or for any other reason (not limited to patent issues),
conditions are imposed on you (whether by court order, agreement or
otherwise) that contradict the conditions of this License, they do not
excuse you from the conditions of this License.  If you cannot
distribute so as to satisfy simultaneously your obligations under this
License and any other pertinent obligations, then as a consequence you
may not distribute the Program at all.  For example, if a patent
license would not permit royalty-free redistribution of the Program by
all those who receive copies directly or indirectly through you, then
the only way you could satisfy both it and this License would be to
refrain entirely from distribution of the Program.

If any portion of this section is held invalid or unenforceable under
any particular circumstance, the balance of the section is intended to
apply and the section as a whole is intended to apply in other
circumstances.

It is not the purpose of this section to induce you to infringe any
patents or other property right claims or to contest validity of any
such claims; this section has the sole purpose of protecting the
integrity of the free software distribution system, which is
implemented by public license practices.  Many people have made
generous contributions to the wide range of software distributed
through that system in reliance on consistent application of that
system; it is up to the author/donor to decide if he or she is willing
to distribute software through any other system and a licensee cannot
impose that choice.

This section is intended to make thoroughly clear what is believed to
be a consequence of the rest of this License.

  8. If the distribution and/or use of the Program is restricted in
certain countries either by patents or by copyrighted interfaces, the
original copyright holder who places the Program under this License
may add an explicit geographical distribution limitation excluding
those countries, so that distribution is permitted only in or among
countries not thus excluded.  In such case, this License incorporates
the limitation as if written in the body of this License.

  9. The Free Software Foundation may publish revised and/or new versions
of the General Public License from time to time.  Such new versions will
be similar in spirit to the present version, but may differ in detail to
address new problems or concerns.

Each version is given a distinguishing version number.  If the Program
specifies a version number of this License which applies to it and "any
later version", you have the option of following the terms and conditions
either of that version or of any later version published by the Free
Software Foundation.  If the Program does not specify a version number of
this License, you may choose any version ever published by the Free Software
Foundation.

  10. If you wish to incorporate parts of the Program into other free
programs whose distribution conditions are different, write to the author
to ask for permission.  For software which is copyrighted by the Free
Software Foundation, write to the Free Software Foundation; we sometimes
make exceptions for this.  Our decision will be guided by the two goals
of preserving the free status of all derivatives of our free software and
of promoting the sharing and reuse of software generally.

			    NO WARRANTY

  11. BECAUSE THE PROGRAM IS LICENSED FREE OF CHARGE, THERE IS NO WARRANTY
FOR THE PROGRAM, TO THE EXTENT PERMITTED BY APPLICABLE LAW.  EXCEPT WHEN
OTHERWISE STATED IN WRITING THE COPYRIGHT HOLDERS AND/OR OTHER PARTIES
PROVIDE THE PROGRAM "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED
OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.  THE ENTIRE RISK AS
TO THE QUALITY AND PERFORMANCE OF THE PROGRAM IS WITH YOU.  SHOULD THE
PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF ALL NECESSARY SERVICING,
REPAIR OR CORRECTION.

  12. IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING
WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MAY MODIFY AND/OR
REDISTRIBUTE THE PROGRAM AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES,
INCLUDING ANY GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING
OUT OF THE USE OR INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED
TO LOSS OF DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY
YOU OR THIRD PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER
PROGRAMS), EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE
POSSIBILITY OF SUCH DAMAGES.

		     END OF TERMS AND CONDITIONS

********************************************************************************
*** END OF HEMODOWNLOADER LICENSE DOCUMENTATION                              ***
********************************************************************************"""

##   START THE MAIN GUI
root.mainloop()
