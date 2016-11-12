###############################################################################
# 91 OG Auto-Scheduler v1.0
#
# Written in Python v3.5.1
#
# Description:   Given a squadron roster/schedule input .csv, this program
#                creates a complete alert schedule for the selected 91 OG
#                squadron.        
###############################################################################

# Imports
import calendar
import csv
import fnmatch
import os
import random
import tkinter as tk
import traceback
from calendar import monthrange
from datetime import date, timedelta
from time import strptime
from tkinter import SUNKEN, Button, StringVar, IntVar, OptionMenu, CENTER
from tkinter import Scale, HORIZONTAL, DISABLED, Toplevel, Checkbutton
from tkinter import font, Frame, FALSE, W, E, N, S, Label, LEFT, RIGHT
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename


# Tooltip class
class CreateToolTip(object):
    # create a tooltip for a given widget
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.close)

    def enter(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tk.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left',
                         background='white', relief='solid', borderwidth=1,
                         font=("times", "11", "normal"))
        label.pack(ipadx=1)

    def close(self, event=None):
        if self.tw:
            self.tw.destroy()


# Main program window
class progWindow(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.title("91 OG Auto Scheduler")
        self.master.resizable(width=FALSE, height=FALSE)
        self.grid(sticky=W + E + N + S, padx=10, pady=10)
        self.file = False
        self.sq = False
        self.years = [date.today().year, date.today().year + 1]
        self.months = list(calendar.month_abbr)
        self.inputLoc = ""
        self.inputLabel = Label(self, text='Select input .csv:', justify=LEFT)
        self.inputLabel.grid(row=0, column=0, sticky=W)
        self.browseLabel = Label(self, text='', relief=SUNKEN, width=30, anchor="w")
        self.browseLabel.grid(row=1, column=0, sticky="ew")
        self.button = Button(self, text="Browse", command=self.load_file)
        self.button.grid(row=1, column=1, sticky="ew")
        self.sqv = StringVar()
        self.sqv.set("Select")
        self.sqLabel = Label(self, text='Select Squadron:', justify=RIGHT)
        self.sqLabel.grid(row=2, column=0, sticky=E)
        self.sqOption = OptionMenu(self, self.sqv, "740", "741", "742", command=self.sqCheck)
        self.sqOption.grid(row=2, column=1, sticky="ew")
        self.yrv = IntVar()
        self.yrv.set(self.years[0])
        self.yearLabel = Label(self, text='Select Year:', justify=RIGHT)
        self.yearLabel.grid(row=3, column=0, sticky=E)
        self.yearEntry = OptionMenu(self, self.yrv, *self.years)
        self.yearEntry.grid(row=3, column=1, sticky="ew")
        self.mv = StringVar()
        self.mv.set(self.months[1])
        self.monthLabel = Label(self, text='Select Month:', justify=RIGHT)
        self.monthLabel.grid(row=4, column=0, sticky=E)
        self.monthEntry = OptionMenu(self, self.mv, *[x for x in self.months if x != ""])
        self.monthEntry.grid(row=4, column=1, sticky="ew")
        self.slider1Label = Label(self, text="Crew Pairing Weight:")
        self.slider1Label.grid(row=6, column=0, sticky="w")
        self.slider2Label = Label(self, text="White Space Weight:")
        self.slider2Label.grid(row=8, column=0, sticky="w")
        self.slider3Label = Label(self, text="Alert Max Percentage Weight:")
        self.slider3Label.grid(row=10, column=0, sticky="w")
        self.slider1Ttp = CreateToolTip(self.slider1Label,
                                        "Increasing will favor crew pairings.")
        self.slider2Ttp = CreateToolTip(self.slider2Label,
                                        "Increasing will favor MCCMs with less white space in the remainder of the month.  \ni.e.:  A MCCM with a large block of leave at the end of the month will be chosen for an \nalert over an MCCM with no leave.")
        self.slider3Ttp = CreateToolTip(self.slider3Label,
                                        "Increasing will favor MCCMs who have a lower percentage of their max alert count dedicated to alert.\ni.e.: an MCCM with 1/8 alerts assigned will be chosen over an MCCM with 4/8 alerts assigned.")
        self.slider1Out = Label(self, justify=CENTER, relief=SUNKEN)
        self.slider1Out.grid(row=7, column=1, sticky="ew")
        self.slider2Out = Label(self, justify=CENTER, relief=SUNKEN)
        self.slider2Out.grid(row=9, column=1, sticky="ew")
        self.slider3Out = Label(self, justify=CENTER, relief=SUNKEN)
        self.slider3Out.grid(row=11, column=1, sticky="ew")
        self.slider1 = Scale(self, from_=1, to=10, orient=HORIZONTAL, length=200, showvalue=0,
                             command=lambda x: self.updateSliderLabel(self.slider1, self.slider1Out))
        self.slider1.grid(row=7, column=0, sticky="w")
        self.slider2 = Scale(self, from_=1, to=10, orient=HORIZONTAL, length=200, showvalue=0,
                             command=lambda x: self.updateSliderLabel(self.slider2, self.slider2Out))
        self.slider2.grid(row=9, column=0, sticky="w")
        self.slider3 = Scale(self, from_=1, to=10, orient=HORIZONTAL, length=200, showvalue=0,
                             command=lambda x: self.updateSliderLabel(self.slider3, self.slider3Out))
        self.slider3.grid(row=11, column=0, sticky="w")
        self.showAdv = Label(self, text="Advanced Options", fg="blue", cursor="hand2")
        self.showAdv.grid(row=12, column=1, sticky="e")
        f = font.Font(self.showAdv, self.showAdv.cget("font"))
        f.configure(underline=True)
        self.showAdv.configure(font=f)
        self.showAdv.bind("<Button-1>", self.showAdvOptions)
        self.tempButton = Button(self, text="Create Template", command=self.create_template)
        self.tempButton.grid(row=13, column=0, sticky="w")
        self.goButton = Button(self, text="Run", width=10, state=DISABLED)
        self.goButton.grid(row=13, column=1, sticky="ew")
        self.adv = advancedOptions()

    # Determine if user has selected a squadron
    def sqCheck(self, sq):
        self.sq = True
        self.readyCheck()

    # Determine if all required inputs have been selected
    def readyCheck(self):
        if self.sq and self.file:
            self.goButton['state'] = 'normal'

    # Updates slider labels when sliders are adjusted
    def updateSliderLabel(self, slider, label):
        label['text'] = str(slider.get())

    # Open file prompt
    def load_file(self):
        fname = askopenfilename(filetypes=(("CSV", "*.csv"),
                                           ("All files", "*.*")))
        if fname:
            self.browseLabel['text'] = os.path.basename(fname)
            self.inputLoc = fname
            self.file = True
            self.readyCheck()
            return

    # Creates a template .csv for user to fill in for input
    def create_template(self):
        template = ["FLIGHT", "NAME", "M/D", "SCP", "ALERTS", "CREW#",
                    "1", "2", "3", '4', '5', '6', '7', '8', '9', '10', '11',
                    '12', '13', '14', '15', '16', '17', '18', '19', '20',
                    '21', '22', '23', '24', '25', '26', '27', '28', '29',
                    '30', '31']
        # write the new CSV file
        outName = asksaveasfilename(filetypes=(("CSV", "*.csv"),
                                               ("All files", "*.*")))
        if outName:
            if not outName.endswith('.csv'):
                outName += '.csv'
            with open(outName, 'w', newline='') as outFile:
                writer = csv.writer(outFile)
                writer.writerow(template)

    # Call advanced options window
    def showAdvOptions(self, event):
        self.adv.top.deiconify()


# Advanced Options window
class advancedOptions():
    def __init__(self):
        self.fltDep = 0
        self.sqDep = 0
        self.b2b = 5
        self.bDays = 1
        self.top = Toplevel()
        self.top.title("Advanced Options")
        self.top.protocol("WM_DELETE_WINDOW", lambda: self.top.withdraw())
        self.top.geometry("270x150")
        self.top.resizable(width=False, height=False)
        self.top.withdraw()
        self.buffer = Label(self.top, text="     ")
        self.buffer.grid(row=0, column=0)
        self.fltDepVar = IntVar()
        self.fltDepVar.set(self.fltDep)
        self.fltDepLab = Label(self.top, text="Flight Deployments:")
        self.fltDepLab.grid(row=0, column=1, sticky="e")
        self.fltDepCB = Checkbutton(self.top, variable=self.fltDepVar,
                                    command=lambda: self.sqflt(self.fltDepCB))
        self.fltDepCB.grid(row=0, column=2, sticky="ew")
        self.sqDepVar = IntVar()
        self.sqDepVar.set(self.sqDep)
        self.sqDepLab = Label(self.top, text="Squadron Deployments:")
        self.sqDepLab.grid(row=1, column=1, sticky="e")
        self.sqDepCB = Checkbutton(self.top, variable=self.sqDepVar,
                                   command=lambda: self.sqflt(self.sqDepCB))
        self.sqDepCB.grid(row=1, column=2, sticky="ew")
        self.b2bLabel = Label(self.top, text="Max number of back-to-backs:")
        self.b2bLabel.grid(row=2, column=1, sticky="e")
        self.b2bVar = IntVar()
        self.b2bVar.set(5)
        self.b2bBox = OptionMenu(self.top,
                                 self.b2bVar,
                                 '1', '2', '3', '4', '5',
                                 '6', '7', '8', '9', '10')
        self.b2bBox.grid(row=2, column=2)
        self.backupsLab = Label(self.top, text="Max number of backups:")
        self.backupsLab.grid(row=3, column=1, sticky="e")
        self.backVar = IntVar()
        self.backVar.set(1)
        self.backups = OptionMenu(self.top,
                                  self.backVar,
                                  '1', '2', '3', '4', '5',
                                  '6', '7', '8', '9', '10')
        self.backups.grid(row=3, column=2)
        self.ok = Button(self.top, text="Ok", command=lambda: self.top.withdraw())
        self.ok.grid(row=4, column=2, sticky="ew")

    def sqflt(self, box):
        if box == self.fltDepCB:
            if self.fltDepVar.get() == 1:
                self.sqDepVar.set(0)
        if box == self.sqDepCB:
            if self.sqDepVar.get() == 1:
                self.fltDepVar.set(0)


# Executes the main program with error handling
def execute():
    try:
        runProgram()
    except Exception as inst:
        tk.messagebox.showerror("Error", "91 OG Auto Scheduler encountered an error. \
                                Please restart the program and try again.\n\n" + str(traceback.print_exc()))


# Actual main program function
def runProgram():
    # Identify globals that will be modified
    global rows, daysInMonth, sq, flightHolder, weekends
    # Get filename from GUI
    filename = window.inputLoc

    # Change working directory to location of csv file
    os.chdir(os.path.dirname(filename))
    with open(filename) as f:
        csvReader = csv.reader(f)
        # Create iterable list of rows in csv
        rows = [row for row in csvReader]
        # Strip white space in each cell and build list of flights
        firstLine = 1
        for row in rows:
            for i, v in enumerate(row):
                row[i] = v.strip()
            if row[flight] not in flights and firstLine != 1 and row[flight] != "":
                flights.append(row[flight])
            firstLine = 0

        # Define year, month, and number of days in month.
        squad = int(window.sqv.get())
        listOfMonths = list(calendar.month_abbr)
        year = int(window.yrv.get())
        month = int(listOfMonths.index(window.mv.get()))
        weekends = [d for d in getWeekends(year, month, [4, 5, 6])]
        daysInMonth = monthrange(year, month)[1]

        # Assign LCCs based on squadron and establish other squadron specific variables
        # flightHolder etermines which flight goes first for flight deployments (in a 0 indexed list)
        sqDep = window.adv.sqDepVar.get()
        if squad == 740:
            plccs = ['a', 'b', 'd', 'e']
            scp = ['c']
            flightHolder = 0
        elif squad == 741:
            plccs = ['f', 'g', 'h', 'j']
            scp = ['i']
            flightHolder = 1
        elif squad == 742:
            plccs = ['k', 'l', 'n', 'o']
            scp = ['m']
            flightHolder = 0
        if sqDep:
            plccs = ['a', 'b', 'd', 'e',
                     'f', 'g', 'h', 'j',
                     'k', 'l', 'n', 'o']
            scp = ['c', 'i', 'm']
            random.shuffle(plccs)
            random.shuffle(scp)
        # Add two items to the end of each row to make room for 't' and 'o' days.
        for row in rows:
            reqLen = daysInMonth + dateOffset + 2
            if len(row) < reqLen:
                reqAdd = reqLen - len(row)
                for i in range(0, reqAdd):
                    row.append('')

        # Erase all dates and rewrite to ensure we only include as many days as are in the month
        for day in range(dateOffset, len(rows[0])):
            rows[0][day] = ""
        for day in range(dateOffset, daysInMonth + dateOffset):
            rows[0][day] = day - 5

        # creates a dictionary to update when an alert is pulled and one to track back to backs
        for i in range(1, len(rows)):
            alertCounter[i] = 0
            b2b[i] = [0, 0]

        # Build a list of all possible alert positions
        aList = []
        for p in ["M", "D"]:
            for a in ["Aa", "Ab", "Ac", "Ad", "Ae",
                      "Af", "Ag", "Ah", "Ai", "Aj",
                      "Ak", "Al", "Am", "An", "Ao", "B1"]:
                aList.append(a + '(' + p + ')')
        # appends pre-designated alerts to alert counter
        for i in range(1, len(rows)):
            aCount = 0
            for a in aList:
                aCount += len(fnmatch.filter(rows[i], a))
            alertCounter[i] = aCount
            # If pre-designated alerts reach or exceed identified max alert
            # count, remove crew member from the alertCounter dictionary to
            # ensure no further alerts are assigned
            if alertCounter[i] >= int(rows[i][alertMax]):
                del alertCounter[i]

        # creates a list of crewnum column
        crew_column = ['placeholder']
        for i in range(1, len(rows)):
            crew_column.append(rows[i][crewNum])

        # Build dictionary of integral crews
        for x in crew_column:
            x1 = [i for i, v in enumerate(crew_column) if v == x]
            x1 = list(x1)
            k = x1[0]
            if len(x1) > 1 and x1[0] not in crewPairColumn:
                v = x1[1]
                crewPairColumn[k] = v
                crewPairColumn[v] = k
            elif len(x1) == 1:
                crewPairColumn[k] = "None"

        # creates a dayx list for every date in the month so that
        # pre-designated alert counts are accounted for and not reassigned
        for x in range(dateOffset, daysInMonth + dateOffset):
            define_a_day(x)

        # Begin process of alert assignment
        if window.adv.fltDepVar.get() == 1:
            r = 2
        else:
            r = 1
        for a in range(r):
            # If this is the second run for flight deployments, turn off flight deployment
            # requirement to attempt to fill empty seats with out of flight options
            if a == 1:
                window.adv.fltDepVar.set(0)
            y = int(window.yrv.get())
            m = strptime(window.mv.get(), '%b').tm_mon
            # Start date for backup and flight deployment cycle (1 Sep 2016)
            sDate = date(2016, 9, 1)
            eDate = date(y, m, monthrange(y, m)[1])
            moDict = getMonths(sDate, eDate)
            bDayInc = 1
            bDays = []
            # Start backups with 740th, then create dictionary of
            # backup responsibility from start date to end of current month.
            # Also assigns flight days based on the same start month.
            sq = 740
            for y in moDict:
                for m in moDict[y]:
                    for d in range(0, monthrange(y, m)[1]):
                        if y == max(moDict) and m == eDate.month:
                            # Assign backup days
                            if sq == squad:
                                bDays.append(d + dateOffset)
                            # Create flight days
                            flightTracker[d + dateOffset] = flightHolder
                        # Increment flight and squadron for flight/backup days
                        flightHolder = getFlight(flightHolder)
                        if sqDep:
                            sq = updateSq()
                        else:
                            bDayInc = updateDay(bDayInc)
            bDays.sort()
            # Assign backup days first, and if squadron deployments selected,
            # assign all alerts.
            for day in bDays:
                for i in range(1, len(rows)):
                    # If it's been 3 days with no alert, reset back to back counter
                    if b2b[i][0] == 3:
                        b2b[i][0] = 0
                        b2b[i][1] = 0
                    b2b[i][0] += 1
                dayx = globals()['day' + str(day - 5)]
                assignAlert(day, "B1", "M", dayx)
                assignAlert(day, "B1", "D", dayx)
                if sqDep:
                    # Start with SCP assignment
                    for s in scp:
                        assignAlert(day, s, "M", dayx, SCP="Y")
                        assignAlert(day, s, "D", dayx, SCP="Y")
                    # Then assign PLCC alerts
                    for lcc in plccs:
                        assignAlert(day, lcc, "M", dayx)
                        assignAlert(day, lcc, "D", dayx)
            # If squadron deployments not selected
            if not sqDep:
                # Create list of dates
                dateList = []
                for i in range(0, daysInMonth):
                    dateList.append(i + dateOffset)

                # Assign standard alerts
                for day in dateList:
                    for i in range(1, len(rows)):
                        # If it's been 3 days with no alert, reset back to back counter
                        if b2b[i][0] == 3:
                            b2b[i][0] = 0
                            b2b[i][1] = 0
                        b2b[i][0] += 1
                    dayx = globals()['day' + str(day - 5)]
                    # Start with SCP assignment
                    for s in scp:
                        assignAlert(day, s, "M", dayx, SCP="Y")
                        assignAlert(day, s, "D", dayx, SCP="Y")
                    # Then assign PLCC alerts
                    for lcc in plccs:
                        assignAlert(day, lcc, "M", dayx)
                        assignAlert(day, lcc, "D", dayx)
            if a == 1:
                # Turn flight deployments back on if it was turned off
                window.adv.fltDepVar.set(1)

        # write the new CSV file
        outName = asksaveasfilename(filetypes=(("CSV", "*.csv"),
                                               ("All files", "*.*")))
        if outName:
            if not outName.endswith('.csv'):
                outName += '.csv'
            with open(outName, 'w', newline='') as outFile:
                writer = csv.writer(outFile)
                for row in rows:
                    writer.writerow(row)
        # Prints integral crew rate statistic
        print(str(integralCount / alertCount * 100) + "% Integral Rate Achieved")
        resetVars()
        print("Program has completed.")


# Increments the month
def incMonth(m):
    if m == 12:
        m = 1
    else:
        m += 1
    return m


# Get a list of months from the start date (Sep 16) to the selected end date
def getMonths(sDate, eDate):
    sm = sDate.month
    em = eDate.month
    rDict = {}
    for y in range(sDate.year, eDate.year + 1):
        rDict[y] = []
        for m in range(sm, 13):
            if y == eDate.year:
                if m > em:
                    break
            rDict[y].append(m)
            sm = incMonth(m)
    return rDict


# Increments squadron
def updateSq():
    if sq == 740:
        return 741
    if sq == 741:
        return 742
    if sq == 742:
        return 740


# Increments backup days (1st, 2nd or 3rd)
def updateDay(day):
    global sq
    if day == 1:
        return 2
    if day == 2:
        return 3
    if day == 3:
        sq = updateSq()
        return 1


# Creates a list representing the values by column corresponding
# to the date in the format dayx where x is the date
def define_a_day(day):
    globals()[str('day') + str(day - 5)] = []
    for i in range(1, len(rows)):
        globals()[str('day') + str(day - 5)].append(rows[i][day])


# identifies weekends so 4 or less alert pullers are never scheduled on them
def getWeekends(year, month, days):
    d = date(year, month, 1)
    while d.month == month:
        if d.weekday() in days:
            yield d.day
        d += timedelta(days=1)


# Create list of eligible crew members who have the least white space
# in their remaining schedule
def getWhiteSpaceRatio(mccmList, day):
    weightedRatio = {}
    for mccm in mccmList:
        whiteSpaceCount = 0
        for d in range(day, daysInMonth + dateOffset + 1):
            if rows[mccm][d].strip() == '':
                whiteSpaceCount += 1
        weightedRatio[mccm] = ((daysInMonth - day + dateOffset) / \
                               whiteSpaceCount) * 1.3 * window.slider2.get()
    if len(weightedRatio) > 0:
        return weightedRatio
    else:
        return False


# Get list of alert count for each eligible mccm
def getAlertCountRatio(mccmList):
    alertCountRatio = {}
    for mccm in mccmList:
        alertCountRatio[mccm] = (alertCounter[mccm] / \
                                 float(rows[mccm][alertMax])) * \
                                -2 * window.slider3.get()
    return alertCountRatio


# Returns crew pair ratio for each eligible mccm
def getCrewPairingRatio(mccmList, day, **optionalParams):
    crewPairingRatio = {}
    for i in mccmList:
        crewPairingRatio[i] = 0
        if type(crewPairColumn[i]) is int and crewPairColumn[i] in alertCounter:
            if b2b[crewPairColumn[i]][1] <= window.adv.b2bVar.get() - 1:
                if rows[crewPairColumn[i]][day] == "" and \
                                rows[crewPairColumn[i]][day + 1].lower() in tDay and \
                                rows[crewPairColumn[i]][day + 2].lower() in oDay:
                    if 'SCP' in optionalParams:
                        if rows[i][scpIndex] == optionalParams['SCP'] and \
                                        rows[crewPairColumn[i]][scpIndex] == optionalParams['SCP']:
                            crewPairingRatio[i] += window.slider1.get() * 0.3
                    else:
                        crewPairingRatio[i] += window.slider1.get() * 0.3
    return crewPairingRatio


# find all eligible mccms for an alert
def findMCCMs(day, position, site, **optionalParams):
    # Get a list of eligible mccms for the alert
    mccm = {}
    print(weekends)
    for i in alertCounter:
        if (site == "B1" and ((int(rows[i][alertMax]) <= 4) or \
                                          len(fnmatch.filter(rows[i], "B1(?)")) >= window.adv.backVar.get())) or \
                ((day - dateOffset + 1) in weekends and int(rows[i][alertMax]) <= 4):
            pass
        elif rows[i][day] == "" and rows[i][day + 1].lower() in tDay and \
                        rows[i][day + 2].lower() in oDay and \
                        b2b[i][1] <= window.adv.b2bVar.get() - 1:
            if rows[i][posIndex] == position:
                if 'SCP' in optionalParams:
                    if rows[i][scpIndex] == optionalParams['SCP']:
                        mccm[i] = 0
                else:
                    mccm[i] = 0
            elif position == "D" and rows[i][posIndex] == "M":
                if 'SCP' in optionalParams:
                    if rows[i][scpIndex] == optionalParams['SCP']:
                        mccm[i] = -999
                else:
                    mccm[i] = -999
    # If we found MCCM possibilities:
    if len(mccm) > 0:
        return mccm
    else:
        return False


# Revise list of eligible mccms to include only those in the flight identified for flight deployments
def flightDeployment(mccm, day):
    global flightTracker
    revisedMccm = {}
    for i in mccm:
        if rows[i][flight] == flights[flightTracker[day]] or rows[i][flight] == "":
            revisedMccm[i] = mccm[i]
    return revisedMccm


# Increment the flight for flight deployments
def getFlight(f):
    if f == len(flights) - 1:
        f = 0
    else:
        f += 1
    return f


# Returns the opposite crew position
def oppositePosition(pos):
    if pos == "M":
        return "D"
    if pos == "D":
        return "M"


# Alert assignment function
def assignAlert(day, site, position, dayx, **optionalParams):
    global alertCount, integralCount
    formatedSite = "A" + site + "(" + position + ")"
    if site == "B1":
        formatedSite = formatedSite[1:]
    # If the alert hasn't already been assigned...
    if formatedSite not in dayx:
        # Find eligible crew members
        if 'SCP' in optionalParams:
            mccm = findMCCMs(day, position, site, SCP=optionalParams['SCP'])
        else:
            mccm = findMCCMs(day, position, site)
        # If we have at least one person to assign the alert to...
        if mccm:
            # If flight deployments were selected, filter list of mccms to those in the selected flight
            if window.adv.fltDepVar.get() == 1 and len(flights) > 0:
                mccm = flightDeployment(mccm, day)
        # If we still have potential mccms...
        if mccm:
            # Rank each mccm based on the slider values to determine who is the best fit for alert
            if 'SCP' in optionalParams:
                cp = getCrewPairingRatio(mccm, day, SCP=optionalParams['SCP'])
            else:
                cp = getCrewPairingRatio(mccm, day)
            ws = getWhiteSpaceRatio(mccm, day)
            ac = getAlertCountRatio(mccm)
            mccmRatios = {}
            for i in mccm:
                mccmRatios[i] = cp[i] + ws[i] + ac[i] + mccm[i]  # +(int(rows[i][alertMax])/8)
            maxVal = max(mccmRatios.values())
            mccmList = []
            for i in mccmRatios:
                if mccmRatios[i] == maxVal:
                    mccmList.append(i)
            # If we have more than one 'best match', shuffle the list and choose one randomly
            random.shuffle(mccmList)
            mccm = mccmList[0]
            # Write the alert into the schedule
            rows[mccm][day] = formatedSite
            if rows[mccm][day + 1] == '':
                rows[mccm][day + 1] = 'T'
            else:
                pass
            if rows[mccm][day + 2] == '':
                rows[mccm][day + 2] = 'O'
            else:
                pass
            if site != "B1":
                alertCount += 1
            dayx.append(formatedSite)
            # Update alert and back to back trackers
            alertCounter[mccm] += 1
            b2b[mccm][0] = 0
            b2b[mccm][1] += 1
            # If we hit this person's max alert count, remove them as an option for future alerts
            if alertCounter[mccm] >= int(rows[mccm][alertMax]):
                del alertCounter[mccm]
            # Assign integral crew partner the other alert position if found.
            formatedSite = "A" + site + "(" + oppositePosition(position) + ")"
            if site == "B1":
                formatedSite = formatedSite[1:]
            if formatedSite not in dayx:
                if type(crewPairColumn[mccm]) is int and \
                                crewPairColumn[mccm] in alertCounter:
                    if rows[crewPairColumn[mccm]][day] == "" and \
                                    rows[crewPairColumn[mccm]][day + 1].lower() in tDay and \
                                    rows[crewPairColumn[mccm]][day + 2].lower() in oDay:
                        if 'SCP' in optionalParams:
                            if rows[crewPairColumn[mccm]][scpIndex] == optionalParams['SCP']:
                                rows[crewPairColumn[mccm]][day] = formatedSite
                                if rows[crewPairColumn[mccm]][day + 1] == '':
                                    rows[crewPairColumn[mccm]][day + 1] = 'T'
                                else:
                                    pass
                                if rows[crewPairColumn[mccm]][day + 2] == '':
                                    rows[crewPairColumn[mccm]][day + 2] = 'O'
                                else:
                                    pass
                                if site != "B1":
                                    integralCount += 1
                                dayx.append(formatedSite)
                                alertCounter[crewPairColumn[mccm]] += 1
                                b2b[crewPairColumn[mccm]][0] = 0
                                b2b[crewPairColumn[mccm]][1] += 1
                                if alertCounter[crewPairColumn[mccm]] >= int(rows[crewPairColumn[mccm]][alertMax]):
                                    del alertCounter[crewPairColumn[mccm]]
                        else:
                            rows[crewPairColumn[mccm]][day] = formatedSite
                            if rows[crewPairColumn[mccm]][day + 1] == '':
                                rows[crewPairColumn[mccm]][day + 1] = 'T'
                            else:
                                pass
                            if rows[crewPairColumn[mccm]][day + 2] == '':
                                rows[crewPairColumn[mccm]][day + 2] = 'O'
                            else:
                                pass
                            if site != "B1":
                                integralCount += 1
                            dayx.append(formatedSite)
                            alertCounter[crewPairColumn[mccm]] += 1
                            b2b[crewPairColumn[mccm]][0] = 0
                            b2b[crewPairColumn[mccm]][1] += 1
                            if alertCounter[crewPairColumn[mccm]] >= int(rows[crewPairColumn[mccm]][alertMax]):
                                del alertCounter[crewPairColumn[mccm]]


# Reset variables to set up the program for another run
def resetVars():
    global rows, alertCounter, b2b, crewPairColumn, alertCount, integralCount
    rows = ''
    alertCounter = {}
    b2b = {}
    crewPairColumn = {}
    alertCount = 0
    integralCount = 0


if __name__ == "__main__":
    # Define globals
    alertCount = 0
    integralCount = 0
    rows = ''
    alertCounter = {}
    b2b = {}
    crewPairColumn = {}
    weightedRatio = {}
    flights = []
    flightTracker = {}
    flightHolder = 0
    daysInMonth = 0
    # Define which columns in csv correspond to needed data. Note: 0 index
    flight = 0
    name = 1
    posIndex = 2
    scpIndex = 3
    alertMax = 4
    crewNum = 5
    # dateOffset for where day 1 of the month begins in csv
    dateOffset = 6
    weekends = []
    # creates lists of types of days that a first or second O day can fall on.
    tDay = ['', 'm2']
    oDay = ['r', 'e', 'lv', '', 'm1', 'm2', 't10', 'sd', 'wd', 'l']

    # Run the program
    window = progWindow()
    window.goButton.configure(command=execute)
    window.mainloop()
