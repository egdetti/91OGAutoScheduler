import os, csv, calendar, fnmatch
from datetime import date, timedelta
import tkinter as tk
from tkinter import SUNKEN, Button, StringVar, IntVar, OptionMenu, CENTER, Entry
from tkinter import Scale, HORIZONTAL, DISABLED, Toplevel, Checkbutton, NORMAL, END
from tkinter import font, Frame, FALSE, W, E, N, S, Label, LEFT, RIGHT, RAISED
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
        self.yearLabel = Label(self, text='Select Year:', justify=RIGHT)
        self.yearLabel.grid(row=3, column=0, sticky=E)
        self.yearEntry = OptionMenu(self, self.yrv, *self.years, command=lambda x:self.adv.backupCalendar.createCalendar(self.yrv.get(),
                                                                                                                         int(list(calendar.month_abbr).index(self.mv.get()))))
        self.yrv.set(self.years[0])
        self.yearEntry.grid(row=3, column=1, sticky="ew")
        self.mv = StringVar()
        self.monthLabel = Label(self, text='Select Month:', justify=RIGHT)
        self.monthLabel.grid(row=4, column=0, sticky=E)
        self.monthEntry = OptionMenu(self, self.mv, *[x for x in self.months if x != ""], command=lambda x:self.adv.backupCalendar.createCalendar(self.yrv.get(),
                                                                                                                         int(list(calendar.month_abbr).index(self.mv.get()))))
        self.mv.set(self.months[1])
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
        self.adv.backupCalendar.createCalendar(self.yrv.get(), int(list(calendar.month_abbr).index(self.mv.get())))

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
            self.loadFlights(fname)
            return

    # Creates a template .csv for user to fill in for input
    def create_template(self):
        template = [["", "", "", "", "", "",
                     "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Mon", "Tue", "Wed", "Thur",
                     "Fri", "Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat",
                     "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Mon",
                     "Tue", "Wed"],
                    ["FLIGHT", "NAME", "M/D", "SCP", "ALERTS", "CREW#",
                    "1", "2", "3", '4', '5', '6', '7', '8', '9', '10', '11',
                    '12', '13', '14', '15', '16', '17', '18', '19', '20',
                    '21', '22', '23', '24', '25', '26', '27', '28', '29',
                    '30', '31']]
        # write the new CSV file
        outName = asksaveasfilename(filetypes=(("CSV", "*.csv"),
                                               ("All files", "*.*")))
        if outName:
            if not outName.endswith('.csv'):
                outName += '.csv'
            with open(outName, 'w', newline='') as outFile:
                writer = csv.writer(outFile)
                for row in template:
                    writer.writerow(row)

    # Call advanced options window
    def showAdvOptions(self, event):
        self.adv.top.deiconify()

    def loadFlights(self, fileName):
        flights = []
        with open(fileName) as f:
            csvReader = csv.reader(f)
            rows = [row for row in csvReader]
            currentLine = 0
            for row in rows:
                if currentLine > 1:
                    if row[0] not in flights and row[0] != "":
                        flights.append(row[0])
                currentLine += 1
        self.adv.fltRotationEntry.delete(0, END)
        self.adv.fltRotationVar.set(",".join(flights))

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
        self.top.config(padx=10, pady=10)
        self.top.resizable(width=False, height=False)
        self.top.withdraw()
        self.buffer = Label(self.top, text="     ")
        self.buffer.grid(row=0, column=0)
        self.backupCalendar = backupCalendar()
        self.fltDepVar = IntVar()
        self.fltDepVar.set(self.fltDep)
        self.fltDepLab = Label(self.top, text="Flight Deployments:")
        self.fltDepLab.grid(row=0, column=1, sticky="e")
        self.fltDepCB = Checkbutton(self.top, variable=self.fltDepVar,
                                    command=lambda: self.sqflt(self.fltDepCB))
        self.fltDepCB.grid(row=0, column=2, sticky="ew")
        self.fltRotationLabel = Label(self.top, text="Flight Rotation:")
        self.fltRotationLabel.grid(row=1, column=1, sticky="e")
        self.fltRotationTtp = CreateToolTip(self.fltRotationLabel,
                                        "Designates ordering of flights for flight deployment alert assignment.\nEnsure that flights are separated by a comma and match the input csv.")
        self.fltRotationVar = StringVar()
        self.fltRotationEntry = Entry(self.top, width=10, state=DISABLED, textvariable=self.fltRotationVar)
        self.fltRotationEntry.grid(row=1, column=2, sticky="ew")
        self.sqDepVar = IntVar()
        self.sqDepVar.set(self.sqDep)
        self.sqDepLab = Label(self.top, text="Squadron Deployments:")
        self.sqDepLab.grid(row=2, column=1, sticky="e")
        self.sqDepCB = Checkbutton(self.top, variable=self.sqDepVar,
                                   command=lambda: self.sqflt(self.sqDepCB))
        self.sqDepCB.grid(row=2, column=2, sticky="ew")
        self.b2bLabel = Label(self.top, text="Max number of back-to-backs:")
        self.b2bLabel.grid(row=3, column=1, sticky="e")
        self.b2bVar = IntVar()
        self.b2bVar.set(5)
        self.b2bBox = OptionMenu(self.top,
                                 self.b2bVar,
                                 '1', '2', '3', '4', '5',
                                 '6', '7', '8', '9', '10')
        self.b2bBox.grid(row=3, column=2)
        self.backupsLab = Label(self.top, text="Max number of backups:")
        self.backupsLab.grid(row=4, column=1, sticky="e")
        self.backVar = IntVar()
        self.backVar.set(1)
        self.backups = OptionMenu(self.top,
                                  self.backVar,
                                  '1', '2', '3', '4', '5',
                                  '6', '7', '8', '9', '10')
        self.backups.grid(row=4, column=2)
        self.customBackupsLabel = Label(self.top, text="Manually select backup days:")
        self.customBackupsLabel.grid(row=5, column=1, sticky="e")
        self.backupCalVar = IntVar()
        self.customBackupsCheck = Checkbutton(self.top, variable=self.backupCalVar, command=self.openCalendar)
        self.customBackupsCheck.grid(row=5, column=2)
        self.numberOfRunsLabel = Label(self.top, text="Number of schedule runs:")
        self.numberOfRunsLabel.grid(row=6, column=1, sticky="e")
        self.numRunsToolTip = CreateToolTip(self.numberOfRunsLabel,
                                        "Will generate the selected number of schedules and choose the one with \nthe least amount of holes for export.\n\nWarning:  Choosing a high number will cause long processing times.")
        self.runVar = IntVar()
        self.runVar.set(1)
        self.numberOfRuns = OptionMenu(self.top,
                                       self.runVar,
                                       1,10,20,30,40,50,
                                       60,70,80,90,100)
        self.numberOfRuns.grid(row=6, column=2)
        self.ok = Button(self.top, text="Ok", command=lambda: self.top.withdraw())
        self.ok.grid(row=7, column=2, sticky="ew")

    def openCalendar(self):
        if self.backupCalVar.get() and self.backupCalendar.month and self.backupCalendar.year:
            self.backupCalendar.top.deiconify()
        else:
            self.backupCalendar.top.withdraw()

    def sqflt(self, box):
        if box == self.fltDepCB:
            if self.fltDepVar.get() == 1:
                self.sqDepVar.set(0)
                self.fltRotationEntry.config(state=NORMAL)
            if self.fltDepVar.get() == 0:
                self.fltRotationEntry.config(state=DISABLED)
        if box == self.sqDepCB:
            if self.sqDepVar.get() == 1:
                self.fltDepVar.set(0)
                self.fltRotationEntry.config(state=DISABLED)


class myLabel(Label):
    def __init__(self, c, t, px, py, r, h, w):
        tk.Label.__init__(self, c, text=t,
                          padx=px, pady=py, relief=r,
                          height=h, width=w)
        self.checked = False

class backupCalendar():
    def __init__(self):
        self.backupDays = []
        self.top = Toplevel()
        self.top.title("Backups")
        self.top.resizable(width=False, height=False)
        self.top.config(bg="#666666", padx=25, pady=25)
        self.top.protocol("WM_DELETE_WINDOW", lambda: self.top.withdraw())
        self.top.withdraw()
        self.defaultbg = None
        self.defaultfg = None
        self.month = None
        self.year = None

    def createCalendar(self, y, m):
        self.month = m
        self.year = y
        self.top.title("Backups - {0} {1}".format(calendar.month_name[m], y))
        for widget in self.top.winfo_children():
            widget.destroy()
        self.backupDays = []
        cal = calendar.monthcalendar(y, m)
        cal2 = []
        for row in range(len(cal) + 1):
            if row == len(cal):
                newRow = cal[row - 1][-1:]
            elif row - 1 >= 0 and row < len(cal):
                newRow = cal[row - 1][-1:] + cal[row][0:-1]
            else:
                newRow = [0] + cal[row][0:-1]
            if sum(newRow) > 0:
                cal2.append(newRow)
        cal = cal2

        for i, h in enumerate(["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]):
            header = Label(self.top, text=h)
            header.config(bg="#666666", fg="#CCCCCC")
            header.grid(row=0, column=i)

        for r in range(len(cal)):
            for c in range(len(cal[r])):
                if cal[r][c] != 0:
                    lab = myLabel(self.top, cal[r][c],
                                  12, 12, RAISED,
                                  1, 1)
                    lab.config(font=("Courier", 12))
                    lab.bind("<Button-1>", self.callback)
                    lab.grid(row=r + 1, column=c)

            self.defaultbg = lab.cget("bg")
            self.defaultfg = lab.cget("fg")

    def callback(self, event):
        lab = event.widget
        if lab.checked:
            lab.config(bg=self.defaultbg,
                       fg=self.defaultfg,
                       relief=RAISED)
            lab.checked = False
            self.backupDays.remove(int(lab.cget("text")))
            self.backupDays.sort()
        else:
            lab.config(bg='#2222AA',
                       fg='#CCCCCC',
                       relief=SUNKEN)
            lab.checked = True
            self.backupDays.append(int(lab.cget("text")))
            self.backupDays.sort()

class schedule():
    def __init__(self, year, month):
        self.year = year
        self.month = month
        self.calendar = [['', '', '', '', '', ''],
                         ['Flight', 'Name', 'M/D', 'SCP', 'Alerts', 'Crew#']]
        self.weekends = [d for d in self.getWeekends(year, month, [4, 5, 6])]
        self.daysInMonth = [d for d in range(1,calendar.mdays[month]+1)]
        for d in self.daysInMonth:
            self.calendar[0].append(calendar.day_abbr[calendar.weekday(year, month, d)])
            self.calendar[1].append(str(d))
        self.mcccHoles = 0
        self.dmcccHoles = 0

    def getWeekends(self, year, month, days):
        d = date(year, month, 1)
        while d.month == month:
            if d.weekday() in days:
                yield d.day
            d += timedelta(days=1)

class mccm():
    def __init__(self, flight, name, position, scp, alertMax, crew_num, schedule):
        self.flight = flight
        self.name = name
        self.position = position
        self.scp = scp
        self.alertMax = alertMax
        self.crew_num = crew_num
        self.schedule = schedule
        self.alertCount = 0
        self.backupCount = 0
        self.crewPartners = []
    def checkBackToBacks(self, day):
        backToBacks = 0
        d = day - 3
        backToBack = True
        while d >= 0 and backToBack:
            if self.schedule[d] != None:
                if fnmatch.fnmatch(self.schedule[d], 'A?(?)') or fnmatch.fnmatch(self.schedule[d], 'B1(?)'):
                    backToBacks += 1
                else:
                    backToBack = False
            d -= 3
        backToBack = True
        d = day + 3
        while d <= len(self.schedule) - 1 and backToBack:
            if self.schedule[d] != None:
                if fnmatch.fnmatch(self.schedule[d], 'A?(?)') or fnmatch.fnmatch(self.schedule[d], 'B1(?)'):
                    backToBacks += 1
                else:
                    backToBack = False
            d += 3
        return backToBacks




class alert():
    def __init__(self, date, site, mccc, dmccc):
        self.date = date
        self.site = site
        self.mccc = mccc
        self.dmccc = dmccc