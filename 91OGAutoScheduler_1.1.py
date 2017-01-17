###############################################################################
# 91 OG Auto-Scheduler v1.0
#
# Written in Python v3.5.1
#
# Description:   Given a squadron roster/schedule input .csv, this program
#                creates a complete alert schedule for the selected 91 OG
#                squadron.        
#
# Authors:       Mike Benedetti              |  Brian Smith
#                michael.benedetti@us.af.mil |  brian.smith.189@us.af.mil
#                916-529-8427                |  701-723-6565
###############################################################################

# Imports
from tkinter import messagebox
import random, traceback, collections
from calendar import monthrange
from classes import *

# Executes the main program with error handling
def execute():
    #Disable 'Run' button and change cursor to 'wait' while program runs
    window.goButton.config(state='disabled')
    window.config(cursor="wait")
    window.update()

    try:
        runProgram()
    except Exception:
        tk.messagebox.showerror("Error", "91 OG Auto Scheduler encountered an error. \
                                Please restart the program and try again.\n\n" + str(traceback.print_exc()))
    #Re-enable 'Run' button and change cursor back to normal
    window.config(cursor="")
    window.goButton.config(state='normal')


# Actual main program function
def runProgram():
    # Identify globals that will be modified
    global rows, daysInMonth, sq, flights, flightHolder, alertList, plccs, scps, mccms, sched, schedules

    # Define variables based on GUI selections
    squad = int(window.sqv.get())
    year = int(window.yrv.get())
    month = int(list(calendar.month_abbr).index(window.mv.get()))
    sqDep = window.adv.sqDepVar.get()
    filename = window.inputLoc
    flightDeployments = window.adv.fltDepVar.get()
    #Run schedule x amount of times based on user selected amount of runs
    for i in range(int(window.adv.runVar.get())):
        print(i)
        #create schedule object
        sched = schedule(year, month)
        daysInMonth = sched.daysInMonth
        flights = []
        with open(filename) as f:
            # Create iterable list of rows in csv
            csvReader = csv.reader(f)
            rows = [row for row in csvReader]

            #Convert csv to mccm objects and add to mccms list
            currentLine = 0
            mccms = []
            #identify length required for each line to ensure T and O days fit for alerts at end of month
            reqLen = len(daysInMonth) + dateOffset + 2
            for row in rows:
                if currentLine > 1:
                    if len(row) < reqLen:
                        reqAdd = reqLen - len(row)
                        for i in range(0, reqAdd):
                            row.append('')
                    #Cleanup any trailing spaces in the input file
                    for i, v in enumerate(row):
                        row[i] = v.strip()
                    #Create list of flights represented in the input file
                    if row[flight] not in flights and row[flight] != "":
                        flights.append(row[flight])
                    #Create mccm objects
                    #Flight Attribute
                    if row[flight] == '': flt = None
                    else: flt = row[flight]
                    #SCP Attribute
                    if row[scpIndex] == "Y": scp = True
                    elif row[scpIndex] == "N": scp = False
                    else: scp = None
                    #Crew # Attribute
                    if row[crewNum] != "": crewN = int(row[crewNum])
                    else: crewN = None
                    mccms.append(mccm(flt,                      #Flight
                                      row[name],                #Last Name
                                      row[posIndex],            #Position
                                      scp,                      #SCP
                                      int(row[alertMax]),       #Alerts
                                      crewN,                    #Crew Number
                                      [None] + row[dateOffset:] #Schedule
                                      ))
                #Create crew parter lists
                for m in mccms:
                    m.crewPartners = [p for p in mccms if p.crew_num == m.crew_num and p.name != m.name]
                currentLine += 1
            if collections.Counter(flights) == collections.Counter(window.adv.fltRotationVar.get().split(',')):
                flights = window.adv.fltRotationVar.get().split(',')
            else:
                messagebox.showwarning("Bad Flight Rotation Input", "Flight Rotations do not match input CSV.  The program will continue\nwith default flight rotations.")
                print("bad")
        # Assign LCCs based on squadron and establish other squadron specific variables
        # flightHolder etermines which flight goes first for flight deployments (in a 0 indexed list)
        if squad == 740:
            plccs = ['a', 'b', 'd', 'e']
            scps = ['c']
        elif squad == 741:
            plccs = ['f', 'g', 'h', 'j']
            scps = ['i']
        elif squad == 742:
            plccs = ['k', 'l', 'n', 'o']
            scps = ['m']
        if sqDep:
            plccs = ['a', 'b', 'd', 'e',
                     'f', 'g', 'h', 'j',
                     'k', 'l', 'n', 'o']
            scps = ['c', 'i', 'm']
        flightHolder = 0
        #randomize order of sites
        random.shuffle(plccs)
        random.shuffle(scps)

        # Build a list of all possible alert positions
        alertDisplayNames =  ["Aa", "Ab", "Ac", "Ad", "Ae",
                              "Af", "Ag", "Ah", "Ai", "Aj",
                              "Ak", "Al", "Am", "An", "Ao", "B1"]

        # builds alert dictionoary and identifies pre-assigned alerts
        alertList = {}
        for d in sched.daysInMonth:
            for a in alertDisplayNames:
                newAlert = alert(d, a[:2], None, None)
                for m in mccms:
                    for p in ["M", "D"]:
                        if m.schedule[d] == "{0}({1})".format(a, p):
                            if a == "B1":
                                m.backupCount += 1
                            else:
                                m.alertCount += 1
                            if p == "M":
                                newAlert.mccc = m.name
                            if p == "D":
                                newAlert.dmccc = m.name
                if d not in alertList:
                    alertList[d] = []
                alertList[d].append(newAlert)

        # Start date for backup and flight deployment cycle (1 Sep 2016)
        sDate = date(2016, 9, 1)
        eDate = date(year, month, monthrange(year, month)[1])
        moDict = getMonths(sDate, eDate)
        bDayInc = 1
        bDays = []
        # Start backups with 740th, then create dictionary of
        # backup responsibility from start date to end of current month.
        # Also assigns flight days based on the same start month.
        sq = 740
        for y in moDict:
            for m in moDict[y]:
                for d in range(1, monthrange(y, m)[1]+1):
                    if y == max(moDict) and m == eDate.month:
                        # Assign backup days
                        if sq == squad:
                            bDays.append(d)
                        # Create flight days
                        flightTracker[d] = flightHolder
                    # Increment flight and squadron for flight/backup days
                    flightHolder = getFlight(flightHolder)
                    if sqDep:
                        sq = updateSq()
                    else:
                        bDayInc = updateDay(bDayInc)
        bDays.sort()
        # Begin process of alert assignment
        if flightDeployments:
            r = 2
        else:
            r = 1
        for a in range(r):
            # If this is the second run for flight deployments, turn off flight deployment
            # requirement to attempt to fill empty seats with out of flight options
            if a == 1:
                window.adv.fltDepVar.set(0)
            # Assign backup days first, and if squadron deployments selected,
            # assign all alerts.
            for day in bDays:
                assignAlert(day, "B1")
                if sqDep:
                    # Start with SCP assignment then do plccs
                    for site in scps + plccs:
                        assignAlert(day, "A" + site)
            # If squadron deployments not selected
            if not sqDep:
                # Assign standard alerts
                for day in sched.daysInMonth:
                    for site in scps + plccs:
                        assignAlert(day, "A" + site)
            if a == 1:
                # Turn flight deployments back on if it was turned off
                window.adv.fltDepVar.set(1)

        # write the new CSV file
        for d in alertList:
            alertHoles = [a for a in alertList[d] if (not a.dmccc or not a.mccc) and a.site[1:] in scps+plccs]
            for a in alertHoles:
                if not a.mccc:
                    sched.mcccHoles += 1
                if not a.dmccc:
                    sched.dmcccHoles += 1
                #print("=====================\nDay:{0}   Site:{1}\nMCCC:{2}\nDMCCC:{3}\n".format(d, a.site, a.mccc, a.dmccc))
        for m in mccms:
            if m.scp: writeSCP = "Y"
            else: writeSCP = "N"
            sched.calendar.append([m.flight, m.name, m.position, writeSCP, m.alertMax, m.crew_num] + m.schedule[1:])
        #print("=======================\nMCCC Holes: {0}\nDMCCC Holes: {1}".format(sched.mcccHoles, sched.dmcccHoles))
        schedules.append(sched)
    schedDict = {s:(s.mcccHoles + s.dmcccHoles) for s in schedules}
    sched = random.choice([s for s in schedDict if schedDict[s] == min(schedDict.values())])
    outName = asksaveasfilename(filetypes=(("CSV", "*.csv"),
                                           ("All files", "*.*")))
    if outName:
        if not outName.endswith('.csv'):
            outName += '.csv'
        with open(outName, 'w', newline='') as outFile:
            writer = csv.writer(outFile)
            for row in sched.calendar:
                writer.writerow(row)
    # Prints integral crew rate statistic
    alertCount1 = 0
    integralCount1 = 0
    for d in alertList:
        test = [a for a in alertList[d] if a.site[1:] in scps+plccs]
        for a in test:
            if a.mccc:
                alertCount1 += 1
                if a.dmccc:
                    if a.dmccc.crew_num == a.mccc.crew_num:
                        integralCount1 += 1
            if a.dmccc:
                alertCount1 += 1
                if a.mccc:
                    if a.dmccc.crew_num == a.mccc.crew_num:
                        integralCount1 += 1
    print(str(integralCount1 / alertCount1 * 100) + "% Integral Rate Achieved (w/o backups)")
    print(str(integralCount / alertCount * 100) + "% Integral Rate Achieved (w/ backups)")
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

# Create list of eligible crew members who have the least white space
# in their remaining schedule
def getWhiteSpaceRatio(mccm, day):
    weightedRatio = {}
    for m in mccm:
        whiteSpaceCount = 0
        for d in sched.daysInMonth[day-1:]:
            if m.schedule[d] == '':
                whiteSpaceCount += 1
        weightedRatio[m] = 2/(whiteSpaceCount / len(sched.daysInMonth[day-1:])) * window.slider2.get()
    if weightedRatio:
        return weightedRatio
    else:
        return False


# Get list of alert count for each eligible mccm
def getAlertCountRatio(mccm):
    alertCountRatio = {}
    for m in mccm:
        alertCountRatio[m] = (m.alertCount / m.alertMax * -2 * window.slider3.get())
    return alertCountRatio


# Returns crew pair ratio for each eligible mccm
def getCrewPairingRatio(mccm, day, site, position):
    Alert = [a for a in alertList[day] if a.site == site][0]
    crewPairingRatio = {}
    for m in mccm:
        crewPairingRatio[m] = 0
        if Alert.dmccc or Alert.mccc:
            if (position == "M" and Alert.dmccc in m.crewPartners) or (position == "D" and Alert.mccc in m.crewPartners):
                crewPairingRatio[m] = window.slider1.get() * .5
        else:
            for partner in m.crewPartners:
                if partner.schedule[day] == '' and partner.schedule[day+1] in tDay and partner.schedule[day+2] in oDay:
                    crewPairingRatio[m] = window.slider1.get() * .5
    return crewPairingRatio

def getPositionRatio(mccm, position):
    positionRatio = {}
    for m in mccm:
        if position == "D" and m.position == "M":
            positionRatio[m] = -999
        else:
            positionRatio[m] = 0
    return positionRatio

# find all eligible mccms for an alert
def findMCCMs(day, position, site):
    backToBacks = window.adv.b2bVar.get()
    backups = int(window.adv.backVar.get())
    # Get all mccms that are not over their alert count and not over their max back to back count
    m = [m for m in mccms if m.alertCount + m.backupCount < m.alertMax and m.checkBackToBacks(day) < backToBacks]
    #If assinging a backup day, and already at max backup count or a lower alert puller, don't consider mccm.
    if site == "B1":
        m = [m for m in m if m.backupCount < backups and m.alertMax > 4]
    #If alert is on a weekend or friday, don't consider low alert pullers
    if day in sched.weekends:
        m = [m for m in m if m.alertMax > 4]
    #If flight deployments selected, ensure appropriate flight filtered
    if window.adv.fltDepVar.get() == 1 and len(flights) > 0:
        m = [m for m in m if m.flight == flights[flightTracker[day]] or not m.flight]
    #If looking for commanders only, filter out deputies
    if position == "M":
        m = [m for m in m if m.position == "M"]
    #Filter list to those with the space available to pull the alert
    m = [m for m in m if m.schedule[day] == "" and m.schedule[day+1] in tDay and m.schedule[day+2] in oDay]
    # Filter to SCP certified only if required
    if site[1:] in scps:
        m = [m for m in m if m.scp]
    # If we found MCCM possibilities:
    if m:
        return m
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
def assignAlert(day, site):
    #print("============================== {0} | {1} =================================".format(day,site))
    global integralCount,alertCount
    for position in ["M", "D"]:
        # If the alert hasn't already been assigned...
        Alert = [a for a in alertList[day] if a.site == site][0]
        assigned = False
        if position == "M":
            if Alert.mccc:
                assigned = True
        if position == "D":
            if Alert.dmccc:
                assigned = True
        if not assigned:
            mccmList = findMCCMs(day, position, site)
            # If we have at least one mccm to assign
            if mccmList:
                #Calculate ratios in order to find a best match
                pr = getPositionRatio(mccmList, position)
                cp = getCrewPairingRatio(mccmList, day, site, position)
                ws = getWhiteSpaceRatio(mccmList, day)
                ac = getAlertCountRatio(mccmList)
                mccmRatios = {}
                for m in mccmList:
                    mccmRatios[m] = cp[m] + ws[m] + ac[m] + pr[m]
                #for m in mccmRatios:
                    #print("=============================")
                    # print("Name:{0}\nPositionRatio:{1}\nCrewPairRatio:{2}\nWhiteSpaceRatio:{3}\nAlertCountRatio:{4}\nTotal:{5}".format(m.name,
                    #                                                                                                                    pr[m],
                    #                                                                                                                    cp[m],
                    #                                                                                                                    ws[m],
                    #                                                                                                                    ac[m],
                    #                                                                                                                    cp[m] + ws[m] + ac[m] + pr[m]))
                # If we have more than one 'best match', choose one randomly
                mccmChoice = random.choice([m for m in mccmRatios if mccmRatios[m] == max(mccmRatios.values())])

                # Write the alert onto mccm's schedule
                mccmChoice.schedule[day] = "{0}({1})".format(site, position)
                if mccmChoice.schedule[day+1] == "":
                    mccmChoice.schedule[day+1] = "T"
                if mccmChoice.schedule[day+2] == "":
                    mccmChoice.schedule[day+2] = "O"
                if site == "B1":
                    mccmChoice.backupCount += 1
                else:
                    mccmChoice.alertCount += 1
                if position == "D":
                    Alert.dmccc = mccmChoice
                elif position == "M":
                    Alert.mccc = mccmChoice
                alertCount += 1
                #Update integral crew rating if crew pairing is together
                if Alert.mccc and Alert.dmccc:
                    if Alert.mccc in Alert.dmccc.crewPartners:
                        integralCount += 2

                # #Check opposite alert position and fill if possible
                # assigned = False
                # if oppositePosition(position) == "D":
                #     if Alert.dmccc: assigned = True
                # if oppositePosition(position) == "M":
                #     if Alert.mccc: assigned = True
                # if not assigned:
                #     #Find potential crew partners
                #     partners = [p for p in mccmChoice.crewPartners if p.alertCount + p.backupCount < p.alertMax]
                #     #Ensure they have space required to allocate the alert
                #     partners = [p for p in partners if p.schedule[day] == "" and p.schedule[day + 1] in tDay and p.schedule[day + 2] in oDay]
                #     #If looking to fill commander seat, ensure they are a commander
                #     if oppositePosition(position) == "M":
                #         partners = [p for p in partners if p.position == "M"]
                #     #And if scp is a factor, ensure they are scp certified
                #     if site[1:] in scps:
                #         partners = [p for p in partners if p.scp]
                #     if partners:
                #         #Choose a partner if more than one
                #         partner = random.choice(partners)
                #         # Assign integral crew partner the other alert position if found.
                #         partner.schedule[day] = "{0}({1})".format(site, oppositePosition(position))
                #         if partner.schedule[day + 1] == "":
                #             partner.schedule[day + 1] = "T"
                #         if partner.schedule[day + 2] == "":
                #             partner.schedule[day + 2] = "O"
                #         if site == "B1":
                #             partner.backupCount += 1
                #         else:
                #             partner.alertCount += 1
                #         if oppositePosition(position) == "D":
                #             Alert.dmccc = partner
                #         elif oppositePosition(position) == "M":
                #             Alert.mccc = partner
                #         integralCount += 2
                #         alertCount += 1

# Reset variables to set up the program for another run
def resetVars():
    global rows, alertCount, integralCount, alertList, mccms, sched, schedules
    rows = None
    alertCount = 0
    integralCount = 0
    alertList = {}
    mccms = []
    sched = None
    schedules = []

if __name__ == "__main__":
    # Define globals
    alertCount = 0
    integralCount = 0
    rows = ''
    alertList = {}
    mccms = []
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
    schedules = []
    sched = None
    # creates lists of types of days that a first or second O day can fall on.
    plccs = []
    scps = []
    tDay = ['', 'm2']
    oDay = ['r', 'e', 'lv', '', 'm1', 'm2', 't10', 'sd', 'wd', 'l']

    # Run the program
    window = progWindow()
    window.goButton.configure(command=execute)
    window.mainloop()
