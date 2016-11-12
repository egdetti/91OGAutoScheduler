import os
from tkinter import *
from tkinter import filedialog, Frame, BOTH
import csv
import fnmatch

os.chdir('C:/Users/bdsmi_000/Desktop')

stats_file = csv.reader(open('STATS.csv', 'r'))

#rows used for formatting into usable data
rows = [row for row in stats_file]

#Establish global variables for calling statistics from GUI
crew_members = []
integral_crew_counts = {}
leave_day = 0
deputy_alerts = 0
commander_alerts = 0
av_cc = 0
av_dc = 0
backups = 0
line_commanders = 0
total_commanders = 0
all_deputies = 0
line_deputies = 0
line_commander_alerts = 0
line_deputy_alerts = 0
inactive_members = 0
crew_percentages = 0
#_________________________________________format input sheet__________________________________
def format_sheet():
    global crew_members
    #trim to the correct date.  Time piece will give you 31 days no matter what.
    for date in rows[0]:
        if '1-' in date:
            EOM_date = rows[0].index(date)
            if EOM_date == 1:
                pass
            else:
                for row in rows:
                    for i in range(EOM_date-1, 31):
                        del row[i]
    #bring alerts to the correct list
    for row in range(len(rows)):
        for i in range(len(rows[1])):
            if rows[row][i] == 'PRP' and '91' not in rows[row][0]:
                if 'A' in rows[row+1][i] or 'B' in rows[row+1][i]:
                    rows[row][i] = rows[row+1][i]
    #create list of crew numbers
    crew_nums = []
    for row in range(len(rows)):
        if '91' in rows[row][0]:
            crew_nums.append(rows[row][0][-3:])
    #create crewmember only list
    crew_members = [row for row in rows if len(row[0])>1 and '91' not in row[0]]
    #deletes dates from rows
    del crew_members[0]
    #insert crew number into crew member list
    for i in range(0, len(crew_members)):
        crew_members[i].insert(1, crew_nums[i])
    #figure out crew member position by looking for commander alerts in their schedule
    for row in range(2, len(rows)):
        global inactive_members
        try:
            if '91'in rows[row+1][0]:
                if '(M)' in rows[row+1][2:] or '/' not in rows[row][1]:
                    rows[row].insert(2, 'M')
                elif '(D)' in rows[row+1][2:]:
                    rows[row].insert(2, 'D')
                else:
                    inactive_members+=1
                    rows[row].insert(2, 'I')
            else:
                pass
        except:
            pass
    #Save file
    with open('STATSOUT.csv', 'w', newline='') as outFile:
        writer = csv.writer(outFile)
        for row in crew_members:
            writer.writerow(row)

#________________________________________________STATS________________________________________________________________________________________________

#average alert counts for mcc and dmcc/# of commanders/# of deputies/#line commander alerts/#of deputy alerts
def baseline_stats():
    global total_commanders
    global line_commander_alerts
    global line_deputy_alerts
    global line_commanders
    global line_deputies
    global all_deputies
    global deputy_alerts
    global av_dc
    global av_cc
    for member in crew_members:
        #counts commanders who are crew paired
        if '/' in member[1] and member[1][2] in ['2', '3', '4', '5', '6', '7', '8', '9'] and member[2] == 'M':
            line_commanders+=1
            for i in member[3:]:
                if 'A' in i:
                    line_commander_alerts+=1
        #counts all deputies in sqaudron
        if member[2] == 'D':
            all_deputies+=1
            for i in member[3:]:
                if 'A' in i:
                    deputy_alerts+=1
        #counts just deputies who are crew paired.  Usef for calculating integral crew%
        if member[2] == 'D' and member[1][2] != '1':
            line_deputies+=1
            for i in member[3:]:
                if 'A' in i:
                    line_deputy_alerts+=1
        #calculates total # of commanders in Sq including leadership
        if member[2]=='M':
            total_commanders+=1
    #figures average alerts by position
    av_cc = line_commander_alerts/line_commanders
    av_dc = deputy_alerts/all_deputies


#finds crew pair, then counts alerts pulled together excluding backup alerts and non line crew members
def integral_alerts():
    global line_deputies
    global integral_crew_counts
    for row in range(0, len(rows)):
        try:
            if crew_members[row][1] == crew_members[row+1][1] and crew_members[row][1][2] in ['2', '3', '4', '5', '6', '7', '8', '9']:
                crew_pair = str(crew_members[row][0]+'/'+ crew_members[row+1][0])
                integral_crew_counts[crew_pair]=0
                for i in range(3, len(crew_members[row])):
                    if crew_members[row][i] == crew_members[row+1][i] and 'A' in crew_members[row][i]:
                        integral_crew_counts[crew_pair] += 1
                    else:
                        pass
            else:
                pass
        except:
            pass


#number of days of leave
def leave_days():
    global leave_day
    for row in rows:
        for i in row:
            if i == 'LV':
                leave_day += 1

#number of backups
def backup_alerts():
    global backups
    for row in rows:
        for i in row:
            if i == 'B1':
                backups +=1

#Crew Pair percentages
def integral_crew_percentages():
    global crew_percentages
    total_integral_alerts = 0
    for k, v in integral_crew_counts.items():
        total_integral_alerts+=v
    total_integral_alerts*2
    crew_percentages = (total_integral_alerts/(line_deputy_alerts+line_commander_alerts))*100

#runs all fucntions
def run_all():
    format_sheet()
    baseline_stats()
    integral_alerts()
    integral_crew_percentages()
    leave_days()
    backup_alerts()

#______________________________Run Lines____________________________________________________________________________________________________________________________________________________

#executes all functions
run_all()

#prints results of global variables
print('Average Line Commander Alerts = '+str(av_cc))
print('Average deputy Alerts = '+str(av_dc))
print('Number of Line Commanders ='+str(line_commanders))
print('Number of total Commanders ='+str(total_commanders))
print('Number of line deputies ='+str(line_deputies))
print('Number of total deputies ='+str(all_deputies))
print('Number of line commander alerts ='+str(line_commander_alerts))
print('Number of total commander alerts ='+str(commander_alerts))
print('Number of deputy alerts ='+str(deputy_alerts))
print('integral crew percentages ='+str(crew_percentages)[0:5]+'%')
print('integral crew alerts by crew:')
for k,v in integral_crew_counts.items():
    print(str(k)+' had '+str(v)+' integral crew alerts')
print('number of inactive crew members ='+str(inactive_members))
print('Number of B1 days ='+str(backups))
print('Number of LV days ='+str(leave_day))




#_______________________________________________GUI_________________________________________________
#
#
# class Stats_GUI:
#    def __init__(self, master):
#         self.master = master
#         master.title("Scheduling Statistics")
#
#         self.label1 = Label(master, text="Input File Location").grid(column=1, row=1)
#         self.input_button = Button(master, text = 'Browse').grid(column=1,row=2)
#         self.label2 = Label(master, text="Output File Location").grid(column=1, row=3)
#         self.output_button = Button(master, text = 'Search').grid(column=1,row=4)
#         self.run_button = Button(master, text= 'RUN', command=format_sheet()).grid(column=2, row=5)
#
#
#
#
#
# root = Tk()
# my_gui = Stats_GUI(root)
# root.mainloop()
#
#


