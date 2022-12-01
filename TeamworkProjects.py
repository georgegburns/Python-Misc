import pandas as pd 
import xlsxwriter 
from dateutil.rrule import DAILY, rrule, MO, TU, WE, TH, FR
def daterange(start_date, end_date):
    return rrule(DAILY, dtstart=start_date, until=end_date, byweekday=(MO, TU, WE, TH, FR))
TeamP = pd.read_excel('C:\\Users\\George.Burns\\Downloads\\Insight Work\\Recurring Tasks\\Reporting Temp\\TeamworkProject.xlsx')
TeamP = TeamP[TeamP['Task list'] == 'Insight']
TeamP = TeamP[~TeamP['Start date'].isnull()]
TeamP = TeamP[~TeamP['Due date'].isnull()]
TeamP['Due date'] = TeamP['Due date'].astype('datetime64[D]')
TeamP = TeamP[['Task name', 'Start date', 'Due date', 'Assigned to', 'Time estimate']].reset_index()
TeamP['Time estimate (Days)'] = (TeamP['Time estimate']/60)
TeamP['Days to Complete'] = TeamP['Due date'] - TeamP['Start date']
Days = []
for i in TeamP.itertuples():
    sd = i[3]
    ed = i[4]
    Dates = daterange(sd,ed)
    Dates = list(Dates)
    if Dates != []:
        Work = len(Dates)
        Days.append(Work)
    else:
        Days.append(1)
TeamP['Days to Complete'] = Days
TeamP['Hours per Day'] = TeamP['Time estimate (Days)']/TeamP['Days to Complete']
TeamP = TeamP.reset_index()
Times = []
Users =[]
Work = []
Number = 0
for i in TeamP.itertuples():
    start_date = i[4]
    end_date = i[5]
    timeline = daterange(start_date, end_date)
    for date in list(timeline):
        Times.append(date)
        Users.append(i[6])
        Work.append(i[10])
    Number += 1
    print('Loop No.: ' + str(Number))
Columns = ['Date','User','Workload']
Workload = pd.DataFrame(zip(Times, Users, Work),columns=Columns)
Workload = Workload.replace('Ollie Williams, George Burns', 'George Burns')
Workload = Workload.replace('Ollie Williams, Kate Lang', 'Kate Lang')
Workload = Workload.groupby(['Date', 'User']).agg({'Workload':'sum'}).reset_index()
out_path = "C:\\Users\\George.Burns\\Downloads\\Insight Work\\Recurring Tasks\\Reporting Temp\\TeamworkProjects.xlsx"
writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
Workload.to_excel(writer, sheet_name='Work')
TeamP.to_excel(writer, sheet_name = 'Check')
writer.save()
print('Document has been successful exported.')