import win32com.client
from jira import JIRA
import datetime
from models.activity import Activity
from datetime import date
from config.config import email_password, JIRA_URL, STATUS_TICKET, PROYECTO, USER_ID

meetings=[]
def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] < '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    print("*******************************************************************")
    return calendar

def get_data_from_calendar(discard):
    begin = date.today() - datetime.timedelta(days=1)
    end = datetime.datetime.now() + datetime.timedelta(days=1)
    cal = get_calendar(begin, end)
    print("********************* GETTING CALENDAR INFO ***********************")
    for metting in cal:
        if not discard in metting.subject:           
            star_date=str(metting.start).split("+")[0]
            end_date=str(metting.end).split("+")[0]
            if star_date.split(" ")[0] == str(date.today()):
                time_start = datetime.datetime.strptime(star_date, '%Y-%m-%d %H:%M:%S')
                time_end = datetime.datetime.strptime(end_date, '%Y-%m-%d %H:%M:%S')
                elapsed=time_end-time_start
                meetings.append(Activity(metting.subject,PROYECTO,str(time_start.date()),elapsed.seconds/3600))
    print("************************* COMPLETED ***********************************")
        
def get_data_from_jira():
    try:
        print("******************** GETTING JIRA INFO ***********************")
        print("Connecting to JIRA: %s" % JIRA_URL)
        jira_options = {'server': JIRA_URL}
        jira = JIRA(options=jira_options, basic_auth=(USER_ID, email_password))
        new_issues = jira.search_issues(f"assignee = currentUser() AND statusCategory = '{STATUS_TICKET}' AND resolution = Unresolved", maxResults=100)
        for issue in new_issues:
            issue_num = issue.key
            issue = jira.issue(issue_num)
            url_ticket=JIRA_URL+"browse/"+issue.key
            meetings.append(Activity(issue.fields.summary,PROYECTO,str(date.today()),7,url_ticket))
        print("************************* COMPLETED ***********************************")
    except Exception as e:
        print("Failed to connect to JIRA: %s" % e)

def get_mettings():
    get_data_from_calendar(discard="Canceled:")
    get_data_from_jira()
    return meetings