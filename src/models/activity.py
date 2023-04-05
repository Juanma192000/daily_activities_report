class Activity:
    def __init__(self, title,proyect,_date, total_hrs,jira_ticket=""):        
        self.title = title
        self.proyect = proyect
        self._date = _date
        self.total_hrs = total_hrs
        self.jira_ticket = jira_ticket