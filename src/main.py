from lib.excel_write import fill_excel_with_meetings
from lib.send_confirm_mail import send_email

if __name__ == "__main__":
    try:
        fill_excel_with_meetings()
        send_email("Home office file has been updated succesfully !!")
    except:
        send_email("An error has occureed while trying to update Home office file.")