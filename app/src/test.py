import win32com.client
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

try:
    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    namespace.Logon()

    appt = outlook.CreateItem(1)  # Try different types for testing
    print("Created item class:", appt.Class)  # 26 = Appointment, 43 = Mail

    appt.Subject = "Test Meeting"
    appt.Start = datetime.now() + timedelta(hours=1)
    appt.End = datetime.now() + timedelta(hours=2)
    print(datetime.now().astimezone().tzinfo)
    appt.Location = "Office"
    appt.RequiredAttendees = "test@example.com"
    appt.MeetingStatus = 1
    appt.Body = "Test meeting body"

    appt.Display()

except Exception as e:
    print("Error:", e)
