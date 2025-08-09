from time import sleep
from office_utils import *

def create_test_appointment():
    """Create a minimal test appointment and return the appointment and inspector objects."""
    try:
        print("Creating test appointment...")
        
        # Initialize Outlook
        outlook = COM.gencache.EnsureDispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()

        # Create appointment
        olAppointmentItem = 1
        appt = outlook.CreateItem(olAppointmentItem)
        
        # Set minimal properties
        appt.Subject = "Test Meeting"
        appt.Body = "Test appointment for Teams integration"
        
        # Display the appointment
        appt.Display()
        
        # Get inspector
        inspector = appt.GetInspector
        
        print("Test appointment created and displayed")
        return appt, inspector
        
    except Exception as e:
        print(f"Error creating test appointment: {e}")
        return None, None

def main():
    print("Teams Meeting Integration Test")
    print("=" * 40)
    
    # Create test appointment
    appt, inspector = create_test_appointment()
    connect_to_meeting_window()
    sleep(1)
    click_teams_meeting_button()

    block = get_teams_meeting_block()
    if block:
        print("=== Captured Teams Block ===")
        print(block)

if __name__ == "__main__":
    main()