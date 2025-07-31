import win32com.client as COM
import pywintypes
from datetime import datetime, timedelta
from office_utils import *
from parse_json import *

def schedule_appointment():
    try:
        # Load data from stdin
        data = read_json()
        form, case_details, lawyer = split_data(data)
        
        # Parse form data
        form_data = {
            'client_title': form.get("clientTitle", ""),
            'client_name': form.get("clientName", ""),
            'client_email': form.get("clientEmail", ""),
            'client_phone': form.get("clientPhone", ""),
            'client_language': form.get("clientLanguage", "English").lower(),
            'location': form.get("location", ""),
            'is_existing_client': form.get("isExistingClient", False),
            'is_ref_barreau': form.get("isRefBarreau", False),
            'is_first_consultation': form.get("isFirstConsultation", False),
            'is_payment_made': form.get("isPaymentMade", False),
            'payment_method': form.get("paymentMethod", ""),
            'appointment_date': form.get("appointmentDate", ""),
            'appointment_time': form.get("appointmentTime", ""),
            'notes': form.get("notes", ""),
            'case_type': form.get("caseType", ""),
            'case_details': case_details or {}
        }
        
        # Parse lawyer data
        lawyer_data = {
            'id': lawyer.get("id", ""),
            'name': lawyer.get("name", ""),
            'email': lawyer.get("email", ""),
            'break_minutes': lawyer.get("breakMinutes", 0)
        }
        
        # Validate the manual time slot
        if not form_data['appointment_date'] or not form_data['appointment_time']:
            raise Exception("Appointment date and time are required.")
            
        # Parse date and time separately to ensure proper datetime object
        appointment_date = datetime.strptime(form_data['appointment_date'], "%Y-%m-%d").date()
        appointment_time = datetime.strptime(form_data['appointment_time'], "%H:%M").time()

        # Create the datetime as normal
        python_start_time = datetime.combine(appointment_date, appointment_time)
        python_end_time = python_start_time + timedelta(hours=1)

        # Convert to COM-compatible time objects
        # This is the critical step that prevents the timezone confusion
        start_time = pywintypes.Time(python_start_time)
        end_time = pywintypes.Time(python_end_time)
        
        # Fetch calendar events and validate the slot
        events = fetch_calendar_events()
        if not is_valid_time_slot(start_time, end_time, events, lawyer_data):
            raise Exception("The selected time slot conflicts with existing appointments.")
        
        # 3. Create meeting draft
        create_meeting_draft(form_data, lawyer_data, start_time, end_time)
        
        print("Meeting draft created in Outlook.")
        
    except Exception as e:
        error_msg = f"Error scheduling appointment: {str(e)}"
        print(error_msg)
        alert_error(error_msg)
        raise

def create_meeting_draft(form_data, lawyer_data, start_time, end_time):
    """Create an Outlook meeting draft."""
    try:
        # Initialize Outlook
        outlook = COM.gencache.EnsureDispatch("Outlook.Application")
        
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()

        olAppointmentItem = 1
        olMeeting = 1
        
        appt = outlook.CreateItem(olAppointmentItem)
        
        # Set basic properties
        try:
            appt.Subject = form_data['client_name']
            appt.Start = start_time
            appt.End = end_time
            appt.Location = form_data['location'].title()
            appt.RequiredAttendees = lawyer_data['email']
            appt.MeetingStatus = olMeeting
            appt.Body = " "  # Required placeholder for WordEditor
        except Exception as e:
            raise Exception(f"Failed to set appointment properties: {str(e)}")
        
        # Force name resolution
        recipients = appt.Recipients
        recipients.ResolveAll()
        
        # Set categories
        try:
            appt.Categories = lawyer_data['name']
        except:
            pass
                       
        # Display appointment to access WordEditor
        appt.Display()
        
        # Get inspector and Word document
        inspector = appt.GetInspector
        word_doc = None
        
        try:
            word_doc = inspector.WordEditor
            if word_doc:
                content_range = word_doc.Content
                content_range.Delete()
                build_appointment_content(word_doc, form_data)
                focus_office_window(appt)       
            else:
                raise Exception("Failed to initialize Word editor - document remained locked")
        except:
            pass
        

            
    except Exception as e:
        raise Exception(f"Failed to create meeting draft: {str(e)}")
        
    finally:
        # Note: Don't release appt object - we want it to remain open
        pass

def build_appointment_content(word_doc, form_data):
    """Build the appointment content using Word's native formatting capabilities."""
    
    # Get the main content range to work with
    doc_range = word_doc.Content
    
    # Start building content section by section
    # Client information section
    client_info = (f"Client:   {form_data['client_title']} {form_data['client_name']}\n"
                  f"Phone:  {form_data['client_phone']}\n"  
                  f"Email:    {form_data['client_email']}\n"
                  f"Lang:     {form_data['client_language'].title()}\n\n")
    
    doc_range.InsertAfter(client_info)
    
    # Make the email address a hyperlink (this is where Word's power shines)
    # Find the email text we just inserted
    email_range = word_doc.Content
    email_range.Find.Text = form_data['client_email']
    if email_range.Find.Execute():
        # Convert to hyperlink
        word_doc.Hyperlinks.Add(
            Anchor=email_range, 
            Address=f"mailto:{form_data['client_email']}"
        )
    
    # Handle client type and pricing information
    if form_data['is_existing_client']:
        existing_text = "Existing Client\n\n"
        doc_range.InsertAfter(existing_text)
        
        # Make "Existing Client" green and bold
        existing_range = word_doc.Content
        existing_range.Find.Text = "Existing Client"
        if existing_range.Find.Execute():
            existing_range.Font.Color = 0x008000  # Green in BGR format
            existing_range.Font.Bold = True
    else:
        # Add pricing information with appropriate highlighting
        if form_data['is_ref_barreau']:
            price_text = "Ref. Barreau ($60+tax)"
            format_type = "underline"
        elif form_data['is_first_consultation']:
            price_text = "First Consultation ($125+tax)" 
            format_type = "yellow_highlight"
        else:
            price_text = "Follow-up ($350+tax)"
            format_type = "gray_highlight"
            
        case_details = get_case_details(
            case_type=form_data['case_type'],
            case_details_dict=form_data['case_details'],
            client_language=form_data['client_language'],
            format_type="detailed_text",
            include_details=True
        )
        
        pricing_text = f"{price_text}: {case_details}\n\n"
        doc_range.InsertAfter(pricing_text)
        
        # Apply highlighting to the price portion
        price_range = word_doc.Content
        price_range.Find.Text = price_text
        if price_range.Find.Execute():
            if format_type == "underline":
                # Word constant for single underline
                price_range.Font.Underline = 1  # wdUnderlineSingle
            elif format_type == "yellow_highlight":
                # Word constant for yellow highlighting
                price_range.HighlightColorIndex = 7  # Yellow
            elif format_type == "gray_highlight":
                # Word constant for gray highlighting  
                price_range.HighlightColorIndex = 16  # Gray 25%
            
        # Add payment status
        payment_check = "✔️" if form_data['is_payment_made'] else "❌"
        payment_text = f"Payment {payment_check}\n"
        if form_data['is_payment_made']:
            payment_text += f"{form_data['payment_method']}\n"
        payment_text += "\n"
        
        doc_range.InsertAfter(payment_text)
    
    # Add notes section
    if form_data['notes']:
        notes_text = f"Notes:\n{form_data['notes']}\n"
        doc_range.InsertAfter(notes_text)
        
        # Make the notes italic
        notes_range = word_doc.Content
        notes_range.Find.Text = form_data['notes']
        if notes_range.Find.Execute():
            notes_range.Font.Italic = True

def fetch_calendar_events(days_ahead=14):
    """Fetch calendar events from Outlook."""
    print("Fetching calendar events...")
    events = []
    
    try:
        outlook_app = COM.Dispatch("Outlook.Application")
        namespace = outlook_app.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar
        items = calendar.Items
        
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        
        now = datetime.now()
        end_date = now + timedelta(days=days_ahead)
        
        filter_str = f"[Start] <= '{end_date.strftime('%m/%d/%Y %H:%M')}' AND [End] >= '{now.strftime('%m/%d/%Y %H:%M')}'"
        restricted_items = items.Restrict(filter_str)
        
        for item in restricted_items:
            if item and not getattr(item, 'AllDayEvent', True):
                try:
                    events.append(item)
                except:
                    continue
                    
        print(f"Fetched {len(events)} calendar appointments.")
        return events
        
    except Exception as e:
        print(f"Error fetching calendar events: {e}")
        return []

def is_valid_time_slot(start_time, end_time, events, lawyer_data):
    """Check if the requested time slot is valid (no conflicts)."""
    # Check for null events
    if any(event is None for event in events):
        error_msg = "Warning: Null calendar items found in events list!"
        print(error_msg)
        alert_error(error_msg)
        return False
        
    # Check for overlaps with lawyer's appointments
    for event in events:
        if not event:
            continue
            
        try:
            # Check if this is the lawyer's appointment
            categories = getattr(event, 'Categories', '')
            if not categories or lawyer_data['name'] not in categories:
                continue
                
            event_start = event.Start
            event_end = event.End
            
            # Apply break buffer
            buffer_start = event_start - timedelta(minutes=lawyer_data['break_minutes'])
            buffer_end = event_end + timedelta(minutes=lawyer_data['break_minutes'])
            
            # Check for overlap
            if not (end_time <= buffer_start or start_time >= buffer_end):
                return False
                
        except:
            continue
            
    return True

if __name__ == "__main__":
    schedule_appointment()