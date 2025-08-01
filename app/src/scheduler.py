from office_utils import *
from parse_json import *
import win32com.client as COM
from datetime import datetime, timedelta

def process_scheduler(data):
    """
    Process appointment scheduling with the provided data.
    Args:
        data (dict): The parsed JSON data containing form information
    Returns:
        dict: Result of the scheduling process
    """
    try:
        # Use passed data instead of reading from stdin
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
            
        # SIMPLIFIED APPROACH: Create datetime strings that COM can parse directly
        # This avoids all the timezone conversion issues
        start_datetime_str = f"{form_data['appointment_date']} {form_data['appointment_time']}"
        
        # Calculate end time (1 hour later)
        temp_datetime = datetime.strptime(start_datetime_str, "%Y-%m-%d %H:%M")
        end_datetime = temp_datetime + timedelta(hours=1)
        end_datetime_str = end_datetime.strftime("%Y-%m-%d %H:%M")
        
        print(f"Scheduling appointment from {start_datetime_str} to {end_datetime_str}")
        
        # Fetch calendar events and validate the slot
        events = fetch_calendar_events(temp_datetime)
        if not is_valid_time_slot(temp_datetime, end_datetime, events, lawyer_data):
            raise Exception("The selected time slot conflicts with existing appointments.")
        
        # Create meeting draft
        create_meeting_draft(form_data, lawyer_data, start_datetime_str, end_datetime_str)
        
        # Return success result
        return {
            "status": "success", 
            "message": "Meeting draft created in Outlook.",
            "appointment_time": start_datetime_str,
            "client": form_data['client_name']
        }
        
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

def fetch_calendar_events(appointment_datetime):
    """
    Fetch calendar events for the specific day of the appointment.
    Returns only events that overlap with that day.
    """
    print(f"Fetching calendar events for {appointment_datetime.strftime('%Y-%m-%d')}...")
    events = []

    try:
        outlook_app = COM.Dispatch("Outlook.Application")
        namespace = outlook_app.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar
        items = calendar.Items

        # Ensure recurring events are included and sorted properly
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        # Create start and end of the day
        day_start = appointment_datetime.replace(hour=0, minute=0, second=0, microsecond=0)
        day_end = appointment_datetime.replace(hour=23, minute=59, second=59, microsecond=999999)

        fmt = "%B %d, %Y %I:%M %p"
        ds = day_start.strftime(fmt)
        de = day_end.strftime(fmt)

        # Overlap filter (start <= day_end AND end >= day_start)
        filter_str = (
            f"[Start] <= '{de}' AND "
            f"[End]   >= '{ds}'"
        )

        restricted = items.Restrict(filter_str)

        # Now parse the final list
        for item in restricted:
            if item and not getattr(item, 'AllDayEvent', True):
                try:
                    events.append(item)
                    print(f"Found appointment: {item.Subject} from {item.Start} to {item.End}")
                except Exception as e:
                    print(f"Warning: Could not process calendar item: {e}")
                    continue

        print(f"Found {len(events)} relevant appointments on {appointment_datetime.strftime('%Y-%m-%d')}")
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
            
            # TODO
                
        except:
            continue
            
    return True

# Backward compatibility
def main():
    data = read_json()
    result = process_scheduler(data)
    if result.get("error"):
        print(f"Error: {result['error']}")
    else:
        print(result.get("message", "Appointment scheduled"))

if __name__ == "__main__":
    main()