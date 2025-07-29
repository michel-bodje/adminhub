import win32com.client as COM
from datetime import datetime, timedelta
import tempfile
import os
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
        
        # 2. Validate the manual time slot
        if not form_data['appointment_date'] or not form_data['appointment_time']:
            raise Exception("Appointment date and time are required.")
            
        # Parse date and time separately to ensure proper datetime object
        appointment_date = datetime.strptime(form_data['appointment_date'], "%Y-%m-%d").date()
        appointment_time = datetime.strptime(form_data['appointment_time'], "%H:%M").time()
        start_time = datetime.combine(appointment_date, appointment_time)
        end_time = start_time + timedelta(hours=1)  # Default 1-hour duration
        
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
            
        # Generate HTML body
        html_body = generate_html_body(form_data)
        
        # Create temp file
        with tempfile.NamedTemporaryFile(
            mode='w',
            suffix='.html',
            delete=False,
            encoding='utf-8') as f:
            f.write(html_body)
            temp_file = f.name
            
        # Display appointment to access WordEditor
        appt.Display()
        
        # Get inspector and Word document
        inspector = appt.GetInspector
        word_doc = inspector.WordEditor
        
        if word_doc:
            # Insert HTML file content
            range_obj = word_doc.Content
            range_obj.InsertFile(temp_file, ConfirmConversions=False, Link=False, Attachment=False)
        else:
            raise Exception("Failed to initialize Word editor for appointment body.")
            
    except Exception as e:
        raise Exception(f"Failed to create meeting draft: {str(e)}")
        
    finally:
        # Cleanup temp file
        if temp_file and os.path.exists(temp_file):
            try:
                os.unlink(temp_file)
            except:
                pass
        # Note: Don't release appt object - we want it to remain open

def generate_html_body(form_data):
    """Generate HTML body for the appointment."""
    body = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
</head>
<body>"""
    
    # Client info
    body += f"""<p>
Client:&nbsp;&nbsp;&nbsp;{form_data['client_title']} {form_data['client_name']}<br>
Phone:&nbsp;&nbsp;{form_data['client_phone']}<br>
Email:&nbsp;&nbsp;&nbsp;&nbsp;<a href="mailto:{form_data['client_email']}">{form_data['client_email']}</a><br>
Lang:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{form_data['client_language'].title()}<br>
</p>"""
    
    if form_data['is_existing_client']:
        body += '<p style="color: green;"><strong>Existing Client</strong></p>'
    else:
        # Pricing details
        if form_data['is_ref_barreau']:
            price_details = "Ref. Barreau ($60+tax)"
            price_style = "text-decoration: underline;"
        elif form_data['is_first_consultation']:
            price_details = "First Consultation ($125+tax)"
            price_style = "background-color: yellow;"
        else:
            price_details = "Follow-up ($350+tax)"
            price_style = "background-color: #d3d3d3;"
            
        case_details_html = get_case_details_html(form_data['case_type'], 
                                                form_data['case_details'])
        
        body += f'<p><span style="{price_style}">{price_details}</span>: {case_details_html}</p>'
        
        # Payment status
        payment_check = "✔️" if form_data['is_payment_made'] else "❌"
        body += f'<p><strong>Payment</strong>  {payment_check}<br>'
        if form_data['is_payment_made']:
            body += form_data['payment_method']
        body += '</p>'
        
    # Notes
    body += f'<p>Notes:<br><span style="font-style: italic">{form_data["notes"]}</span></p>'
    body += '\n</body>\n</html>'
    
    return body

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

def get_case_details_html(case_type, case_details):
    """Generate HTML for case details based on case type."""
    if not case_type or not case_details:
        return ""
        
    def get_detail(key):
        return case_details.get(key, "")
        
    def conflict_check(done):
        return "✔️" if done else "❌"
    
    case_type = case_type.lower()
    
    if case_type == "divorce":
        conflict_done = case_details.get("conflictSearchDone", False)
        return (f"Divorce / Family Law<br><br>"
                f"<strong>Spouse Name:</strong> {get_detail('spouseName')}<br>"
                f"Conflict Search Done? {conflict_check(conflict_done)}")
                
    elif case_type == "estate":
        conflict_done = case_details.get("conflictSearchDone", False)
        return (f"Successions / Estate Law<br><br>"
                f"<strong>Deceased Name:</strong> {get_detail('deceasedName')}<br>"
                f"<strong>Executor Name:</strong> {get_detail('executorName')}<br>"
                f"Conflict Search Done? {conflict_check(conflict_done)}")
                
    elif case_type == "employment":
        return (f"Employment Law<br><br>"
                f"<strong>Employer Name:</strong> {get_detail('employerName')}")
                
    elif case_type == "contract":
        return (f"Contract Law<br><br>"
                f"<strong>Other Party:</strong> {get_detail('otherPartyName')}")
                
    elif case_type == "mandates":
        return (f"Regimes de Protection / Mandates<br><br>"
                f"<strong>Mandate Details:</strong> {get_detail('mandateDetails')}")
                
    elif case_type == "business":
        return (f"Business Law<br><br>"
                f"<strong>Business Name:</strong> {get_detail('businessName')}")
                
    elif case_type == "common":
        return get_detail("commonField")
        
    elif case_type in ["defamations", "real_estate", "name_change", 
                       "adoptions", "assermentation"]:
        case_names = {
            "defamations": "Defamations",
            "real_estate": "Real Estate", 
            "name_change": "Name Change",
            "adoptions": "Adoptions",
            "assermentation": "Assermentation"
        }
        return case_names.get(case_type, "")
        
    return ""

if __name__ == "__main__":
    schedule_appointment()