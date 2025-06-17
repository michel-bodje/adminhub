using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

class Scheduler
{
    static string baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..");
    static string jsonPath = Path.Combine(baseDir, "app", "data.json");

    [STAThread] // Required for clipboard
    static void Main(string[] args)
    {
        try
        {
            // 1. Load data.json
            if (!File.Exists(jsonPath))
                throw new FileNotFoundException("data.json not found.", jsonPath);
            JObject root = JObject.Parse(File.ReadAllText(jsonPath));
            var form = ParseFormState(root);

            // 2. Validate minimal inputs
            if (string.IsNullOrEmpty(form.ClientName) || string.IsNullOrEmpty(form.LawyerName))
                throw new Exception("Missing client or lawyer info.");
            if (string.IsNullOrEmpty(form.ClientEmail) || !Util.IsValidEmail(form.ClientEmail))
                throw new Exception("Invalid client email.");
            if (string.IsNullOrEmpty(form.Location))
                throw new Exception("Missing location type.");
            if (!Util.IsValidPhoneNumber(form.ClientPhone))
                throw new Exception("Invalid client phone number.");

            // 3. Determine scheduling mode
            bool hasManual = !string.IsNullOrEmpty(form.AppointmentDate)
                         && !string.IsNullOrEmpty(form.AppointmentTime);
            List<Slot> validSlots = hasManual
                ? FindManualSlot(form)
                : FindAutoSlots(form);

            if (validSlots == null || validSlots.Count == 0)
                throw new Exception("No available slots found.");

            // 4. Pick a slot
            Slot selected = PickSlot(validSlots);

            // 5. Create meeting draft
            CreateMeetingDraft(form, selected);

            Console.WriteLine("Meeting draft created in Outlook.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error scheduling appointment: " + ex.Message);
            Console.WriteLine("StackTrace: " + ex.StackTrace);

            string errorMsg = "Error scheduling appointment:\n" + ex.Message + "\n\n"
                + "StackTrace:\n" + ex.StackTrace;

            System.Windows.Forms.MessageBox.Show(
                errorMsg,
                "Error",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error
            );
            System.Windows.Forms.Clipboard.SetText(errorMsg);
        }
    }

    // ---- Data structures ----
    class FormState
    {
        // From form:
        public string ClientTitle;
        public string ClientName;
        public string ClientEmail;
        public string ClientPhone;
        public string ClientLanguage; // "Français" or other
        public string Location;       // e.g. "office", "phone", "teams"
        public bool IsExistingClient;
        public bool IsRefBarreau;
        public bool IsFirstConsultation;
        public bool IsPaymentMade;
        public string PaymentMethod;
        public string AppointmentDate; // "yyyy-MM-dd"
        public string AppointmentTime; // "HH:mm"
        public string Notes;
        public string CaseType;        // if present

        // From lawyer object in JSON:
        public string LawyerId;
        public string LawyerName;
        public string LawyerEmail;
        public TimeSpan WorkingStart;  // parsed from workingHours.start
        public TimeSpan WorkingEnd;    // parsed from workingHours.end
        public int BreakMinutes;       // buffer between appointments
        public int MaxDailyAppointments;
        public string[] Specialties;

        // From caseDetails object in JSON:
        public Dictionary<string, object> CaseDetails;
    }

    class Slot
    {
        public DateTime Start;
        public DateTime End;
        public string Location;
    }

    // ---- Parsing JSON ----
    static FormState ParseFormState(JObject root)
    {
        var f = root["form"];
        var form = new FormState
        {
            ClientTitle = (string)f["clientTitle"],
            ClientName = (string)f["clientName"],
            ClientEmail = (string)f["clientEmail"],
            ClientPhone = (string)f["clientPhone"],
            ClientLanguage = (string)f["clientLanguage"],
            Location = (string)f["location"],
            IsExistingClient = (bool?)f["isExistingClient"] ?? false,
            IsRefBarreau = (bool?)f["isRefBarreau"] ?? false,
            IsFirstConsultation = (bool?)f["isFirstConsultation"] ?? false,
            IsPaymentMade = (bool?)f["isPaymentMade"] ?? false,
            PaymentMethod = (string)f["paymentMethod"],
            AppointmentDate = (string)f["appointmentDate"],
            AppointmentTime = (string)f["appointmentTime"],
            Notes = (string)f["notes"],
            CaseType = (string)f["caseType"]
        };

        var lawyer = root["lawyer"];
        if (lawyer != null)
        {
            form.LawyerId = (string)lawyer["id"];
            form.LawyerName = (string)lawyer["name"];
            form.LawyerEmail = (string)lawyer["email"];

            // Parse working hours "9:00" etc.
            var wh = lawyer["workingHours"];
            if (wh != null)
            {
                string startStr = (string)wh["start"];
                string endStr = (string)wh["end"];
                TimeSpan ws, we;
                if (!TimeSpan.TryParse(startStr, out ws))
                    throw new Exception("Invalid workingHours.start:" + startStr);
                if (!TimeSpan.TryParse(endStr, out we))
                    throw new Exception("Invalid workingHours.end:" + endStr);
                form.WorkingStart = ws;
                form.WorkingEnd = we;
            }
            else
            {
                // Defaults if missing
                form.WorkingStart = TimeSpan.FromHours(9);
                form.WorkingEnd = TimeSpan.FromHours(17);
            }

            form.BreakMinutes = (int?)lawyer["breakMinutes"] ?? 0;
            form.MaxDailyAppointments = (int?)lawyer["maxDailyAppointments"] ?? int.MaxValue;

            var specs = lawyer["specialties"] as JArray;
            if (specs != null)
                form.Specialties = specs.Select(t => (string)t).ToArray();
            else
                form.Specialties = Array.Empty<string>();
        }
        else
        {
            throw new Exception("Missing lawyer object in JSON.");
        }

        // Parse caseDetails as Dictionary<string, object> (compat version)
        var caseDetailsDict = new Dictionary<string, object>();
        var cd = root["caseDetails"] as JObject;
        if (cd != null)
        {
            foreach (var prop in cd.Properties())
            {
            caseDetailsDict[prop.Name] = prop.Value != null ? prop.Value.ToObject<object>() : null;
            }
        }
        form.CaseDetails = caseDetailsDict;

        return form;
    }

    // ---- Manual scheduling ----
    static List<Slot> FindManualSlot(FormState form)
    {
        if (string.IsNullOrEmpty(form.AppointmentDate) || string.IsNullOrEmpty(form.AppointmentTime))
            throw new Exception("Manual date/time missing.");

        DateTime start = DateTime.ParseExact(
            form.AppointmentDate + " " + form.AppointmentTime,
            "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        DateTime end = start.AddHours(1); // or adjust duration if variable

        var slot = new Slot { Start = start, End = end, Location = form.Location };
        var events = FetchCalendarEvents();
        if (!IsValidSlot(slot, events, form))
            throw new Exception("The selected time slot is not valid.");
        return new List<Slot> { slot };
    }

    // ---- Auto scheduling ----
    static List<Slot> FindAutoSlots(FormState form)
    {
        var events = FetchCalendarEvents();
        // Generate raw slots for next RANGE_IN_DAYS days:
        var raw = GenerateSlots(form, events);
        // Validate each
        var valid = new List<Slot>();
        foreach (var slot in raw)
        {
            if (IsValidSlot(slot, events, form))
                valid.Add(slot);
        }
        if (valid.Count == 0)
            throw new Exception("No available slots found in the next period.");
        return valid;
    }

    static Slot PickSlot(List<Slot> slots)
    {
        DateTime now = DateTime.Now;
        // Next strictly future start
        foreach (var s in slots)
        {
            if (s.Start > now)
                return s;
        }
        return slots[0];
    }

    // ---- Outlook meeting draft ----
static void CreateMeetingDraft(FormState form, Slot slot)
{
    Outlook.Application outlookApp = null;
    Outlook.AppointmentItem appt = null;
    string tempFile = null;

    try
    {
        // Initialize Outlook
        outlookApp = new Outlook.Application();
        appt = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);

        // Set basic properties
        appt.Subject = form.ClientName;
        appt.Start = slot.Start;
        appt.Duration = 60;
        appt.Location = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(slot.Location);
        appt.RequiredAttendees = form.LawyerEmail;
        appt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
        appt.Body = " "; // Required placeholder for WordEditor

        // Force name resolution
        Outlook.Recipients recipients = appt.Recipients;
        recipients.ResolveAll();

        // Set categories
        try { appt.Categories = form.LawyerName; } catch { }

        // Generate HTML
        string htmlBody = GenerateHtmlBody(form);
        tempFile = Path.GetTempFileName();
        File.WriteAllText(tempFile, htmlBody);

        // Initialize Word Editor
        appt.Display(); // Must display before accessing WordEditor

        Outlook.Inspector inspector = appt.GetInspector;
        WindowFocus.ShowWithFocus(inspector); // Force focus

        Word.Document wordDoc = inspector.WordEditor as Word.Document;

        if (wordDoc != null)
        {
            Word.Range range = wordDoc.Content;
            range.InsertFile(tempFile, ConfirmConversions: false, Link: false, Attachment: false);
        }
        else
        {
            throw new Exception("Failed to initialize Word editor for appointment body.");
        }

        // Keep the appointment visible
        Marshal.ReleaseComObject(inspector);
    }
    catch (Exception ex)
    {
        throw new Exception("Failed to create meeting draft: " + ex.Message);
    }
    finally
    {
        // Cleanup
        if (tempFile != null && File.Exists(tempFile))
            File.Delete(tempFile);
        
        // Note: Don't release appt object here - we want it to remain open
        if (outlookApp != null)
            Marshal.ReleaseComObject(outlookApp);
    }
}

    // ---- HTML body generation ----
    static string GenerateHtmlBody(FormState form)
    {
        string formattedPhone = Util.FormatPhoneNumber(form.ClientPhone);
        UpdateJsonPhone(formattedPhone);

        // Build body HTML
        string body = "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n</head>\n<body>";
        body += "<p>" +
            "Client:&nbsp;&nbsp;&nbsp;" + form.ClientTitle + " " + form.ClientName + "<br>" +
            "Phone:&nbsp;&nbsp;" + formattedPhone + "<br>" +
            "Email:&nbsp;&nbsp;&nbsp;&nbsp;<a href=\"mailto:" + form.ClientEmail + "\">" + form.ClientEmail + "</a><br>" +
            "Lang:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + form.ClientLanguage + "</p>";

        if (form.IsExistingClient)
        {
            body += "<p style=\"color: green;\"><strong>Existing Client</strong></p>";
        }
        else
        {
            string priceDetails;
            string priceStyle = "";
            if (form.IsRefBarreau)
            {
                priceDetails = "Ref. Barreau ($60+tax)";
                priceStyle = "text-decoration: underline;";
            }
            else if (form.IsFirstConsultation)
            {
                priceDetails = "First Consultation ($125+tax)";
                priceStyle = "background-color: yellow;";
            }
            else
            {
                priceDetails = "Follow-up ($350+tax)";
                priceStyle = "background-color: #d3d3d3;";
            }

            string detailsHtml = GetCaseDetails(form.CaseType, form.CaseDetails);

            body += "<p><span style=\"" + priceStyle + "\">" +
                    priceDetails + "</span>: " + detailsHtml + "</p>";
            body += "<p><strong>Payment</strong>  " + (form.IsPaymentMade ? "✔️" : "❌") + "<br>";
            if (form.IsPaymentMade)
                body += form.PaymentMethod;
            body += "</p>";
        }
        body += "<p>Notes:<br><span style=\"font-style: italic\">" + form.Notes + "</span></p>";
        body += "\n</body>\n</html>";
        return body;
    }

    // ---- Calendar fetch ----
    static List<Outlook.AppointmentItem> FetchCalendarEvents(int daysAhead = 14)
    {
        Console.WriteLine("Entering fetch");
        var list = new List<Outlook.AppointmentItem>();
        Outlook.Application outlookApp = new Outlook.Application();
        Outlook.NameSpace ns = outlookApp.GetNamespace("MAPI");
        Outlook.MAPIFolder calendar = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
        Outlook.Items items = calendar.Items;

        items.IncludeRecurrences = true;
        items.Sort("[Start]");

        DateTime now = DateTime.Now;
        DateTime end = now.AddDays(daysAhead);

        Console.WriteLine("Building filter...");
        string filter = string.Format("[Start] <= '{0:g}' AND [End] >= '{1:g}'", end, now);
        Outlook.Items restrictedItems = items.Restrict(filter);

        Console.WriteLine("Filtering visible appointments...");
        System.Collections.IEnumerator enumerator = restrictedItems.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Outlook.AppointmentItem item = enumerator.Current as Outlook.AppointmentItem;
            
            if (item != null && !item.AllDayEvent)
            {
                try { list.Add(item); }
                catch { continue; /* skip broken ones */}
            }
        }

        Console.WriteLine("Fetched " + list.Count + " visible appointments.");
        return list;
    }


    // ---- Slot generation ----
    static List<Slot> GenerateSlots(FormState form, List<Outlook.AppointmentItem> events)
    {
        var slots = new List<Slot>();
        DateTime now = DateTime.Now;
        DateTime endDate = now.AddDays(14).Date;
        TimeSpan slotDur = TimeSpan.FromHours(1);

        TimeSpan workStart = form.WorkingStart;
        TimeSpan workEnd = form.WorkingEnd;
        // Lunch break
        TimeSpan lunchStart = TimeSpan.FromHours(13);
        TimeSpan lunchEnd = TimeSpan.FromHours(14);

        // Count existing appointments per day to enforce maxDailyAppointments
        var eventsByDay = events
            .Select(ev => {
                DateTime s = ev.Start;
                return new { ev, Day = s.Date };
            })
            .GroupBy(x => x.Day)
            .ToDictionary(g => g.Key, g => g.Count());

        for (DateTime day = now.Date; day <= endDate; day = day.AddDays(1))
        {
            // Skip weekends if desired
            if (day.DayOfWeek == DayOfWeek.Saturday || day.DayOfWeek == DayOfWeek.Sunday)
                continue;

            // Check daily appointment limit
            int cnt = 0;
            if (eventsByDay.ContainsKey(day))
                cnt = eventsByDay[day];
            if (cnt >= form.MaxDailyAppointments)
                continue;

            // Unavailability that day
            if (IsLocationUnavailableOnDay(form, day))
                continue;

            // Determine starting cursor
            DateTime cursor = day.Add(workStart);
            if (day.Date == now.Date)
            {
                DateTime rounded = RoundUpToNextHalfHour(now);
                if (rounded.TimeOfDay > workStart)
                    cursor = day.Date.Add(rounded.TimeOfDay);
            }
            DateTime lastStart = day.Add(workEnd) - slotDur;

            while (cursor <= lastStart)
            {
                // Skip lunch
                if (cursor.TimeOfDay >= lunchStart && cursor.TimeOfDay < lunchEnd)
                {
                    cursor = cursor.AddMinutes(30);
                    continue;
                }

                var slot = new Slot { Start = cursor, End = cursor.Add(slotDur), Location = form.Location };
                slots.Add(slot);
                cursor = cursor.AddMinutes(30);
            }
        }
        return slots;
    }

    // ---- Slot validation ----
    static bool IsValidSlot(Slot slot, List<Outlook.AppointmentItem> events, FormState form)
    {
        if (events.Any(e => e == null))
        {
            string msg = "Warning: Null calendar items found in events list!";
            Console.WriteLine(msg);
            System.Windows.Forms.MessageBox.Show(
                msg,
                "Calendar Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return false;
        }
        
        // 1. Overlap check + breakMinutes buffer
        foreach (var ev in events.Where(ev => ev != null && !string.IsNullOrEmpty(ev.Categories) && ev.Categories.Contains(form.LawyerName)))
        {
            DateTime s, e;
            try { s = ev.Start; e = ev.End; }
            catch { continue; }
            
            // apply break buffer
            DateTime bufStart = s.AddMinutes(-form.BreakMinutes);
            DateTime bufEnd = e.AddMinutes(form.BreakMinutes);
            if (!(slot.End <= bufStart || slot.Start >= bufEnd))
                return false;
        }
        // 2. Enforce maxDailyAppointments: already done in generation by skipping days with >= limit

        // 3. Other refined rules: e.g. shared office, proximity rules—port here if needed
        
        // TODO

        return true;
    }

    // Example unavailability check per lawyer
    static bool IsLocationUnavailableOnDay(FormState form, DateTime day)
    {
        if (form.LawyerId == "DH" && form.Location.Equals("office", StringComparison.OrdinalIgnoreCase)
            && day.DayOfWeek == DayOfWeek.Monday)
            return true;
        if (form.LawyerId == "TG" && form.Location.Equals("office", StringComparison.OrdinalIgnoreCase)
            && day.DayOfWeek == DayOfWeek.Friday)
            return true;
        // Add more rules as needed
        // Default: no unavailability
        return false;
    }

    // ---- Helpers ----
    static DateTime RoundUpToNextHalfHour(DateTime dt)
    {
        int minutes = dt.Minute;
        int add = (minutes % 30 == 0 && dt.Second == 0 && dt.Millisecond == 0) 
                  ? 0 
                  : (30 - minutes % 30);
        var res = dt.AddMinutes(add);
        return new DateTime(res.Year, res.Month, res.Day, res.Hour, res.Minute / 30 * 30, 0);
    }

    static void UpdateJsonPhone(string phoneNumber)
    {
        // Update the JSON file with the formatted phone number
        try
        {
            if (File.Exists(jsonPath))
            {
                var json = JObject.Parse(File.ReadAllText(jsonPath));
                if (json["form"] != null)
                {
                    json["form"]["clientPhone"] = phoneNumber;
                    File.WriteAllText(jsonPath, json.ToString());
                }
            }
        }
        catch (Exception ex)
        {
            // Log or ignore, but don't block meeting creation
            Console.WriteLine("Warning: Failed to update data.json with formatted phone: " + ex.Message);
        }
    }
    
    // ---- Case details logic ----
    static string GetCaseDetails(string caseType, Dictionary<string, object> caseDetails)
    {
        if (string.IsNullOrEmpty(caseType) || caseDetails == null)
            return "";

        switch (caseType)
        {
            case "divorce":
                return "Divorce / Family Law<br><br>" +
                       "<strong>Spouse Name:</strong> " + (caseDetails.ContainsKey("spouseName") ? caseDetails["spouseName"] : "") + "<br>" +
                       "Conflict Search Done? " + (caseDetails.ContainsKey("conflictSearchDone") && caseDetails["conflictSearchDone"] is bool && (bool)caseDetails["conflictSearchDone"] ? "✔️" : "❌");
            case "estate":
                return "Successions / Estate Law<br><br>" +
                       "<strong>Deceased Name:</strong> " + (caseDetails.ContainsKey("deceasedName") ? caseDetails["deceasedName"] : "") + "<br>" +
                       "<strong>Executor Name:</strong> " + (caseDetails.ContainsKey("executorName") ? caseDetails["executorName"] : "") + "<br>" +
                       "Conflict Search Done? " + (caseDetails.ContainsKey("conflictSearchDone") && caseDetails["conflictSearchDone"] is bool && (bool)caseDetails["conflictSearchDone"] ? "✔️" : "❌");
            case "employment":
                return "Employment Law<br><br>" +
                       "<strong>Employer Name:</strong> " + (caseDetails.ContainsKey("employerName") ? caseDetails["employerName"] : "");
            case "contract":
                return "Contract Law<br><br>" +
                       "<strong>Other Party:</strong> " + (caseDetails.ContainsKey("otherPartyName") ? caseDetails["otherPartyName"] : "");
            case "defamations":
                return "Defamations";
            case "real_estate":
                return "Real Estate";
            case "name_change":
                return "Name Change";
            case "adoptions":
                return "Adoptions";
            case "mandates":
                return "Regimes de Protection / Mandates<br><br>" +
                       "<strong>Mandate Details:</strong> " + (caseDetails.ContainsKey("mandateDetails") ? caseDetails["mandateDetails"] : "");
            case "business":
                return "Business Law<br><br>" +
                       "<strong>Business Name:</strong> " + (caseDetails.ContainsKey("businessName") ? caseDetails["businessName"] : "");
            case "assermentation":
                return "Assermentation";
            case "common":
                return caseDetails.ContainsKey("commonField") ? caseDetails["commonField"].ToString() : "";
            default:
                return "";
        }
    }
}
