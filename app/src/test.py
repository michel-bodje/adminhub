import win32com.client as COM
import pythoncom
from time import sleep

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

def method_3_keyboard_shortcuts(appt, inspector):
    """Method 3: Keyboard shortcuts"""
    print("\n=== METHOD 3: Keyboard Shortcuts ===")
    
    try:
        import win32gui
        import win32con
        import win32api
        
        # Find Outlook window
        window_handle = win32gui.FindWindow("rctrl_renwnd32", None)
        if not window_handle:
            print("Could not find Outlook window")
            return False
            
        print("Found Outlook window, setting focus...")
        win32gui.SetForegroundWindow(window_handle)
        sleep(1)
        
        print("Executing Alt+H, Y, 1, Enter...")
        
        # Alt + H
        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
        win32api.keybd_event(ord('H'), 0, 0, 0)
        win32api.keybd_event(ord('H'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
        sleep(0.5)
        
        # Y
        win32api.keybd_event(ord('Y'), 0, 0, 0)
        win32api.keybd_event(ord('Y'), 0, win32con.KEYEVENTF_KEYUP, 0)
        sleep(0.3)
        
        # 1
        win32api.keybd_event(ord('1'), 0, 0, 0)
        win32api.keybd_event(ord('1'), 0, win32con.KEYEVENTF_KEYUP, 0)
        sleep(0.3)
        
        # Enter
        win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
        win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        print("Keyboard sequence executed")
        return True
        
    except Exception as e:
        print(f"Method 3 error: {e}")
        return False

def check_com_addins(inspector):
    """Get COM add-ins list from Outlook inspector"""
    if hasattr(inspector.Application, 'COMAddIns'):
        return inspector.Application.COMAddIns
    return None

def find_teams_addin(addins):
    """Find the Microsoft Teams COM add-in"""
    for i in range(1, addins.Count + 1):
        try:
            addon = addins.Item(i)
            description = getattr(addon, 'Description', '')
            prog_id = getattr(addon, 'ProgId', '')
            connected = getattr(addon, 'Connect', False)

            print(f"Add-in {i}: {description} ({prog_id}) - Connected: {connected}")

            if 'teams' in description.lower() or 'teams' in prog_id.lower():
                print(f"  FOUND TEAMS ADD-IN: {description}")
                if not connected:
                    try:
                        addon.Connect = True
                        print("  Connected Teams add-in")
                    except Exception as e:
                        print(f"  Failed to connect: {e}")
                return addon
        except Exception as e:
            print(f"  Error accessing add-in {i}: {e}")
    return None

def m5A_try_addin_methods(addon_object):
    """Method 5A: Try known methods directly on add-in object"""
    try:
        methods = [m for m in dir(addon_object) if not m.startswith('_')]
        print(f"  Available methods: {methods[:10]}...")
        common_methods = [
            'CreateTeamsMeeting', 'AddTeamsMeeting', 'AddOnlineMeeting',
            'CreateOnlineMeeting', 'EnableOnlineMeeting', 'SetOnlineMeeting',
            'OnConnection', 'OnBeginShutdown', 'AddMeeting'
        ]
        for method_name in common_methods:
            if hasattr(addon_object, method_name):
                print(f"  Found method: {method_name}")
                try:
                    method = getattr(addon_object, method_name)
                    if callable(method):
                        print(f"    Trying {method_name}()...")
                        result = method()
                        print(f"    SUCCESS: {method_name}() returned: {result}")
                        return True
                except Exception as e:
                    print(f"    Failed {method_name}(): {e}")
    except Exception as e:
        print(f"  Could not explore add-in methods: {e}")
    return False

def m5B_try_property_trigger(appt):
    """Method 5B: Trigger Teams add-in by modifying appointment"""
    try:
        original_subject = appt.Subject
        appt.Subject = original_subject + " [Teams]"
        appt.Save()
        print("  Modified appointment subject and saved")
        sleep(1)
        body = getattr(appt, 'Body', '')
        if 'teams.microsoft.com' in body.lower() or 'join microsoft teams meeting' in body.lower():
            print("  SUCCESS: Teams meeting link detected!")
            return True
    except Exception as e:
        print(f"  Property trigger failed: {e}")
    return False

def m5C_try_ribbon_trigger(inspector):
    """Method 5C: Trigger Teams add-in via ribbon command bars"""
    try:
        command_bars = inspector.CommandBars
        teams_commands = [
            'TeamsAddinRibbonCallback',
            'AddTeamsMeetingCallback', 
            'OnlineMeetingCallback',
            'TeamsAddin.AddMeeting'
        ]
        for command in teams_commands:
            try:
                command_bars.ExecuteMso(command)
                print(f"    SUCCESS: Executed {command}")
                return True
            except:
                pass
    except Exception as e:
        print(f"  Ribbon trigger failed: {e}")
    return False

def m5D_try_direct_com_instantiation(teams_addon, appt):
    """Method 5D: try COM.Dispatch on ProgId and attempt selected methods"""
    try:
        print("  Trying direct COM instantiation...")
        teams_com = COM.Dispatch(teams_addon.ProgId)
        if teams_com:
            print(f"  Created COM object: {type(teams_com)}")

            # Inspect object
            members = [m for m in dir(teams_com) if not m.startswith('_')]
            print(f"  COM object has {len(members)} members")
            for m in members:
                print(f"    {m}")

            # Try to get type info
            try:
                typeinfo = teams_com.GetTypeInfo()
                attr = typeinfo.GetTypeAttr()
                print(f"  Type kind: {attr.typekind}, Functions: {attr.cFuncs}")

                for i in range(attr.cFuncs):
                    funcdesc = typeinfo.GetFuncDesc(i)
                    names = typeinfo.GetNames(funcdesc.memid)
                    print(f"  Method: {names[0]}")
                    if len(names) > 1:
                        print(f"    Parameters: {names[1:]}")
                    print(f"    InvKind: {funcdesc.invkind}, Params: {funcdesc.cParams}")

            except pythoncom.com_error as ce:
                print(f"  Could not retrieve type info: {ce}")

            # List of methods you want to try — comment out ones you don’t want
            methods_to_try = [
                'Ribbon_ButtonClick',
            ]
            for method_name in methods_to_try:
                if hasattr(teams_com, method_name):
                    try:
                        method = getattr(teams_com, method_name)
                        print(f"    Trying {method_name}()...")
                        try:
                            result = method()  # will fail if params needed
                            print(f"    SUCCESS: {method_name}() returned: {result}")
                        except Exception as e:
                            print(f"    {method_name}() failed: {e}")
                    except Exception as e:
                        print(f"    Could not retrieve {method_name}: {e}")
                else:
                    print(f"    {method_name} not found on object")

    except Exception as e:
        print(f"  Direct COM instantiation failed: {e}")
    return False

def method_5_com_addins(appt, inspector):
    """Main controller for Method 5"""
    print("\n=== METHOD 5: COM Add-ins ===")
    try:
        addins = check_com_addins(inspector)
        if not addins:
            print("COMAddIns not available")
            return False
        print(f"Found {addins.Count} COM add-ins")
        teams_addon = find_teams_addin(addins)
        if not teams_addon:
            print("No Teams add-in found")
            return False

        print(f"\nWorking with Teams add-in: {teams_addon.ProgId}")
        try:
            addon_object = teams_addon.Object
            if addon_object:
                if m5A_try_addin_methods(addon_object):
                    return True
                if m5B_try_property_trigger(appt):
                    return True
                if m5C_try_ribbon_trigger(inspector):
                    return True
            else:
                print("  Could not get add-in object")
        except Exception as e:
            print(f"  Could not access add-in object: {e}")

        if m5D_try_direct_com_instantiation(teams_addon, appt):
            return True
        return False

    except Exception as e:
        print(f"Method 5 error: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("Teams Meeting Integration Test")
    print("=" * 40)
    
    # Create test appointment
    appt, inspector = create_test_appointment()
    if not appt or not inspector:
        print("Failed to create test appointment")
        return
    
    print("\nAppointment created. You can now test different methods...")
    print("The appointment window should be open. Check it after each method.")
    
    # Test methods - comment/uncomment as needed
    
    # Method 3: Keyboard shortcuts
    # method_3_keyboard_shortcuts(appt, inspector)
       
    # Method 5: COM Add-ins
    method_5_com_addins(appt, inspector)
    
    print("\nTest completed. Check the appointment window for Teams meeting link.")
    print("Press Enter to close...")
    input()
    
    # Clean up
    try:
        appt.Close(0)  # olDiscard
    except:
        pass

if __name__ == "__main__":
    main()