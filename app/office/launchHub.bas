Attribute VB_Name = "launchHub"
Option Explicit

Sub LaunchHub()
    Shell "python app/web/launch.py", "", vbNormalFocus
End Sub
