Attribute VB_Name = "launchHub"
Option Explicit

Sub LaunchHub()
    Shell "python \\AMNAS\amlex\Admin\Scripts\lawhub\app\web\launch.py", "", vbNormalFocus
End Sub
