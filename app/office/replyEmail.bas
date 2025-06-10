Attribute VB_Name = "replyEmail"
Option Explicit

Sub CreateReplyEmail()
    Dim clientEmail As String
    Dim clientLanguage As String
    Dim lawyerName As String
    Dim mail As Outlook.MailItem
    Dim htmlBody As String
    Dim signaturePath As String
    Dim signatureHTML As String
    Dim fso As Object

    ' === Get values from UserForm ===
    With frmMain ' Replace with your actual form name
        clientEmail = .txtClientEmail.Value
        clientLanguage = .cmbClientLanguage.Value
        lawyerName = .cmbLawyerName.Value
    End With

    ' === Load the appropriate template ===
    If clientLanguage = "Français" Then
        htmlBody = LoadTemplate("fr/Reply.html")
    Else
        htmlBody = LoadTemplate("en/Reply.html")
    End If

    ' === Replace placeholder ===
    htmlBody = Replace(htmlBody, "{{lawyerName}}", lawyerName)

    ' === Load default Outlook signature (optional) ===
    signaturePath = Environ("APPDATA") & "\Microsoft\Signatures\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(signaturePath & "default.htm") Then
        signatureHTML = fso.OpenTextFile(signaturePath & "default.htm", 1).ReadAll
    Else
        signatureHTML = fso.OpenTextFile(signaturePath & "Michel Assi-Bodje (admin@amlex.ca).htm", 1).ReadAll
    End If

    ' === Create and display the email ===
    Set mail = Application.CreateItem(olMailItem)
    With mail
        .To = clientEmail
        .Subject = IIf(clientLanguage = "Français", "Réponse - Allen Madelin", "Reply - Allen Madelin")
        .HTMLBody = htmlBody & "<br><br>" & signatureHTML
        .Display
    End With
End Sub

Function LoadTemplate(templateName As String) As String
    Dim templatePath As String
    Dim fso As Object, file As Object
    templatePath = ThisOutlookSession.CurrentProjectPath & "\Shared\templates\" & templateName
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(templatePath, 1)
    LoadTemplate = file.ReadAll
    file.Close
End Function
