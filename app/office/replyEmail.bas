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
    Dim jsonPath As String
    Dim jsonText As String
    Dim json As Object
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    jsonPath = "\\AMNAS\amlex\Admin\Scripts\lawhub\app\data.json"
    jsonText = fso.OpenTextFile(jsonPath, 1).ReadAll
    Set json = JsonConverter.ParseJson(jsonText)

    ' === Extract data from JSON ===
    clientEmail = json("form")("clientEmail")
    clientLanguage = json("form")("clientLanguage")
    lawyerName = json("lawyer")("name")

    ' === Load the appropriate template ===
    If clientLanguage = "English" Then
        htmlBody = LoadTemplate("en/Reply.html")
    Else
        htmlBody = LoadTemplate("fr/Reply.html")
    End If

    ' === Replace placeholder ===
    htmlBody = Replace(htmlBody, "{{lawyerName}}", lawyerName)

    ' === Load default Outlook signature (optional) ===
    signaturePath = Environ("APPDATA") & "\Microsoft\Signatures\"
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
    templatePath = "\\AMNAS\amlex\Admin\Scripts\lawhub\templates\" & templateName
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(templatePath, 1)
    LoadTemplate = file.ReadAll
    file.Close
End Function
