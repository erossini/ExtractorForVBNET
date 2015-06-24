# ExtractorForVBNET

This class helps you to read a template file, extract and remove a particular tag and replace some content in the file.

Scenario
========
You want to send an email from a template. In the template you have some fields to replace with a real data.
For example you want to replace the field {FirstName} with a real name
In the template you have a tag for the title of your email.

Template Example
================
```
<title>Do you know {FirstName} {LastName}?</title>
<a href="mailto:enrico@westhill.co.uk?subject=Yes, i know {RefereeFirstName} {RefereeLastName}&body=Hi, my reference for {FirstName} is...">
Click here</a> if you know {FirstName} and send to us a reference...
```

Example code
===
```
    Private Sub SendRegistrationEmail(ID As String, strEmail As String, strFirstName As String, strLastName As String)
        Dim ext As New Extractor
        Dim strBody As String = ext.ReadTemplate(Server.MapPath("~/Registration/Template/RegistrationEmail.html"))

        Dim params As New Dictionary(Of String, String)
        params.Add("{FirstName}", strFirstName)
        params.Add("{LastName}", strFirstName)
        strBody = ext.ReplaceContent(strBody, params)

        Dim strSubject As String = ext.ExtractTagAndRemove(strBody, "title")

        SendEmail(strEmail, strSubject, strBody, ID)
    End Sub
```
