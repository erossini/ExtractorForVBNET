Imports System.Collections.Generic

''' <summary>
''' This class helps you to read a template file, extract and remove a particular tag and replace some content in the file.
''' 
''' Scenario
''' ========
''' You want to send an email from a template. In the template you have some fields to replace with a real data.
''' For example you want to replace the field {FirstName} with a real name
''' In the template you have a tag for the title of your email.
''' 
''' Template Example
''' ================
''' <title>Do you know {FirstName} {LastName}?</title>
''' <a href="mailto:enrico@westhill.co.uk?subject=Yes, i know {RefereeFirstName} {RefereeLastName}&body=Hi, my reference for {FirstName} is...">
''' Click here</a> if you know {FirstName} and send to us a reference.
''' 
''' </summary>
''' <remarks></remarks>
Public Class Extractor

    ''' <summary>
    ''' Reads a tag from the string strBody, return the tag content and the new content for strBody
    ''' </summary>
    ''' <param name="strBody">Full document to check</param>
    ''' <param name="strTag">Custom tag to replace</param>
    ''' <returns>Content tag and byref strBody</returns>
    ''' <remarks></remarks>
    Public Function ExtractTagAndRemove(ByRef strBody As String, strTag As String) As String
        Dim rtn As String = ""
        Dim regex As New Text.RegularExpressions.Regex("<\s*" & strTag & "[^>]*>(.*?)<\s*/\s*" & strTag & ">")

        For Each m As Text.RegularExpressions.Match In regex.Matches(strBody)
            rtn += m.Value.Trim
        Next
        strBody = regex.Replace(strBody, "<\s*" & strTag & "[^>]*>(.*?)<\s*/\s*" & strTag & ">", String.Empty)
        rtn = rtn.Replace("<" & strTag & ">", "").Replace("</" & strTag & ">", "")

        Return rtn
    End Function

    ''' <summary>
    ''' Replaces in a text each item in the dictionary with its value
    ''' </summary>
    ''' <param name="strText">Full document to check</param>
    ''' <param name="params">Param Dictionary for all tags</param>
    ''' <returns>A new string with replacing tags with values</returns>
    ''' <remarks></remarks>
    Public Function ReplaceContent(strText As String, params As Dictionary(Of String, String)) As String
        Dim rtn As String = strText

        For Each param In params
            rtn = rtn.Replace(param.Key, param.Value)
        Next

        Return rtn
    End Function

    ''' <summary>
    ''' Reads a template from a path
    ''' </summary>
    ''' <param name="pthTemplate">Full template path</param>
    ''' <returns>File content</returns>
    ''' <remarks></remarks>
    Public Function ReadTemplate(pthTemplate As String) As String
        Dim rtn As String = ""

        Dim oRead As System.IO.StreamReader
        oRead = IO.File.OpenText(pthTemplate)
        rtn = oRead.ReadToEnd()
        oRead.Close()

        Return rtn
    End Function

End Class