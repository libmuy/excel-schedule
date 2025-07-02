

Public Function GetRedmineIssueProgress(issueId As String, repoId As Integer) As Double
    Dim redmineUrl As String
    Dim apiKey As String
    
    ' Get Redmine URL and API Key from the sheet
    GetRedmineRepo repoId, redmineUrl, apiKey

    If redmineUrl = "" Or apiKey = "" Then
        GetRedmineIssueProgress = -1 ' Indicate error
        Exit Function
    End If

    Dim xmlHttp As New MSXML2.XMLHTTP60
    Dim xmlDoc As New MSXML2.DOMDocument60
    Dim doneRatioNode As IXMLDOMNode
    Dim doneRatio As String
    
    Dim requestUrl As String
    requestUrl = redmineUrl & "issues/" & issueId & ".xml?key=" & apiKey
    
    xmlHttp.Open "GET", requestUrl, False
    xmlHttp.setRequestHeader "Content-Type", "text/xml"
    xmlHttp.send
    
    If xmlHttp.Status = 200 Then
        xmlDoc.LoadXML xmlHttp.responseText
        
        ' Select the <done_ratio> node inside <issue>
        Set doneRatioNode = xmlDoc.SelectSingleNode("//issue/done_ratio")
        If Not doneRatioNode Is Nothing Then
            doneRatio = doneRatioNode.Text
            GetRedmineIssueProgress = CDbl(doneRatio) / 100
        Else
            GetRedmineIssueProgress = -1 ' Not found
        End If
    Else
        GetRedmineIssueProgress = -1 ' Error
    End If
    
    Set xmlHttp = Nothing
    Set xmlDoc = Nothing
End Function



Sub GetRedmineRepo(id As Integer, ByRef url As String, ByRef apiKey As String)
    Dim r As Range

    Set r = Range("REDMINE_REPO")
    url = r.Offset(id, 1).Value
    apiKey = r.Offset(id, 2).Value
End Sub

