Public Function GetRedmineIssueProgress(issueId As String, repoId As Integer) As Double
    Dim redmineUrl As String
    Dim apiKey As String
    
    ' Get Redmine URL and API Key from the sheet
    GetRedmineRepo repoId, redmineUrl, apiKey

    If redmineUrl = "" Or apiKey = "" Then
        GetRedmineIssueProgress = -1 ' Indicate error
        Exit Function
    End If

    Dim xmlDoc As MSXML2.DOMDocument60
    Dim doneRatioNode As IXMLDOMNode
    Dim doneRatio As String

    If Not FetchRedmineIssueXml(issueId, repoId, xmlDoc) Then
        GetRedmineIssueProgress = -1
        Exit Function
    End If

    Set doneRatioNode = xmlDoc.SelectSingleNode("//issue/done_ratio")
    If Not doneRatioNode Is Nothing Then
        doneRatio = doneRatioNode.Text
        GetRedmineIssueProgress = CDbl(doneRatio) / 100
    Else
        GetRedmineIssueProgress = -1
    End If

    Set xmlDoc = Nothing
End Function

' Common function to fetch Redmine issue XML document
Private Function FetchRedmineIssueXml(issueId As String, repoId As Integer, ByRef xmlDoc As MSXML2.DOMDocument60) As Boolean
    Dim redmineUrl As String
    Dim apiKey As String
    Dim xmlHttp As New MSXML2.XMLHTTP60
    Dim requestUrl As String

    FetchRedmineIssueXml = False
    GetRedmineRepo repoId, redmineUrl, apiKey

    If redmineUrl = "" Or apiKey = "" Then Exit Function

    requestUrl = redmineUrl & "issues/" & issueId & ".xml?key=" & apiKey

    On Error GoTo RedmineError
    xmlHttp.Open "GET", requestUrl, False
    xmlHttp.setRequestHeader "Content-Type", "text/xml"
    xmlHttp.send

    If xmlHttp.Status = 200 Then
        Set xmlDoc = New MSXML2.DOMDocument60
        xmlDoc.LoadXML xmlHttp.responseText
        FetchRedmineIssueXml = True
    End If

    Set xmlHttp = Nothing
    Exit Function

RedmineError:
    Set xmlHttp = Nothing
    Set xmlDoc = Nothing
End Function

' Get start and end date from Redmine issue
Public Function GetRedmineIssueStartEndDate(issueId As String, repoId As Integer, ByRef startDate As Date, ByRef endDate As Date) As Boolean
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim startNode As IXMLDOMNode
    Dim endNode As IXMLDOMNode

    GetRedmineIssueStartEndDate = False
    If Not FetchRedmineIssueXml(issueId, repoId, xmlDoc) Then Exit Function

    Set startNode = xmlDoc.SelectSingleNode("//issue/start_date")
    Set endNode = xmlDoc.SelectSingleNode("//issue/due_date")

    If Not startNode Is Nothing Then
        startDate = CDate(startNode.Text)
    Else
        startDate = 0
    End If

    If Not endNode Is Nothing Then
        endDate = CDate(endNode.Text)
    Else
        endDate = 0
    End If

    GetRedmineIssueStartEndDate = Not (startDate = 0 And endDate = 0)
    Set xmlDoc = Nothing
End Function



Sub GetRedmineRepo(id As Integer, ByRef url As String, ByRef apiKey As String)
    Dim r As Range

    Set r = Range("REDMINE_REPO")
    url = r.Offset(id, 1).Value
    apiKey = r.Offset(id, 2).Value
End Sub

