Attribute VB_Name = "modGoogleImages"
Public Const FORM_CAPTION = "Google Image Browser - "

Public iCurImageOffset As Integer
Public bStop As Boolean

Sub Main()
    'First we are going to load our form, and check to see
    'that we are currently connected to the internet
    Load frmMain
    Dim objTemp As Object
    
    bStop = False
    
    On Error Resume Next
    MkDir App.Path & "\pics"
    On Error GoTo 0
    
    'Disable all the controls while we check the connection
    For Each objTemp In frmMain.Controls
        On Error Resume Next     'Don't bitch if it's not
        objTemp.Enabled = False  'a control with this property;
        On Error GoTo 0          'ImageList for example
    Next
    
    'Display our form
    frmMain.Show
    DoEvents
    
    'Set the form caption status
    SetCaptionStatus "Checking Internet Connection"
    
    'Check for an internet connection
    On Error GoTo ErrorConnect
    frmMain.webSearch.Navigate "http://www.google.com/"
    Do Until frmMain.webSearch.ReadyState = READYSTATE_COMPLETE And frmMain.webSearch.Busy = False
        DoEvents
    Loop
    On Error GoTo 0
    
    'Set the minimum width and height combos to 200
    frmMain.cmbMinHeight.ListIndex = 3
    frmMain.cmbMinWidth.ListIndex = 3
    
    'Set the status to show that we ready to begin
    SetCaptionStatus "Ready"
    
    'Re-enable all the controls
    For Each objTemp In frmMain.Controls
        On Error Resume Next     'Don't bitch if it's not
        objTemp.Enabled = True   'a control with this property;
        On Error GoTo 0          'ImageList for example
    Next
    
    'Put the focus on the search terms field
    frmMain.txtSearchTerms.SetFocus
    
Exit Sub

ErrorConnect:
    MsgBox "There was an error connecting to Google!"
    Exit Sub
End Sub

'Function to stop this thing~!
Public Sub StopNow()
    bStop = True
End Sub

'This function changes our status
Public Sub SetCaptionStatus(sStatus As String)
    frmMain.Caption = FORM_CAPTION & sStatus
End Sub

'This function will search Google and start indexing the images
Public Sub SearchNow()
    'First we are going to make sure that a search criteria was entered
    If frmMain.txtDomain.Text = "" And frmMain.txtSearchTerms.Text = "" Then
        MsgBox "Please enter at least one search term, or a site to search!"
        GoTo CancelSearch
    End If
    
    'Set the image offset for the query
    iCurImageOffset = 0
    
    'Set status to searching
    SetCaptionStatus "Searching"
    StartSearch
    
    'We need to build the search url
    Dim sSearchURL As String
    sSearchURL = sBuildSearchURL
    
    frmMain.sbStatus.SimpleText = "Loading Images"
    
    'Grab the first page from google
    Dim sSearchSource As String
    frmMain.webSearch.Navigate sSearchURL
    Do Until frmMain.webSearch.ReadyState = READYSTATE_COMPLETE And frmMain.webSearch.Busy = False
        DoEvents
    Loop
    sSearchSource = frmMain.webSearch.Document.body.innerText
       
    'Get the number of images from this query,
    'Google will return a maximum of 1000, 20 per page
    Dim iNumImages As Integer
    iNumImages = iGetTotalResults(sSearchSource)
       
    'Set the status to show the number of found results
    frmMain.sbStatus.SimpleText = iNumImages & " Images Found (Max 1000, Limited by Google)"
    
    'Now we are going to cycle through the results pages and
    'save the images
    Dim iCurPage As Integer
    Dim iEndPage As Integer
    iCurPage = 0
    iEndPage = iNumImages Mod 20
    iEndPage = iEndPage + Int(iNumImages / 20)
    For iCurPage = 0 To iEndPage - 1
        If bStop = False Then GetImages iCurPage
    Next
    
    'Set status to ready
    SetCaptionStatus "Ready"
    EndSearch
    frmMain.sbStatus.SimpleText = "Ready"
    
CancelSearch:
End Sub

'This function gets the images on the current page and displays the thumbnails
Public Sub GetImages(iCurPage As Integer)
    If bStop = True Then Exit Sub
    Dim uiData() As Byte
    Dim sPageSource As String
    
    'Build the URL
    iCurImageOffset = 20 * iCurPage
    Dim sSearchURL As String
    sSearchURL = sBuildSearchURL
    
    frmMain.sbStatus.SimpleText = "Page: " & iCurPage
    
    'Grab the next page from google
    frmMain.webSearch.Navigate sSearchURL
    Do Until frmMain.webSearch.ReadyState = READYSTATE_COMPLETE And frmMain.webSearch.Busy = False
        DoEvents
    Loop
    
    'Grab the html from the page
    sPageSource = frmMain.webSearch.Document.body.innerHtml
    
    'Cycle through and load each image
    Dim iUrlStart As Integer
    Dim iUrlEnd As Integer
    Dim iWidthStart As Integer
    Dim iWidthEnd As Integer
    Dim iHeightStart As Integer
    Dim iHeightEnd As Integer
    ReDim uiData(500000)
    
NextImage:
    
    If bStop = True Then Exit Sub
    iUrlStart = InStr(iUrlStart + 1, sPageSource, "imgurl=") + 7
    iUrlEnd = InStr(iUrlStart, sPageSource, "&amp;imgrefurl=")
    iWidthStart = InStr(iUrlStart, sPageSource, "&amp;h=") + 7
    iWidthEnd = InStr(iUrlStart, sPageSource, "&amp;w=")
    iHeightStart = InStr(iUrlStart, sPageSource, "&amp;w=") + 7
    iHeightEnd = InStr(iUrlStart, sPageSource, "&amp;sz=")
    
    Dim iWidth As Integer
    Dim iHeight As Integer
    Dim sImageUrl As String
    
    iWidth = Int(Mid(sPageSource, iWidthStart, (iWidthEnd - iWidthStart)))
    iHeight = Int(Mid(sPageSource, iHeightStart, (iHeightEnd - iHeightStart)))
    sImageUrl = Mid(sPageSource, iUrlStart, (iUrlEnd - iUrlStart))
    
    'Check that the picture is big enough
    If iWidth < Int(frmMain.cmbMinHeight.Text) Or _
        iHeight < Int(frmMain.cmbMinWidth.Text) Then
        Exit Sub
    End If
    
    uiData() = ""
    sImageUrl = Replace(sImageUrl, "%2520", "%20")
    uiData() = frmMain.inetImages.OpenURL("http://" & sImageUrl, icByteArray)
    
    Do While frmMain.inetImages.StillExecuting = True
        DoEvents
    Loop
    
    Dim sImgData As String
    If UBound(uiData) > 0 Then
        sImgData = StrConv(uiData, vbUnicode)
    Else
        sImgData = " <"
    End If
    
    If InStr(UCase(sImgData), "</") = 0 Then
    
        Open App.Path & "\tmp" & Right(sImageUrl, 4) For Binary Access Write As #1
            Put #1, , uiData()
        Close #1
        
        frmMain.picDisp.Picture = LoadPicture(App.Path & "\tmp" & Right(sImageUrl, 4))
        
        Dim sPicName As String
        sPicName = Replace(Now, ":", "_")
        sPicName = Replace(sPicName, "/", "-")
        Open App.Path & "\pics\" & sPicName & Right(sImageUrl, 4) For Binary Access Write As #1
            Put #1, , uiData()
        Close #1
        
    End If
    
    sImgData = ""
    uiData = ""
    sImgData = ""
    If bStop = True Then Exit Sub
    If InStr(iUrlStart + 1, sPageSource, "imgurl=") > 0 Then GoTo NextImage
    
End Sub

'This function disables the search controls and sets the hourglass
Public Sub StartSearch()
    frmMain.txtDomain.Enabled = False
    frmMain.txtSearchTerms.Enabled = False
    frmMain.cmdSearch.Caption = "Stop!"
    Screen.MousePointer = vbHourglass
End Sub

'This function enables the search controls and sets the pointer
Public Sub EndSearch()
    frmMain.txtDomain.Enabled = True
    frmMain.txtSearchTerms.Enabled = True
    frmMain.cmdSearch.Caption = "Go!"
    Screen.MousePointer = vbNormal
End Sub

'This function will tell us the number of total results for a query
Public Function iGetTotalResults(sSearchSource As String) As Integer
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim sTmp As String
    Dim lCount As Long
    iStart = InStr(sSearchSource, " of about ") + 10
    iEnd = InStr(iStart, sSearchSource, "for")
    sTmp = Mid(sSearchSource, iStart, (iEnd - iStart))
    lCount = CLng(sTmp)
    If lCount > 1000 Then
        iGetTotalResults = 1000
    Else
        iGetTotalResults = Int(lCount)
    End If
End Function

'This function will generate the url for our search
Public Function sBuildSearchURL() As String
    Dim sNewURL As String
    sNewURL = "http://images.google.com/images?hl=en&lr=&ie=UTF-8&svnum=20&oe=UTF-8&safe=off&start=" & iCurImageOffset & "&sa=N&filter=0&q="
    If Len(frmMain.txtDomain.Text) > 0 Then
        sNewURL = sNewURL & "site:" & URLEncode(frmMain.txtDomain.Text) & "+"
    End If
    If Len(frmMain.txtSearchTerms.Text) > 0 Then
        sNewURL = sNewURL & URLEncode(frmMain.txtSearchTerms.Text)
    End If
    sBuildSearchURL = sNewURL
End Function

'This function will URLEncode a passed string
Public Function URLEncode(sString As String) As String
    Dim iChar As Integer
    Dim sNewString As String
    Dim iCurChar As Integer
    For iChar = 1 To Len(sString)
        iCurChar = Asc(Mid(sString, iChar, 1))
        If iCurChar < 36 Or iCurChar = 37 Or (iCurChar > 41 And _
            iCurChar < 45) Or (iCurChar > 58 And iCurChar < 61) Or _
            iCurChar = 62 Or iCurChar = 63 Or icurcar = 64 Or _
            (iCurChar > 90 And iCurChar < 95) Or iCurChar = 96 Or _
            iCurChar = 124 Or iCurChar > 126 Then
            sNewString = sNewString & "%" & Hex(iCurChar)
        Else
            sNewString = sNewString & Mid(sString, iChar, 1)
        End If
    Next
    sNewString = Replace(sNewString, "%20", "+")
    URLEncode = sNewString
End Function
