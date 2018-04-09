
Sub PhosphonetScraper()
    ProcessPhosphonetQuery
End Sub

Sub ProcessPhosphonetQuery()
    ' State
    Dim ie As Object
    Dim href_list() As String
    Dim site_name_list() As String
    Dim size As Integer
    Dim curr_row As Integer
    Dim query As String
    query = InputBox("Enter Protein Name", "Input", "")
    
    ' clear active worksheet
    Application.ActiveSheet.UsedRange.ClearContents

    ' write header information into sheet
    Range("A1").Value = "Protein: " & query
    Range("A3").Value = "Site"
    Range("B3").Value = "Type of GRK"
    Range("C3").Value = "Score"
    Range("D3").Value = "Rank"

    ' set the current row where we start writing out kinase records to row 4
    curr_row = 4
    
    ' create browser context
    Set ie = CreateObject("internetexplorer.application")
    
    ' load phosphonet.ca
    Application.StatusBar = "Navigating to Phosphonet.ca..."
    IE_NavigateAndWait ie, "http://www.phosphonet.ca/"
    
    ' enter our query into the search box and press search
    Application.StatusBar = "Performing search for '" & query & "'"
    ie.document.getElementById("tbSearch").Value = query
    ie.document.getElementById("btnSearch").Click
    IE_WaitReady ie
    
    ' store a list of all the hrefs we want to visit once we navigate away we wont have a reference anymore
    size = ie.document.getElementById("PhosphoSiteTable").getElementsByTagName("tbody")(0).rows.Length - 1 - 3
    ReDim href_list(size)
    ReDim site_name_list(size)
    
    'Setup the href_list'
    For i = 0 To UBound(href_list)
        site_name_list(i) = ie.document.getElementById("PhosphoSiteTable").getElementsByTagName("tbody")(0).rows(i + 3).getElementsByClassName("pSiteNameCol")(0).innerHtml
        href_list(i) = ie.document.getElementById("PhosphoSiteTable").getElementsByTagName("tbody")(0).rows(i + 3).Cells(27).getElementsByTagName("a")(0).href
    Next i
    
    ' for each Kinase Predictor page'
    For i = 0 To UBound(href_list)
        Application.StatusBar = "Processing " & query & " Kinase Predictor " & i & "/" & UBound(href_list) & " [" & href_list(i) & "]"
        
        ' navigate to the site's kinase info page
        IE_NavigateAndWait ie, href_list(i)
        
        ' for each kinase record
        For j = 1 To ie.document.getElementsByClassName("table-KinaseInfo")(0).getElementsByTagName("tbody")(0).rows.Length - 1
            
            ' get the GRK
            grk = ie.document.getElementsByClassName("table-KinaseInfo")(0).getElementsByTagName("tbody")(0).rows(j).Cells(1).getElementsByTagName("a")(0).innerHtml
            
            If InStr(1, grk, "BARK1") Or InStr(1, grk, "BARK2") Or InStr(1, grk, "GPRK4") Or InStr(1, grk, "GPRK5") Or InStr(1, grk, "GPRK6") Then
                
                ' transform "Protein Kinase Match" to a rank
                kinase_rank = ie.document.getElementsByClassName("table-KinaseInfo")(0).getElementsByTagName("tbody")(0).rows(j).Cells(0).innerHtml
                kinase_rank = Mid(kinase_rank, 8)
                kinase_rank = Left(kinase_rank, Len(kinase_rank) - 1)
                
                Dim siteName As String
                siteName = site_name_list(i)
                If InStr(1, siteName, "S") Or InStr(1, siteName, "T") Then
                    ' write out record
                    Application.ActiveSheet.Cells(curr_row, 1).Value = siteName
                    Application.ActiveSheet.Cells(curr_row, 2).Value = grk
                    Application.ActiveSheet.Cells(curr_row, 3).Value = ie.document.getElementsByClassName("table-KinaseInfo")(0).getElementsByTagName("tbody")(0).rows(j).Cells(3).innerHtml
                    Application.ActiveSheet.Cells(curr_row, 4).Value = kinase_rank
                End If
                
                ' next row!
                curr_row = curr_row + 1
                
            End If
        Next j
    Next i
    
    ' clean up ie context
    ie.Quit
    Set ie = Nothing
        
    ' done
    Application.StatusBar = "Complete"
    MsgBox "Report complete for search '" & query & "'"
    
End Sub

Sub IE_NavigateAndWait(ie As Object, url As String)
    ie.navigate url
    IE_WaitReady ie
End Sub

Sub IE_WaitReady(ie As Object)
    Do Until (ie.readyState = 4 And Not ie.Busy)
        DoEvents
    Loop
End Sub






