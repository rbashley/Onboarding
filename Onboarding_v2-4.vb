Sub Copy_NewHireQuery()
    ' Unprotect the worksheet to make changes
    ActiveWorkbook.Worksheets("New Agents").Unprotect Password:="Secret123"
    ' Clear the specified range
    ActiveWorkbook.Worksheets("New Agents").Range("B:L").Clear
    ' Select and copy data
    Range("D1:N2001").SpecialCells(12).Select
    Selection.Copy
    ' Paste the copied data as values
    ActiveWorkbook.Worksheets("New Agents").Range("B1:L1").PasteSpecial Paste:=xlPasteValues
    ' Re-protect the worksheet
    ActiveWorkbook.Worksheets("New Agents").Protect Password:="Secret123"
    ' Activate another worksheet
    ActiveWorkbook.Worksheets("AddEmployee").Activate
End Sub

Sub Export_CSV()
    Application.DisplayAlerts = False
    Dim FilePath As String
    Dim User As String
    Dim CurrentSheet As Worksheet
    Dim FolderExist As String

    Set CurrentSheet = ActiveWorkbook.ActiveSheet
    User = Environ$("UserName")
    DirPath = "C:\users\" & User & "\Desktop\Onboarding\"
    FilePath = DirPath & CurrentSheet.Name & ".csv"
    FolderExist = Dir(DirPath, vbDirectory)

    ' Create directory if it doesn't exist
    If FolderExist = "" Then
        MkDir (DirPath)
    End If

    ' Create a temporary sheet
    ActiveWorkbook.Worksheets.Add.Name = "Temp_Sheet"
    CurrentSheet.Activate
    ActiveWorkbook.ActiveSheet.Range("A1:I1").Resize( _
        Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row, _
        Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column _
    ).Select
    Selection.Copy
    ActiveWorkbook.Worksheets("Temp_Sheet").Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveWorkbook.Worksheets("Temp_Sheet").Activate

    ' Save as CSV and delete the temporary sheet
    ActiveWorkbook.ActiveSheet.SaveAs Filename:=FilePath, FileFormat:=xlCSV, CreateBackup:=False
    CurrentSheet.Activate
    ActiveWorkbook.Worksheets("Temp_Sheet").Delete

    Application.DisplayAlerts = True
End Sub

Sub Refresh_Data()
    ActiveWorkbook.RefreshAll
End Sub

Sub Refresh_Ext()
    ' Collection to store user data
    Dim users As New Collection
    Dim Rng As Range, c As Range, rSel As Range
    Dim i As Integer

    Set Rng = Range("E2:E1001")
    Set rSel = Nothing

    ' Collect unique values
    For Each c In Rng
        If c.Value <> "" Then
            users.Add (c.Value)
        End If
    Next c

    ' Activate and process the worksheet
    ActiveWorkbook.Worksheets("New Hire Query").Activate
    ActiveWorkbook.ActiveSheet.ListObjects(1).AutoFilter.ShowAllData

    For Each u In users
        i = Range("G2:G1001").Cells.Find(What:=u).Row
        Cells(i, 6).Copy
        ActiveWorkbook.Worksheets("New Agents").Activate
        i = Range("E2:E1001").Cells.Find(What:=u).Row
        Cells(i, 4).Activate
        ActiveCell.PasteSpecial Paste:=xlPasteValues
    Next u
End Sub

Sub Jabber_Export()
    Dim Tmplts As New Collection
    Dim Hdrs As Range, slrws As Range
    Dim cx As Integer, cy As Integer, i As Integer
    Dim curValue As String, contains As Boolean

    Set Hdrs = Range("A2")
    cx = 1
    cy = 1
    i = 0

    ' Find "Template" header
    While i < 1
        cx = cx + 1
        If Cells(cy, cx) <> "" Then
            If Cells(cy, cx) = "Template" Then i = 1
        Else
            MsgBox "Template Header not found"
            i = 1
        End If
    Wend

    ' Collect template values
    i = 0
    cy = cy + 1
    While i < 1
        curValue = Cells(cy, cx).Value
        contains = False
        cy = cy + 1
        If curValue <> " " Then
            For j = 1 To Tmplts.Count
                If Tmplts(j) = curValue Then contains = True
            Next j
            If Not contains Then Tmplts.Add (curValue)
        Else
            i = 1
        End If
    Wend

    ' Process templates
    cx = 14
    Set slrws = Range("B1:N1")
    For Each T In Tmplts
        i = 0
        cy = 2
        While i < 1
            If Cells(cy, 15).Value <> " " Then
                If T = Cells(cy, 15).Value Then
                    Set slrws = Union(Range(Cells(cy, 2), Cells(cy, 14)), slrws)
                End If
            Else
                i = 1
            End If
            cy = cy + 1
        Wend

        ' Export CSV for each template
        slrws.Select
        slrws.Copy
        shtnm = Replace(T & "_JabberUpload", " ", "_")
        ActiveWorkbook.Sheets.Add.Name = shtnm
        ActiveWorkbook.Sheets(shtnm).Activate
        Range("A1").PasteSpecial xlPasteValues
        Export_CSV
        ActiveWorkbook.Sheets("JabberImport").Activate

        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets(shtnm).Delete
        Application.DisplayAlerts = True

        Set slrws = Range("B1:N1")
    Next T
End Sub