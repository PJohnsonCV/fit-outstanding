' Paul Johnson
' pjohnson.cv@gmail.com
' October, 2021

' A visual basic macro for Excel to read CSV files and remove lines for matched sample IDs. For best (intended) user experience, use a barcode scanner, but typing sample IDs also works.

' Ensure trailing slash is included e.g.: "\\network-share\subdir\"
Const defaultPath = "<YOUR PATH HERE>"
' File name exactly as exported i.e.: "fit.csv" and "fit" are two different files
Const defaultFITFile = "fit"
Const defaultTURFITFile = "turfit"

' Enables macro calls without GUI elements - just a cell change on Main
' This is called whenever a change happens on the worksheet, so it's important that macro work is encapsulated in an IF
Private Sub Worksheet_Change(ByVal Target As Range)
  ' B3 is the large yellow merged cell on Main and the only cell that we're interested in being changed to cause an action
  ' Activaing 3,2 (B3) at the end of this check makes sure the input field is always selected
  If Not Intersect(Target, Range("B3")) Is Nothing Then
    ' Not interested in deletes, only entered data
    If Trim(Cells(3, 2)) <> "" Then
      ' If the file's last modified dates haven't been output, gather data otherwise the rest of the macro can't function
      ' (no data = no joy)
      If Cells(3, 10) = "" And Cells(4, 10) = "" Then
        Call CheckData
      End If
    
      ' Thought a running counter and the last ID searched might be useful, has no impact on the function
      Cells(7, 5) = Cells(3, 2)           ' Last ID to search for, i.e. this value in B3
      Cells(8, 5) = Cells(8, 5).Value + 1 ' Incrementing counter
    
      ' Results of CompareData
      foundFit = CompareData("fit", Cells(3, 2))
      foundTurFit = CompareData("turfit", Cells(3, 2))
      ' Append this sample ID to a list of sample IDs not found on either list
      If foundFit = False Or foundTurFit = False Then
        ' Find the bottom of the list by going to an arbitarily low cell index and count backwards
        ' Put the current ID in the next row
        iRow = Range("B10000").End(xlUp).Row + 1
        Cells(iRow, 2) = Cells(3, 2)
        ' Identify whether its not in FIT or TURFIT by adding a X in the appropriate column
        ' Obviously this should be an X in both if not found in either
        If foundFit = False Then
          Cells(iRow, 5) = "X"
        End If
        If foundTurFit = False Then
          Cells(iRow, 6) = "X"
        End If
        ' Not sure why this is here as it'll overwrite everything below A12 with A12?
        ' Commented it out for github, but this exists in production. Will have to double check!
        'If iRow > 12 Then
        '  Range("A12:A12").AutoFill Destination:=Range("A12:A" & iRow)
        'End If
      End If
    End If
  End If
  ' If enter is pressed (or newline from a barcode read) make sure that the selected cell doesn't default to the next row,
  ' select the 'input' cell instead
  Cells(3, 2).Activate
End Sub

' Returns the date the file was last modified, using the built-in FSO
Function FileLMDate(fName)
  Dim fs, f, s
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.GetFile(fName)
  FileLMDate = f.DateLastModified
End Function

' Check for the LIMS generated files and perform functions if available
Sub CheckData()
  ' These were hard coded before converting to consts for public release
  fit = defaultPath & defaultFITFile
  turfit = defaultPath & defaultTURFITFile
  
  ' Don't alert users to the deletion of sheets OR the inability to delete
  ' Less user input, the better, and prevents crashing the macro if sheet not present to delete
  On Error Resume Next
  Application.DisplayAlerts = False
  Sheets("fit").Delete
  Sheets("turfit").Delete
  Application.DisplayAlerts = True
  
  ' Get the date each file was last modified
  ' (would throw an exception if the file was not there)
  lastFit = Sheet1.FileLMDate(fit)
  lastTurfit = Sheet1.FileLMDate(turfit)
  ' Absolutely, 100% DEFINITELY restore the alert system before anything else!
  On Error GoTo 0
  
  ' Output the last modified dates
  Cells(3, 10) = lastFit
  Cells(4, 10) = lastTurfit
  
  ' Process data, if found
  If lastFit <> "" Then
    Call GetData(fit, "fit")
    TidyData ("fit")
  Else
    Sheets.Add ("fit")
  End If
  If lastTurfit <> "" Then
    Call GetData(turfit, "turfit")
    TidyData ("turfit")
  Else
    Sheets.Add ("turfit")
  End If
  
  ' Reset the user's view to Main
  Sheets("Main").Select
End Sub

' Literally copy the file as a new worksheet (Excel behaviour usually means the worksheet will have the same name as the file but safer to hardcode a name)
' fName should reflect the concatenated path+file consts, sName can be whatever you like
Sub GetData(fName, sName)
  wbMain = ActiveWorkbook.Name
    
  Workbooks.OpenText Filename:=fName, Origin:= _
    xlMSDOS, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote _
    , ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False _
    , Space:=True, Other:=False
  Sheets(sName).Select
  Sheets(sName).Move After:=Workbooks(wbMain).Sheets(Workbooks(wbMain).Sheets.Count)
End Sub

' Remove some output we don't need (header and footer rows)
' The output used to be different between file types, it was tidied up within the LIMS hence the commented out lines that remain from original development.
' Magic numbers, go!
Sub TidyData(sName)
  iHead = 7
  'If sName = "fit" Then
    ' A10000 is a row that will never be filled by the LIMS output, I always go large and work back to the last filled cell
    lastRow = Sheets(sName).Range("A10000").End(xlUp).Row
    ' -3 here is a known number of rows with footer data that can be deleted
    Sheets(sName).Rows(lastRow - 3 & ":" & lastRow).Delete Shift:=xlUp
  '  iHead = 9
  'End If
  
  Sheets(sName).Rows("1:" & iHead).Delete Shift:=xlUp
  ' Make the data user redable
  Sheets(sName).Columns("A:K").Select
  Sheets(sName).Columns("A:K").EntireColumn.AutoFit
End Sub

' Point the macro at a sheet and give it a sample number to find, it will return True (found) or False (not found)
' We can "hide" the found samples, but choose not to delete them just incase. So they get coloured backgrounds.
Function CompareData(sName, sampFind)
  cd = False
  iCol = 2
  'If sName = "fit" Then
    iCol = 1
  'End If
  
  ' Slow search: row by row, manually - I now know using Excel's search is an option in VBA
  ' Doesn't exit the For on match, could add an exit to speed up the search, but the performance is negligible for the number of samples on a list
  lastRow = Sheets(sName).Range("A10000").End(xlUp).Row
  For iRow = 1 To lastRow
    ' Cleanse the current cell - IDs are formatted X.YY,1234567.Z
    ' We only want the 7-digit number
    sampNum = Sheets(sName).Cells(iRow, iCol)
    sampNum = Trim(Mid(sampNum, 6, 7))
    ' Ensure the ID searched for is also free from invisible spaces
    sFind = Trim(Str(sampFind))
    ' When matched, colour the cells to mask the value without destroying the list
    If sampNum = sFind Then
      Sheets(sName).Range("A" & iRow & ":K" & iRow).Interior.Color = 16
      cd = True
    End If
  Next
  CompareData = cd
End Function