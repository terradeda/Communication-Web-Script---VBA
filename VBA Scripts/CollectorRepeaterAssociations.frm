VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CollectorRepeaterAssociations 
   Caption         =   "Collector <-> Repeater Association Plotter"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   OleObjectBlob   =   "CollectorRepeaterAssociations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CollectorRepeaterAssociations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************
'      Global Variables
'****************************

    'Used to temporarily store CSV File Locations
    Dim filepath As String

    'Worksheet Variables
    Dim mainWS As Worksheet
    Dim ColWS As Worksheet
    Dim RepWS As Worksheet

    
    
'
'
'****************************
'      USERFORM CODE
'****************************

Private Sub CollectorBrowseButton_Click()
   filepath = Application.GetOpenFilename(FileFilter:="CSV Files (*.CSV), *.CSV", Title:="Select Collector File")
   
   If VarType(filepath) = vbString And filepath <> "False" Then
        ColTextBox.Value = filepath
   End If

End Sub

Private Sub RepeaterBrowseButton_Click()
    filepath = Application.GetOpenFilename(FileFilter:="CSV Files (*.CSV), *.CSV", Title:="Select Repeator File")
    
    If VarType(filepath) = vbString And filepath <> "False" Then
        RepTextBox.Value = filepath
    End If
    
End Sub

Private Sub AssociationListBrowseButton_Click()
   filepath = Application.GetOpenFilename(FileFilter:="CSV Files (*.CSV), *.CSV", Title:="Select Association List")
   
   If VarType(filepath) = vbString And filepath <> "False" Then
        AssocTextBox.Value = filepath
   End If
   
End Sub

Private Sub UserForm_initialize()

    SheetFrame.Visible = False
    ColTextBox.Enabled = False
    RepTextBox.Enabled = False
    AssocTextBox.Enabled = False

End Sub

Private Sub newDataCheckBox_Click()

    If newDataCheckBox.Value = True Then
        SheetFrame.Visible = True
        ColTextBox.Value = Null
        RepTextBox.Value = Null
        AssocTextBox.Value = Null
        
    Else
        SheetFrame.Visible = False
    End If

End Sub

Private Sub CancelButton_Click()
    End
End Sub

Private Sub RunButton_Click()
    
    'Used To Create New Sheets during data import
    Dim SheetName As String
    
    'Used to check if the worksheets already exist (When not importing)
    Dim wsCheck As Worksheet
    
    'These variables are used to check if the sheets contained in this workbook are formatted correctly.
    Dim numColRep As Integer
    Dim numColCol As Integer
    Dim numColAssoc As Integer
    
    Dim colHeaders() As Variant
    Dim repHeaders() As Variant
    Dim assocHeaders() As Variant
    
    Dim wrongHeaders As Boolean
    
    colHeaders = Array("CollectorID", "SecondaryID", "Latitude", "Longitude", "Repeaters_DailyActual", "Repeaters_DailyManaged", "Endpoints_DailyActual", "Endpoints_DailyManaged", "AvgNumEndpointsHurd", "Date")
    repHeaders = Array("ItronRepeaterID", "RepeaterId", "Latitude", "Longitude", "Active", "DailyActual", "DailyManaged", "NumTSErrEP", "RefDateTime")
    assocHeaders = Array("ITronCollectorId", "ITronRepeaterId", "DailyMaxRSSI", "DailyAvgRSSI", "ReadCoeffBitmap", "NumMessages", "Rank", "ReportList", "ManagementList", "recordDateTime")
       
    'Worksheet name variables
    Dim mainWSName As String
    Dim colWSName As String
    Dim RepWSName As String
    
    'These variables contain the names of each of the three sheets
    mainWSName = "Col-Rep Assoc"
    colWSName = "Collectors"
    RepWSName = "Repeaters"
       
       
       
       
    If newDataCheckBox.Value = True Then
        'Check to make sure all the inputs have been filled in
        If RepTextBox.Value = "" Or ColTextBox.Value = "" Or AssocTextBox.Value = "" Then
            
            MsgBox "One or more of the fields above are not filled in"
            Exit Sub
        End If
        
        '***************************************
        'Create New WorkSheets and import Data
        '***************************************
        
        
        'Import Collector Data
        SheetName = colWSName
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Sheets(SheetName).Delete
    
        Application.DisplayAlerts = True
        On Error GoTo 0
        Worksheets.Add.name = SheetName
        Set ColWS = ActiveWorkbook.Sheets(SheetName)
        
        Call ImportWorksheet(SheetName, ColTextBox.Value)
        
        'Import Repeater Data
        SheetName = RepWSName
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Sheets(SheetName).Delete
    
        Application.DisplayAlerts = True
        On Error GoTo 0
        Worksheets.Add.name = SheetName
        Set RepWS = ActiveWorkbook.Sheets(SheetName)
        
        Call ImportWorksheet(SheetName, RepTextBox.Value)
        
        'Import Association Data
        SheetName = mainWSName
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Sheets(SheetName).Delete
    
        Application.DisplayAlerts = True
        On Error GoTo 0
        Worksheets.Add.name = SheetName
        Set mainWS = ActiveWorkbook.Sheets(SheetName)
        
        Call ImportWorksheet(SheetName, AssocTextBox.Value)
        
    Else
        'Check to make sure the correct are in the worksheet
        
        'Check if 'Collectors' Sheet Exists
        
        If WorksheetExists(colWSName) = False Then
            MsgBox "There is no 'Collectors' sheet contained within this workbook!" & vbCrLf & "Import New Data"
            newDataCheckBox.Value = True
            Exit Sub
        ElseIf WorksheetExists(RepWSName) = False Then
            MsgBox "There is no 'Repeaters' sheet contained within this workbook!" & vbCrLf & "Import New Data"
            newDataCheckBox.Value = True
            Exit Sub
        ElseIf WorksheetExists(mainWSName) = False Then
            MsgBox "There is no 'Association' sheet contained within this workbook!" & vbCrLf & "Import New Data"
            newDataCheckBox.Value = True
            Exit Sub
        End If
                   
        'Initialize the worksheet variables
        Set ColWS = ActiveWorkbook.Sheets(colWSName)
        Set RepWS = ActiveWorkbook.Sheets(RepWSName)
        Set mainWS = ActiveWorkbook.Sheets(mainWSName)
          
    End If
     
    '************************************
    '          SHEET VALIDATION
    '************************************
    'Check Sheets to make sure they are the Un-Altered Results of the SQL Queries)
     
    numColCol = ColWS.Cells(1, Columns.Count).End(xlToLeft).Column
    numColRep = RepWS.Cells(1, Columns.Count).End(xlToLeft).Column
    numColAssoc = mainWS.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Make sure the existing sheets have the correct number of column headers
    If numColCol <> 10 Then
        MsgBox "The 'Collectors' sheet does not have the right number of columns!" & vbCrLf & "Import New Data"
        newDataCheckBox.Value = True
        Exit Sub
    ElseIf numColRep <> 9 Then
        MsgBox "The 'Repeaters' sheet does not have the right number of columns!" & vbCrLf & "Import New Data"
        newDataCheckBox.Value = True
        Exit Sub
    ElseIf numColAssoc <> 10 Then
        MsgBox "The 'Association' sheet does not have the right number of columns!" & vbCrLf & "Import New Data"
        newDataCheckBox.Value = True
        Exit Sub
    End If
     
    'Check the headers of each sheet to check for a match
    wrongHeaders = False
    For i = 1 To numColCol
        If ColWS.Cells(1, i).Value <> colHeaders(i - 1) Then
            wrongHeaders = True
        End If
    Next i
    
    If wrongHeaders = True Then
        MsgBox "The 'Collectors' sheet column headers are wrong!" & vbCrLf & "Import New Data"
        newDataCheckBox.Value = True
        Exit Sub
    End If
    
    wrongHeaders = False
    For i = 1 To numColRep
        If RepWS.Cells(1, i).Value <> repHeaders(i - 1) Then
            wrongHeaders = True
        End If
    Next i
    
    If wrongHeaders = True Then
        MsgBox "The 'Repeaters' sheet column headers are wrong!" & vbCrLf & "Import New Data"
        newDataCheckBox.Value = True
        Exit Sub
    End If
    
    wrongHeaders = False
    For i = 1 To numColAssoc
        If mainWS.Cells(1, i).Value <> assocHeaders(i - 1) Then
            wrongHeaders = True
        End If
    Next i
    
    If wrongHeaders = True Then
        MsgBox "The 'Associations' sheet column headers are wrong!" & vbCrLf & "Import New Data"
        newDataCheckBox.Value = True
        Exit Sub
    End If
    
    '***************************************
    '      CALL EXPORT KML SUBROUTINE
    '***************************************
    CollectorRepeaterAssociations.Hide

    'Call GenerateKML subroutine from the ExportKML Module
    Call ExportKML.generateKML

    End
     
End Sub
Public Function WorksheetExists(ByVal WorksheetName As String) As Boolean

On Error Resume Next
WorksheetExists = (Sheets(WorksheetName).name <> "")
On Error GoTo 0

End Function

Sub ImportWorksheet(mainWS As String, copyWS As String)
' This macro will import a file into this workbook
'Taken From 'http://www.zerrtech.com/content/excel-vba-open-csv-file-and-import'
    
    Sheets(mainWS).Select
    
    With ThisWorkbook.Worksheets(mainWS).QueryTables.Add(Connection:="TEXT;" & copyWS, Destination:=ThisWorkbook.Worksheets(mainWS).Range("A1"))
        .name = Filename
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .Refresh BackgroundQuery:=False
    End With
    

End Sub


