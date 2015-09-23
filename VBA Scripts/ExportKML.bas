Attribute VB_Name = "ExportKML"
Sub generateKML()
'
' Create KML File for Repeater/Collector Associations
' Written by David Terrade - Aug 21st 2015
'


'******************************************************
'               KML FILE - OBJECT SYNTAX
'******************************************************

    '-------------------------
    ' CREATE KML FILE SYNTAX
    '-------------------------

    Dim myFile As String

    'Set output file details
    myFile = Application.ActiveWorkbook.Path + "\Col-Rep Associations.kml"
       
    '-------------------------
    '  HEADER/FOOTER SYNTAX
    '-------------------------
    
        Dim docName As String
        Dim Header As String
        Dim Footer As String
    
        'Define Document Name
        docName = "My KML Import"
    
        'Create KML File Header String
        Header = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf & _
                 "<kml xmlns=" & Chr(34) & "http://earth.google.com/kml/2.0" & Chr(34) & ">" & vbCrLf & _
                 "<Document>" & vbCrLf & _
                 "  <name>" & docName & "</name>"
                 
                          
                 
        'Create KML File Footer String
        Footer = "</Document>" & vbCrLf & "</kml>"
        
        
      '-------------------------
      '  INITIAL ZOOM LOCATION
      '-------------------------
        Dim zoomLoc As String
        
        zoomLoc = "<LookAt>" & vbCrLf & _
                  "         <longitude>" & "-75.682778" & "</longitude>" & vbCrLf & _
                  "         <latitude>" & "45.364444" & "</latitude>" & vbCrLf & _
                  "         <range>50000</range>" & vbCrLf & _
                  "</LookAt>"
                  

    '-------------------------
    '    NEW FOLDER SYNTAX
    '-------------------------
        Dim folderName As String
        
        Dim Folder1 As String
        Dim Folder2 As String
        Dim Folder3 As String
    
        Folder1 = "<Folder>" & vbCrLf & "    <name>"
        Folder2 = "</name>"
        Folder3 = "</Folder>"
        
       
    '-------------------------
    '   LINE STYLES SYNTAX
    '-------------------------
        
        Dim hurdLineStyle As String
        Dim managedLineStyle As String
        
        Dim styleColor As String
        Dim styleWidth As String
        Dim styleOpacity As String
        
        styleColor = "64781E00"
        styleWidth = "5"
        
        hurdLineStyle = "<Style id =" & Chr(34) & "hurdLineStyle" & Chr(34) & ">" & vbCrLf & _
                 "  <LineStyle>" & vbCrLf & _
                 "    <color>" & styleColor & "</color>" & vbCrLf & _
                 "    <width>" & styleWidth & "</width>" & vbCrLf & _
                 "  </LineStyle>" & vbCrLf & _
                 "</Style>"
                 
        styleColor = "FF14F00A"
        styleWidth = "5"
        
        managedLineStyle = "<Style id =" & Chr(34) & "managedLineStyle" & Chr(34) & ">" & vbCrLf & _
                 "  <LineStyle>" & vbCrLf & _
                 "    <color>" & styleColor & "</color>" & vbCrLf & _
                 "    <width>" & styleWidth & "</width>" & vbCrLf & _
                 "  </LineStyle>" & vbCrLf & _
                 "</Style>"

    '-------------------------
    '   POINT STYLES SYNTAX
    '-------------------------
        Dim repstyle As String
        Dim colstyle As String
        
        repstyle = "<Style id =" & Chr(34) & "RepStyle" & Chr(34) & ">" & vbCrLf & _
                 " <IconStyle>" & vbCrLf & _
                 "    <scale>0.5</scale>" & vbCrLf & _
                 "    <Icon> <href>Rep.png</href> </Icon>" & vbCrLf & _
                 "  </IconStyle>" & vbCrLf & _
                 "  <LabelStyle>" & vbCrLf & _
                 "    <scale>0</scale>" & vbCrLf & _
                 "  </LabelStyle>" & vbCrLf & _
                 "</Style>"
                 
        colstyle = "<Style id =" & Chr(34) & "ColStyle" & Chr(34) & ">" & vbCrLf & _
                 " <IconStyle>" & vbCrLf & _
                 "    <scale>2</scale>" & vbCrLf & _
                 "    <Icon> <href>Col.png</href> </Icon>" & vbCrLf & _
                 "  </IconStyle>" & vbCrLf & _
                 "</Style>"

    '-------------------------
    ' PLACEMARK OBJECT SYNTAX
    '-------------------------
    ' USE AS FOLLOWS:
    '       - "Placemark1 & pmName & Placemark2 & longitudeValue & ", " & latitudeValue & Placemark3 & pmDescription & Placemark4"
        
        Dim Placemark1 As String
        Dim Placemark2 As String
        Dim Placemark3 As String
        Dim Placemark4 As String
        Dim Placemark5 As String
        
        'Create KML File Placemark1 String
        Placemark1 = "  <Placemark>" & vbCrLf & "    <name>"
        
        'Create KML File Placemark2 String
        Placemark2 = "</name>" & vbCrLf & "    <styleUrl>#"
        
        Placemark3 = "</styleUrl>" & vbCrLf & "    <Point>" & vbCrLf & "     <coordinates>"
                 
        'Create KML File Placemark3 String
        Placemark4 = ",0</coordinates>" & vbCrLf & "    </Point>" & vbCrLf & "    <description><![CDATA["
        
        'Create KML File Placemark3 String
        Placemark5 = "]]></description>" & vbCrLf & "  </Placemark>"
    
    '-------------------------
    '    DRAW LINE SYNTAX
    '-------------------------
    
    Dim lat1 As Double
    Dim long1 As Double
    Dim lat2 As Double
    Dim long2 As Double
    
    Dim Color As String
    
    Dim line1 As String
    Dim line2 As String
    Dim line3 As String
    Dim line4 As String
    Dim Line5 As String
    
    line1 = "  <Placemark>" & vbCrLf & "    <name>"
    line2 = "</name>" & vbCrLf & "    <styleUrl>#"
    line3 = "</styleUrl>" & vbCrLf & "    <LineString> " & vbCrLf & "      <coordinates> "
    line4 = "</coordinates> " & vbCrLf & "    </LineString>" & vbCrLf & "    <description><![CDATA["
    Line5 = "]]></description>" & vbCrLf & "  </Placemark>"
    
    '-------------------------


    '******************************************************
    '                     MAIN CODE
    '******************************************************
    
    '---------------------------------
    '   DECLARE/INITIALIZE VARIABLES
    '---------------------------------
      'Worksheet variables
        Dim CollectorWS As Worksheet
        Dim RepeaterWS As Worksheet
        Dim mainWS As Worksheet
        
        
        Set CollectorWS = ThisWorkbook.Sheets("Collectors")
        Set RepeaterWS = ThisWorkbook.Sheets("Repeaters")
        Set mainWS = ThisWorkbook.Sheets("Col-Rep Assoc")
    
      'numRow Variables
        Dim numRowsC As Long
        Dim numRowsR As Long
        Dim numRowsM As Long

        numRowsC = CollectorWS.UsedRange.Rows.Count
        numRowsR = RepeaterWS.UsedRange.Rows.Count
        numRowsM = mainWS.UsedRange.Rows.Count
        
      'Placemark Attribute Variables
      
        Dim pmName As String
        Dim longitudeValue As String
        Dim latitudeValue As String
        Dim pmDescription As String
        
       'Line Attribute Variables
       
       
       Dim ColID As String
       Dim ColRow As Long
       Dim ColLat As String
       Dim ColLong As String
       
       Dim RepId As String
       Dim RepRow As Long
       Dim RepLat As String
       Dim RepLong As String
       
    

    '---------------------------------
    '    WRITE HEADERS AND STYLES
    '---------------------------------
             
    'Open Output File
    Open myFile For Output As #1
    
    'Write header to file
    outputText = Header
    Print #1, outputText
    
    outputText = zoomLoc
    Print #1, outputText
    
    outputText = hurdLineStyle
    Print #1, outputText
    
    outputText = managedLineStyle
    Print #1, outputText
    
    outputText = repstyle
    Print #1, outputText
    
    outputText = colstyle
    Print #1, outputText
    
   
    
    '---------------------------------
    '    WRITE COLLECTOR PLACEMARKS
    '---------------------------------
    
     'Start A folder
    folderName = "Collectors"
    outputText = Folder1 & folderName & Folder2
    Print #1, outputText

    For i = 2 To numRowsC

        'Define Placemark Attributes
        pmName = "Col ID: " & CollectorWS.Cells(i, 1).Value
        longitudeValue = CollectorWS.Cells(i, 4).Value
        latitudeValue = CollectorWS.Cells(i, 3).Value
        pmDescription = "<b>Repeater Stats</b>" & vbCrLf & _
                        "<br/>     Daily Actuals: " & CollectorWS.Cells(i, 5).Value & vbCrLf & _
                        "<br/>     Daily Managed: " & CollectorWS.Cells(i, 6).Value & vbCrLf & _
                        "<br/><b>Endpoint Stats</b>" & vbCrLf & _
                        "<br/>     Daily Actuals: " & CollectorWS.Cells(i, 7).Value & vbCrLf & _
                        "<br/>    Daily Managed: " & CollectorWS.Cells(i, 8).Value & vbCrLf & _
                        "<br/>   Average Managed: " & CollectorWS.Cells(i, 8).Value

        'Create KML code for a collector Placemark
        outputText = Placemark1 & pmName & Placemark2 & "ColStyle" & Placemark3 & longitudeValue & ", " & latitudeValue & Placemark4 & pmDescription & Placemark5

        'Print Placemark(s)
        Print #1, outputText


    Next i

    'End A folder
    outputText = Folder3
    Print #1, outputText


    '---------------------------------
    '    WRITE REPEATER PLACEMARKS
    '---------------------------------

    'Start A folder
    folderName = "Repeaters"
    outputText = Folder1 & folderName & Folder2
    Print #1, outputText

    For i = 2 To numRowsR

        'Define Placemark Attributes
        pmName = "Rep ID: " & RepeaterWS.Cells(i, 1).Value
        longitudeValue = RepeaterWS.Cells(i, 4).Value
        latitudeValue = RepeaterWS.Cells(i, 3).Value
        pmDescription = "Active: " & RepeaterWS.Cells(i, 5).Value & vbCrLf & _
                        "Daily Actual: " & RepeaterWS.Cells(i, 6).Value & vbCrLf & _
                        "Daily Managed: " & RepeaterWS.Cells(i, 7).Value & vbCrLf & _
                        "Num TS Errors Btwn EPs: " & RepeaterWS.Cells(i, 8).Value & vbCrLf & _
                        "Reference Date-Time: " & RepeaterWS.Cells(i, 9).Value



        'Create KML code for a collector Placemark
        outputText = Placemark1 & pmName & Placemark2 & "RepStyle" & Placemark3 & longitudeValue & ", " & latitudeValue & Placemark4 & pmDescription & Placemark5

        'Print Placemark(s)
        Print #1, outputText


    Next i

    'End A folder
    outputText = Folder3
    Print #1, outputText

    
    '---------------------------------
    ' WRITE COLLECTOR/REPEATER LINES
    '---------------------------------
    
    'Start A folder for Collector/Repeater Associations
    folderName = "Collector/Repeater Associations"
    outputText = Folder1 & folderName & Folder2
    Print #1, outputText
    
    
    Dim y As Long
    Dim x As Long
    
    Dim CurrentColID As String
    
    y = 2
    
    Do While y < numRowsM

        'Re-Initialize nested loop counter/index
        x = y
        'Set New Collector ID
        ColID = mainWS.Cells(y, 1).Value
        
        'Start A folder for the Collector
        folderName = ColID
        outputText = Folder1 & folderName & Folder2
        Print #1, outputText
        

        'Loop Through each row of mainWS until the collector ID changes
        Do While mainWS.Cells(x, 1).Value = ColID

            RepId = mainWS.Cells(x, 2).Value
            
            '...........................
            ' Match Col/Rep Btwn Sheets
            '...........................
            
            'Set variables to 0
            ColRow = 0
            RepRow = 0
            
            'Find The Row Corresponding to the Collector in the collector Worksheet
            For i = 2 To numRowsC
                If CollectorWS.Cells(i, 1).Value = ColID Then
                    ColRow = i
                    Exit For
                End If
            Next i
            
            '!!!!!!!!!!!!!!!!!!!!!NOTE!!!!!!!!!!!!!!!!!!!!!
            'THIS METHOD OF SEARCHING IS ACCEPTABLE BECAUSE
            'OF THE LOW NUMBER OF REPEATERS AND COLLECTORS
            'IN THE SYSTEM
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            
            'Find The Row Corresponding to the Collector in the collector Worksheet
            For i = 2 To numRowsR
                If RepeaterWS.Cells(i, 1).Value = RepId Then
                    RepRow = i
                    Exit For
                End If
            Next i
            
            '...........................
            ' Construct Line KML Code
            '...........................
            
            ColLat = CollectorWS.Cells(ColRow, 3).Value
            ColLong = CollectorWS.Cells(ColRow, 4).Value
            
            RepLat = RepeaterWS.Cells(RepRow, 3).Value
            RepLong = RepeaterWS.Cells(RepRow, 4).Value
            
            'Set Line Attributes
            pmName = "Col: " & ColID & " to Rep: " & RepId
            pmDescription = "Rank: " & mainWS.Cells(x, 7).Value & vbCrLf & _
                            "Max RSSI: " & mainWS.Cells(x, 3).Value & vbCrLf & _
                            "Avg. RSSI: " & mainWS.Cells(x, 4).Value & vbCrLf & _
                            "Channel Bitmap: " & returnBinaryStr(mainWS.Cells(x, 5).Value) & vbCrLf & _
                            "Num Messages: " & mainWS.Cells(x, 6).Value & vbCrLf & _
                            "On Report List: " & mainWS.Cells(x, 8).Value & vbCrLf & _
                            "On Management List: " & mainWS.Cells(x, 9).Value
                            
            'Check to see if the Repeater is on the Collectors management list
            If mainWS.Cells(x, 9).Value = "True" Then
                outputText = line1 & pmName & line2 & "managedLineStyle" & line3 & Format(ColLong, "#0.000000") & "," & Format(ColLat, "#0.000000") & ",100 " & Format(RepLong, "#0.000000") & "," & Format(RepLat, "#0.000000") & ",100 " & line4 & pmDescription & Line5
            Else
                outputText = line1 & pmName & line2 & "hurdLineStyle" & line3 & Format(ColLong, "#0.000000") & "," & Format(ColLat, "#0.000000") & ",100 " & Format(RepLong, "#0.000000") & "," & Format(RepLat, "#0.000000") & ",100 " & line4 & pmDescription & Line5
            End If
            'Print Line
            Print #1, outputText
            
            'Increment Counter
            x = x + 1
            
        Loop
        
        'End folder for specific collectors associations
        outputText = Folder3
        Print #1, outputText
        
        'Assign Y = to X (when new CollectorID Starts)
        y = x


    Loop
    

    'End A folder
    outputText = Folder3
    Print #1, outputText
 

    'Write footer to file
    outputText = Footer
    Print #1, outputText
    
    Close #1
    
    
    MsgBox "KMP File Created and can be found at: " & vbCrLf & myFile
    
End Sub

'Opens a navigation window for user to select the file location and return the address
Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    
    Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
    
End Function

'Returns the Path to the Desktop
Function GetDesktop() As String
    Dim oWSHShell As Object

    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    Set oWSHShell = Nothing
    
End Function

'Return a binary String Representation of an integer
Function returnBinaryStr(inpt As Long) As String
    
    Dim temp As Long
    Dim n As Integer
    
    Dim output As String
    
   n = 23
   
    Do While inpt > 0
        temp = 2 ^ n
        If (inpt - (temp)) >= 0 Then
            output = output + "1"
            inpt = inpt - temp
        Else
            output = output + "0"
        End If
        
        n = n - 1
        
    Loop
    
    returnBinaryStr = output

End Function


