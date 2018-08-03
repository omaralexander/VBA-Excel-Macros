'We leave these as public so both the first and second run can have access to it, and reduce code clutter'


Public counterForPreferred As Integer


Public splitsCount As Integer


Public aquisitionCount As Integer


Public tickerSymbolCount As Integer


Public delistingCount As Integer


Public exchangeOfferCount As Integer


Public idNumberChangeCount As Integer


Public stockdividendCount As Integer


Public debtRedemptionCallCount As Integer


Public spinOffCount As Integer


Public rightsCount As Integer


Public mergersCount As Integer


Public drSinkFundCount As Integer


Public drPutCount As Integer


Public drCallCount As Integer


Public muniRefundCount As Integer


Public optionConversionCount As Integer


Public reclassificationOfSharesCount As Integer

Public exchangeOfferingCount As Integer

'This downloads the file from Bloomberg '

 Private Sub getBloomFile()

Dim fso As Scripting.FileSystemObject
Dim fol As Scripting.Folder
Dim currentCount As Integer

Set fso = CreateObject("Scripting.FileSystemObject")
Set fol = fso.GetFolder("C:\Program Files (x86)\blp\data\")

currentCount = fol.Files.Count

blp = DDEInitiate("Winblp", "bbk")

Call DDEExecute(blp, "<blp-1>" & "CACT" & "<GO>")
Call DDEExecute(blp, "<blp-1>" & "31" & "<GO>")
Call DDEExecute(blp, "<blp-1>" & "1" & "<GO>")

Call DDETerminate(blp)

checkWhenFileIsDone Now, currentCount


End Sub
Sub BBFileRun()

getFiles ("Bloom A.xlsx")

End Sub

    
End Sub
Private Sub getFiles(bloomType As String)

Dim fso As Scripting.FileSystemObject
Dim fol As Scripting.Folder
Dim fdr As Scripting.Folder
Dim fil As Scripting.File
Dim flc As Scripting.Folders
Dim subFolder As Scripting.Folder

Dim symbolsFolder As String
Dim BBFileLocation As String
Dim CAFileLocation As String
Dim bloomLocation As String

Dim corporateActionLocation As String

Dim bloomALocation As String

Dim bloombergFile

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)


With fd
.AllowMultiSelect = False
.Title = "Select folder"
.Filters.Clear
.InitialFileName = ""
    If .Show = True Then
    
    corporateActionLocation = .SelectedItems(1)
    
    End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fol = fso.GetFolder(corporateActionLocation)
    Set flc = fol.SubFolders

For Each fil In fol.Files

If fil.Name = "Securities Lookup Reference.xlsx" Then

symbolsFolder = fil.Path

End If

If fil.Name = bloomType Then

bloomLocation = fil.Path

End If


If fil.Name = "Bloom A.xlsx" Then
bloomALocation = fil.Path
End If

If InStr(1, fil.Name, "BB - ") = 1 Then

    BBFileLocation = fil.Path
    
    End If

        If InStr(1, fil.Name, "CA - ") = 1 Then
            
            CAFileLocation = fil.Path
                
            End If

    Next fil
    
If bloomType = "Bloom A.xlsx" Then

startMovingValuesToBB bloomLocation, BBFileLocation, CAFileLocation, symbolsFolder

End If

End Sub


Private Function findBloombergFile(time As Date, fileCount As Integer)

Dim brake As Integer
brake = 0

Dim fso As Scripting.FileSystemObject
Dim fol As Scripting.Folder
Dim fdr As Scripting.Folder
Dim fil As Scripting.File

Set fso = CreateObject("Scripting.FileSystemObject")
Set fol = fso.GetFolder("C:\Program Files (x86)\blp\data\")

While brake = 0

  If fol.Files.Count > fileCount Then
 
    
    brake = 4000
  
  End If
  
Wend
For Each fil In fol.Files
  
        If Format(fil.DateLastModified, "MM/DD/YYYY HH:MM:SS AM/PM") > time Then
                         
                findBloombergFile fil.Path

        End If
    Next fil
  
  End Function

Private Sub startMovingValuesToBB(bloomLocation As String, secondWorkbookLocation As String, CAWorkbook As String, symbolLocation As String)

Dim firstWorkBook As Workbook
Dim secondWorkbook As Workbook

Set firstWorkBook = Workbooks.Open(bloomLocation)
Set secondWorkbook = Workbooks.Open(secondWorkbookLocation)

secondWorkbook.Sheets("BB File Conversion").Cells.AutoFilter

BBFileCleanUp

firstWorkBook.Sheets(1).range("A6:U6", "A2333:U2333").Cut secondWorkbook.Sheets("BB File Conversion").range("B3:Z3", "B2333:Z2333")

'Worksheets("GIM2 Symbols").Range("A1:N1", "A2333:N2333").Interior.Color = RGB(146, 208, 80)

Worksheets("Holdings").range("A1:L1", "A2333:A2333").Interior.Color = RGB(146, 208, 80)
             
            Worksheets("BB File Conversion").range("A2").Formula = "=IF(B2=""Ticker Symbol Change"",$T2,IF(B2 = ""Splits"",$C2,IF(B2 = ""Option Conversions"",$C2,$Q2)))"
            Worksheets("BB File Conversion").range("A2").AutoFill Destination:=Worksheets("BB File Conversion").range("A2:A2333")

            Worksheets("BB File Conversion").range("Q2").Formula = "=IF(B2=""Ticker Symbol Change"",MID(H2,14,LEN(H2)-7),IF(MID(C2,LEN(C2)-3,4)=""Corp"",MID(C2,1,LEN(C2)-5),IF(MID(C2,LEN(C2)-3,4)=""Muni"",MID(C2,1,LEN(C2)-5),IF(MID(C2,LEN(C2)-9,10)="" US Equity"",MID(C2,1,LEN(C2)-10),IF(MID(C2,LEN(C2)-5,6)=""Equity"",MID(C2,1,LEN(C2)-7),IF(MID(C2,LEN(C2)-2,3)=""Pfd"",MID(C2,1,LEN(C2)-4)))))))"
            Worksheets("BB File Conversion").range("Q2").AutoFill Destination:=Worksheets("BB File Conversion").range("Q2:Q2333")

            Worksheets("BB File Conversion").range("R2").Formula = "=RIGHT(H2,LEN(H2)-13)"
            Worksheets("BB File Conversion").range("R2").AutoFill Destination:=Worksheets("BB File Conversion").range("R2:R2333")

            Worksheets("BB File Conversion").range("S2").Formula = "=RIGHT(Q2,2)"
            Worksheets("BB File Conversion").range("S2").AutoFill Destination:=Worksheets("BB File Conversion").range("S2:S2333")

            Worksheets("BB File Conversion").range("T2").Formula = "=IF(S2=""US"",LEFT(Q2,LEN(Q2)-3),Q2)"
            Worksheets("BB File Conversion").range("T2").AutoFill Destination:=Worksheets("BB File Conversion").range("T2:T2333")
 
'Clear out the extra N/A Values'

Dim finalValueOfBBSheet As Integer

For i = 2 To 2333

If Worksheets("BB File Conversion").range("B" & i).Value = "" Then Exit For

Next i

finalValueOfBBSheet = i


Worksheets("BB File Conversion").range("A" & finalValueOfBBSheet & ":X" & finalValueOfBBSheet, "A5000" & ":X5000").ClearContents
Worksheets("GIM2 Symbols").range("A" & finalValueOfBBSheet & ":X" & finalValueOfBBSheet, "A5000" & ":X5000").ClearContents
Worksheets("Holdings").range("A" & finalValueOfBBSheet & ":L" & finalValueOfBBSheet, "A5000" & ":L5000").ClearContents

Worksheets("GIM2 Symbols").range("A" & finalValueOfBBSheet & ":X" & finalValueOfBBSheet, "A5000" & ":X5000").Interior.Color = RGB(255, 255, 255)
Worksheets("GIM2 Symbols").range("A" & finalValueOfBBSheet & ":L" & finalValueOfBBSheet, "A5000" & ":L5000").Interior.Color = RGB(255, 255, 255)

Dim symbolsWorkbook As Workbook


Set symbolsWorkbook = Workbooks.Open(symbolLocation)

With secondWorkbook.Sheets("GIM2 Symbols").range("A" & finalValueOfBBSheet & ":X" & finalValueOfBBSheet, "A5000" & ":X5000").Borders
.LineStyle = xlContinuous
.Weight = xlThin
.ColorIndex = 15
End With

With secondWorkbook.Sheets("Holdings").range("A" & finalValueOfBBSheet & ":L" & finalValueOfBBSheet, "A5000" & ":L5000").Borders
.LineStyle = xlContinuous
.Weight = xlThin
.ColorIndex = 15
End With

With symbolsWorkbook.Sheets(1).range("A2" & ":A" & finalValueOfBBSheet).Borders
.LineStyle = xlContinuous
.Weight = xlThin
.ColorIndex = 15
End With

symbolsWorkbook.Sheets(1).range("A2:A2333").ClearContents

symbolsWorkbook.Sheets(1).range("A2" & ":A" & finalValueOfBBSheet).Value = secondWorkbook.Sheets("BB File Conversion").range("A2" & ":A" & finalValueOfBBSheet).Value
symbolsWorkbook.Sheets(1).range("A2" & ":A" & finalValueOfBBSheet).Interior.Color = RGB(255, 255, 255)


firstWorkBook.Close False



counterForPreferred = 1
getPreferredSymbols secondWorkbook
counterForPreferred = 1

For clearNA = 2 To 2333
    
    If secondWorkbook.Sheets("GIM2 Symbols").range("B" & clearNA).Value = "" Then Exit For
    
    If IsError(secondWorkbook.Sheets("GIM2 Symbols").range("A" & clearNA).Value) = True And IsError(secondWorkbook.Sheets("GIM2 Symbols").range("L" & clearNA).Value) = True Then
        
        secondWorkbook.Sheets("GIM2 Symbols").range("A" & clearNA & ":P" & clearNA).ClearContents
        
        End If
    
Next clearNA

'shift the blank values up'
secondWorkbook.Sheets("GIM2 Symbols").range("A2" & ":P" & clearNA).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp


End Sub
Private Function getPreferredSymbols(BBWorkbook As Workbook)

    For i = 2 To 2333

            If IsError(BBWorkbook.Sheets("BB File Conversion").range("C" & i).Value) = False Then
                    
                If BBWorkbook.Sheets("BB File Conversion").range("C" & i).Value = "" Then Exit For
                    
                    If InStrRev(BBWorkbook.Sheets("BB File Conversion").range("C" & i).Value, " Pfd") <> 0 Then
                                                 
                             counterForPreferred = counterForPreferred + 1
                             
                             BBWorkbook.Sheets("Preferreds").range("B" & counterForPreferred & ":F" & counterForPreferred).Value = BBWorkbook.Sheets("BB File Conversion").range("A" & i & ":E" & i).Value
                        
                    End If
            
            End If
        
    Next i
        


End Function

Sub CAFileRun()

'Auto saving the BB and CA workbooks automatically'

For Each Wb In Application.Workbooks

If InStr(1, Wb.Name, "CA - ") = 1 Or InStr(1, Wb.Name, "BB -") = 1 Then
    
    Wb.Save
    
End If

Next Wb

Dim fso As Scripting.FileSystemObject
Dim fol As Scripting.Folder
Dim fdr As Scripting.Folder
Dim fil As Scripting.File
Dim flc As Scripting.Folders
Dim subFolder As Scripting.Folder

Dim CAFileLocation As String
Dim BBFileLocation As String

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)


With fd
.AllowMultiSelect = False
.Title = "Select folder"
.Filters.Clear
.InitialFileName = ""
    If .Show = True Then
    
    corporateActionLocation = .SelectedItems(1)
    
    End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fol = fso.GetFolder(corporateActionLocation)
    Set flc = fol.SubFolders


For Each fil In fol.Files

        If InStr(1, fil.Name, "CA - ") = 1 Then
            
            CAFileLocation = fil.Path
                
            End If
            
            If InStr(1, fil.Name, "BB - ") = 1 Then

    BBFileLocation = fil.Path
    
    End If

    Next fil


Dim CAWorkbook As Workbook
Dim secondWorkbook As Workbook
Dim startValue As Integer

Set secondWorkbook = Workbooks.Open(BBFileLocation)
Set CAWorkbook = Workbooks.Open(CAFileLocation)


CAWorkbook.Activate

cleanUpCAFile

'remove duplicates'

secondWorkbook.Sheets("GIM2 Symbols").range("A2:N2333").RemoveDuplicates Columns:=Array(1, 2)

'remove values inside of dupe 1 and dupe 2'

secondWorkbook.Sheets("GIM2 Symbols").range("O2:P2", "O2333:P2333").ClearContents

'Start looking at loop values'

For duploop = 2 To 2333

If secondWorkbook.Sheets("GIM2 Symbols").range("M" & duploop).Value = "Duplicated!" Then

    findDupeValues secondWorkbook.Sheets("GIM2 Symbols").range("B" & duploop).Value, secondWorkbook, secondWorkbook.Sheets("GIM2 Symbols").range("A" & duploop).Value, CInt(duploop)

End If
Next duploop


For lastRow = 2 To 2333

If secondWorkbook.Sheets("GIM2 Symbols").range("J" & lastRow).Value = "" Then Exit For

Next lastRow

'Move any option conversions to the end of the GIM2 Sheet'

With secondWorkbook.Sheets("BB File Conversion").range("B2:B2334")
    Set optionLocationInBB = .find("Option Conversions", LookIn:=xlValues)
        If Not optionLocationInBB Is Nothing Then
            
            firstAddress = optionLocationInBB.Address
        
            Do
        
            optionRow = Split(optionLocationInBB.Address, "$B$")(1)
        
            secondWorkbook.Sheets("GIM2 Symbols").range("C" & lastRow & ":J" & lastRow).Value = secondWorkbook.Sheets("BB File Conversion").range("C" & optionRow & ":H" & optionRow).Value
            
            secondWorkbook.Sheets("GIM2 Symbols").range("B" & lastRow).Value = secondWorkbook.Sheets("GIM2 Symbols").range("C" & lastRow).Value
            
            secondWorkbook.Sheets("GIM2 Symbols").range("J" & lastRow).Value = "A"
            
            secondWorkbook.Sheets("GIM2 Symbols").range("I" & lastRow).Value = "USD"
            
            lastRow = lastRow + 1

            Set optionLocationInBB = .FindNext(optionLocationInBB)
            
            If optionLocationInBB Is Nothing Then
                 
                 GoTo DoneFinding
                 
                 End If
                 
                 Loop While optionLocationInBB.Address <> firstAddress
                 
                 End If
                
DoneFinding:
                 
                 End With
                 
                 

'Move any splits to the end of the GIM2 Sheet'
                 
With secondWorkbook.Sheets("BB File Conversion").range("B2:B2334")
    Set splitsLocationInBB = .find("Splits", LookIn:=xlValues)
        If Not splitsLocationInBB Is Nothing Then
            
            firstAddress = splitsLocationInBB.Address
        
            Do
        
            splitsRow = Split(splitsLocationInBB.Address, "$B$")(1)
        
            secondWorkbook.Sheets("GIM2 Symbols").range("C" & lastRow & ":J" & lastRow).Value = secondWorkbook.Sheets("BB File Conversion").range("C" & splitsRow & ":H" & splitsRow).Value
            
            secondWorkbook.Sheets("GIM2 Symbols").range("B" & lastRow).Value = secondWorkbook.Sheets("GIM2 Symbols").range("C" & lastRow).Value
            
            secondWorkbook.Sheets("GIM2 Symbols").range("J" & lastRow).Value = "A"
            
            secondWorkbook.Sheets("GIM2 Symbols").range("I" & lastRow).Value = "USD"
            
            lastRow = lastRow + 1

            Set splitsLocationInBB = .FindNext(splitsLocationInBB)
            
            If splitsLocationInBB Is Nothing Then
                 
                 GoTo SplitsDoneFinding
                 
                 End If
                 
                 Loop While splitsLocationInBB.Address <> firstAddress
                 
                 End If
                
SplitsDoneFinding:
                 
                 End With


lastRowForTheBBSheet = lastRow - 1


'Changing color to green to not mess up second run'
secondWorkbook.Sheets("GIM2 Symbols").range("A1:N1", "A" & lastRow & ":N" & lastRow).Interior.Color = RGB(146, 208, 80)

'Autofill formulas on secondWorkbook to avoid errors and also to keep splits and option conversions bundled with Group'

    secondWorkbook.Worksheets("GIM2 Symbols").range("L2").Formula = "=TEXT(VLOOKUP('GIM2 Symbols'!B2,'BB File Conversion'!$A:$Q,5,0),""MM/DD/YY"")"
    secondWorkbook.Worksheets("GIM2 Symbols").range("L2").AutoFill Destination:=secondWorkbook.Worksheets("GIM2 Symbols").range("L2" & ":L" & lastRow)
    
     secondWorkbook.Worksheets("GIM2 Symbols").range("A2").Formula = "=K2"
    secondWorkbook.Worksheets("GIM2 Symbols").range("A2").AutoFill Destination:=secondWorkbook.Worksheets("GIM2 Symbols").range("A2" & ":A" & lastRow)

    secondWorkbook.Worksheets("GIM2 Symbols").range("K2").Formula = "=VLOOKUP(B2,'BB File Conversion'!A:B,2,0)"
    secondWorkbook.Worksheets("GIM2 Symbols").range("K2").AutoFill Destination:=secondWorkbook.Worksheets("GIM2 Symbols").range("K2" & ":K" & lastRow)

    secondWorkbook.Worksheets("GIM2 Symbols").range("M2").Formula = "=IF(COUNTIF('BB File Conversion'!A:A,B2)>1, ""Duplicated!"",""OK"")"
    secondWorkbook.Worksheets("GIM2 Symbols").range("M2").AutoFill Destination:=secondWorkbook.Worksheets("GIM2 Symbols").range("M2" & ":M" & lastRow)

    secondWorkbook.Worksheets("GIM2 Symbols").range("N2").Formula = "=""'""&C2&""',"""
    secondWorkbook.Worksheets("GIM2 Symbols").range("N2").AutoFill Destination:=secondWorkbook.Worksheets("GIM2 Symbols").range("N2" & ":N" & lastRow)

    
    

'Pass the values to the count'

splitsCount = findLastValueFor("Splits", CAWorkbook, 5)

aquisitionCount = findLastValueFor("Acquisitions", CAWorkbook, 3)

tickerSymbolCount = findLastValueFor("Ticker Symbol Changes", CAWorkbook, 3)

delistingCount = findLastValueFor("Delistings", CAWorkbook, 5)

exchangeOfferCount = findLastValueFor("Exchange Offer", CAWorkbook, 4)

idNumberChangeCount = findLastValueFor("ID Number Change", CAWorkbook, 3)

stockdividendCount = findLastValueFor("Stock Dividend", CAWorkbook, 3)

debtRedemptionCallCount = findLastValueFor("DR Call", CAWorkbook, 4)

spinOffCount = findLastValueFor("Spinoffs", CAWorkbook, 3)

rightsCount = findLastValueFor("Rights", CAWorkbook, 4)

mergersCount = findLastValueFor("Mergers", CAWorkbook, 4)

drSinkFundCount = findLastValueFor("DR Sink Fund", CAWorkbook, 4)

drPutCount = findLastValueFor("DR Put", CAWorkbook, 4)

drCallCount = findLastValueFor("DR Call", CAWorkbook, 4)

muniRefundCount = findLastValueFor("Muni Refund", CAWorkbook, 4)

optionConversionCount = findLastValueFor("Option Conversions", CAWorkbook, 4)

reclassificationOfSharesCount = findLastValueFor("Reclassification Of Shares", CAWorkbook, 5)

exchangeOfferingCount = findLastValueFor("Exchange Offer", CAWorkbook, 4)


For i = 2 To lastRowForTheBBSheet

'Check if the date and the corporate action are N/A'

    If IsError(secondWorkbook.Sheets("GIM2 Symbols").range("A" & i).Value) = False And IsError(secondWorkbook.Sheets("GIM2 Symbols").range("L" & i).Value) = False Then
    
    For filterLoopStart = 2 To 2333
        
        If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = "" Then Exit For
        
            If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = secondWorkbook.Sheets("GIM2 Symbols").range("A" & i) Then
                
                If secondWorkbook.Sheets("CA Interested").range("B" & filterLoopStart).Value = "A" Then
                                
                 passValuesToCA secondWorkbook.Sheets("GIM2 Symbols").range("A" & i).Value, CInt(i), CAWorkbook, secondWorkbook
                
                
                        End If
                
                    End If
        
                Next filterLoopStart
    
            End If
             
    Next i


For i = 2 To lastRowForTheBBSheet
    
    If IsError(secondWorkbook.Sheets("GIM2 Symbols").range("O" & i).Value) = False And IsError(secondWorkbook.Sheets("GIM2 Symbols").range("L" & i).Value) = False And IsError(secondWorkbook.Sheets("GIM2 Symbols").range("P" & i).Value) = False Then
    
    Select Case secondWorkbook.Sheets("GIM2 Symbols").range("O" & i).Value

Case Is = ""

Case Is <> ""
    
    'If both the O and P columns are populated'
                
         For filterLoopStart = 2 To 2333
        
            If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = "" Then Exit For
                       
                If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = secondWorkbook.Sheets("GIM2 Symbols").range("O" & i).Value Then
                
                    If secondWorkbook.Sheets("CA Interested").range("B" & filterLoopStart).Value = "A" Then
                                
                    passValuesToCA secondWorkbook.Sheets("GIM2 Symbols").range("O" & i).Value, CInt(i), CAWorkbook, secondWorkbook
                
                        End If
                
                    End If
                    
                    'O does not meet criteria so we look at P'
                If secondWorkbook.Sheets("GIM2 Symbols").range("P" & i).Value <> "" Then
                
                    If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = secondWorkbook.Sheets("GIM2 Symbols").range("P" & i).Value Then
                        
                        If secondWorkbook.Sheets("CA Interested").range("B" & filterLoopStart).Value = "A" Then
                                
                            passValuesToCA secondWorkbook.Sheets("GIM2 Symbols").range("P" & i).Value, CInt(i), CAWorkbook, secondWorkbook
                
                            End If
                
                        End If
                    
                    End If
                    
            Next filterLoopStart
            
Case Is <> ""

'Contains value, we then check if it meets the criteria'
     
     For filterLoopStart = 2 To 2333
        
            If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = "" Then Exit For
                       
                If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = secondWorkbook.Sheets("GIM2 Symbols").range("O" & i).Value Then
                
                    If secondWorkbook.Sheets("CA Interested").range("B" & filterLoopStart).Value = "A" Then
                                
                    passValuesToCA secondWorkbook.Sheets("GIM2 Symbols").range("O" & i).Value, CInt(i), CAWorkbook, secondWorkbook
                
                            Exit For
                            
                        End If
                
                    End If
                
                Next filterLoopStart
                                
                End Select
            
            End If
        
        Next i
        
    'Prepping count for holdings column so we add 3 to create space'
        
    splitsCount = splitsCount + 3
         
    aquisitionCount = aquisitionCount + 3

    tickerSymbolCount = tickerSymbolCount + 3

    delistingCount = delistingCount + 3

    exchangeOfferCount = exchangeOfferCount + 3

    idNumberChangeCount = idNumberChangeCount + 3

    stockdividendCount = stockdividendCount + 3

    debtRedemptionCallCount = debtRedemptionCallCount + 3

    spinOffCount = spinOffCount + 3

    rightsCount = rightsCount + 3

    mergersCount = mergersCount + 3

    drSinkFundCount = drSinkFundCount + 3

    drPutCount = drPutCount + 3

    drCallCount = drCallCount + 3

    muniRefundCount = muniRefundCount + 3

    optionConversionCount = optionConversionCount + 3

    reclassificationOfSharesCount = reclassificationOfSharesCount + 3
    
    exchangeOfferingCount = exchangeOfferingCount + 3
    
        'Insert the column of the header row of the holdings'
        
        For i = 1 To CAWorkbook.Worksheets.Count
        
            Select Case CAWorkbook.Sheets(i).Name
            
            Case Is = "Splits"
            insertEmptyRows 2, splitsCount, CAWorkbook.Sheets(i).Name, CAWorkbook
            
            CAWorkbook.Sheets(i).range("A" & splitsCount & ":L" & splitsCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
            CAWorkbook.Sheets(i).range("A" & splitsCount & ":L" & splitsCount).Interior.Color = RGB(146, 208, 80)
            
           
            Case Is = "Acquisitions"
            insertEmptyRows 2, aquisitionCount, CAWorkbook.Sheets(i).Name, CAWorkbook
            
            CAWorkbook.Sheets(i).range("A" & aquisitionCount & ":L" & aquisitionCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
            CAWorkbook.Sheets(i).range("A" & aquisitionCount & ":L" & aquisitionCount).Interior.Color = RGB(146, 208, 80)
        
            Case Is = "Ticker Symbol Changes"
            
            insertEmptyRows 2, tickerSymbolCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                      
            CAWorkbook.Sheets(i).range("A" & tickerSymbolCount & ":L" & tickerSymbolCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
            CAWorkbook.Sheets(i).range("A" & tickerSymbolCount & ":L" & tickerSymbolCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "ID Number Change"
                
                insertEmptyRows 2, idNumberChangeCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & idNumberChangeCount & ":L" & idNumberChangeCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & idNumberChangeCount & ":L" & idNumberChangeCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "Stock Dividend"
                
                insertEmptyRows 2, stockdividendCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & stockdividendCount & ":L" & stockdividendCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & stockdividendCount & ":L" & stockdividendCount).Interior.Color = RGB(146, 208, 80)
               
            Case Is = "Spinoffs"
                
                insertEmptyRows 2, spinOffCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & spinOffCount & ":L" & spinOffCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & spinOffCount & ":L" & spinOffCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "Rights"
                
                insertEmptyRows 2, rightsCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & rightsCount & ":L" & rightsCount).Interior.Color = RGB(146, 208, 80)
            CAWorkbook.Sheets(i).range("A" & rightsCount & ":L" & rightsCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             
            Case Is = "Mergers"
                
                insertEmptyRows 2, mergersCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & mergersCount & ":L" & mergersCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & mergersCount & ":L" & mergersCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "DR Sink Fund"
                
                insertEmptyRows 2, drSinkFundCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & drSinkFundCount & ":L" & drSinkFundCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & drSinkFundCount & ":L" & drSinkFundCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "DR Put"
                
                insertEmptyRows 2, drPutCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & drPutCount & ":L" & drPutCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & drPutCount & ":L" & drPutCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "DR Call"
                
                insertEmptyRows 2, drCallCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & drCallCount & ":L" & drCallCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & drCallCount & ":L" & drCallCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "Muni Refund"
                
                insertEmptyRows 2, muniRefundCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & muniRefundCount & ":L" & muniRefundCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & muniRefundCount & ":L" & muniRefundCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "Option Conversions"
                
                insertEmptyRows 2, optionConversionCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & optionConversionCount & ":L" & optionConversionCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & optionConversionCount & ":L" & optionConversionCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "Reclassification of Shares"
                
                insertEmptyRows 2, reclassificationOfSharesCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & reclassificationOfSharesCount & ":L" & reclassificationOfSharesCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
            CAWorkbook.Sheets(i).range("A" & reclassificationOfSharesCount & ":L" & reclassificationOfSharesCount).Interior.Color = RGB(146, 208, 80)
             
            Case Is = "Exchange Offer"
                
                insertEmptyRows 2, exchangeOfferCount, CAWorkbook.Sheets(i).Name, CAWorkbook
                
            CAWorkbook.Sheets(i).range("A" & exchangeOfferCount & ":L" & exchangeOfferCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
             CAWorkbook.Sheets(i).range("A" & exchangeOfferCount & ":L" & exchangeOfferCount).Interior.Color = RGB(146, 208, 80)
                         
            Case Is = "Delistings"
                 
             insertEmptyRows 2, delistingCount, CAWorkbook.Sheets(i).Name, CAWorkbook
             
            CAWorkbook.Sheets(i).range("A" & delistingCount & ":L" & delistingCount).Value = secondWorkbook.Sheets("Holdings").range("A1:L1").Value
            CAWorkbook.Sheets(i).range("A" & delistingCount & ":L" & delistingCount).Interior.Color = RGB(146, 208, 80)
             
                End Select
           
        Next i
        
               
        'removes the GIM2 Duplicates CA Workbook here'
        
    For removeDupeFromWorksheet = 1 To CAWorkbook.Worksheets.Count
        
    Select Case CAWorkbook.Worksheets(removeDupeFromWorksheet).Name
        
            Case Is = "Splits"
                  Dim splitsBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown

                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then splitsBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                 CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & splitsBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes
                                 
        Case Is = "Acquisitions"
                
                    Dim aqBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then aqBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                  CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & aqBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes
                
            Case Is = "Ticker Symbol Changes"
            
                Dim tickerSymbolBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then tickerSymbolBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                 CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & tickerSymbolBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes
  
            Case Is = "ID Number Change"
            
                 Dim idNumberBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then idNumberBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                  CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & idNumberBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes
                
            Case Is = "Stock Dividend"
            
                 Dim stockdivBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then stockdivBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & stockdivBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes
                
                 
            Case Is = "Spinoffs"
            
                 Dim spinoffBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then spinoffBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                   CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & spinoffBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes
                
            Case Is = "Rights"
            
                Dim rightsBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then rightsBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                    CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & rightsBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            Case Is = "Mergers"
            
                 Dim mergersBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then mergersBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                  CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & mergersBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            Case Is = "DR Sink Fund"
            
                 Dim drSinkFundBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then drSinkFundBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                     CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & drSinkFundBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes
                

            Case Is = "DR Put"
            
                 Dim drPutBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then drPutBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                
                 CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & drPutBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            Case Is = "DR Call"
            
                 Dim drCallBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then drCallBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                 CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & drCallBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            Case Is = "Muni Refund"
            
                Dim muniRefundBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then muniRefundBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                   CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & muniRefundBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            Case Is = "Option Conversions"
            
                Dim optionConversionBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then optionConversionBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                   CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & optionConversionBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            Case Is = "Reclassification of Shares"
            
                 Dim reclassBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then reclassBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                 CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & reclassBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            Case Is = "Exchange Offer"
            
                 Dim exchangeOfferBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then exchangeOfferBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
                  CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & exchangeOfferBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            Case Is = "Delistings"
                 Dim delistingBlackBar As Integer
            
            For testLoop = 3 To 2333
                        
                        If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                                                
                Next testLoop
                CAWorkbook.Sheets(removeDupeFromWorksheet).Cells.UnMerge
                CAWorkbook.Sheets(removeDupeFromWorksheet).range("A3" & ":R" & testLoop - 1).RemoveDuplicates Columns:=Array(1, 2)
               CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & testLoop + 2 & ":R" & testLoop + 2).Insert Shift:=xlDown
                
                For bottomLoop = testLoop + 3 To 2333
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(0, 0, 0) Then delistingBlackBar = bottomLoop
                    
                    If CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & bottomLoop).Interior.Color = RGB(146, 208, 80) Then Exit For
                
                Next bottomLoop
             CAWorkbook.Sheets(removeDupeFromWorksheet).range("A" & delistingBlackBar & ":R" & bottomLoop - 1).RemoveDuplicates Columns:=Array(1, 2), header:=xlYes

            End Select
     
        Next removeDupeFromWorksheet
               
    
        'Change the count so that holdings will read updated numbers'
        
    splitsCount = splitsCount + 1
        
    aquisitionCount = aquisitionCount + 1

    tickerSymbolCount = tickerSymbolCount + 1

    delistingCount = delistingCount + 1

    exchangeOfferCount = exchangeOfferCount + 1

    idNumberChangeCount = idNumberChangeCount + 1

    stockdividendCount = stockdividendCount + 1

    debtRedemptionCallCount = debtRedemptionCallCount + 1

    spinOffCount = spinOffCount + 1

    rightsCount = rightsCount + 1

    mergersCount = mergersCount + 1

    drSinkFundCount = drSinkFundCount + 1

    drPutCount = drPutCount + 1

    drCallCount = drCallCount + 1

    muniRefundCount = muniRefundCount + 1

    optionConversionCount = optionConversionCount + 1

    reclassificationOfSharesCount = reclassificationOfSharesCount + 1
    
    exchangeOfferingCount = exchangeOfferingCount + 1
        
        
        
        'start passing the holdings values'
        
        
        For holdingsLoop = 2 To 2333
            
            If secondWorkbook.Sheets("Holdings").range("A" & holdingsLoop).Value = "" Then Exit For
                    
                    'Making sure the date and the reorg type is not n/a'
                    
                        If IsError(secondWorkbook.Sheets("Holdings").range("M" & holdingsLoop).Value) = False And IsError(secondWorkbook.Sheets("Holdings").range("N" & holdingsLoop).Value) = False Then
                        
                                passHoldingsValuesToCA secondWorkbook.Sheets("Holdings").range("M" & holdingsLoop).Value, CInt(holdingsLoop), CAWorkbook, secondWorkbook
                                
                End If
                
        Next holdingsLoop
          
          'pass dupe values for holdings'
          
          moveDupeValues secondWorkbook, CAWorkbook
          

   'After values have been passed we set the count equal to zero to prep for second CA run of the day'
    
    splitsCount = 0
    
    aquisitionCount = 0

    tickerSymbolCount = 0

    delistingCount = 0

    exchangeOfferCount = 0

    idNumberChangeCount = 0

    stockdividendCount = 0

    debtRedemptionCallCount = 0

    spinOffCount = 0

    rightsCount = 0

    mergersCount = 0

    drSinkFundCount = 0

    drPutCount = 0

    drCallCount = 0

    muniRefundCount = 0

    optionConversionCount = 0

    reclassificationOfSharesCount = 0
    
    exchangeOfferingCount = 0

    'We check if the dupes meet the filter criteria, before passing them '
    
End Sub
Private Function insertEmptyRows(howMany As Integer, startingFrom As Integer, sheetName As String, CAWorkbook As Workbook)

    For i = 1 To howMany
        'CAWorkbook.Sheets(sheetName).Cells.AutoFilter
            CAWorkbook.Sheets(sheetName).range("A" & startingFrom & ":R" & startingFrom).Insert Shift:=xlDown
        Next i
   
End Function

Private Function checkIfGIM2SymbolMeetsCriteria(secondWorkbook As Workbook, CAWorkbook As Workbook, GIM2CorporateAction As String, LocationOfTheCell As Integer, lastRowOnBBSheet As Integer)

For i = 2 To lastRowOnBBSheet

If secondWorkbook.Sheets("GIM2 Symbols").range("A" & i).Value = "" Then Exit For

'Check if the date and the corporate action are N/A'

If IsError(secondWorkbook.Sheets("GIM2 Symbols").range("A" & i).Value) = False And IsError(secondWorkbook.Sheets("GIM2 Symbols").range("L" & i).Value) = False Then
      
    'Check if GIM2 Symbols meets the filter criteria'
    
    For filterLoopStart = 2 To 2333
        
        If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = "" Then Exit For
        
            If secondWorkbook.Sheets("CA Interested").range("A" & filterLoopStart).Value = secondWorkbook.Sheets("GIM2 Symbols").range("A" & i) Then
                
                If secondWorkbook.Sheets("CA Interested").range("B" & filterLoopStart).Value = "A" Then
                                
                 passValuesToCA secondWorkbook.Sheets("GIM2 Symbols").range("A" & i).Value, CInt(i), CAWorkbook, secondWorkbook
                
                
                        End If
                
                    End If
        
                Next filterLoopStart
    
            End If
             
    Next i




End Function

Private Function findLocationOfTheCell(symbolString As String, actionString As String, BBWorkbook As Workbook)

For i = 2 To 2333
    If IsError(BBWorkbook.Sheets("BB File Conversion").range("A" & i).Value) = False Then
        If BBWorkbook.Sheets("BB File Conversion").range("A" & i).Value = symbolString Then
            If BBWorkbook.Sheets("BB File Conversion").range("B" & i).Value = actionString Then Exit For
                End If
                    End If
            
    If BBWorkbook.Sheets("BB File Conversion").range("B" & i).Value = "" Then Exit For


    Next i

    findLocationOfTheCell = i
       
End Function

Private Sub cleanUpCAFile()

Const dummyFinalCell As Integer = 2333
Dim endingTopValue As Integer
Dim startingTopValue As Integer
Dim bottomEndingValue As Integer

Dim worksheetCount As Integer
Dim i As Integer

worksheetCount = ActiveWorkbook.Worksheets.Count

For i = 1 To worksheetCount
        If ActiveWorkbook.Worksheets(i).Name <> "Recon" Then
        
        For startingValue = 1 To dummyFinalCell
                
                If Worksheets(ActiveWorkbook.Worksheets(i).Name).range("A" & startingValue).Interior.Color = RGB(255, 255, 255) And Worksheets(ActiveWorkbook.Worksheets(i).Name).range("A" & startingValue).Value <> "Symbols ending in ""Q"" represent bankruptcies. Book conversions for these types of symbols to the issue of the new security listed. (Do not change any names to ""_OLD"")" Then Exit For
                    
Next
                    
    startingTopValue = startingValue

        For startingAtCell = startingTopValue To dummyFinalCell

            If Worksheets(ActiveWorkbook.Worksheets(i).Name).range("A" & startingAtCell).Interior.Color = RGB(0, 0, 0) Then Exit For
        
        Next
    
            endingTopValue = startingAtCell - 1

    For startingBottomValue = endingTopValue + 1 To dummyFinalCell
    
        If Worksheets(ActiveWorkbook.Worksheets(i).Name).range("A" & startingBottomValue).Interior.Color = RGB(255, 255, 255) And Worksheets(ActiveWorkbook.Worksheets(i).Name).range("A" & startingBottomValue).Value <> "Symbols ending in ""Q"" represent bankruptcies. Book conversions for these types of symbols to the issue of the new security listed. (Do not change any names to ""_OLD"")" Then Exit For
     
    Next
        
        bottomEndingValue = startingBottomValue
        
        Worksheets(ActiveWorkbook.Worksheets(i).Name).range("A" & bottomEndingValue & ":BQ" & dummyFinalCell).Clear

'Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("A" & startingTopValue & ":BQ" & startingTopValue, "A" & endingTopValue & ":BQ" & endingTopValue).Cut Worksheets(ActiveWorkbook.Worksheets(I).Name).Range("A" & bottomEndingValue & ":BQ" & bottomEndingValue, "A" & dummyFinalCell & ":BQ" & dummyFinalCell)
              
              Worksheets(ActiveWorkbook.Worksheets(i).Name).range("A" & startingTopValue & ":BQ" & endingTopValue).Cut Worksheets(ActiveWorkbook.Worksheets(i).Name).range("A" & bottomEndingValue)
    End If
    
  Next i

'Format the date in the CA File'

Worksheets("Mergers").range("J:J").NumberFormat = "MM/DD/YYYY"
Worksheets("Ticker Symbol Changes").range("L:L").NumberFormat = "MM/DD/YYYY"
Worksheets("DR Call").range("J:J").NumberFormat = "MM/DD/YYYY"
Worksheets("ID Number Change").range("M:M").NumberFormat = "MM/DD/YYYY"
Worksheets("Rights").range("K:K").NumberFormat = "MM/DD/YYYY"
Worksheets("Exchange Offer").range("K:K").NumberFormat = "MM/DD/YYYY"

End Sub
Public Function findLastValueFor(sheetName As String, CAWorkbook As Workbook, startFrom As Integer)

CAWorkbook.Activate

For i = startFrom To 2333
     
     If CAWorkbook.Worksheets(sheetName).range("A" & i + 1).Value = "" Then
    
                    findLastValueFor = i
                
                Exit For
            
            End If
        
        Next i
    
End Function

Private Sub BBFileCleanUp()

'Create arbitrary final number to find range'
Const dummyFinalCell As Integer = 2333

'Create variable to hold amount of worksheets in file'
Dim worksheetCount As Integer

'Variable to hold count of which worksheet loop it's on'
Dim i As Integer

'Pass value of amount of worksheets'
worksheetCount = ActiveWorkbook.Worksheets.Count

'IMPORTANT: GIM2 Symbols need to be deleted first before anything else'

    'Finding the range of "GIM2 Symbols" using the color of the cell as the loop exit  '
     For gettingGreenValueRange = 2 To dummyFinalCell
        
        If Worksheets("GIM2 Symbols").range("J" & gettingGreenValueRange).Interior.Color = RGB(255, 255, 255) Then Exit For
    
    Next gettingGreenValueRange
    
    'Start looping and deleting the cells'
    For i = 2 To 2333
    
    If Worksheets("GIM2 Symbols").range("J" & i).Value = "" Then Exit For
    
    Next i
    
    'Can't use this code below because it woult delete all the dates'
    
     'Worksheets("GIM2 Symbols").Range("A" & I & ":AC" & I, "A2333:AC2333").ClearContents

   
    For startValueInGIM = gettingGreenValueRange - 1 To 2 Step -1
    
    If IsError(Worksheets("GIM2 Symbols").range("L" & startValueInGIM).Value) = False Then
        
       If Worksheets("GIM2 Symbols").range("L" & startValueInGIM).Value <> " " Then

            If Format(CDate(Worksheets("GIM2 Symbols").range("L" & startValueInGIM).Value), "MM/DD/YYYY") >= Format(Date, "MM/DD/YYYY") Then
            'We keep it '
    Else

    'If an "mismatch error" appears, that means there was an n/a date that appeared'
    
    'We delete the row'
    
    Worksheets("GIM2 Symbols").range("A" & startValueInGIM & ":AC" & startValueInGIM).Delete
           
         End If
                End If
                    End If
                        Next startValueInGIM
                       
                            
    Worksheets("GIM2 Symbols").range("A2").Formula = "=K2"
    Worksheets("GIM2 Symbols").range("A2").AutoFill Destination:=Worksheets("GIM2 Symbols").range("A2:A2333")

    Worksheets("GIM2 Symbols").range("K2").Formula = "=VLOOKUP(B2,'BB File Conversion'!A:B,2,0)"
    Worksheets("GIM2 Symbols").range("K2").AutoFill Destination:=Worksheets("GIM2 Symbols").range("K2:K2333")


    Worksheets("GIM2 Symbols").range("L2").Formula = "=TEXT(VLOOKUP('GIM2 Symbols'!B2,'BB File Conversion'!$A:$Q,5,0),""MM/DD/YY"")"
    Worksheets("GIM2 Symbols").range("L2").AutoFill Destination:=Worksheets("GIM2 Symbols").range("L2:L2333")

    Worksheets("GIM2 Symbols").range("M2").Formula = "=IF(COUNTIF('BB File Conversion'!A:A,B2)>1, ""Duplicated!"",""OK"")"
    Worksheets("GIM2 Symbols").range("M2").Interior.Color = RGB(50, 205, 50)
    Worksheets("GIM2 Symbols").range("M2").AutoFill Destination:=Worksheets("GIM2 Symbols").range("M2:M2333")

    Worksheets("GIM2 Symbols").range("N2").Formula = "=""'""&C2&""',"""
    Worksheets("GIM2 Symbols").range("N2").AutoFill Destination:=Worksheets("GIM2 Symbols").range("N2:N2333")

'Starting to loop through worksheets. Using the worksheet name, we perform tasks relative to each sheet'
For i = 1 To worksheetCount
    
   Select Case ActiveWorkbook.Worksheets(i).Name
   
    Case Is = "BB File Conversion"
    
            Worksheets("BB File Conversion").Cells.AutoFilter
            Worksheets("BB File Conversion").range("A3" & ":AC3", "A2333" & ":AC2333").Clear
            
            Worksheets("BB File Conversion").range("A2").Formula = "=IF(B2=""Ticker Symbol Change"",$T2,IF(B2 = ""Splits"",$C2,IF(B2 = ""Option Conversions"",$C2,$Q2)))"
            Worksheets("BB File Conversion").range("A2").AutoFill Destination:=Worksheets("BB File Conversion").range("A2:A2333")

            Worksheets("BB File Conversion").range("Q2").Formula = "=IF(B2=""Ticker Symbol Change"",MID(H2,14,LEN(H2)-7),IF(MID(C2,LEN(C2)-3,4)=""Corp"",MID(C2,1,LEN(C2)-5),IF(MID(C2,LEN(C2)-3,4)=""Muni"",MID(C2,1,LEN(C2)-5),IF(MID(C2,LEN(C2)-9,10)="" US Equity"",MID(C2,1,LEN(C2)-10),IF(MID(C2,LEN(C2)-5,6)=""Equity"",MID(C2,1,LEN(C2)-7),IF(MID(C2,LEN(C2)-2,3)=""Pfd"",MID(C2,1,LEN(C2)-4)))))))"
            Worksheets("BB File Conversion").range("Q2").AutoFill Destination:=Worksheets("BB File Conversion").range("Q2:Q2333")

            Worksheets("BB File Conversion").range("R2").Formula = "=RIGHT(H2,LEN(H2)-13)"
            Worksheets("BB File Conversion").range("R2").AutoFill Destination:=Worksheets("BB File Conversion").range("R2:R2333")

            Worksheets("BB File Conversion").range("S2").Formula = "=RIGHT(Q2,2)"
            Worksheets("BB File Conversion").range("S2").AutoFill Destination:=Worksheets("BB File Conversion").range("S2:S2333")

            Worksheets("BB File Conversion").range("T2").Formula = "=IF(S2=""US"",LEFT(Q2,LEN(Q2)-3),Q2)"
            Worksheets("BB File Conversion").range("T2").AutoFill Destination:=Worksheets("BB File Conversion").range("T2:T2333")

                       
    Case Is = "Preferreds"

    Worksheets("Preferreds").range("A2:Q2333").ClearContents
    
    Case Is = "Holdings"
    
            Worksheets("Holdings").range("A2:L2", "A2333:L2333").Clear
            
    Case Is = "Exceptions"
    
    'If needed, place Exceptions code in here'
    
        End Select
        
    Next i
     
End Sub
Private Function findingNextGreenBottomRow(CAWorkbook As Workbook, sheetName As String, headerToDisregard As Integer)

Dim greenBarCounter As Integer

greenBarCounter = 0

For findingGreenBar = 1 To 2333
    
    If greenBarCounter = 2 Then Exit For
    
    If CAWorkbook.Sheets(sheetName).range("A" & findingGreenBar).Interior.Color = RGB(146, 208, 80) Then
        
        greenBarCounter = greenBarCounter + 1

    End If

Next findingGreenBar

    For nextEmptyCell = findingGreenBar To 2333
    
        If CAWorkbook.Sheets(sheetName).range("A" & nextEmptyCell).Value = "" Then Exit For
        
    Next nextEmptyCell
    
    findingNextGreenBottomRow = nextEmptyCell
       
End Function

Private Function findingNextBottomRow(CAWorkbook As Workbook, sheetName As String, headerToDisregard As Integer)
    
    For findingBlackBar = 5 To 2333
        If CAWorkbook.Sheets(sheetName).range("A" & findingBlackBar).Interior.Color = RGB(0, 0, 0) Then Exit For
    Next findingBlackBar
    
    For searchingForBottom = headerToDisregard + findingBlackBar To 2333
    
        If CAWorkbook.Sheets(sheetName).range("A" & searchingForBottom + 1).Interior.Color = RGB(146, 208, 80) Then
            CAWorkbook.Sheets(sheetName).range("A" & searchingForBottom & ":L" & searchingForBottom).Insert Shift:=xlShiftDown
        End If
    
        If CAWorkbook.Sheets(sheetName).range("A" & searchingForBottom).Value = "" Then Exit For
 
    Next searchingForBottom
        
        findingNextBottomRow = searchingForBottom

End Function

Private Sub findDupeValues(symbol As String, BBWorkbook As Workbook, corporateAction As String, LocationOfTheCell As Integer)

'Create the new tab'
Dim counter As Integer
Dim temporaryWorksheet As Worksheet
Set temporaryWorksheet = BBWorkbook.Sheets.Add(After:=BBWorkbook.Sheets(BBWorkbook.Sheets.Count))

temporaryWorksheet.Name = "TEMP"

counter = 1

For i = 2 To 2333

If IsError(BBWorkbook.Sheets("BB File Conversion").range("A" & i).Value) = False Then

If BBWorkbook.Sheets("BB File Conversion").range("A" & i).Value = "" Then Exit For

'create a new tab, paste the values into the tab, remove duplicates, pull the uniqe values that are not equal to the tickerCorporateAction string, paste them into the GIM2 Conversion, then delete the tab'

If BBWorkbook.Sheets("BB File Conversion").range("A" & i).Value = symbol Then

 BBWorkbook.Sheets("TEMP").range("A" & counter).Value = BBWorkbook.Sheets("BB File Conversion").range("B" & i).Value

    counter = counter + 1
    
        End If
    End If
Next i

'Remove all duplicates so that we can compare the corporate actions easier'
BBWorkbook.Sheets("TEMP").range("A1:A255").RemoveDuplicates Columns:=Array(1), header:=xlNo

'Compare the corporate actions against the ones we have and find the ones that are not equal '
For startLoop = 2 To 2333
    If BBWorkbook.Sheets("TEMP").range("A" & startLoop).Value = "" Then Exit For
        If BBWorkbook.Sheets("TEMP").range("A" & startLoop).Value <> corporateAction Then
            If BBWorkbook.Sheets("GIM2 Symbols").range("O" & LocationOfTheCell).Value = "" Then
                BBWorkbook.Sheets("GIM2 Symbols").range("O" & LocationOfTheCell).Value = BBWorkbook.Sheets("TEMP").range("A" & startLoop).Value
                    
                    Else
                
                BBWorkbook.Sheets("GIM2 Symbols").range("P" & LocationOfTheCell).Value = BBWorkbook.Sheets("TEMP").range("A" & startLoop).Value
            
            End If
            
    End If

Next startLoop

'We delete the tab because we're finished with it for this symbol'

Application.DisplayAlerts = False
BBWorkbook.Sheets("TEMP").Delete
Application.DisplayAlerts = True

          
End Sub

Private Sub moveDupeValues(BBWorkbook As Workbook, CAWorkbook As Workbook)

'Start looping through the holdings tab to check which values to move. If an N/A is thrown it is assumed that there is no dupe that needs to be moved'

    For i = 2 To 2333
    
    If IsError(BBWorkbook.Sheets("Holdings").range("A" & i).Value) = False Then
        
        If BBWorkbook.Sheets("Holdings").range("A" & i).Value = "" Then Exit For
            
            If IsError(BBWorkbook.Sheets("Holdings").range("O" & i).Value) = False Then
                
                'There is a value inside of dupe 1 we need to move'
                
                    passHoldingsValuesToCA BBWorkbook.Sheets("Holdings").range("O" & i).Value, CInt(i), CAWorkbook, BBWorkbook
                
                End If
                
                If IsError(BBWorkbook.Sheets("Holdings").range("P" & i).Value) = False Then
                
                'There is a value inside of dupe 2 we need to move'

                    passHoldingsValuesToCA BBWorkbook.Sheets("Holdings").range("P" & i).Value, CInt(i), CAWorkbook, BBWorkbook
                    
                    End If
                
                If IsError(BBWorkbook.Sheets("Holdings").range("Q" & i).Value) = False Then
                    
                    ' There is a value inside of dupe 3 that we need to move'

                    passHoldingsValuesToCA BBWorkbook.Sheets("Holdings").range("Q" & i).Value, CInt(i), CAWorkbook, BBWorkbook
                
                End If
                
        End If
    
    Next i
           
End Sub



Private Function findNextEmptyRowForTemplate(templateWorkbook As Workbook)


    For i = 1 To 2333
            
            If templateWorkbook.Sheets(1).range("A" & i).Value = "" Then Exit For
        
    Next i
    
     findNextEmptyRowForTemplate = i
    
End Function
Private Function extractNumbers(s As String) As Double
    
  ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Or Mid(s, i, 1) = "." Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.'
    If retval = "" Then extractNumbers = 0 Else extractNumbers = retval
    


End Function


Private Function convertAndDivideNumber(s As String) As Double

    Dim i As Integer        ' Counter for character position. '
    Dim leftSide As String
    Dim rightSide As String
    Dim foundLeftSideValues As Boolean
    Dim foundRightSideValues As Boolean
    
    ' Initialise return string to empty '
    
    leftSide = ""
    rightSide = ""
    
    foundLeftSideValues = False
    foundRightSideValues = False

'Find values on the left side'
If s <> "" Then

    For i = 1 To Len(s)
        
        If foundLeftSideValues = True Then Exit For
        
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
        
        'Start loop for left side'
        
        leftSide = leftSide + Mid(s, i, 1)
        
        For leftSideLoop = i + 1 To Len(s)
        
        If Mid(s, leftSideLoop, 1) >= "0" And Mid(s, leftSideLoop, 1) <= "9" Or Mid(s, leftSideLoop, 1) = "." Then
                    
                    leftSide = leftSide + Mid(s, leftSideLoop, 1)
                    
                    Else
                        
                        foundLeftSideValues = True
                        
                            Exit For
                    
                        End If
    
                Next leftSideLoop
    
            End If
    
        Next i
        
End If

        
 If s <> "" Then
 
    For f = leftSideLoop To Len(s)
        
        If foundRightSideValues = True Then Exit For
        
        If Mid(s, f, 1) >= "0" And Mid(s, f, 1) <= "9" Then
            
            rightSide = rightSide + Mid(s, f, 1)
                
                For rightSideLoop = f + 1 To Len(s)
                
                        If rightSideLoop = Len(s) Then foundRightSideValues = True
                        
                     If Mid(s, rightSideLoop, 1) >= "0" And Mid(s, rightSideLoop, 1) <= "9" Or Mid(s, rightSideLoop, 1) = "." Then
                                                
                        rightSide = rightSide + Mid(s, rightSideLoop, 1)
                        
                            Else
                            
                            foundRightSideValues = True
                            
                            Exit For

                     End If

            Next rightSideLoop
  
                End If
    
    Next f

End If

        If s <> "" Then

        convertAndDivideNumber = CDbl(rightSide) / CDbl(leftSide)

        Else

        convertAndDivideNumber = 0
        
        End If
        
End Function

Private Function seperateString(s As String) As String

Dim resultingString As String
resultingString = ""

For i = 1 To Len(s)
    If Mid(s, i, 1) = ":" Then Exit For
Next i
         
For f = i + 1 To Len(s)
    resultingString = resultingString + Mid(s, f, 1)
Next f

If InStrRev(resultingString, "US") <> 0 Or InStrRev(resultingString, "Equity") <> 0 Then

resultingString = Split(resultingString, "US")(0)

If InStrRev(resultingString, "Equity") <> 0 Then

resultingString = Split(resultingString, "Equity")(0)

    End If

End If

seperateString = resultingString

End Function

Sub generateEmail()

Dim fso As Scripting.FileSystemObject
Dim fol As Scripting.Folder
Dim fdr As Scripting.Folder
Dim fil As Scripting.File
Dim flc As Scripting.Folders
Dim subFolder As Scripting.Folder

Dim CAPostedAndPendingString As String
Dim CAWorkbookString As String

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)


With fd
.AllowMultiSelect = False
.Title = "Select folder"
.Filters.Clear
.InitialFileName = ""
    If .Show = True Then
    
    corporateActionLocation = .SelectedItems(1)
    
    End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fol = fso.GetFolder(corporateActionLocation)
    Set flc = fol.SubFolders
    
    Set outlookEmail = CreateObject("Outlook.Application")
    Set Email = outlookEmail.createItem(olMailItem)
    


For Each fil In fol.Files

        If InStr(1, fil.Name, "CA Posted and Pending") = 1 Then
            
            CAPostedAndPendingString = fil.Path
                
            End If
            
            If InStr(1, fil.Name, "CA - ") = 1 Then

            CAWorkbookString = fil.Path
    
    End If

    Next fil


Dim CAPostedAndPendingWorkbook As Workbook
Dim CAWorkbook As Workbook
Dim startValue As Integer
Dim rangeToCopy As range

Set CAWorkbook = Workbooks.Open(CAWorkbookString)
Set CAPostedAndPendingWorkbook = Workbooks.Open(CAPostedAndPendingString)

recipent = ""
Subject = Format(Date, "MM/DD/YYYY")

'We are clearing out the CA Posted and Pending file'

CAPostedAndPendingWorkbook.Sheets("Posted").range("A2:J383").ClearContents
CAPostedAndPendingWorkbook.Sheets("Pending").range("A2:J70").ClearContents

'Autofil the formulas back in '

CAPostedAndPendingWorkbook.Sheets("Pending").range("C2").Formula = "=VLOOKUP(LEFT(A2,4),Coverage!A:D,3,FALSE)"
CAPostedAndPendingWorkbook.Sheets("Posted").range("C2").Formula = "=VLOOKUP(LEFT(A2,4),Coverage!A:D,3,FALSE)"

'Any tab that has bloom status in it, then it can have a pending inside of it'
For tabLoop = 1 To CAWorkbook.Worksheets.Count

Dim Merger As String
Dim Conversion As String
Dim rights As String
Dim nameChange As String
Dim spinOff As String

Merger = "Merger"
Conversion = "Conversion"
rights = "Rights Issue"
nameChange = "Name Change"
spinOff = "Spin-off"

'This loops and finds all the values within the Bloom Status column that contain either (Posted) or (Pending)'

Select Case CAWorkbook.Sheets(tabLoop).Name

            Case Is = "Acquisitions"
                Dim acquisitionsGreenBar As Integer
                    acquisitionsGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                acquisitionsGreenBar = acquisitionsGreenBar + 1
                                If acquisitionsGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                    Debug.Print ("Stopping at " & i)
                                End If
                            End If
                            If acquisitionsGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Merger, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Merger, False
                                End If
                             ElseIf acquisitionsGreenBar = 1 Then
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Merger, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Merger, True
                              End If
                            End If
                        Next i
                        
            Case Is = "Ticker Symbol Changes"
                Dim tickerSymbolGreenBar As Integer
                    tickerSymbolGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                tickerSymbolGreenBar = tickerSymbolGreenBar + 1
                                If tickerSymbolGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If tickerSymbolGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", nameChange, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", nameChange, False
                                End If
                             ElseIf tickerSymbolGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", nameChange, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", nameChange, True
                              End If
                            End If
                        Next i
                                        
            
            Case Is = "ID Number Change"
                Dim idNumberGreenBar As Integer
                    idNumberGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                idNumberGreenBar = idNumberGreenBar + 1
                                If idNumberGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If idNumberGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", nameChange, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", nameChange, False
                                End If
                             ElseIf idNumberGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", nameChange, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", nameChange, True
                              End If
                            End If
                        Next i


            Case Is = "Spinoffs"
                Dim spinOffsGreenBar As Integer
                    spinOffsGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                spinOffsGreenBar = spinOffsGreenBar + 1
                                If spinOffsGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If spinOffsGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", spinOff, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", spinOff, False
                                End If
                             ElseIf spinOffsGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", spinOff, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", spinOff, True
                              End If
                            End If
                        Next i

                             
            Case Is = "Rights"
                Dim rightsGreenBar As Integer
                    rightsGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                rightsGreenBar = rightsGreenBar + 1
                                If rightsGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If rightsGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", rights, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", rights, False
                                End If
                             ElseIf rightsGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", rights, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", rights, True
                              End If
                            End If
                        Next i
 
            Case Is = "Mergers"
                Dim mergersGreenBar As Integer
                    mergersGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                mergersGreenBar = mergersGreenBar + 1
                                If mergersGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If mergersGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Merger, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Merger, False
                                End If
                             ElseIf mergersGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Merger, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Merger, True
                              End If
                            End If
                        Next i

            Case Is = "DR Sink Fund"
                Dim drSinkGreenBar As Integer
                    drSinkGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                drSinkGreenBar = drSinkGreenBar + 1
                                If drSinkGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If drSinkGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, False
                                End If
                             ElseIf drSinkGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, True
                              End If
                            End If
                        Next i
                        
            Case Is = "DR Put"
                Dim drPutGreenBar As Integer
                    drPutGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                drPutGreenBar = drPutGreenBar + 1
                                If drPutGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If drPutGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, False
                                End If
                             ElseIf drPutGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, True
                              End If
                            End If
                        Next i
                        
            Case Is = "DR Call"
                Dim drCallGreenBar As Integer
                    drCallGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                drCallGreenBar = drCallGreenBar + 1
                                If drCallGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If drCallGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, False
                                End If
                             ElseIf drCallGreenBar = 0 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, True
                              End If
                            End If
                        Next i
                        
            Case Is = "Muni Refund"
                Dim muniRefundGreenBar As Integer
                    muniRefundGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                muniRefundGreenBar = muniRefundGreenBar + 1
                                If muniRefundGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If muniRefundGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, False
                                End If
                             ElseIf muniRefundGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, True
                              End If
                            End If
                        Next i
                        
            Case Is = "Option Conversions"
                Dim optionsGreenBar As Integer
                    optionsGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                optionsGreenBar = optionsGreenBar + 1
                                If optionsGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If optionsGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, False
                                End If
                             ElseIf optionsGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, True
                              End If
                            End If
                        Next i
                        
            Case Is = "Reclassification of Shares"
                Dim reclassGreenBar As Integer
                    reclassGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                reclassGreenBar = reclassGreenBar + 1
                                If reclassGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If reclassGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, False
                                End If
                             ElseIf reclassGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, True
                              End If
                            End If
                        Next i
                        
            Case Is = "Exchange Offer"
                Dim exchangeOfferGreenBar As Integer
                    exchangeOfferGreenBar = 0
                        For i = 4 To 2333
                            If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Interior.Color = RGB(146, 208, 80) Then
                                exchangeOfferGreenBar = exchangeOfferGreenBar + 1
                                If exchangeOfferGreenBar = 1 Then
                                    If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i - 1).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value = "" And CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i + 1).Value = "" Then Exit For
                                End If
                            End If
                            If exchangeOfferGreenBar = 0 Then
                                If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                    moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, False
                                 ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                 moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, False
                                End If
                             ElseIf exchangeOfferGreenBar = 1 Then
                            'We start looking at the bottom values now'
                              If CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Pending" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "pending" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Pending", Conversion, True
                              ElseIf CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "Posted" Or CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("I" & i).Value = "posted" Then
                                moveToCAPendingAndPosted CAWorkbook, CAPostedAndPendingWorkbook, CAWorkbook.Sheets(tabLoop).Name, CInt(i), CAWorkbook.Sheets(CAWorkbook.Sheets(tabLoop).Name).range("A" & i).Value, "Posted", Conversion, True
                              End If
                            End If
                        Next i
                
End Select

Next tabLoop

For lastValueInPending = 2 To 2333

    If CAPostedAndPendingWorkbook.Worksheets("Pending").range("A" & lastValueInPending).Value = "" Then Exit For
    
Next lastValueInPending

For lastValueInPosted = 2 To 2333

    If CAPostedAndPendingWorkbook.Worksheets("Posted").range("A" & lastValueInPosted).Value = "" Then Exit For
    
Next lastValueInPosted

CAPostedAndPendingWorkbook.Worksheets("Pending").range("C2").AutoFill Destination:=CAPostedAndPendingWorkbook.Worksheets("Pending").range("C2" & ":C" & lastValueInPending - 1)
CAPostedAndPendingWorkbook.Worksheets("Posted").range("C2").AutoFill Destination:=CAPostedAndPendingWorkbook.Worksheets("Posted").range("C2" & ":C" & lastValueInPosted - 1)

Set rangeToCopy = CAPostedAndPendingWorkbook.Sheets(1).range("A1" & ":L" & lastValueInPosted).SpecialCells(xlCellTypeVisible)

With Email
        .To = recipent
        .Subject = Subject
        .HTMLBody = "POSTED " & RangetoHTML(rangeToCopy) & "PENDING" & RangetoHTML(CAPostedAndPendingWorkbook.Sheets(2).range("A1" & ":L" & lastValueInPending).SpecialCells(xlCellTypeVisible))
        .display
End With

'Returning a double value inside of the file'
        
        
End Sub


Private Function RangetoHTML(rng As range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function