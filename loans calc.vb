Option Explicit

Private Sub Workbook_Open()
With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

RefreshAll

If Sheets(2).Cells(Rows.Count, 1).End(xlUp).Row > 1 Then Sheets(2).Range("A1").CurrentRegion.Offset(1, 0).ClearContents

    Dim Msg, Style, Title, Response, errorMsg
    Msg = "Do you want to UPDATE?"
    Style = vbYesNo
    Title = "Initial or Update?"
    Response = MsgBox(Msg, Style, Title)

    If Response = vbYes Then
        updateProcess = True
        updateStudentAid
    Else
        updateProcess = False
    End If

importData
loans

With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

End Sub

Option Explicit
Global updateProcess As Boolean

Sub importData()
Dim importData$
Dim sTemp() As Variant
Dim c As Long, r As Long, i As Long

If Not updateProcess Then
    importData = "M:\Financial Aid\Auto Package\AGO\20-21\pkgData.txt"
Else
    importData = "M:\Financial Aid\Auto Package\AGO\20-21\rePkgData.txt"
End If

Sheets(2).Activate

With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With

'First we will create an Array from the pkgData text file
Workbooks.OpenText Filename:=importData, Tab:=True
    Application.DisplayAlerts = False
    'This method starts the array at 1 instead of 0
    sTemp = Range("A1:N" & Range("A" & Rows.Count).End(xlUp).Row)
ActiveWorkbook.Close
Application.DisplayAlerts = True

'Now we will begin to populate the values we need from the Array
'First 3 columns we can easily populate
Range("A2", Range("A2").Offset(UBound(sTemp, 1) - 1, 2)).Value = sTemp
'Now we have to loop through the Array to get what we want

'We're going to extract cohort
Range("D2").Activate
i = 0
For r = LBound(sTemp, 1) To UBound(sTemp, 1)
    ActiveCell.Offset(i, 0).Value = sTemp(r, 10)
    i = i + 1
Next r

'Now we'll extract package units
Range("E2").Activate
i = 0
For r = LBound(sTemp, 1) To UBound(sTemp, 1)
    For c = 11 To 14
        ActiveCell.Offset(i, c - 11).Value = sTemp(r, c)
    Next c
    i = i + 1
Next r

'Now we can delete the array to free up memory and write other formulas
Erase sTemp
FormulaLoad

With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub

Option Explicit
Sub FormulaLoad()
Dim lr As Integer
Dim halfT As Range, prog As Range, level As Range, levelUp As Range, sem As Range, semUp As Range, semUpTerm As Range, dataTable As Range
Dim dataData As Range, yrElig As Range, yrSub As Range, yrSubUp As Range, yrUnsub As Range, yrUnsubUp As Range, awards As Range, match As Range

lr = Range("A1").End(xlDown).Row

Set halfT = Sheets(2).Range("I2:I" & lr)
Set prog = Sheets(2).Range("J2:J" & lr)
Set level = Sheets(2).Range("K2:K" & lr)
Set levelUp = Sheets(2).Range("L2:L" & lr)
Set sem = Sheets(2).Range("M2:M" & lr)
Set semUp = Sheets(2).Range("N2:N" & lr)
Set semUpTerm = Sheets(2).Range("O2:O" & lr)
Set dataData = Sheets(2).Range("P2:T" & lr)
Set yrElig = Sheets(2).Range("U2:V" & lr)
Set yrSub = Sheets(2).Range("W2:W" & lr)
Set yrSubUp = Sheets(2).Range("X2:X" & lr)
Set yrUnsub = Sheets(2).Range("Y2:Y" & lr)
Set yrUnsubUp = Sheets(2).Range("Z2:Z" & lr)
Set awards = Sheets(2).Range("AA2:AA" & lr)
Set match = Sheets(2).Range("AB2:AB" & lr)
Set dataTable = Sheets(2).Range("A2:AB" & lr)

halfT.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],Table_Tablix2[[StudentID]:[FullTimeHours]],4,FALSE)+0,99)"
prog.FormulaR1C1 = "=IFERROR(LEFT(VLOOKUP(RC[-9],Table_Tablix2[[StudentID]:[GPAGrouping]],3,FALSE),1),""W"")"
level.FormulaR1C1 = "=IF(RC[-1]=""G"",""G"",VLOOKUP(RC[-7],Level,2,TRUE))"
levelUp.FormulaR1C1 = "=IF(RC[-2]=""G"",""G"",IF(RC[1]<2,RC[-1],MAX(RC[-1],VLOOKUP(RC[-8]+RC[-7],Level,2,TRUE),VLOOKUP(RC[-8]+RC[-7]+RC[-6],Level,2,TRUE),VLOOKUP(RC[-8]+RC[-7]+RC[-6]+RC[-5],Level,2,TRUE))))"
sem.FormulaR1C1 = "=COUNTIF(RC[-8]:RC[-5],"">=""&RC[-4])"
semUp.FormulaR1C1 = "=IF(RC[-3]<>RC[-2],MAX(IF(RC[-9]>0,IF(VLOOKUP(RC[-10]+RC[-9],Level,2,TRUE)=RC[-2],RC[-1]-1,RC[-1]-2),IF(VLOOKUP(RC[-10]+RC[-8],Level,2,TRUE)=RC[-2],RC[-1]-1,RC[-1]-2)),1),0)"
semUpTerm.FormulaR1C1 = "=IF(RC[-4]<>RC[-3],IF(VLOOKUP(RC[-11]+RC[-10],Level,2,TRUE)=RC[-3],2,IF(VLOOKUP(RC[-11]+RC[-10]+RC[-9],Level,2,TRUE)=RC[-3],3,4)),0)"
dataData.FormulaR1C1 = "=IFERROR(INDEX(Table_Tablix2[[D/I/W]:[AGG Remain]],MATCH(RC1,Table_Tablix2[StudentID],0),MATCH(R1C,Table_Tablix2[[#Headers],[D/I/W]:[AGG Remain]],0)),""W"")"
yrElig.FormulaR1C1 = "=IF(OR(RC16=""W"",RC[-8]=0),0,IF(AND(RC11=4,RC13=1),ROUND(MIN(VLOOKUP(RC16&RC[-10],YrGross,2,FALSE),RC20)*(RC6/24),0),MIN(VLOOKUP(RC16&RC[-10],YrGross,2,FALSE),RC20)))"
yrSub.FormulaR1C1 = "=IF(RC16=""W"",0,IF(AND(RC[-12]=4,RC[-10]=1),MIN(ROUND(VLOOKUP(RC[-12],YrSub,2,FALSE)*(RC[-17]/12),0),RC17,RC19,RC20),MIN(VLOOKUP(RC[-12],YrSub,2,FALSE),RC17,RC19,RC20))/IF(RC[-10]=1,IF(RC[-12]<5,2,1),1))"
yrSubUp.FormulaR1C1 = "=IF(RC16=""W"",0,IF(RC[-12]=RC[-13],0,MIN(VLOOKUP(RC[-12],YrSub,2,FALSE)-RC[-1],RC[-7]-RC[-1],RC[-5]-RC[-1],RC[-4]-RC[-1])))"
yrUnsub.FormulaR1C1 = "=IF(RC16=""W"",0,MIN(RC[-7]-RC[-2]-RC[-1],RC[-5]-RC[-2]-RC[-1],IF(AND(RC[-14]=4,RC[-12]=1),RC[-4]-RC[-2],RC[-4]/IF(RC[-12]=1,2,1)-RC[-2])))"
yrUnsubUp.FormulaR1C1 = "=IF(RC16=""W"",0,IF(RC[-15]=RC[-14],0,MIN(RC[-8]-RC[-3]-RC[-2]-RC[-1],RC[-6]-RC[-3]-RC[-2]-RC[-1],RC[-4]-RC[-3]-RC[-2]-RC[-1])))"
awards.FormulaR1C1 = "=SUM(SUM(COUNTIF(RC[-4],"">0""),COUNTIF(RC[-2],"">0""))*RC[-14],SUM(COUNTIF(RC[-3],"">0""),COUNTIF(RC[-1],"">0""))*RC[-13])"

Range("A1").CurrentRegion.Value = Range("A1").CurrentRegion.Value

If updateProcess Then
    'Now we make data static and clear temporary contents (units) from Match column
    match.FormulaR1C1 = "=AND(EXACT(SUMIFS(Aid_Amount,Aid_ID,RC[-27],Aid_Type,""FA SUB*""),SUM(RC[-5],RC[-4])),EXACT(SUMIFS(Aid_Amount,Aid_ID,RC[-27],Aid_Type,""FA UNSUB*""),SUM(RC[-3],RC[-2])),COUNTIFS(Aid_ID,RC[-27],Aid_Cat,""70 Stafford Loan"")=RC[-1])"
    match.Value = match.Value
    
    'Now we can delete any student who's calculated loan amounts match what is in CAMS
    Application.DisplayAlerts = False
    
    Sheets(2).AutoFilterMode = False
    Range("A1").CurrentRegion.AutoFilter
    
    If Application.WorksheetFunction.CountIf(match, True) > 1 Then
        With match
            'Filter down to those who match and delete those rows
            .AutoFilter Field:=28, Criteria1:=True
            'Delete those Rows
            .Resize(.Rows.Count).SpecialCells(xlCellTypeVisible).Rows.Delete
        End With
    Else
        'If all students match
        If Application.WorksheetFunction.CountIf(match, True) = lr - 1 Then dataTable.Rows.Delete
    End If
    
    'Now we unfilter
    With Worksheets(2)
        If .FilterMode = True Then .ShowAllData
    End With
    
    Application.DisplayAlerts = True
End If

End Sub

Option Explicit

Sub loans()
'Declare variables
Dim lr As Byte, x As Byte, y As Byte, z As Byte
Dim Awards() As Variant
Dim dimension1 As Long, dimension2 As Long, dim1 As Long, dim2 As Byte, origFee As Single, amount As Integer, sem As Byte
Dim term As String, term1 As String, term2 As String, term3 As String, term4 As String, award As String
Dim addLoans As Boolean

Sheets(2).Activate

With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With

term1 = "SU-20"
term2 = "FA-20"
term3 = "SP-21"
term4 = "SU-21"

lr = Cells(Rows.Count, 1).End(xlUp).Row

'Rows in Array
dimension1 = Application.WorksheetFunction.Sum(Range("AA2:AA" & Range("AA1").End(xlDown).Row))
'Columns in Array (ID, Term, AwardType, Amount, NetAmount, OrigFeeAmount) which is 6
dimension2 = 6

If dimension1 < 1 Then
    MsgBox "No Awards to Export"
    Exit Sub
End If

'We declared it as a variant because we want a dynamic multi dimensional array so now that we now the rows and columns we can re-dimension the array
ReDim Awards(1 To dimension1, 1 To dimension2)
dim1 = 1

'First we do CP
'I need to double loop row & column with dim1 & dim2
For x = 0 To (lr - 2)
    For y = 0 To 3
        If Range("W2").Offset(x, y) > 0 Then
            'First we determine the award based on the column we're in
            Select Case y
                Case 0
                    award = "FA SUB"
                    origFee = 1 - 0.01059
                Case 1
                    award = "FA SUB 2"
                    origFee = 1 - 0.01062
                Case 2
                    award = "FA UNSUB"
                    origFee = 1 - 0.01059
                Case 3
                    award = "FA UNSUB 2"
                    origFee = 1 - 0.01062
            End Select
                'Now we loop through the terms to calculate the awards
                For z = 5 To 8
                    addLoans = True
                    Select Case z
                        Case 5
                            term = term1
                                If y Mod 2 = 0 Then
                                    sem = Cells(x + 2, 13)
                                Else
                                    addLoans = False
                                End If
                            amount = Round(Cells(x + 2, y + 23) / sem, 0)
                        Case 6
                            term = term2
                                If y Mod 2 = 0 Then
                                    sem = Cells(x + 2, 13)
                                Else
                                    If Cells(x + 2, 15) < 3 Then
                                        sem = Cells(x + 2, 14)
                                    Else
                                        addLoans = False
                                    End If
                                End If
                            amount = Round(Cells(x + 2, y + 23) / sem, 0)
                        Case 7
                            term = term3
                                If y Mod 2 = 0 Then
                                    sem = Cells(x + 2, 13)
                                Else
                                    If Cells(x + 2, 15) < 4 Then
                                        sem = Cells(x + 2, 14)
                                    Else
                                        addLoans = False
                                    End If
                                End If
                                    If Cells(x + 2, 5) > 0 Then
                                        If Cells(x + 2, 13) = 3 Then
                                            amount = Cells(x + 2, y + 23) - Awards(dim1 - 2, 5) - Awards(dim1 - 1, 5)
                                        Else
                                            amount = Cells(x + 2, y + 23) - Awards(dim1 - 1, 5)
                                        End If
                                    Else
                                        amount = Round(Cells(x + 2, y + 23) / sem, 0)
                                    End If
                        Case 8
                            term = term4
                            'you are right here
                            If y Mod 2 = 0 Then
                                Select Case Cells(x + 2, 13)
                                    Case 1
                                        amount = Cells(x + 2, y + 23)
                                    Case 2
                                        amount = Cells(x + 2, y + 23) - Awards(dim1 - 1, 4)
                                    Case 3
                                        amount = Cells(x + 2, y + 23) - Awards(dim1 - 2, 4) - Awards(dim1 - 1, 4)
                                End Select
                            Else
                                Select Case Cells(x + 2, 14)
                                    Case 1
                                        amount = Cells(x + 2, y + 23)
                                    Case 2
                                        amount = Cells(x + 2, y + 23) - Awards(dim1 - 1, 4)
                                    Case 3
                                        amount = Cells(x + 2, y + 23) - Awards(dim1 - 2, 4) - Awards(dim1 - 1, 4)
                                End Select
                            End If
                    End Select
                        If addLoans Then
                            'Now if term enrollment is greater than HT we calculate award
                            If Cells(x + 2, z) >= Cells(x + 2, 9) Then
                                'First is the StudentID
                                Awards(dim1, 1) = Range("A" & x + 2)
                                'Then is the Term
                                Awards(dim1, 2) = term
                                'Then is the AwardType
                                Awards(dim1, 3) = award
                                'Then is the Amount
                                Awards(dim1, 4) = amount
                                'Then the NetAmount
                                Awards(dim1, 5) = Application.WorksheetFunction.RoundUp(Awards(dim1, 4) * origFee, 0)
                                'Then the OrigFeeAmount
                                Awards(dim1, 6) = Awards(dim1, 4) - Awards(dim1, 5)
                                dim1 = dim1 + 1
                            End If
                        End If
                Next z
        End If
    Next y
Next x

'Now we can export the aid to the Export Table
Worksheets(2).Range("AC2", Worksheets(2).Range("AC2").Offset(UBound(Awards, 1) - 1, UBound(Awards, 2) - 1)).Value = Awards
Erase Awards

With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub
