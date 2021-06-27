VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub OneYearStockData():
    
    For Each yearly_sheet In Worksheets
        
        Dim LastRowNum As Long
        Dim RowCounter As Long
        Dim PercentChange As Double
        Dim j As Long
        Dim Total As Double
        'Dim i As Long
        
        'Setting column headers
        yearly_sheet.Cells(1, 9).Value = "Ticker"
        yearly_sheet.Cells(1, 10).Value = "Yearly Change"
        yearly_sheet.Cells(1, 11).Value = "Percent Change"
        yearly_sheet.Cells(1, 12).Value = "Total Stock Volume"
        
        'Last row of data
        LastRowNum = yearly_sheet.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRowNum)
        RowCounter = 2
        j = 2
        Total = 0
        
        For i = 2 To LastRowNum
            
                If yearly_sheet.Cells(i + 1, 1).Value <> yearly_sheet.Cells(i, 1).Value Then
                    
                    
                    'set ticker value in column 9
                    yearly_sheet.Cells(RowCounter, 9).Value = yearly_sheet.Cells(i, 1).Value
                    'MsgBox (yearly_sheet.Cells(RowCounter, 1).Value)
                    
                    'Setting yearly change in column 10 (Close - Open)
                    yearly_sheet.Cells(RowCounter, 10).Value = yearly_sheet.Cells(i, 6).Value - yearly_sheet.Cells(j, 3).Value
                    
                        'Formatting
                        If yearly_sheet.Cells(RowCounter, 10).Value < 0 Then
                        'setting less than 0, red
                            yearly_sheet.Cells(RowCounter, 10).Interior.ColorIndex = 3
                        
                        Else
                            yearly_sheet.Cells(RowCounter, 10).Interior.ColorIndex = 4
                        
                        End If
                        
                    If yearly_sheet.Cells(j, 3).Value <> 0 Then
                        'Setting percentage change in column 11
                        PercentChange = ((yearly_sheet.Cells(i, 6).Value - yearly_sheet.Cells(j, 3).Value) / yearly_sheet.Cells(j, 3).Value)
                        yearly_sheet.Cells(RowCounter, 11).Value = Format(PercentChange, "Percent")
                    Else
                        yearly_sheet.Cells(RowCounter, 11).Value = Format(0, "Percent")
                    End If
                    
                    'Setting total stock volume in column 12
                        yearly_sheet.Cells(RowCounter, 12).Value = WorksheetFunction.Sum(yearly_sheet.Cells(j, 7), yearly_sheet.Cells(i, 7))
                         
    
                    RowCounter = RowCounter + 1
                    j = i + 1
                End If
        Next i
    Next yearly_sheet
End Sub

