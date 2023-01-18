Attribute VB_Name = "Module1"

Sub NickLeisenringStockMacro()
Attribute NickLeisenringStockMacro.VB_ProcData.VB_Invoke_Func = " \n14"
'For Each ws In Worksheets
     
        'Create the Labels for the Columns and a few rows
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'Current Worksheet, needed to run over multiple sheets
        Dim WorksheetName As String
      
        'Two counters to work their way through the rows
        Dim i As Long
        Dim j As Long
        j = 2
        Dim Count As Long
        Count = 2

        Dim LastA As Long
        Dim LastI As Long
        Dim PChange As Double
        Dim GrIncr As Double
        Dim GrDecr As Double
        Dim GrVol As Double
        
        'Find name of worksheet
        WorksheetName = ws.Name
          
        'Find the final row of data
        LastA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Create a loop that goes from 2nd row to the last
            For i = 2 To LastA
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                ws.Cells(Count, 9).Value = ws.Cells(i, 1).Value
    
                ws.Cells(Count, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    If ws.Cells(Count, 10).Value < 0 Then
                
                    'Sets color to red
                    ws.Cells(Count, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Sets color to green
                    ws.Cells(Count, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                    PChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    ws.Cells(Count, 11).Value = Format(PChange, "Percent")
                    
                    Else
                    
                    ws.Cells(Count, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate total volume
                ws.Cells(Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                Count = Count + 1
                j = i + 1
                
                End If
            
            Next i
            
        'Find the last set of data in the 9th column
        LastI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Create my cells for the summaries
        GrVol = ws.Cells(2, 12).Value
        GrIncr = ws.Cells(2, 11).Value
        GrDecr = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastI
            
                'Condition to find greatest total value
                If ws.Cells(i, 12).Value > GrVol Then
                GrVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GrVol = GrVol
                
                End If
                
                'Condition to find Greatest Increase
                If ws.Cells(i, 11).Value > GrIncr Then
                GrIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GrIncr = GrIncr
                
                End If
                
                'Condition to find Greatest Decrease
                If ws.Cells(i, 11).Value < GrDecr Then
                GrDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GrDecr = GrDecr
                
                End If
                
            'Fill out the summaries with our findings
            ws.Cells(2, 17).Value = Format(GrIn, "Percent")
            ws.Cells(3, 17).Value = Format(GrDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GrVol, "Scientific")
            
            Next i
            
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
