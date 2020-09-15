# VBA-challenge
Repository for Vba Homework


'-------------------------------------------
'MY NAME IS LATYR THIAO
'AND 
'HERE IS MY WORK
'-------------------------------------------


Sub WallStreetSummary1()
    
'------------------------
'Do it for each worksheet
    For Each ws In Worksheets
'-------------------------------
    
  
    '----------------------------
    '1st summary
    '----------------------------
        'Find the last row for the Ticker column (source data)
        Dim Source_Table_LastRow As LongLong
        Source_Table_LastRowTicker = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Find the the 1 first row to be poulated (summary data)
        Dim Summary_Table_Row As LongLong
        Summary_Table_Row = 2

        'Create headers & assign header names for the summary columns
        Dim Header_Ticker As String
        ws.Range("K" & Summary_Table_Row - 1) = "Ticker"

        Dim Header_YearlyChange As String
        ws.Range("L" & Summary_Table_Row - 1) = "Yearly Change"

        Dim Header_PercentChange As String
        ws.Range("M" & Summary_Table_Row - 1) = "Percent Change"

        Dim Header_TotalStockVolume As String
        ws.Range("N" & Summary_Table_Row - 1) = " Total Stock Volume"
 
        'Define summary Columns to be retrieved
        Dim Summary_Ticker As String
    
        Dim Summary_YearlyChange As Double
        Summary_YearlyChange = 0
    
        Dim Summary_PercentChange As Double
        Summary_PercentChange = 0
    
        Dim Summary_TotalStockVolume As LongLong
        Summary_TotalStockVolume = 0
    
        Dim Summary_Greatest_Increase_Percent As Double
        Summary_Greatest_Increase_Percent = 0
    
        Dim Summary_Greatest_Decrease_Percent As Double
        Summary_Greatest_Decrease_Percent = 0
    
        Dim Summary_Greatest_Total_Volume As LongLong
        Summary_Greatest_Total_Volume = 0
        
        Dim TickerRowTracker As LongLong
        TickerRowTracker = 2
        
    
        'Create a loop to very the conditions
        For i = Summary_Table_Row To Source_Table_LastRowTicker
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
                'Set Ticker
                Summary_Ticker = ws.Cells(i, 1).Value
            
                'Calculate volume
                Summary_TotalStockVolume = Summary_TotalStockVolume + ws.Cells(i, 7).Value
            
                'Calculate Yearly Change
                Summary_YearlyChange = ((ws.Cells(i, 3).Value) - (ws.Cells(TickerRowTracker, 3).Value))
            
                'Calculate Percent Change
                If ws.Cells(i, 3) = 0 Then
                    Summary_PercentChange = 0
                Else
                    Summary_PercentChange = (Summary_YearlyChange / ws.Cells(i, 3).Value) '* 100
                End If
                    
                   
                'Summary_Ticker
              
                'Print ticker
                ws.Range("K" & Summary_Table_Row) = Summary_Ticker
            
                'Print volume
                ws.Range("N" & Summary_Table_Row) = Summary_TotalStockVolume
            
                'Print Yearly Change
                ws.Range("L" & Summary_Table_Row) = Summary_YearlyChange
            
                ' Print Percent Change
                ws.Range("M" & Summary_Table_Row) = Summary_PercentChange
             
                'Move to next row
                Summary_Table_Row = Summary_Table_Row + 1
            
                'Reset Values
                Summary_TotalStockVolume = 0
                Summary_YearlyChange = 0
                Summary_PercentChange = 0
                TickerRowTracker = i + 1
            
            
            Else
                'Add Total Stock Volume to Summary
                Summary_TotalStockVolume = Summary_TotalStockVolume + ws.Cells(i, 7).Value
            
                'Add Yearly Change
                Summary_YearlyChange = ((ws.Cells(i, 3).Value) - (ws.Cells(TickerRowTracker, 3).Value))
            
                'Add Percent Change
                If ws.Cells(i, 3) = 0 Then
                    Summary_PercentChange = 0
                
                Else
                    Summary_PercentChange = (Summary_YearlyChange / ws.Cells(i, 3).Value) '* 100
                End If
                
               

            End If
    
        
        Next i

         

    '-------------------------
    '2nd summary
    '-------------------------
        'Find the last row for the 1st summary that is used a data source for the 2nd summary
        Dim Source_Table2_LastRow As LongLong
        Source_Table2_LastRowTicker = ws.Cells(Rows.Count, "K").End(xlUp).Row

        'Find the the 1 first row to be poulated (summary data)
        Dim Summary_Table2_Row As LongLong
        Summary_Table2_Row = 2

        'Define output values
        Dim Summary_Max_Increase_Percent As Double
        Dim Summary_Max_Decrease_Percent As Double
        Dim Summary_Max_Total_Volume As LongLong
    

        ' Define headers
        ws.Range("R" & Summary_Table2_Row - 1) = "Ticker"
        ws.Range("Q" & Summary_Table2_Row) = " Greatest Increase Percent"
        ws.Range("Q" & Summary_Table2_Row + 1) = " Greatest Decrease Percent  "
        ws.Range("Q" & Summary_Table2_Row + 2) = " Greatest Total Volume"
        ws.Range("S" & Summary_Table2_Row - 1) = " Value"

        'Create a loop
        For i = Summary_Table2_Row To Source_Table2_LastRowTicker
        
            'Calculate the Minimum and Maximum for the summary
            Summary_Max_Increase_Percent = Application.WorksheetFunction.Max(ws.Range("M" & Summary_Table2_Row & ":M" & (Source_Table2_LastRowTicker)))
        
            Summary_Max_Decrease_Percent = Application.WorksheetFunction.Min(ws.Range("M" & Summary_Table2_Row & ":M" & (Source_Table2_LastRowTicker)))
            Summary_Max_Total_Volume = Application.WorksheetFunction.Max(ws.Range("N" & Summary_Table2_Row & ":N" & (Source_Table2_LastRowTicker)))
            '==> ws.Range("M" & Summary_Table2_Row & ":M" & (Source_Table2_LastRowTicker ))
            
            'Print values and %
            ws.Range("S2") = Summary_Max_Increase_Percent
            ws.Range("S3") = Summary_Max_Decrease_Percent
            ws.Range("S4") = Summary_Max_Total_Volume

        
        Next i
  
  
    '------------------------
    'Step 3
    '-------------------------

        'Find the last row for the 1st summary that is used a data source for the 2nd summary
        Dim Source_Table3_LastRow As LongLong
        Source_Table3_LastRowTicker = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
        'Find the the 1 first row to be poulated (summary data)
        Dim Summary_Table3_Row As LongLong
        Summary_Table3_Row = 2
    
        Dim Greatest_Increase_Percent_Value3 As Double
        Greatest_Increase_Percent_Value3 = ws.Range("S2")
        'MsgBox (Greatest_Increase_Percent_Value3 * 100)
    
        Dim Greatest_Decrease_Percent_Value3 As Double
        Greatest_Decrease_Percent_Value3 = ws.Range("S3")
        'MsgBox (Greatest_Decrease_Percent_Value3 * 100)
    
        Dim Greatest_Volume_Value3 As LongLong 'do not accept integer or long
        Greatest_Volume_Value3 = ws.Range("S4")
        'MsgBox (Greatest_Volume_Value)
    
        Dim SourcePercent3 As Double
        Dim SourceVolume3 As LongLong
        Dim Ticker3 As String
        
        'Ticker = ws.Range("K" & Summary_Table3_Row & ":K" & (Source_Table3_LastRowTicker))
    
        For i = Summary_Table3_Row To Source_Table3_LastRowTicker
            
                'Compare values to the range
                SourcePercent3 = ws.Range("M" & i).Value
                SourceVolume3 = ws.Range("N" & i).Value
                Ticker3 = ws.Range("K" & i).Value
              
              
                'Verify the condition #1
                If Greatest_Increase_Percent_Value3 = SourcePercent3 Then
                ws.Range("R2") = Ticker3
                End If
            
                'Verify the condition #2
                If Greatest_Decrease_Percent_Value3 = SourcePercent3 Then
                ws.Range("R3") = Ticker3
                End If
            
                'Verify the condition #3
                If Greatest_Volume_Value3 = SourceVolume3 Then
                ws.Range("R4") = Ticker3
                End If
            
            
                If ws.Range("M" & i) > 0 Then
                    ws.Range("M" & i).Interior.ColorIndex = 4
            
                ElseIf ws.Range("M" & i) < 0 Then
                    ws.Range("M" & i).Interior.ColorIndex = 3
                    
                ElseIf ws.Range("M" & i) = 0 Then
                    ws.Range("M" & i).Interior.ColorIndex = 7
                   
                
                Else
                    ws.Range("M" & i).Interior.ColorIndex = 2
                
            End If
            
             
        Next i

        '-------------------------------------------------------------
        '' Assign the right color code ==> Red if <0 and Green if > 0
        'For i = Int(Summary_Table_Row) To Int(Source_Table_LastRowTicker)
        
            'If ws.Range("M" & i) > 0 Then
                'ws.Range("M" & i).Interior.ColorIndex = 4
            
            'ElseIf ws.Range("M" & i) < 0 Then
                'ws.Range("M" & i).Interior.ColorIndex = 3
                
            'Else
                'ws.Range("M" & i).Interior.ColorIndex = 2
                
            'End If
        'Next i
        '--------------------------------------------------------------
        
        
        MsgBox ("Task complete !")
        
'------------------------
'Closing it for each worksheet
    Next ws
'------------------------

End Sub






