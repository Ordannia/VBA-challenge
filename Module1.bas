Attribute VB_Name = "Module1"
Sub QuarterlyLoop()
'Loops Through Worksheets
'Dims
Dim ws As Worksheet
For Each ws In Worksheets

'----------Formatting-----------------------------
    'Titles Ticker column
    ws.Cells(1, 9).Value = "Ticker"
   
    'Titles Quarterly Change column
    ws.Cells(1, 10).Value = "Quarterly Change"
   
    'Titles Percentage Change column
    ws.Cells(1, 11).Value = "Percentage Change"
   
    'Titles Total Stock Volume
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Titles Ticker2
    ws.Cells(1, 16).Value = "Ticker"
    
    'Titles Value2
    ws.Cells(1, 17).Value = "Value"
   
    'Titles Greatest % Increase
    ws.Cells(2, 15).Value = "Greatest % Increase"
    
    'Titles Greatest % Decrease
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    
    'Titles Greatest Total Volume
    ws.Cells(4, 15).Value = "Greatest Total Volume"
'-------------------------------------------------
   
    'Loops Through Rows
   
'-------Dims--------------------------------------

    Dim i As Long
   
    Dim Ticker As String
   
    Dim OpenPrice As Double
    Dim ClosePrice As Double
   
    Dim TotalVolume As Double
    TotalVolume = 0
   
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
   
'-------------------------------------------------

    'Calculate Last Row
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
'-------------------------------------------------
   
    'Loop Rows
    For i = 2 To Lastrow
       
'------------If Then cell does not equal one above---------------

        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
       
        'Ticker Move
            Ticker = ws.Cells(i, 1).Value
        ws.Cells(Summary_Table_Row, 9) = Ticker
       
            'Set OpenPrice
            OpenPrice = ws.Cells(i, 3).Value
           
            'Add TotalVolume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
           
           
'------------Else If Then cell does not equal one below----------

        'ElseIf
        ElseIf ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
       
            'Set ClosePrice
            ClosePrice = ws.Cells(i, 6).Value
           
'-----------Calculate Quarterly Change----------------------------------------------------

        ws.Cells(Summary_Table_Row, 10).Value = (ClosePrice - OpenPrice)
       
'-----------Set Color to Green (Positive) or Red (Negative)-------------------------------

            If (ws.Cells(Summary_Table_Row, 10).Value > 0) Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            ElseIf (ws.Cells(Summary_Table_Row, 10).Value < 0) Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If
       
'-----------Calculate Percentage Change---------------------------------------------------

        ws.Cells(Summary_Table_Row, 11).Value = ((ClosePrice - OpenPrice) / OpenPrice)
       
            'Set Column to percentage
            ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
            
'------------TotalVolume Calculation------------------------------------------------------
           
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
       
'-----------Post to Summary Table----------------------------------------------------------

        ws.Cells(Summary_Table_Row, 12).Value = TotalVolume
        
'----------Add to line---------------------------------------------------------------------

        'Add to Summary_Table_Row
       Summary_Table_Row = Summary_Table_Row + 1
       
'---------Reset Values--------------------------------------------------
       TotalVolume = 0
       
        'Else for Total Volume
        Else
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        

       
    'End if
    End If
   
    'Next i
    Next i
   
   
    'Greatest Increase and Decrease
    Dim j As Long
    j = 2
    Dim Highest As Double
    Highest = 0
    Dim Lowest As Double
    Lowest = 0
    Dim Value2 As Double
   
   
    'New Last Row
    LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
       
    'For j
    For j = 2 To LastRow2
   
        'Ifs and Thens Highest
        If (ws.Cells(j, 11).Value > Highest) Then
            Highest = ws.Cells(j, 11).Value
       
        'Moves Ticker over
        ws.Cells(2, 16).Value = ws.Cells(j, 9)
       
        'Moves and Formats Highest Over
        ws.Cells(2, 17).Value = Highest
        ws.Cells(2, 17).NumberFormat = "0.00%"
       
        End If
       
        'Ifs and Thens Lowest
        If (ws.Cells(j, 11).Value < Lowest) Then
            Lowest = ws.Cells(j, 11).Value
       
        'Moves Ticker over
        ws.Cells(3, 16).Value = ws.Cells(j, 9)
       
        'Moves and Formats Lowest over
        ws.Cells(3, 17).Value = Lowest
        ws.Cells(3, 17).NumberFormat = "0.00%"
       
        End If
       
        'Ifs and Thens Total Volume
        If (ws.Cells(j, 12).Value > Volume2) Then
            Volume2 = ws.Cells(j, 12).Value
           
        'Moves Ticker over
        ws.Cells(4, 16) = ws.Cells(j, 9)
       
        'Moves Volume over
        ws.Cells(4, 17).Value = Volume2
       
        End If
   
    'Next j
    Next j
   
'Next Worksheet loop
Next ws

End Sub
