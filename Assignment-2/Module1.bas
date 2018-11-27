Attribute VB_Name = "Module1"
'*********************************************************************************************************************************************************
 'Below script will loop through all the stocks and take the following info.

  '* Yearly change from what the stock opened the year at to what the closing price was.

  '* The percent change from the what it opened the year at to what it closed.

  '* The total Volume of the stock

  '* Ticker symbol
  
  '*Have conditional formatting that will highlight positive change in green and negative change in red.
  
  '*Locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
'************************************************************************************************************************************************************


Sub stock_Analyser()

    '********* Defining all the variables used below**********************
    
    Dim ws_cnt As Integer, ws_variable As Integer
    Dim openingvalue As Double, closingvalue, percentvalu, sumvolume As Double
    Dim lastRow As Long, llastrow As Long
    Dim lastcol As Integer
    Dim tickr As String
   
    '**********************************************************************
    'Determine the number of worksheets in the workbook
    '***********************************************************************
    
    ws_cnt = ActiveWorkbook.Worksheets.Count
      
    For ws_variable = 1 To ws_cnt
    
    
        Worksheets(ws_variable).Activate
        
        With ActiveSheet
        
            'Get last row and column in the spreadsheet
            
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            'clear contents from Previous run
            
            Range("I1:P" & lastRow).ClearContents
            
            lastcol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            
            'Get the first tickr and opening value
            
            tickr = Worksheets(ws_variable).Cells(2, 1).Value
            openingvalue = Worksheets(ws_variable).Cells(2, 3).Value
            
            'initialise sum of volume =0
            sumvolume = 0
            
            Range(Cells(1, lastcol + 2), Cells(1, lastcol + 5)).Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock volume")
            '******************* Call to build the summary table*********************************
            
            
            Call PopulateValues(lastRow, openingvalue, tickr, ws_variable, lastcol)
            
            '***********Getting the greatest % increase, Greatest % decrease, Greatest total volume and the corresponding ticker*********************
            
            Call CalculateValues(lastcol)
            
           '****** Formatting **********************************
           
           Call Formatting(lastcol, ws_variable)
        End With
             
    Next ws_variable
    MsgBox "Successfully completed !!!"
    
End Sub

'****************** This procedure build the summary table with the tickr, yearly change , percentage change and the total stock volume
'*****************************************************************************************************************************************


Sub PopulateValues(lastRow As Long, openingvalue As Double, tickr As String, ws_variable As Integer, lastcol As Integer)
    Dim j As Integer
    j = 2
    For i = 2 To lastRow
        'if the tickr symbol is same sum the volume
        
        If Worksheets(ws_variable).Cells(i, 1).Value = tickr Then
        
            sumvolume = sumvolume + Worksheets(ws_variable).Cells(i, 7)
        
        Else
        
            'update the ticket in the summary table with the volume and move to the next tickr
            
            Worksheets(ws_variable).Cells(j, lastcol + 2).Value = tickr
            Worksheets(ws_variable).Cells(j, lastcol + 5).Value = sumvolume
            
            'get the closing value of the ticker by looking at the previous row
            
            closingvalue = Worksheets(ws_variable).Cells(i - 1, 6).Value
            'Calculate the yearly change
            Worksheets(ws_variable).Cells(j, lastcol + 3).Value = closingvalue - openingvalue
            
            'Calculate the percentage change.if Opening value =0 percentage change is 0
            If openingvalue = 0 Then
                percentvalu = 0
            Else
                percentvalu = (closingvalue - openingvalue) / openingvalue
            End If
        
            Worksheets(ws_variable).Cells(j, lastcol + 4) = percentvalu
        
        
            'Conditional Formatting for rise or decline
            If Worksheets(ws_variable).Cells(j, 11).Value < 0 Then
                Cells(j, 11).Interior.ColorIndex = 3
            Else
                Worksheets(ws_variable).Cells(j, 11).Interior.ColorIndex = 4
            End If
        
            'Gather the opening value of the next tickr and increment the summary row counter and rest the sumvolume
            openingvalue = Worksheets(ws_variable).Cells(i, 3).Value
            sumvolume = 0
            tickr = Worksheets(ws_variable).Cells(i, 1).Value
            j = j + 1
            sumvolume = sumvolume + Worksheets(ws_variable).Cells(i, 7)
        
        End If
    Next i
    

End Sub


Sub CalculateValues(lastcol As Integer)
    Dim llastrow As Integer
    Dim maxpercentvalue As Double, minpercentvalue As Double, maxtotalvolume As Double
    Dim maxtickr As String, mintickr As String, maxvoltckr As String
    
    llastrow = Range("K" & Rows.Count).End(xlUp).Row
    
    '*****************************************************************************************************
    'Getting the Max Percent increase ,Decrease and the max total volume and the associated tickr
    '*****************************************************************************************************
    
    maxpercentvalue = Application.WorksheetFunction.Max(Range("K" & 2 & ":K" & llastrow))
    minpercentvalue = Application.WorksheetFunction.Min(Range("K" & 2 & ":K" & llastrow))
    maxtotalvolume = Application.WorksheetFunction.Max(Range("L" & 2 & ":L" & llastrow))
    
    maxtickr = Application.WorksheetFunction.Index(Range("I" & 2 & ":I" & llastrow), Application.Match(maxpercentvalue, Range("K" & 2 & ":K" & llastrow), 0), 0)
    mintickr = Application.WorksheetFunction.Index(Range("I" & 2 & ":I" & llastrow), Application.Match(minpercentvalue, Range("K" & 2 & ":K" & llastrow), 0), 0)
    maxvoltckr = Application.WorksheetFunction.Index(Range("I" & 2 & ":I" & llastrow), Application.Match(maxtotalvolume, Range("L" & 2 & ":L" & llastrow), 0), 0)
    
    Cells(2, lastcol + 8).Value = maxtickr
    Cells(3, lastcol + 8).Value = mintickr
    Cells(4, lastcol + 8).Value = maxvoltckr
    
    Range("K" & 2 & ":K" & llastrow).NumberFormat = "0.00%"
            
    Cells(2, lastcol + 9).Value = Format(maxpercentvalue, "0.00%")
    Cells(3, lastcol + 9).Value = Format(minpercentvalue, "0.00%")
    
    Cells(4, lastcol + 9).Value = maxtotalvolume
    
End Sub

'****************** Formatting the cells per the requirement ****************************/
Sub Formatting(lastcol As Integer, ws_variable As Integer)
          Dim labelarray As Variant

            Range(Cells(1, lastcol + 8), Cells(1, lastcol + 9)).Value = Array("Ticker", "Value")
            labelarray = VBA.Array("Greatest % Increase", "Greatest % decrease", "Greatest Total volume")
            labelarray = Application.WorksheetFunction.Transpose(labelarray)
            Range("N2:N4").Value = labelarray
            ' Adjusting the Formatting section to match the layout requirement
            Worksheets(ws_variable).Columns("L").NumberFormat = "General"
            Worksheets(ws_variable).UsedRange.Columns.AutoFit
End Sub

