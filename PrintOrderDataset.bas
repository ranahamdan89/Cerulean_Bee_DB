Sub Step06_Insert_PrintOrder()

'For 95% of our artworkorder_project generate between 1 and 7 printorders ...

Dim R As Integer, N As Integer, I As Integer
Dim strSQL As String
'Dim rs As Recordset

Set rs = CurrentDb.OpenRecordset("Sheet1")

Do While Not rs.EOF

    'See if this ArtworkOrder_Project is in the 95% ...
    R = Rnd
    
    If R < 0.95 Then 'Yes ...
    
        'See how many printorders to generate ...
        N = Int(7 * Rnd) + 1
        
        For I = 1 To N
        
            strSQL = "Insert Into PrintOrder( PrintDate, ApparelOrderDate, Art_FilmDate,DateDelivered,PrintOrderDate,DueDate,Art_SlideDate,SetUpCharge,Deposit,Discount,ArtworkOrderID ) "
            strSQL = strSQL & "Values(" & rs("ScheduledPrintDate") & ", "
            
            'Generate a random TripNo between lowest TripNo (should be 1) and highest (should be 101) ...
            R = Int(101 * Rnd) + 1
            
            strSQL = strSQL & R & ", "
			
			Randomize
    GetRndDate = DateAdd("d", Int((DateDiff("d", dtStartDate, dtEndDate) + 1) * Rnd), dtStartDate)
        
            'Generate a random NumberOfPeople between 1 and 12 ...
            R = Int(12 * Rnd) + 1
            
            strSQL = strSQL & R & ")"
        
            Debug.Print strSQL
            CurrentDb.Execute strSQL
        
        Next I
     
    End If
    
    rs.MoveNext
    
Loop

rs.Close

Debug.Print "***Step06_Insert_Registration***"

End Sub