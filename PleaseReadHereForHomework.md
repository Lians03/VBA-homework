Hi Mike,

I hope this message finds you well. I wanted to reach out to you in advance regarding my VBA assignment. Unfortunately, the past couple of weeks have been extremely busy for me at work, and I've found the VBA and Python tasks to be quite challenging.

As I  spent a lot of time reviewing the in-class activities, I realized that I haven't had enough time to complete my VBA assignment thoroughly. There are several bugs in the code that I'm unable to fix, and unfortunately, I haven't been able to secure a tutoring session to seek guidance.

I want to apologize for any inconvenience this may cause and appreciate your understanding. I managed to run the first half of the code with Pankaj, and it worked on the alphabet sheet. However, I encountered difficulties when attempting to run it on the Multiple_year_stock_data sheet. Regrettably, it even caused Excel to crash. Just to be cautious, I have attached all of my code below. I hope it works on your computer.

Once again, I apologize for any trouble this may cause, and I truly appreciate your support.

Best regards,
shan

Sub yearstock():
                        
    Dim i As Long
    Dim counter As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim summary As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim ws As Worksheet
    
                        
    Set ws = ActiveSheet
                        
                        
    counter = 0
    summary = 2
                        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        
    lastrow = ws.Range("A1").End(xlDown).Row
                        
    For i = 2 To lastrow
                            
        If (ws.Cells(i, 1).value = ws.Cells(i + 1, 1).value) Then
            counter = counter + ws.Cells(i, 7).value
                    
        Else
            counter = counter + ws.Cells(i, 7).value
            ws.Cells(summary, 9).value = ws.Cells(i, 1).value
            ws.Cells(summary, 12).value = counter
                                
            counter = 0
            summary = summary + 1
        End If
                            
            If (Cells(i, 1).value = Cells(i + 1, 1).value) Then
            openprice = ws.Cells(i, 3).value
            closeprice = ws.Cells(i, 6).value
                                
            Cells(summary, 10).value = closeprice - openprice
            percent_change = (closeprice - openprice) / openprice * 100
            ws.Cells(summary, 11) = percent_change & "%"
                            
            summary = summary + 1
                                                  
        End If
            If Max(Cells(i, 11).value) = ws.Cells(2, 17) Then
                Max(Cells(i, 9).value) = ws.Cells(2, 16)
            If Min(Cells(i, 11).value) = ws.Cells(3, 17) Then
               Min(Cells(i, 9).value) = ws.Cells(3, 16)
            If Max(Cells(i, 12).value) = ws.Cells(3, 17) Then
                Max(Cells(i, 9).value) = ws.Cells(3, 16)
            
            
        End If
       Next i
       
    End Sub
    
