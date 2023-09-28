Attribute VB_Name = "Module1"
Sub stockChallenege():

    ' variable to refernce last row
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' loop through each sheet
    For Each ws In Worksheets
        ' variable to refernce last row
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        ' process actions in current worksheet
        ' ticker name
        Dim tickerName As String
        'stock volume
        Dim totalVolume As Long
        'initialize
        totalVolume = 0
        'summary row refernce
        Dim summaryRow As Long
        summaryRow = 2
    
    
    sheetName = ws.Name
    
    'display name test
    'MsgBox (sheetName)
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
    
    Next ws
    

End Sub
