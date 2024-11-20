Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean) 


    Dim Plan1 As Worksheet
    Dim Plan2 As Worksheet
    Dim somaPar As Long
    
    
    Set Plan1 = Sheets("Planilha1")
    
    If Target.Value Mod 2 = 0 Then
    
    Target.Interior.Color = vbGreen
    somaPar = Plan1.Range("B1").Value + Target.Value
    Plan1.Range("B1").Value = somaPar
    
    
    
    Else
    
    Target.Interior.Color = vbRed
    
    End If
    

    
    'ActiveCell.Interior.Color = vbYellow
    'Target.Interior.Color = vbYellow
    
End Sub

em[[VBA]]
