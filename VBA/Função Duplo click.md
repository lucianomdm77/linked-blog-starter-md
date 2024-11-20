[[vba/VBA]]


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Dim Plan1 As Worksheet
    Dim Plan2 As Worksheet
    
    Set Plan1 = Sheets("Planilha1")
    
    If Target.Value Mod 2 = 0 Then
    
    Target.Interior.Color = vbGreen
    
    Else
    
    Target.Interior.Color = vbRed
    
    End If
    

    
    'ActiveCell.Interior.Color = vbYellow
    'Target.Interior.Color = vbYellow
    
End Sub
