# VBA Sum Divisible

Jednoduchý VBA program pro sčítání čísel v rozsahu A1:AN na základě jejich dělitelnosti.

## Zdrojový kód

```vba
Sub SumDivisible()
    Dim n As Long, d As Integer, i As Long, s As Double
    
    n = InputBox("Zadejte N (počet řádků):")
    d = InputBox("Zadejte dělitele (vlastnost):")
    s = 0

    For i = 1 To n
        If Cells(i, 1).Value Mod d = 0 Then
            s = s + Cells(i, 1).Value
        End If
    Next i

    MsgBox "Součet: " & s
End Sub
