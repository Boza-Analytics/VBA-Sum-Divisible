Sub SumDivisible()
    Dim n As Long, d As Integer, i As Long
    Dim s As Double, v As Variant

    n = InputBox("N:")
    d = InputBox("Divisor:")
    s = 0

    For i = 1 To n
        v = Cells(i, 1).Value
        If IsNumeric(v) And v <> "" Then
            If v Mod d = 0 Then s = s + v
        End If
    Next i

    MsgBox "Sum: " & s
End Sub
