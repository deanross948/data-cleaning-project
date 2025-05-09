Attribute VB_Name = "Module1"
Sub FixAllEncodings()
    Dim ws As Worksheet
    Dim cell As Range
    Dim cleaned As String

    Dim specialCaps As Variant
    specialCaps = Array("sql", "sas", "aws", "spss", "r", "bi", "etl", "vba", "api", "ai")

    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim rng As Range
    Set rng = ws.Range("M1:AK" & ws.Cells(ws.Rows.Count, "M").End(xlUp).Row)

    For Each cell In rng
        If VarType(cell.Value) = vbString Then
            cleaned = Replace(cell.Value, "_", " ")
            cleaned = CapitalizeWithExceptions(cleaned, specialCaps)
            If cleaned <> cell.Value Then
                cell.Value = cleaned
            End If
        End If
    Next cell

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Text cleaned in columns M to AK.", vbInformation
End Sub

Function CapitalizeWithExceptions(text As String, exceptions As Variant) As String
    Dim words() As String
    Dim result As String
    Dim i As Integer

    words = Split(LCase(text), " ")
    For i = LBound(words) To UBound(words)
        If IsInArray(words(i), exceptions) Then
            words(i) = UCase(words(i))
        Else
            words(i) = UCase(Left(words(i), 1)) & Mid(words(i), 2)
        End If
    Next i
    CapitalizeWithExceptions = Join(words, " ")
End Function

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If val = arr(i) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function


Function FixEncoding(txt As String) As String
    Dim cleaned As String
    cleaned = txt
    cleaned = Replace(cleaned, "â€“", "–")
    cleaned = Replace(cleaned, "â€”", "—")
    cleaned = Replace(cleaned, "â€˜", "‘")
    cleaned = Replace(cleaned, "â€™", "’")
    cleaned = Replace(cleaned, "â€œ", "“")
    cleaned = Replace(cleaned, "â€?", "”")
    cleaned = Replace(cleaned, "â€¦", "…")
    cleaned = Replace(cleaned, "â€", "”")
    cleaned = Replace(cleaned, "â€¢", "•")
    cleaned = Replace(cleaned, "â„¢", "™")
    cleaned = Replace(cleaned, "â‚¬", "€")
    cleaned = Replace(cleaned, "Â", "")
    cleaned = Replace(cleaned, "Ã©", "é")
    cleaned = Replace(cleaned, "Ã¨", "è")
    cleaned = Replace(cleaned, "Ã¢", "â")
    cleaned = Replace(cleaned, "Ãª", "ê")
    cleaned = Replace(cleaned, "Ã®", "î")
    cleaned = Replace(cleaned, "Ã´", "ô")
    cleaned = Replace(cleaned, "Ã»", "û")
    cleaned = Replace(cleaned, "Ã¶", "ö")
    cleaned = Replace(cleaned, "Ã¤", "ä")
    cleaned = Replace(cleaned, "Ã¼", "ü")
    cleaned = Replace(cleaned, "Ã ", "à")
    cleaned = Replace(cleaned, "Ã¡", "á")
    cleaned = Replace(cleaned, "Ã±", "ñ")
    cleaned = Replace(cleaned, "Ã", "Ã")
    FixEncoding = cleaned
End Function

