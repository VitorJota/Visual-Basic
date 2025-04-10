Attribute VB_Name = "fCelulasVazias"

Function VerificarCelulasVazias(sht As Worksheet, celulas As Variant) As Boolean
    
    Dim celula As Variant
    Dim celulaVazia As Boolean
    celulaVazia = False
    
    For Each celula In celulas
        If sht.Range(celula).Value = "" Then
            sht.Range(celula).Interior.Color = RGB(255, 192, 203)
            MsgBox "A célula " & celula & " está vazia."
            celulaVazia = True
        End If
        
    Next celula
    'sht.Range(celula).Interior.Color = RGB(0, 0, 0)
    VerificarCelulasVazias = celulaVazia
End Function
