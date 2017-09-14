Attribute VB_Name = "modCobaia"
Public Function ENumero(ByVal valor As String) As Boolean
    Dim resultado As Boolean
    resultado = IsNumeric(valor)
    ENumero = resultado
End Function

Public Sub Teste()
    Dim valor1 As Variant, valor2 As Variant
    
    valor1 = InputBox("Digite um valor a ser somado", "Valor 1", "0")
    valor2 = InputBox("Digite um valor a ser somado", "Valor 2", "0")
    
    If Not ENumero(valor1) Or Not ENumero(valor2) Then
        MsgBox "Ambos os valores precisam ser números"
    Else
        MsgBox CInt(valor1) + CInt(valor2)
    End If
End Sub
