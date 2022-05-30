Attribute VB_Name = "ConecBD"
Dim CncDB As New ADODB.Connection
Sub Main()
CncDB.Open "Cad", "usuario", ""
If CncDB.State = adStateOpen Then
    MsgBox "Conexão Ativa"
Else
    MsgBox "Não conectado"
End If

CncDB.Close
End Sub

End Sub



