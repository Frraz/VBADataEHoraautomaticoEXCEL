Private Sub Worksheet_Change(ByVal Target As Range)
    Dim AlteracaoColuna As Long
    Dim MonitorarColunas As Range
    
    AlteracaoColuna = 10 ' Coluna J (ajuste conforme necessário)
    
    ' Define o intervalo de colunas monitoradas (toda a planilha por padrão)
    Set MonitorarColunas = Me.UsedRange

    ' Verifica se a célula alterada está dentro do intervalo e não é a coluna de registro
    If Not Intersect(Target, MonitorarColunas) Is Nothing And Target.Column <> AlteracaoColuna Then
        Application.EnableEvents = False
        Me.Cells(Target.Row, AlteracaoColuna).Value = Now ' Registra data e hora
        Application.EnableEvents = True
    End If
End Sub
