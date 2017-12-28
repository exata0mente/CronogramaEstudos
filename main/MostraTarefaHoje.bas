Option Explicit

'Autor: Ricardo Bezerra
'Descricao: Esta funcao vasculha uma tabela formatada preenchida com dados de projetos e retornar qual as Tarefas agendadas para o dia proposto

'TO DO LIST: Ao encontrar a Tag "P", zerar os índices para iniciar o preenchimento da coluna A"n"

Sub MostraTarefasHoje2()
    
    Dim Painel As Object
    Dim ListaProjetos As Object 'Variavel que tera informacao sobre a localizacao de celulas
    Dim InicioLinhaProjetos, FimLinhaProjetos As Integer
    Dim InicioColunaProjetos, FimColunaProjetos As Integer
    Dim IndiceLinhaProjeto, IndiceColunaProjeto As Integer
    Dim IndiceLinhaPainel, IndiceColunaPainel As Integer
    Dim VerificaTag As String
    Dim Tarefa As clsTarefas
    Dim EscreveProjeto, EscreveTarefa As Boolean
                
    Set Painel = Worksheets("Painel")
    Set ListaProjetos = Worksheets("Projetos")
    Set Tarefa = New clsTarefas
    
    'Inicio da Lista de Projetos
    InicioLinhaProjetos = 3 - 1
    InicioColunaProjetos = 1
    IndiceLinhaPainel = 3
    IndiceColunaPainel = 1
    
    'Fim da Lista de Projetos
    ListaProjetos.Activate
    With ListaProjetos
        With Range("B99999")
            .Select
            .End(xlUp).Select
        End With
    End With
    
    FimLinhaProjetos = ActiveCell.Row
    
    'Painel.Activate
    
    IndiceColunaProjeto = InicioColunaProjetos
    For IndiceLinhaProjeto = InicioLinhaProjetos To FimLinhaProjetos
        VerificaTag = ListaProjetos.Cells(IndiceLinhaProjeto, InicioColunaProjetos)
        Select Case VerificaTag
        Case "P"
            'Set Tarefa = Nothing
            EscreveProjeto = False
            EscreveTarefa = False
            Tarefa.NomeProjeto = ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 1)
        Case "T"
            Tarefa.NomeTarefa = ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 1)
        Case "ST"
            Tarefa.NomeSubTarefa = ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 1)
            If ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 2) = Painel.Cells(1, 2) Then
                If Not EscreveProjeto Then
                    Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel) = Tarefa.NomeProjeto
                    EscreveProjeto = True
                    IndiceLinhaPainel = IndiceLinhaPainel + 1
                    IndiceColunaPainel = IndiceColunaPainel + 1
                End If
                If Not EscreveTarefa Then
                    Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel) = Tarefa.NomeTarefa
                    Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 2) = "Data Início"
                    Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 3) = "Data Fim"
                    Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 4) = "Faltam"
                    Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 5) = "%Concluido"
                    Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 6) = "Status"
                    EscreveTarefa = True
                    IndiceLinhaPainel = IndiceLinhaPainel + 1
                    IndiceColunaPainel = IndiceColunaPainel + 1
                End If
            Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel) = Tarefa.NomeSubTarefa
            Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 1) = ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 2)
            Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 2) = ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 3)
            Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 3) = ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 4)
            Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 4) = ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 5)
            Painel.Cells(IndiceLinhaPainel, IndiceColunaPainel + 5) = ListaProjetos.Cells(IndiceLinhaProjeto, IndiceColunaProjeto + 6)
            IndiceLinhaPainel = IndiceLinhaPainel + 1
            End If
        End Select
    Next
    Painel.Activate
    'Set Tarefa = Nothing
End Sub



