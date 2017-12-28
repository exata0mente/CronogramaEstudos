VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTarefas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private aNome As String
Private aDataInicio As Date
Private aDataFim As Date
Private aConcluido As Double
Private aSituacao As String
Private aNomeProjeto As String
Private aNomeTarefa As String
Private aNomeSubTarefa As String

Property Get Nome() As String
    Nome = aNome
End Property
Property Let Nome(pNome As String)
    aNome = pNome
End Property

Property Get DataInicio() As Date
    DataInicio = aDataInicio
End Property
Property Let DataInicio(pDataInicio As Date)
    aDataInicio = pDataInicio
End Property

Property Get DataFim() As Date
    DataFim = aDataFim
End Property
Property Let DataFim(pDataFim As Date)
    aDataFim = pDataFim
End Property

Property Get Concluido() As Double
    Concluido = aConcluido
End Property
Property Let Concluido(pConcluido As Double)
    aConcluido = pConcluido
End Property

Property Get Situacao() As String
    Situacao = aSituacao
End Property
Property Let Situacao(pSituacao As String)
    aSituacao = pSituacao
End Property

Property Get NomeProjeto() As String
    NomeProjeto = aNomeProjeto
End Property
Property Let NomeProjeto(pNomeProjeto As String)
    aNomeProjeto = pNomeProjeto
End Property

Property Get NomeTarefa() As String
    NomeTarefa = aNomeTarefa
End Property

Property Let NomeTarefa(pNomeTarefa As String)
    aNomeTarefa = pNomeTarefa
End Property

Property Get NomeSubTarefa() As String
    NomeSubTarefa = aNomeSubTarefa
End Property

Property Let NomeSubTarefa(pNomeSubTarefa As String)
    aNomeSubTarefa = pNomeSubTarefa
End Property

Public Sub DiasRestantes()
    Dim DiasRest As Date
    
    DiasRest = (DataFim - DataInicio)
    
End Sub
