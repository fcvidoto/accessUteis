VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "accessUteis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'preenche uma listbox de 2 colunas do Access
Public Sub preencheListBox_2Columns(ByRef lista As ListBox, _
                                    dados As Variant, _
                                    cabecalho As String)
  Dim i As Integer
  
  lista.RowSource = cabecalho
  For i = LBound(dados) To UBound(dados)
      If IsNull(lista.RowSource) Then
          lista.RowSource = dados(i, 1) & " - " & dados(i, 2)
      Else
        
        'se o valor de row source for vazio
        If lista.RowSource = "" Then
          lista.RowSource = dados(i, 1) & ";" & dados(i, 2)
        
        Else
          lista.RowSource = lista.RowSource & ";" & dados(i, 1) & ";" & dados(i, 2)
        End If
        
      End If
  Next
  
End Sub

'preenche uma listbox de 1 coluna do Access
Public Sub preencheListBox_1Columns(ByRef lista As ListBox, _
                                    dados As Variant, _
                                    cabecalho As String)
  Dim i As Integer
  
  lista.RowSource = cabecalho
  For i = LBound(dados) To UBound(dados)
      If IsNull(lista.RowSource) Then
          lista.RowSource = dados(1)
      Else
        
        'se o valor de row source for vazio
        If lista.RowSource = "" Then
          lista.RowSource = dados(i)
        
        Else
          lista.RowSource = lista.RowSource & ";" & dados(i)
        End If
        
      End If
  Next
  
End Sub

'retorna o array dos relatorios selecionados
Public Function getListValues(lista As ListBox, coluna As Integer) As Variant

  Dim r As Long
  Dim k As Integer
  Dim listaArray() As String
  
  With lista
    For r = 0 To .ListCount - 1
      If .Selected(r) Then
        k = k + 1
        ReDim Preserve listaArray(1 To k)
        listaArray(k) = .Column(coluna, r)
        'strMsg = strMsg & ", " & .Column(1, r)
      End If
    Next r
  End With
  
  getListValues = listaArray 'return
  
End Function

