VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufAnalistas 
   Caption         =   "Analistas"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   OleObjectBlob   =   "ufAnalistas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufAnalistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDireita_Click()
    If Not (lbTodos.ListIndex = -1) Then
        lbAtivos.AddItem (lbTodos.List(lbTodos.ListIndex))
        ufRoteador3.listaAtvAnalistas.Add lbTodos.List(lbTodos.ListIndex), ufRoteador3.indiceGeral
        ufRoteador3.listaTdsAnalistas.Remove lbTodos.List(lbTodos.ListIndex)
        ufRoteador3.indiceGeral = ufRoteador3.indiceGeral + 1
        lbTodos.RemoveItem (lbTodos.ListIndex)
    End If
End Sub

Private Sub btnEsquerda_Click()
    If Not (lbAtivos.ListIndex = -1) Then
        lbTodos.AddItem (lbAtivos.List(lbAtivos.ListIndex))
        ufRoteador3.listaAtvAnalistas.Remove lbAtivos.List(lbAtivos.ListIndex)
        ufRoteador3.listaTdsAnalistas.Add lbAtivos.List(lbAtivos.ListIndex), ufRoteador3.indiceGeral
        ufRoteador3.indiceGeral = ufRoteador3.indiceGeral + 1
        lbAtivos.RemoveItem (lbAtivos.ListIndex)
    End If
End Sub
Private Sub UserForm_Initialize()
    For Each it In ufRoteador3.listaTdsAnalistas.Keys()
        lbTodos.AddItem (it)
    Next
    For Each it In ufRoteador3.listaAtvAnalistas.Keys()
        lbAtivos.AddItem (it)
    Next
    If lbTodos.ListCount > 0 Then
        lbTodos.ListIndex = 0
    End If
End Sub

Private Sub UserForm_Terminate()
    Call ufRoteador3.atualizaAnalistas
    If (ufRoteador3.lbAnalistas.ListCount > 0) Then
        ufRoteador3.lbAnalistas.ListIndex = 0
    End If
End Sub
