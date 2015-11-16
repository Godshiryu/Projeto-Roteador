VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufRoteador3 
   Caption         =   "Roteador"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11220
   OleObjectBlob   =   "ufRoteador3.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufRoteador3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filaEmail(200) As Outlook.MailItem 'Fila que armazena os e-mails
Public indiceFIFO As Integer 'Indice da fila
Public listaAnalistas As Object
Public indiceAnalistas As Integer
Public listaTdsAnalistas As Object 'Lista de todos os analistas
Public listaAtvAnalistas As Object 'Lista dos analistas ativos
Public indiceGeral As Integer 'Indice das listas de todos e dos analistas ativos
Public pstBase As Outlook.MAPIFolder 'Pasta base onde est� os diret�rios que ser�o analisados
Public pstEspecifica As Outlook.MAPIFolder 'Pasta que utiliza a base e seta uma espec�fica para an�lise
Public pstInbox As Outlook.MAPIFolder 'Pasta da caixa de entrada.
Public firstCaixa As Integer 'Flag que verifica se � a primeira atualizada da caixa

Private Sub UserForm_Initialize() 'M�todo de inicializa��o, todos os processos iniciais s�o feitos aqui
    indiceFIFO = 0 'Declara o indice FIFO como 0
    firstCaixa = 0 'Declara a firstCaixa como 0, pois ainda n�o atualizou a caixa
    indiceGeral = 0 'Declara o indice Geral como 0
    Set listaTdsAnalistas = CreateObject("Scripting.Dictionary") 'Instancia��o da lista de todos os analistas
    Set listaAtvAnalistas = CreateObject("Scripting.Dictionary") 'Instancia��o da lista dos analistas ativos
    'Seta pasta base:
    Set pstBase = Application.GetNamespace("MAPI").Folders("treta").Folders("teste").Folders("Email")
    'Seta caixa de entrada
    Set pstInbox = Application.GetNamespace("MAPI").Folders("treta").Folders("teste").Folders("Email").Folders("A1")
    'Set pstInbox = Application.GetNamespace("MAPI").Folders("noc@algartech.com").Folders("Pastas").Folders("5 - RESOLU��O").Folders("Renan Gonzales")
    'Set pstInbox = Application.GetNamespace("MAPI").Folders("noc@algartech.com").Folders("Caixa de Entrada")
    
    'LEMBRETE: Fazer um m�todo aqui para adicionar funcion�rios por bloco de notas:
    listaTdsAnalistas.Add "Renan Gonzales", indiceGeral
    indiceGeral = indiceGeral + 1
    listaTdsAnalistas.Add "Ismael Fagundes", indiceGeral
    indiceGeral = indiceGeral + 1
End Sub

Private Sub btnAnalistas_Click() 'Bot�o Analistas
    ufAnalistas.Show 'Mostra janelas de Analistas
End Sub

Private Sub btnAtualizaCaixa_Click() 'Bot�o Atualizar Caixa
        Call Me.atualizaLb 'Chama a fun��o de atualiza��o de listBox
End Sub

Private Sub btnRotear_Click() 'Bot�o ROTEAR
    Dim proximo As String 'Vari�vel para auxiliar em qual � o analista que receber� o e-mail
    Dim auxiliar As String 'Vari�vel para auxiliar a reposi��o do listBox
    Dim email As Outlook.MailItem 'Vari�vel para auxiliar na passada de par�metro para a fun��o removeFila
    If Not (lbEmails.ListIndex = -1 Or lbAnalistas.ListIndex) = -1 Then 'Se falta alguma op��o do LB para selecionar pula para o else
        proximo = lbAnalistas.List(lbAnalistas.ListIndex) 'Seta o valor de pr�ximo o analista selecionado no listBox
        Set pstEspecifica = pstBase.Folders(proximo) 'Seta pasta espec�fica do pr�ximo analista;
        For Each mail In pstInbox.Items 'Percorre a caixa de entrada para verificar qual o objeto(e-mail) a ser movido
            If mail.Subject = lbEmails.List(lbEmails.ListIndex, 0) Then 'Checa se o e-mail da caixa � o selecionado
                Set email = mail 'Altera o valor de email para mail, para remover o e-mail da fila, se n�o fizer isto vai dar erro por passar por refer�ncia
                Call Me.removeFila(email) 'Chamada da fun��o para remover o e-mail da fila
                mail.Move pstEspecifica 'Move o e-mail para a pasta espec�fica da pessoa
                Exit For 'For�a a parada do for caso o e-mail j� tenha sido encontrado
            End If
        Next
        
        If lbAnalistas.ListIndex = 0 Then
            auxiliar = lbAnalistas.List(lbAnalistas.TopIndex)
            listaAtvAnalistas.Remove auxiliar
            lbAnalistas.RemoveItem (0)
            listaAtvAnalistas.Add auxiliar, indiceGeral
            indiceGeral = indiceGeral + 1
            lbAnalistas.AddItem (auxiliar)
            lbAnalistas.ListIndex = 0
        End If
        
        If lbEmails.ListCount > 0 Then 'Seleciona automaticamente o primeiro e-mail da lista ap�s roteamento, caso existir algum
            lbEmails.ListIndex = lbEmails.TopIndex
        End If
        
        If lbAnalistas.ListCount > 0 Then 'Seleciona automaticamente o primeiro e-mail da lista ap�s roteamento, caso existir algum
            lbAnalistas.ListIndex = lbAnalistas.TopIndex
        End If
        
        Call Me.atualizaLb 'Chama a fun��o de atualizar caixa de entrada
    Else
        If lbEmails.ListIndex = -1 And lbAnalistas.ListIndex Then 'Esta parte verifica o que est� faltando selecionar
            MsgBox ("Selecione um e-mail e para qual analista vai ser roteado")
        Else
            If lbEmails.ListIndex = -1 Then
                MsgBox ("Selecione um e-mail a ser roteado")
            End If
            If lbAnalistas.ListIndex = -1 Then
                MsgBox ("Selecione para qual analista vai ser roteado")
            End If
        End If
    End If
End Sub
Public Sub insereFila(oMail As Outlook.MailItem) 'Fun��o de inserir um objeto na fila
    Set filaEmail(indiceFIFO) = oMail 'Adiciona objeto na fila
    indiceFIFO = indiceFIFO + 1 'Incrementa indice da fila
End Sub
Public Sub removeFila(oMail As Outlook.MailItem) 'Fun��o de remover objeto da fila
    If Not indiceFIFO = 0 Then 'Checa se a fila � vazia
        For i = 0 To indiceFIFO - 1 'Percorre toda a fila
            If (oMail = filaEmail(i)) Then 'Checa se o objeto passado por par�metro � o objeto da itera��o atual
                For j = i To indiceFIFO - 1 'Se for, percorre com todos os objetos para tr�s
                   Set filaEmail(j) = filaEmail(j + 1)
                Next j
                indiceFIFO = indiceFIFO - 1 'Decrementa indice da fila
                Exit Sub 'Finaliza a fun��o, n�o � mais necess�rio continuar procurando
            End If
        Next i
    End If
End Sub

Public Function existeFila(oMail As Outlook.MailItem, flag As Boolean) 'Fun��o que checa se existe um objeto na fila
    For i = 0 To indiceFIFO - 1 'Percorre toda a fila
        If oMail = filaEmail(i) Then 'Checa se o e-mail passado � o da itera��o atual da fila
            flag = True 'Retorna flag como True
            Exit Function 'Se encontrou, n�o � necess�rio continuar procurando. Finaliza a fun��o.
        End If
    Next i
    flag = False 'Chega aqui apenas se n�o encontrar, ent�o retorna flag como False
End Function
Public Sub atualizaLb()
    lbEmails.Clear 'Limpa listBox de Emails.
    Dim oMail As Outlook.MailItem 'Declara item de e-mail.
    Dim flag As Boolean 'Flag usada para checar se chegou algum e-mail que n�o est� na fila.
    Dim flagaux As Boolean 'Flag usada para auxiliar na checagem de existir algum e-mail na fila que n�o est� na caixa.

    pstInbox.Items.Sort "[Recebido em]", False 'Ordena a caixa de entrada pela data recebida.
    
    'Checa se � a primeira vez que atualiza a caixa de e-mails.
    If firstCaixa = 0 Then
        For Each oMail In pstInbox.Items 'Percorre toda a caixa de e-mail
            Call Me.insereFila(oMail) 'Insere na fila todos os e-mails
        Next
        firstCaixa = 1 'Seta flag 1 para a vari�vel que verifica se � a primeira da caixa
    End If
    
    'Checa se chegou algum e-mail que n�o est� na fila.
    For Each oMail In pstInbox.Items 'Percorre toda a caixa de e-mail
        Call Me.existeFila(oMail, flag) 'Verifica se existe na fila o e-mail da iteracao
        If (flag = False) Then 'Se n�o existir, ent�o:
            Call Me.insereFila(oMail) 'Insere na fila
        End If
    Next
    'Checa se existe algum e-mail na fila que n�o est� mais na caixa.
    For i = indiceFIFO - 1 To 0 Step -1 'Percorre a fila do fim para o come�o
        flagaux = False 'Coloca como Falso esta flag
        For Each oMail In pstInbox.Items
            If (oMail = filaEmail(i)) Then
                flagaux = True
            End If
        Next
        If flagaux = False Then
            Call Me.removeFila(filaEmail(i))
        End If
    Next i
    
    'Atualiza listBox de e-mails com os itens que est�o na fila.
    If indiceFIFO > 0 Then
        For i = 0 To indiceFIFO - 1
            'lbEmails.AddItem (filaEmail(i).Subject & " - " & filaEmail(i).CreationTime)
            lbEmails.AddItem (filaEmail(i).Subject)
            lbEmails.List(lbEmails.ListCount - 1, 1) = filaEmail(i).CreationTime
        Next i
    End If
    
    'Seleciona automaticamente o primeiro item, caso exista
    If lbEmails.ListCount > 0 Then
        lbEmails.ListIndex = 0
    End If
End Sub
Public Sub atualizaAnalistas()
    lbAnalistas.Clear
    For Each it In listaAtvAnalistas
        lbAnalistas.AddItem (it)
    Next
End Sub
