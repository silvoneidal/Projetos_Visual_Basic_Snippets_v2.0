VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11130
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11130
   ScaleWidth      =   12150
   Begin VB.TextBox txtMensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   10440
      Width           =   3255
   End
   Begin VB.ListBox listSnippets 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   11085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   11400
      Top             =   10080
   End
   Begin VB.TextBox txtSnippet 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   11100
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   0
      Width           =   8655
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Begin VB.Menu mSnippet 
         Caption         =   "Snippet"
         Begin VB.Menu mAbrir 
            Caption         =   "Abrir"
         End
         Begin VB.Menu mSalvar 
            Caption         =   "Salvar"
            Enabled         =   0   'False
         End
         Begin VB.Menu mExcluir 
            Caption         =   "Excluir"
            Enabled         =   0   'False
         End
         Begin VB.Menu mRenomear 
            Caption         =   "Renomear"
         End
      End
      Begin VB.Menu mColor 
         Caption         =   "Color"
         Begin VB.Menu mBlack 
            Caption         =   "Black"
            Checked         =   -1  'True
         End
         Begin VB.Menu mWhite 
            Caption         =   "White"
         End
      End
      Begin VB.Menu mHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reference: Microsoft XML, V3.0 - usada para trabalhar com o arquivo.xml

Option Explicit

Dim Color As String
Dim filePathXml As String
Dim filePathHelp As String

Private Sub Form_Load()
   ' Titulo do formulário
   Me.Caption = App.Title & "_v" & App.Major & "." & App.Minor & " by DALÇÓQUIO AUTOMAÇÃO"
   
   ' Localização dos arquivos
   filePathXml = App.Path & "\snippets.xml"
   filePathHelp = App.Path & "\help.html"
   
   ' ToolTipText
   listSnippets.ToolTipText = "Duplo click para copiar."
   
   ' Mensagem de texto
   txtMensagem.Visible = False
   txtMensagem.Text = "Snippet copiado com sucesso..." & vbCrLf & _
                      "Use (Ctrl+V) no local desejado."
   
   ' Largura inicial do formulário
   Me.Width = 3600
   
   ' Carrega lista de snippets
   Call ListaSnippetXML
   
   ' Ordem Alfabética para lista de snippets
   Call OrdenarListBoxAlfabeticamente(listSnippets)
   
   'Recupera os valores em config.ini
   Color = ReadIniValue(App.Path & "\Config.ini", "VARIAVEIS", "Color")
   
   ' Atualiza Color do Formulário
   If Color = "Black" Then Call mBlack_Click
   If Color = "White" Then Call mWhite_Click
   
End Sub

Private Sub Timer1_Timer()
   ' Fecha texto de mensagem
   txtMensagem.Visible = False
   Timer1.Enabled = False

End Sub

Private Sub mAbrir_Click()
   If Me.Width = 3600 Then
      Me.Width = 12250 ' open
      mAbrir.Caption = "Fechar"
      mSalvar.Enabled = True
      mExcluir.Enabled = True
   Else
      Me.Width = 3600 ' close
      mAbrir.Caption = "Abrir"
      mSalvar.Enabled = False
      mExcluir.Enabled = False
   End If

End Sub

Private Sub mSalvar_Click()
   Dim snippetText As String
   Dim snippetName As String
   
   snippetName = listSnippets.List(listSnippets.ListIndex)
   
   ' Verifica se ah texto para snippet
   If txtSnippet.Text = Empty Then
      MsgBox "Digite um texto para o snippet antes de salvar.", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
      
   ' Verifica se ha snippet selecionado
   If listSnippets.SelCount > 0 Then ' ou listSnippets.ListIndex >= 0
      ' Confirmação do usuário
      Dim response As VbMsgBoxResult
      response = MsgBox("Deseja salvar no snippet: " & snippetName & " ?", vbYesNo + vbQuestion, "DALÇÓQUIO AUTOMAÇÃO")
      If response = vbYes Then
         GoTo SNIPPET_SELECT
      Else
         GoTo SNIPPET_NEW
      End If
   Else
      GoTo SNIPPET_NEW
   End If
   
SNIPPET_SELECT:
   ' Remove o snippet do arquivo.xml
   Call RemoverSnippetXML(snippetName)
   
   'Adiciona novamente o snippet ao arquivo.xml
   snippetText = txtSnippet.Text
   Call AdicionarSnippetXML(snippetName, snippetText)
   
   ' Carrega lista de snippets
   Call ListaSnippetXML
   
   ' Confirmação de que o snippet foi excluido
    MsgBox "Snippet: " & snippetName & " salvo com sucesso...", , "DALÇÓQUIO AUTOMAÇÃO"
   Exit Sub


SNIPPET_NEW:
   snippetName = InputBox("Digite um nome para o snippet:", "DALÇÓQUIO AUTOMAÇÃO")
   ' verifica se o nome do snippet já existe
   If checkName(snippetName) = True Then
      MsgBox "Nome para snippet já existente !!!", vbExclamation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
   
   ' Verifica se tem nome para o snippet
   If snippetName <> Empty Then
      snippetText = txtSnippet.Text
      
      ' Adiciona o nome do snippet à lista
      listSnippets.AddItem snippetName
      
      ' Salva o texto do snippet em arquivo.xml
      Call AdicionarSnippetXML(snippetName, snippetText)
      
      ' Carrega lista de snippets
      Call ListaSnippetXML
      
      ' Confirmação de que o snippet foi excluido
      MsgBox "Snippet: " & snippetName & " salvo com sucesso...", , "DALÇÓQUIO AUTOMAÇÃO"
   Else
      MsgBox "Nome para snippet em branco ou cancelado.", vbExclamation, "DALÇÓQUIO AUTOMAÇÃO"
   End If
      
End Sub

Private Sub mExcluir_Click()
    Dim snippetName As String
    
    ' Verifica se ah snippet selecionado
    If listSnippets.SelCount = 0 Then ' ou If listSnippets.ListIndex >= 0 Then
        MsgBox "Nenhum snippet selecionado para excluir", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        Exit Sub
    End If
    
    ' Verifica nome do snippet selecionado
    snippetName = listSnippets.List(listSnippets.ListIndex)

    ' Confirmação do usuário
    Dim response As VbMsgBoxResult
    response = MsgBox("Tem certeza de que deseja excluir o snippet: " & snippetName & " ?", vbYesNo + vbQuestion, "DALÇÓQUIO AUTOMAÇÃO")

    If response = vbYes Then
        ' Remove o snippet do ListBox
        listSnippets.RemoveItem listSnippets.ListIndex

        ' Remove o snippet do arquivo.xml
        Call RemoverSnippetXML(snippetName)

        ' Limpa o TextBox
        txtSnippet.Text = Empty
        
        ' Carrega lista de snippets
        Call ListaSnippetXML
        
        ' Confirmação de que o snippet foi excluido
        MsgBox "Snippet: " & snippetName & " excluido com sucesso...", , "DALÇÓQUIO AUTOMAÇÃO"
    End If
   
End Sub

Private Sub mRenomear_Click()
   Dim snippetTemp As String
   Dim snippetName As String
   Dim snippetText As String
   
   ' Verifica se ah snippet selecionado
    If listSnippets.SelCount = 0 Then ' ou If listSnippets.ListIndex >= 0 Then
        MsgBox "Nenhum snippet selecionado para renomear", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        Exit Sub
    End If
    
    ' Verifica nome do snippet selecionado
    snippetName = listSnippets.List(listSnippets.ListIndex)
    ' Guarda temporáriamente o nome atual do snippet
    snippetTemp = snippetName
    
    ' Mensagem para o usuário
    snippetName = InputBox("Digite um novo nome para o snippet:", "DALÇÓQUIO AUTOMAÇÃO", snippetName)
   ' verifica se o nome do snippet já existe
   If checkName(snippetName) = True Then
      MsgBox "Nome para snippet já existente !!!", vbExclamation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
   
   ' Remove o snippet do arquivo.xml
   Call RemoverSnippetXML(snippetTemp)
   
   'Adiciona novamente o snippet ao arquivo.xml
   snippetText = txtSnippet.Text
   Call AdicionarSnippetXML(snippetName, snippetText)
   
   ' Carrega lista de snippets
   Call ListaSnippetXML
   
   ' Confirmação de que o snippet foi excluido
    MsgBox "Snippet: " & snippetTemp & " para " & snippetName & " renomeado com sucesso...", , "DALÇÓQUIO AUTOMAÇÃO"

End Sub

Private Sub mBlack_Click()
   ' Color Black
   mBlack.Checked = True
   mWhite.Checked = False
   Color = "Black"
   listSnippets.BackColor = vbBlack ' cor de fundo
   listSnippets.ForeColor = vbWhite  ' cor do texto
   txtSnippet.BackColor = vbBlack ' cor de fundo
   txtSnippet.ForeColor = vbWhite  ' cor do texto
   WriteIniValue App.Path & "\Config.ini", "VARIAVEIS", "Color", Color
   
 End Sub
 
Private Sub mWhite_Click()
   ' Color White
   mWhite.Checked = True
   mBlack.Checked = False
   Color = "White"
   listSnippets.BackColor = vbWhite ' cor de fundo
   listSnippets.ForeColor = vbBlack  ' cor do texto
   txtSnippet.BackColor = vbWhite ' cor de fundo
   txtSnippet.ForeColor = vbBlack  ' cor do texto
   WriteIniValue App.Path & "\Config.ini", "VARIAVEIS", "Color", Color

End Sub

Private Sub mHelp_Click()
    ' Abre o arquivo HTML no navegador padrão
    Shell "rundll32.exe url.dll,FileProtocolHandler " & filePathHelp, vbNormalFocus
End Sub

Private Sub listSnippets_Click()
   Dim snippetName As String
   snippetName = listSnippets.List(listSnippets.ListIndex)
   
   ' Obtém o texto do snippet selecionado
   Dim snippetText As String
   Call BuscarSnippetXML(snippetName, snippetText)
   txtSnippet.Text = snippetText
   
End Sub

Private Sub listSnippets_DblClick()
   ' Verifica se snippet selecionado
   If listSnippets.SelCount = 0 Then ' ou listSnippets.ListIndex >= 0
      MsgBox "Nenhum snippet selecionado para copiar", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
      Exit Sub
   End If
      
   Dim snippetName As String
   snippetName = listSnippets.List(listSnippets.ListIndex)
   
   ' Obtém o texto do snippet selecionado
   Dim snippetText As String
   Call BuscarSnippetXML(snippetName, snippetText)
   'txtSnippet.Text = snippetText
   
   ' Copia o texto do snippet para a área de transferência
   Clipboard.Clear
   Clipboard.SetText snippetText
   
   Timer1.Enabled = True
   txtMensagem.Visible = True
   'MsgBox "O snippet foi copiado para a área de transferência (Ctrl+V para colar).", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
   
End Sub

Function checkName(itemName As String) As Boolean
   
   Dim itemExists As Boolean
   itemExists = False
   
   Dim i As Integer
   For i = 0 To listSnippets.ListCount - 1
       If listSnippets.List(i) = itemName Then
           ' Um item com o mesmo nome foi encontrado
           itemExists = True
           Exit For
       End If
   Next i
   
   If itemExists Then
       checkName = True ' já existe snippet com este nome
   Else
       checkName = False ' não existe snippet com este nome
   End If
   
End Function

Private Sub OrdenarListBoxAlfabeticamente(lstBox As ListBox)
    Dim arrItens() As String
    Dim i As Integer

    ' Armazena os itens do ListBox em um array
    ReDim arrItens(lstBox.ListCount - 1)
    For i = 0 To lstBox.ListCount - 1
        arrItens(i) = lstBox.List(i)
    Next i

    ' Ordena o array em ordem alfabética
    Call QuickSort(arrItens, 0, UBound(arrItens))

    ' Limpa o ListBox
    lstBox.Clear

    ' Adiciona os itens ordenados de volta ao ListBox
    For i = 0 To UBound(arrItens)
        lstBox.AddItem arrItens(i)
    Next i
End Sub

Private Sub QuickSort(arr() As String, left As Integer, right As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim pivot As String
    Dim Temp As String

    i = left
    j = right
    pivot = arr((left + right) \ 2)

    While i <= j
        While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Wend
        While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Wend
        If i <= j Then
            Temp = arr(i)
            arr(i) = arr(j)
            arr(j) = Temp
            i = i + 1
            j = j - 1
        End If
    Wend

    If left < j Then
        QuickSort arr, left, j
    End If
    If i < right Then
        QuickSort arr, i, right
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FUNÇÕES PARA ARQUIVO.XML
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Busca o conteúdo do snippet do arquivo.xml
Private Sub BuscarSnippetXML(ByVal snippetName As String, ByRef snippetText As String)
    Dim xmlDoc As Object ' Objeto XML
    Dim xmlRoot As Object ' Elemento raiz
    Dim xmlItems As Object ' Elementos item
    Dim xmlItem As Object ' Elemento item
    Dim i As Integer ' Índice do item
    
    ' Verifica se o arquivo XML existe
    If Dir(filePathXml) <> "" Then
        ' Carrega o arquivo XML
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
        xmlDoc.async = False
        xmlDoc.preserveWhiteSpace = True ' Preserva espaços em branco
        xmlDoc.Load filePathXml
        
        ' Obtém o elemento raiz
        Set xmlRoot = xmlDoc.documentElement
        
        ' Obtém a lista de elementos item
        Set xmlItems = xmlRoot.getElementsByTagName("Item")
        
        ' Procura o item pelo snippetName
        For i = 0 To xmlItems.length - 1
            Set xmlItem = xmlItems.Item(i)
            If xmlItem.firstChild.Text = snippetName Then
                ' Obtém o conteúdo do snippet
                Dim xmlContent As Object
                Set xmlContent = xmlItem.getElementsByTagName("Content").Item(0)
                
                ' Obtém o conteúdo do snippet com linhas e identação
                Dim lines As String
                lines = xmlContent.xml
                
                ' Remove os prefixos e sufixos do CDATA
                Const cdataPrefix As String = "<Content><![CDATA["
                Const cdataSuffix As String = "]]></Content>"
                If left(lines, Len(cdataPrefix)) = cdataPrefix And right(lines, Len(cdataSuffix)) = cdataSuffix Then
                    lines = Mid(lines, Len(cdataPrefix) + 1, Len(lines) - Len(cdataPrefix) - Len(cdataSuffix))
                End If
                
                ' Define o snippetText com o conteúdo formatado
                snippetText = lines
                Exit For
            End If
        Next i
    End If
End Sub

'Lista snippets do arquivo.xml para ListBox
Private Sub ListaSnippetXML()
    Dim xmlDoc As Object ' Objeto XML
    Dim xmlRoot As Object ' Elemento raiz
    Dim xmlItems As Object ' Elementos item
    Dim xmlItem As Object ' Elemento item
    Dim snippetName As String ' Nome do snippet
        
    ' Verifica se o arquivo XML existe
    If Dir(filePathXml) = "" Then
         ' Arquivo não existe, cria um novo arquivo vazio
         Set xmlDoc = CreateObject("MSXML2.DOMDocument")
         Set xmlRoot = xmlDoc.createElement("Snippets")
         xmlDoc.appendChild xmlRoot
         xmlDoc.save filePathXml
         xmlDoc.async = False
         ' Limpa o ListBox
         listSnippets.Clear
    Else
        ' Carrega o arquivo XML
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
        xmlDoc.async = False
        xmlDoc.Load filePathXml
        
        ' Obtém o elemento raiz
        Set xmlRoot = xmlDoc.documentElement
         
        ' Obtém a lista de elementos item
        Set xmlItems = xmlRoot.childNodes
        
        ' Limpa o ListBox
        listSnippets.Clear
         
        ' Adiciona os itens ao ListBox
        For Each xmlItem In xmlItems
            If xmlItem.nodeName = "Item" Then
                If Not xmlItem.firstChild Is Nothing Then
                    snippetName = xmlItem.firstChild.Text
                    listSnippets.AddItem snippetName
                End If
            End If
        Next xmlItem
    End If
    
    ' Ordem Alfabética para lista de snippets
   Call OrdenarListBoxAlfabeticamente(listSnippets)
    
End Sub


' Adiciona snippet ao arquivo.xml
Private Sub AdicionarSnippetXML(ByVal snippetName As String, ByVal snippetText As String)
    Dim xmlDoc As Object ' Objeto XML
    Dim xmlRoot As Object ' Elemento raiz
    Dim xmlItem As Object ' Elemento item
    Dim xmlText As Object ' Texto do item
    Dim xmlContent As Object ' Conteúdo do item
    Dim xmlContentCDATA As Object ' Seção de dados CDATA
    
    ' Cria um novo documento XML ou carrega o existente
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.preserveWhiteSpace = True
    
    ' Verifica se o arquivo XML já existe
    If Dir(filePathXml) = "" Then
        ' Se o arquivo não existe, cria o elemento raiz
        Set xmlRoot = xmlDoc.createElement("Snippets")
        xmlDoc.appendChild xmlRoot
    Else
        ' Se o arquivo existe, carrega o XML existente
        xmlDoc.Load filePathXml
        Set xmlRoot = xmlDoc.documentElement
    End If
    
    ' Cria um novo elemento item
    Set xmlItem = xmlDoc.createElement("Item")
    Set xmlText = xmlDoc.createElement("Text")
    Set xmlContent = xmlDoc.createElement("Content")
    
    ' Define o texto do item
    xmlText.Text = snippetName
    
    ' Define o conteúdo do item usando a seção de dados CDATA
    Set xmlContentCDATA = xmlDoc.createCDATASection(snippetText)
    
    ' Adiciona o elemento text e content ao elemento item
    xmlItem.appendChild xmlText
    xmlItem.appendChild xmlContent
    xmlContent.appendChild xmlContentCDATA
    
    ' Adiciona o elemento item ao elemento raiz
    xmlRoot.appendChild xmlItem
    
    ' Salva o documento XML
    xmlDoc.save filePathXml
End Sub

'Remover snippet do arquivo.xml
Private Sub RemoverSnippetXML(ByVal snippetName As String)
    Dim xmlDoc As Object ' Objeto XML
    Dim xmlRoot As Object ' Elemento raiz
    Dim xmlItems As Object ' Elementos item
    Dim xmlItem As Object ' Elemento item
    Dim i As Integer ' Índice do item a ser excluído

    ' Carrega o arquivo XML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load filePathXml

    ' Obtém o elemento raiz
    Set xmlRoot = xmlDoc.documentElement

    ' Obtém a lista de elementos item
    Set xmlItems = xmlRoot.getElementsByTagName("Item")

    ' Procura pelo item a ser excluído
    For i = 0 To xmlItems.length - 1
        Set xmlItem = xmlItems.Item(i)
        ' Verifica se o nome corresponde ao item a ser excluído
        'If xmlItem.getAttribute("Name") = snippetName Then
        If xmlItem.firstChild.Text = snippetName Then
            ' Remove o item do XML
            xmlRoot.removeChild xmlItem
            Exit For
        End If
    Next i

    ' Salva o documento XML
    xmlDoc.save filePathXml

End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BLOCO DE NOTAS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Sub txtSnippets_DblClick()
'   Dim projectPath As String
'   projectPath = App.Path
'
'   Shell "explorer.exe " & projectPath, vbNormalFocus
'
'End Sub


