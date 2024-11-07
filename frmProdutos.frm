VERSION 5.00
Begin VB.Form frmProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de aprodutos"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10590
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&BUSCAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   6660
      TabIndex        =   6
      Top             =   5595
      Width           =   2625
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&GRAVAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   555
      TabIndex        =   4
      Top             =   5550
      Width           =   2625
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&EXCLUIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3645
      TabIndex        =   5
      Top             =   5565
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      Height          =   5220
      Left            =   165
      TabIndex        =   7
      Top             =   180
      Width           =   10245
      Begin VB.CommandButton cmdBuscaCodigoBarras 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   8145
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   870
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   435
         TabIndex        =   11
         Top             =   945
         Width           =   1755
      End
      Begin VB.TextBox txtPreco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   450
         TabIndex        =   3
         Top             =   3885
         Width           =   3240
      End
      Begin VB.TextBox txtDescricao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   450
         TabIndex        =   2
         Top             =   2520
         Width           =   9555
      End
      Begin VB.TextBox txtCodigoBarras 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2670
         TabIndex        =   0
         Top             =   930
         Width           =   5235
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   465
         TabIndex        =   12
         Top             =   555
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço R$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   450
         TabIndex        =   10
         Top             =   3480
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   450
         TabIndex        =   9
         Top             =   2130
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de barras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2655
         TabIndex        =   8
         Top             =   540
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'PESQUISAR PRODUTOS PELO CODIGO DE BARRAS
Private Sub cmdBuscaCodigoBarras_Click()
   If txtCodigoBarras.Text = "" Then
   MsgBox "Codigo de barras não informado!", vbInformation
   txtCodigoBarras.SetFocus
   Exit Sub
   End If
   
   Dim cmd As New ADODB.Command
   Dim rs As ADODB.Recordset 'Recordset = Conjunto de registros
   
   cmd.ActiveConnection = conexao
   
   cmd.CommandText = "SELECT * FROM produtos WHERE codigobarras=?"
   cmd.Parameters.Append cmd.CreateParameter("codigobarras", adVarChar, adParamInput, 20, txtCodigoBarras.Text) ' codigobarras
   
   cmd.Execute
   
   Set rs = cmd.Execute
   
   If rs.EOF Then 'EOF(End Of File) = Fim do arquivo
   MsgBox "Nenhumproduto encontrado", vbExclamation
   Exit Sub
   End If
   
   MsgBox "Produto encontrado com sucesso", vbInformation
   
   txtID.Text = rs("id")
   txtCodigoBarras = rs("codigobarras")
   txtDescricao = rs("descricao")
   txtPreco = rs("preco")
   
End Sub

'--------------------------'ADICIONAR PRODUTOS---------------------------------------'

Private Sub cmdGravar_Click()

   'Validação do formulário
   If txtCodigoBarras.Text = "" Then
      MsgBox "Código de barras não informado!", vbExclamation
      txtCodigoBarras.SetFocus
      Exit Sub
   End If
   
   If txtDescricao.Text = "" Then
      MsgBox "Descrição não informada!", vbExclamation
      txtDescricao.SetFocus
      Exit Sub
   End If
   
   If txtPreco.Text = "" Or Not IsNumeric(txtPreco.Text) Then
      MsgBox "Preço inválido!", vbExclamation
      txtPreco.Text = ""
      txtPreco.SetFocus
      Exit Sub
   End If
   
   If ExisteProdutoCodigoBarras() = True Then
      MsgBox "Codigo de barras já existente", vbExclamation
      txtCodigoBarras.SetFocus
      Exit Sub
   End If
   
   'Fim Validação do formulário
   
   On Error GoTo NomeDoErroAqui ' Inicia o tratamento de erros

   Dim sql As String
   Dim cmd As New ADODB.Command
    
   ' Define o comando SQL usando parâmetros
   sql = "INSERT INTO produtos (codigobarras, descricao, preco) VALUES (?, ?, ?)"
    
   'Sem o With ficaria assim
   'cmd.CommandText = "INSERT INTO produtos (codigobarras, descricao, preco) VALUES (?, ?, ?)"
   'cmd.Parameters.Append cmd.CreateParameter("codigobarras", adVarChar, adParamInput, 20, txtCodigoBarras.Text) ' codigobarras
   'cmd.Parameters.Append cmd.CreateParameter("descricao", adVarChar, adParamInput, 100, txtDescricao.Text) ' descricao
   'cmd.Parameters.Append cmd.CreateParameter("preco", adDouble, adParamInput, , CDbl(txtPreco.Text)) ' preco
   'cmd.Execute
    
   With cmd
      .ActiveConnection = conexao
      .CommandText = sql
      ' Adiciona parâmetros             '(Nome, TipoDeDado, TipoDeParametro, Tamanho, Valor)
      .Parameters.Append .CreateParameter("codigobarras", adVarChar, adParamInput, 20, txtCodigoBarras.Text) ' codigobarras
      .Parameters.Append .CreateParameter("descricao", adVarChar, adParamInput, 100, txtDescricao.Text) ' descricao
      .Parameters.Append .CreateParameter("preco", adDouble, adParamInput, , CDbl(txtPreco.Text)) ' preco
      .Execute
    End With
    
    MsgBox "Produto adicionado com sucesso!"
    
    ' Limpa os campos do formulário
    txtCodigoBarras.Text = ""
    txtDescricao.Text = ""
    txtPreco.Text = ""

    Exit Sub

NomeDoErroAqui:
    MsgBox "Erro ao adicionar produto: " & Err.Description ' Exibe mensagem de erro
      
   End Sub

   
    '---------------------------' PESQUISAR PRODUTOS (READ , LER)--------------------------------------'

   Private Function ExisteProdutoCodigoBarras() As Boolean
      Dim sql As String
      Dim cmd As New ADODB.Command
      Dim rs As ADODB.Recordset 'Recordset = Conjunto de registros
      cmd.ActiveConnection = conexao
      sql = "SELECT 1 FROM PRODUTOS WHERE codigobarras = ?"
      cmd.CommandText = sql
      cmd.Parameters.Append cmd.CreateParameter("codigobarras", adVarChar, adParamInput, 20, txtCodigoBarras.Text) ' codigobarras
      Set rs = cmd.Execute
      ExisteProdutoCodigoBarras = Not rs.EOF  'EOF(End Of File) = Fim do arquivo
   End Function

   '--------------------------------------------------------------------------------------------------'










   '---------------------------' EXCLUIR PRODUTOS--------------------------------------'
   
   Private Sub cmdExcluir_Click()

   On Error GoTo ErrorHandler ' Inicia o tratamento de erros
   
   Dim cmd As New ADODB.Command
   
      cmd.ActiveConnection = conexao
      cmd.CommandText = "DELETE FROM produtos WHERE codigobarras = ?"
      ' Adiciona parâmetro
      cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 20, txtCodigoBarras.Text) ' codigobarras
      cmd.Execute
      
   MsgBox "Deseja realmente excluir o registro?", vbYesNo + vbQuestion
  
   MsgBox "Produto excluído com sucesso!"
   Exit Sub
   
ErrorHandler:
       MsgBox "Erro ao excluir produto: " & Err.Description
End Sub

'-----------------------------------------------------------------'

   
   


   




