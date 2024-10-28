VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PBV-Login"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&CANCELAR"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   4980
      Width           =   2625
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&LOGIN"
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
      Left            =   1710
      TabIndex        =   2
      Top             =   4965
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3900
      Left            =   135
      TabIndex        =   4
      Top             =   840
      Width           =   9015
      Begin VB.TextBox txtSenha 
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
         IMEMode         =   3  'DISABLE
         Left            =   225
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2295
         Width           =   8430
      End
      Begin VB.TextBox txtUsuario 
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
         Left            =   270
         TabIndex        =   0
         Top             =   765
         Width           =   8430
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
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
         Left            =   270
         TabIndex        =   7
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usu·rio"
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
         Left            =   270
         TabIndex        =   6
         Top             =   375
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PONTO DE VENDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   210
      TabIndex        =   5
      Top             =   195
      Width           =   8865
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
   End
End Sub

Private Sub cmdLogin_Click()

   If txtUsuario.Text = "" Then
   MsgBox "Campo Usu√°rio n√£o informado!", vbExclamation
   txtUsuario.SetFocus
   Exit Sub
   End If
   
   If txtSenha.Text = "" Then
   MsgBox "Campo Senha n√£o informado!", vbExclamation
   txtSenha.SetFocus
   Exit Sub
   End If
   
   Dim rs As ADODB.Recordset
   Dim cmd As New ADODB.Command
   
   Set cmd.ActiveConnection = conexao
   cmd.CommandText = "select id from usuarios where usuario =? and senha =?"
   cmd.Parameters.Append cmd.CreateParameter("usuario", adVarChar, adParamInput, 30, txtUsuario.Text)
   cmd.Parameters.Append cmd.CreateParameter("senha", adVarChar, adParamInputOutput, 20, txtSenha.Text)
   Set rs = cmd.Execute
   
   'Set rs = conexao.Execute("select * from usuarios where usuario = '" & txtUsuario.Text & "' and senha = '" & txtSenha.Text & "'")'
   
   If rs.EOF Then
      MsgBox "Usu·rio ou senha incorretos!", vbExclamation
      Exit Sub
   End If
   
   usuarioLogado = rs("id")
   MDI.Show
   Unload Me
   
End Sub
