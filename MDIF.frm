VERSION 5.00
Begin VB.MDIForm MDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PDV - Menu Principal"
   ClientHeight    =   5370
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8625
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuProdutos 
         Caption         =   "Produtos"
      End
      Begin VB.Menu mnuCadastrosSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuários"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuProdutos_Click()
   frmProdutos.Show
End Sub

