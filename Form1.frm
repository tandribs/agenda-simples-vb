VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNomeCad 
      Height          =   285
      Left            =   810
      TabIndex        =   17
      Top             =   2100
      Width           =   3075
   End
   Begin VB.CommandButton cmdPosterior 
      Caption         =   ">"
      Height          =   315
      Left            =   2370
      TabIndex        =   16
      Top             =   6420
      Width           =   285
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<"
      Height          =   315
      Left            =   2010
      TabIndex        =   15
      Top             =   6420
      Width           =   285
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   285
      Left            =   2340
      TabIndex        =   14
      Top             =   5850
      Width           =   675
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   5850
      Width           =   675
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   285
      Left            =   1530
      TabIndex        =   12
      Top             =   5850
      Width           =   675
   End
   Begin VB.TextBox txtRedesCad 
      Height          =   285
      Left            =   780
      TabIndex        =   11
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox txtEmailCad 
      Height          =   285
      Left            =   810
      TabIndex        =   10
      Top             =   4530
      Width           =   3105
   End
   Begin VB.TextBox txtTelCad 
      Height          =   255
      Left            =   810
      TabIndex        =   9
      Top             =   3720
      Width           =   3105
   End
   Begin VB.TextBox txtEndCad 
      Height          =   255
      Left            =   810
      TabIndex        =   8
      Top             =   2880
      Width           =   3075
   End
   Begin VB.TextBox txtIDCad 
      Height          =   285
      Left            =   3570
      TabIndex        =   7
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label lblRedesCad 
      Caption         =   "Redes Sociais:"
      Height          =   285
      Left            =   810
      TabIndex        =   6
      Top             =   4980
      Width           =   1335
   End
   Begin VB.Label lblEmailCad 
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   810
      TabIndex        =   5
      Top             =   4140
      Width           =   705
   End
   Begin VB.Label lblTelCad 
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   810
      TabIndex        =   4
      Top             =   3300
      Width           =   705
   End
   Begin VB.Label lblEndCad 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   810
      TabIndex        =   3
      Top             =   2490
      Width           =   735
   End
   Begin VB.Label lblNomeCad 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   810
      TabIndex        =   2
      Top             =   1710
      Width           =   585
   End
   Begin VB.Label lblIDCad 
      Caption         =   "ID:"
      Height          =   255
      Left            =   3180
      TabIndex        =   1
      Top             =   1680
      Width           =   285
   End
   Begin VB.Label lblMeusContatos 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "MEUS CONTATOS"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   570
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type FData
    IDCad As Integer
    NomeCad As String * 30
    EndCad As String * 50
    TelCad As String * 15
    EmailCad As String * 50
    RedesCad As String * 50
End Type

Dim Agenda As FData
Dim FileName As String
Dim IDNum As Integer
Dim FF As Integer

Private Sub txtIDCad_GotFocus()
    txtNomeCad.SetFocus
End Sub

Private Sub txtNomeCad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
    txtNomeCad.Text = UCase(txtNomeCad.Text)
    txtEndCad.SetFocus
    End If
End Sub

Private Sub txtEndCad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
    txtTelCad.SetFocus
    End If
End Sub

Private Sub txtTelCad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
    txtTelCad.Text = Format(txtTelCad.Text, "(00) 00000-0000")
    txtEmailCad.SetFocus
    End If
End Sub

Private Sub txtEmailCad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
    txtRedesCad.SetFocus
    End If
End Sub

Private Sub txtRedesCad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
    cmdSalvar.SetFocus
    End If
End Sub

Private Sub Form_Load()
    FF = FreeFile
    FileName = "Agenda" & ".txt"
    IDNum = 1
End Sub

Private Sub cmdLimpar_Click()
    txtNomeCad.Text = ""
    txtEndCad.Text = ""
    txtTelCad.Text = ""
    txtEmailCad.Text = ""
    txtRedesCad.Text = ""
    
    Open App.Path & "\" & FileName For Random Access Read Write As #FF Len = Len(Agenda)
    IDNum = (LOF(FF) / Len(Agenda)) + 1
    Close #FF
End Sub

Private Sub cmdSalvar_Click()
    If Len(txtNomeCad.Text) > 0 Then
        Open App.Path & "\" & FileName For Random Access Read Write As #FF Len = Len(Agenda)
        Agenda.IDCad = IDNum
        Agenda.NomeCad = txtNomeCad.Text
        Agenda.EndCad = txtEndCad.Text
        Agenda.TelCad = txtEndCad.Text
        Agenda.EmailCad = txtEmailCad.Text
        Agenda.RedesCad = txtRedesCad.Text
        Put #FF, IDNum, Agenda
        Close #FF
        IDNum = IDNum + 1
        MsgBox ("Registro Salvo no Arquivo."), vbInformation
        cmdLimpar_Click
    Else
        MsgBox ("Digite os dados antes de salvar."), vbInformation
    End If
    cmdSair.SetFocus
End Sub

Private Sub cmdSair_Click()
    End
End Sub

Private Sub cmdPosterior_Click()
    Dim TAM As Long
    Open App.Path & "\" & FileName For Random Access Read Write As #FF Len = Len(Agenda)
    TAM = LOF(FF) / Len(Agenda)
    If IDNum > 0 Then
        If IDNum > 0 And IDNum < TAM Then
            IDNum = IDNum + 1
        End If
        Get #FF, IDNum, Agenda
        txtIDCad.Text = Agenda.IDCad
        txtNomeCad.Text = Agenda.NomeCad
        txtEndCad.Text = Agenda.EndCad
        txtTelCad.Text = Agenda.TelCad
        txtEmailCad.Text = Agenda.EmailCad
        txtRedesCad.Text = Agenda.RedesCad
    End If
    Close #FF
End Sub

Private Sub cmdAnterior_Click()
    Open App.Path & "\" & FileName For Random Access Read Write As #FF Len = Len(Agenda)
    If IDNum > 1 Then
        IDNum = IDNum - 1
        If IDNum > 0 Then
            Get #FF, IDNum, Agenda
            txtIDCad.Text = Agenda.IDCad
            txtNomeCad.Text = Agenda.NomeCad
            txtEndCad.Text = Agenda.EndCad
            txtTelCad.Text = Agenda.TelCad
            txtEmailCad.Text = Agenda.EmailCad
            txtRedesCad.Text = Agenda.RedesCad
        End If
    End If
    Close #FF
End Sub
