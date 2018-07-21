VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Calculadora Substituição Tributária"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   827.519
   ScaleMode       =   0  'User
   ScaleWidth      =   1247
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "&SAIR"
      Height          =   615
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8760
      Width           =   2055
   End
   Begin VB.CommandButton cmdLimpar 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "&LIMPAR"
      Height          =   615
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalcular 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "&CALCULAR"
      Height          =   615
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox txtMva 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "%"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   8880
      Width           =   4335
   End
   Begin VB.TextBox txtAliquotaIcmsExterna 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "%"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   7080
      Width           =   4335
   End
   Begin VB.TextBox txtAliquotaEntradaIcms 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "%"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   720
      TabIndex        =   2
      Top             =   5280
      Width           =   4335
   End
   Begin VB.TextBox txtAliquotaIpi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "%"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   3600
      Width           =   4335
   End
   Begin VB.TextBox txtvalorProduto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """R$""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label lblValorProdSubst 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   11760
      TabIndex        =   13
      Top             =   5040
      Width           =   5895
   End
   Begin VB.Label lblBaseSubst 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   6000
      TabIndex        =   12
      Top             =   8880
      Width           =   4335
   End
   Begin VB.Label lblValorIcmsExt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   6000
      TabIndex        =   11
      Top             =   7080
      Width           =   4335
   End
   Begin VB.Label lblValorIcmsEnt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   6000
      TabIndex        =   10
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Label lblValorIpi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   6000
      TabIndex        =   9
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label lblValorSubst 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   2040
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcular_Click() 'CÁLCULOS MATEMÁTICOS'

lblValorIpi.Caption = Format(Round(txtvalorProduto.Text * (txtAliquotaIpi / 100), 2), "R$0,0.00")
lblValorIcmsEnt.Caption = Format(Round(txtvalorProduto.Text * (txtAliquotaEntradaIcms.Text / 100), 2), "R$0,0.00")
lblValorIcmsExt.Caption = Format(Round(txtvalorProduto.Text * (txtAliquotaIcmsExterna.Text / 100), 2), "R$0,0.00")
lblBaseSubst.Caption = Format(Round((((txtvalorProduto.Text + CDbl(lblValorIpi)))) * (txtMva.Text / 100) + (txtvalorProduto.Text + CDbl(lblValorIpi)), 2), "R$0,0.00")
lblValorSubst.Caption = Format(Round((CDbl(lblBaseSubst) * (txtAliquotaIcmsExterna.Text / 100)) - CDbl(lblValorIcmsEnt), 2), "R$0,0.00")
lblValorProdSubst.Caption = Format(Round(txtvalorProduto.Text + CDbl(lblValorIpi) + CDbl(lblValorSubst), 2), "R$0,0.00")

cmdSair.SetFocus

End Sub


Private Sub cmdLimpar_Click()

limpar

txtvalorProduto.SetFocus

End Sub
Sub limpar() 'FUNÇÃO PARA LIMPAR AS LABELS E TEXTBOXES'

txtvalorProduto.Text = ""
txtAliquotaIpi.Text = ""
txtAliquotaEntradaIcms.Text = ""
txtAliquotaIcmsExterna.Text = ""
txtMva.Text = ""
lblBaseSubst = ""
lblValorIcmsEnt = ""
lblValorIcmsExt = ""
lblValorIpi = ""
lblValorProdSubst = ""
lblValorSubst = ""

End Sub



Private Sub cmdSair_Click()

fechar 'CHAMA A FUNÇÃO PARA FECHAR O FORM'

End Sub

Sub fechar() 'CRIA A MSGBOX PARA CONFIRMAR A AÇÃO DO USUÁRIO'

confirma = MsgBox("Tem certeza que deseja sair?", vbYesNo, "Sair")

If confirma = vbYes Then
    Unload Me
Else
    limpar
    txtvalorProduto.SetFocus
End If

End Sub


Private Sub lblBaseSubst_Change()

lblBaseSubst.ForeColor = vbRed 'MUDAR A COR DA LETRA'

End Sub

Private Sub lblValorIcmsEnt_Change()

lblValorIcmsEnt.ForeColor = vbRed 'MUDAR A COR DA LETRA'

End Sub

Private Sub lblValorIcmsExt_Change()

lblValorIcmsExt.ForeColor = vbRed 'MUDAR A COR DA LETRA'

End Sub

Private Sub lblValorIpi_Change()

lblValorIpi.ForeColor = vbRed 'MUDAR A COR DA LETRA'

End Sub

Private Sub lblValorProdSubst_Change()

lblValorProdSubst.ForeColor = vbRed 'MUDAR A COR DA LETRA'

End Sub

Private Sub lblValorSubst_Change()

lblValorSubst.ForeColor = vbRed 'MUDAR A COR DA LETRA'

End Sub

Private Sub txtAliquotaEntradaIcms_Change()

txtAliquotaEntradaIcms.ForeColor = vbBlue 'MUDAR A COR DA LETRA'

End Sub

Private Sub txtAliquotaEntradaIcms_KeyPress(KeyAscii As Integer) 'VALIDAR SE O USUÁRIO ESTÁ DIGITANDO NÚMEROS'
If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(txtAliquotaEntradaIcms.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtAliquotaEntradaIcms_keydown(keyCode As Integer, shift As Integer) 'MUDA O FOCO DO CURSOR PARA A PRÓXIMA TEXTBOX'

If keyCode = vbKeyReturn Then
    If txtAliquotaEntradaIcms = "" Then
        MsgBox "FAVOR DIGITAR O VALOR DA ALÍQUOTA DE ICMS DE ENTRADA", vbCritical, "ALERTA"
        txtAliquotaEntradaIcms.SetFocus
    Else
        txtAliquotaIcmsExterna.SetFocus
    End If
End If

End Sub

Private Sub txtAliquotaIcmsExterna_Change()

txtAliquotaIcmsExterna.ForeColor = vbBlue

End Sub

Private Sub txtAliquotaIcmsExterna_KeyPress(KeyAscii As Integer) 'VALIDAR SE O USUÁRIO ESTÁ DIGITANDO NÚMEROS'
If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(txtAliquotaIcmsExterna.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtAliquotaIcmsExterna_keydown(keyCode As Integer, shift As Integer) 'MUDA O FOCO DO CURSOR PARA A PRÓXIMA TEXTBOX'

If keyCode = vbKeyReturn Then
    If txtAliquotaIcmsExterna = "" Then
        MsgBox "FAVOR DIGITAR O VALOR DA ALÍQUOTA DE ICMS DE SAÍDA", vbCritical, "ALERTA"
        txtAliquotaIcmsExterna.SetFocus
    Else
        txtMva.SetFocus
    End If
End If

End Sub

Private Sub txtAliquotaIpi_Change() 'MUDAR A COR DA LETRA'

txtAliquotaIpi.ForeColor = vbBlue

End Sub

Private Sub txtAliquotaIpi_KeyPress(KeyAscii As Integer) 'VALIDAR SE O USUÁRIO ESTÁ DIGITANDO NÚMEROS'
If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(txtAliquotaIpi.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtAliquotaIpi_KeyDown(keyCode As Integer, shift As Integer) 'MUDA O FOCO DO CURSOR PARA A PRÓXIMA TEXTBOX'

If keyCode = vbKeyReturn Then
    If txtAliquotaIpi = "" Then
        txtAliquotaIpi = "0"
        txtAliquotaEntradaIcms.SetFocus
    Else
        txtAliquotaEntradaIcms.SetFocus
    End If
End If

End Sub

Private Sub txtBaseCalcSubst_Change() 'MUDAR A COR DA LETRA'

txtBaseCalcSubst.ForeColor = vbRed

End Sub

Private Sub txtMva_Change() 'MUDAR A COR DA LETRA'

txtMva.ForeColor = vbBlue

End Sub

Private Sub txtMva_KeyPress(KeyAscii As Integer) 'VALIDAR SE O USUÁRIO ESTÁ DIGITANDO NÚMEROS'
If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(txtMva.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtMva_KeyDown(keyCode As Integer, shift As Integer) 'MUDA O FOCO DO CURSOR PARA A PRÓXIMA TEXTBOX'

If keyCode = vbKeyReturn Then
    If txtMva = "" Or txtMva = "0" Then
        MsgBox "FAVOR DIGITAR O VALOR DA ALÍQUOTA DE MVA", vbCritical, "ALERTA"
        txtMva = ""
    Else
        cmdCalcular.SetFocus
    End If
End If

End Sub

Private Sub txtValorIcmsEntrada_Change() 'MUDAR A COR DA LETRA'
 
txtValorIcmsEntrada.ForeColor = vbRed

End Sub

Private Sub txtValorIcmsExterna_Change() 'MUDAR A COR DA LETRA'

txtValorIcmsExterna.ForeColor = vbRed

End Sub

Private Sub txtValorIpi_Change() 'MUDAR A COR DA LETRA'

txtValorIpi.ForeColor = vbRed

End Sub

Private Sub txtvalorProduto_Change() 'MUDAR A COR DA LETRA'

txtvalorProduto.ForeColor = vbBlue

End Sub

Private Sub txtvalorProduto_KeyPress(KeyAscii As Integer) 'VALIDAR SE O USUÁRIO ESTÁ DIGITANDO NÚMEROS'
If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(txtvalorProduto.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtvalorProduto_KeyDown(keyCode As Integer, shift As Integer) 'MUDA O FOCO DO CURSOR PARA A PRÓXIMA TEXTBOX'

If keyCode = vbKeyReturn Then
    If txtvalorProduto = "" Or txtvalorProduto = "0" Then
        MsgBox "FAVOR DIGITAR O VALOR DO PRODUTO!", vbCritical, "ALERTA"
        txtvalorProduto = ""
        txtvalorProduto.SetFocus
    Else
        txtAliquotaIpi.SetFocus
    End If
End If
End Sub

Private Sub txtValorProdutoSubst_Change() 'MUDAR A COR DA LETRA'

txtValorProdutoSubst.ForeColor = vbRed

End Sub

Private Sub txtValorSubst_Change() 'MUDAR A COR DA LETRA'

txtValorSubst.ForeColor = vbRed

End Sub
