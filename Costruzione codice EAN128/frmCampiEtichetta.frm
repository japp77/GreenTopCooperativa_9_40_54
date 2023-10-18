VERSION 5.00
Begin VB.Form frmCampiEtichetta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CAMPI ETICHETTA"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   8025
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmCampiEtichetta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    CARICA_CAMPI
End Sub
Private Sub CARICA_CAMPI()
Dim rs As ADODB.Recordset
Dim I As Long


Set rs = New ADODB.Recordset

rs.Open "RV_POTMPStampaEtichetteRighe", Cn.InternalConnection


Me.List1.Clear

For I = 0 To rs.Fields.Count - 1
    Me.List1.AddItem rs.Fields(I).Name
    Me.List1.ItemData(Me.List1.NewIndex) = I
Next

rs.Close
Set rs = Nothing
End Sub



Private Sub List1_DblClick()
    frmMain.txtCampoEtichetta.Text = Me.List1.Text
    Unload Me
End Sub
