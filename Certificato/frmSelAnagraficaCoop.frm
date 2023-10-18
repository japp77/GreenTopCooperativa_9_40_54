VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelAnagraficaCoop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleziona anagrafica cooperativa"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelAnagraficaCoop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13150
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMove      =   0   'False
      ColumnsHeaderHeight=   20
   End
End
Attribute VB_Name = "frmSelAnagraficaCoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        Griglia_DblClick
    End If
End Sub

Private Sub Form_Load()
    LINK_ANA_COOP_SEL = 0
    GET_GRIGLIA
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA_PROCESSI
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POIEAnagraficaCooperativaDaLibroSoci "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
If Len(Trim(frmMain.txtCodiceAnaCoop.Text)) > 0 Then
    sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + Trim(frmMain.txtCodiceAnaCoop) + "%")
End If
If Len(Trim(frmMain.txtAnaCoop.Text)) > 0 Then
    sSQL = sSQL & " AND DenominazioneCompleta LIKE " + fnNormString("%" + Trim(frmMain.txtAnaCoop.Text) + "%")
End If

sSQL = sSQL & " ORDER BY DenominazioneCompleta"

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic

With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    
    .ColumnsHeader.Clear
    
    .ColumnsHeader.Add "IDAnagraficaFatturazione", "IDAnagraficaFatturazione", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Codice", "Codice", dgchar, True, 1500, dgAlignleft
    .ColumnsHeader.Add "DenominazioneCompleta", "Ragione sociale", dgchar, True, 7000, dgAlignleft
    
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With

Cn.CursorLocation = OLDCursor
Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Griglia_DblClick()
On Error GoTo ERR_Griglia_DblClick
    LINK_ANA_COOP_SEL = fnNotNullN(Me.Griglia.AllColumns("IDAnagraficaFatturazione").Value)
    Unload Me
Exit Sub
ERR_Griglia_DblClick:
    MsgBox Err.Description, vbCritical, "Griglia_DblClick"
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_Form_Unload
    rsGriglia.Close
    Set rsGriglia = Nothing
Exit Sub
ERR_Form_Unload:
    MsgBox Err.Description, vbCritical, "Form_Unload"
End Sub
