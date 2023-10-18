VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmNumeroConf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELEZIONA NUMERO CONFEZIONI"
   ClientHeight    =   3000
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNumeroConf.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5318
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
Attribute VB_Name = "frmNumeroConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub GET_GRIGLIA_PROCESSI()
On Error GoTo ERR_GET_GRIGLIA_PROCESSI
Dim sSQL As String
Dim sSQL_WHERE As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = "SELECT *  "
sSQL = sSQL & "FROM RV_POIEDistintaBasaConf "
sSQL = sSQL & "WHERE IDArticolo=" & frmMain.CDArticolo.KeyFieldID
sSQL = sSQL & " AND IDArticoloImballo=" & frmMain.CDImballo.KeyFieldID

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection

With Me.GrigliaCorpo
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
    
   
    .ColumnsHeader.Add "IDRV_PODistintaBaseRigheConf", "IDRV_PODistintaBaseRigheConf", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDRV_PODistintaBase", "IDRV_PODistintaBase", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDArticoloImballo", "IDArticoloImballo", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDArticolo", "IDArticoloPadre", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignRight
    
    .ColumnsHeader.Add "NumeroConfezioni", "N� conf.", dgchar, True, 2000, dgAlignleft
    Set cl = .ColumnsHeader.Add("Tara", "Tara per conf.", dgDouble, True, 2000, dgAlignRight, True, True, False)
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    
    .ColumnsHeader.Add "IDArticoloAssociato", "IDArticoloAssociato", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "CodiceArticoloAssociato", "Codice art. conf.", dgchar, True, 2000, dgAlignleft
    .ColumnsHeader.Add "ArticoloAssociato", "Articolo conf.", dgchar, True, 2500, dgAlignleft

    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
    
End With

Cn.CursorLocation = OLDCursor

Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA_PROCESSI"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GrigliaCorpo_DblClick
    End If
    
End Sub

Private Sub Form_Load()
        GET_GRIGLIA_PROCESSI
End Sub

Private Sub GrigliaCorpo_DblClick()
    If ((Me.GrigliaCorpo.Recordset.EOF) And (Me.GrigliaCorpo.Recordset.BOF)) Then Exit Sub
    
    frmMain.txtNumeroConfImballo.Value = fnNotNullN(Me.GrigliaCorpo.AllColumns("NumeroConfezioni").Value)
    frmMain.txtTaraConfImballo.Value = fnNotNullN(Me.GrigliaCorpo.AllColumns("Tara").Value)
    IDImballoPrimario = fnNotNullN(Me.GrigliaCorpo.AllColumns("IDArticoloAssociato").Value)
    CodiceImballoPrimario = fnNotNull(Me.GrigliaCorpo.AllColumns("CodiceArticoloAssociato").Value)
    DescrizioneImballoPrimario = fnNotNull(Me.GrigliaCorpo.AllColumns("ArticoloAssociato").Value)
    CostoConfezione = fnNotNullN(Me.GrigliaCorpo.AllColumns("CostoConfezione").Value)
    
    Unload Me
    
    
End Sub
