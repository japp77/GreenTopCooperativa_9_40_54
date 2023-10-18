VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Begin VB.Form frmSelAnagraficaSocio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleziona anagrafica socio"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelAnagraficaSocio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   19230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Seleziona"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16680
      TabIndex        =   14
      ToolTipText     =   "Elimina tutti i filtri e riesegue la lista "
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RICERCA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Refresh della lista in base ai filtri impostati"
      Top             =   8640
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TUTTI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      ToolTipText     =   "Elimina tutti i filtri e riesegue la lista "
      Top             =   8640
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "FILTRI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   9135
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox Check1 
         Caption         =   "Utilizza varietà"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtDescrizione 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtCodice 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin DMTDATETIMELib.dmtDate txtDataRevoca 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboSocioDiretto 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   4560
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboNonCompLibroSoci 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   5280
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboVarieta 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboFamiglia 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Famiglia articolo contratto"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Varietà articolo contratto"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Ragione sociale"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Codice"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Non compilare nel libro soci"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   5040
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Data revoca"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Socio diretto"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   2535
      End
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   8535
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   15055
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
Attribute VB_Name = "frmSelAnagraficaSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub INIT_CONTROLLI()
    
    With Me.cboSocioDiretto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POSiNo"
        .DisplayField = "SiNo"
        .SQL = "SELECT * FROM RV_POSiNO"
    End With
    
    With Me.cboNonCompLibroSoci
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POSiNo"
        .DisplayField = "SiNo"
        .SQL = "SELECT * FROM RV_POSiNO"
    End With
    
    'Famiglia
    With Me.cboFamiglia
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_PO01_FamigliaProdotti"
        .DisplayField = "FamigliaProdotti"
        .SQL = "SELECT * FROM RV_PO01_FamigliaProdotti "
        .Fill
    End With
    
    'Varieta
    With Me.cboVarieta
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_PO01_Varieta"
        .DisplayField = "Varieta"
        .SQL = "SELECT * FROM RV_PO01_Varieta WHERE IDRV_PO01_FamigliaProdotti=" & frmMain.LINK_FAMIGLIA_ART_CONTRATTO & " ORDER BY Varieta"
        .Fill
    End With
    
End Sub

Private Sub Command1_Click()
    GET_GRIGLIA
End Sub

Private Sub Command2_Click()
    Me.txtDataRevoca.Value = 0
    Me.txtCodice.Text = ""
    Me.txtDescrizione.Text = ""
    Me.cboSocioDiretto.WriteOn 0
    Me.cboNonCompLibroSoci.WriteOn 0
    Me.Check1.Value = vbUnchecked
    GET_GRIGLIA
End Sub

Private Sub Command3_Click()
    Griglia_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then
        Griglia_DblClick
    End If
End Sub

Private Sub Form_Load()
    INIT_CONTROLLI
    LINK_ANA_SOCIO_SEL = 0
    
    Me.txtDataRevoca.Text = DateAdd("m", -NumeroMesiPerDataRevocaCertificato, Date)
    Me.txtCodice.Text = frmMain.txtCodiceAnaSocio.Text
    Me.txtDescrizione.Text = frmMain.txtAnaSocio.Text
    Me.cboSocioDiretto.WriteOn 0
    Me.cboNonCompLibroSoci.WriteOn 0
    
    Me.cboSocioDiretto.Enabled = frmMain.CDSocioFatt.KeyFieldID = 0
    Me.cboFamiglia.WriteOn fnNotNullN(frmMain.LINK_FAMIGLIA_ART_CONTRATTO)
    Me.cboVarieta.WriteOn fnNotNullN(frmMain.LINK_VARIETA_ART_CONTRATTO)
    
    If frmMain.LINK_VARIETA_ART_CONTRATTO > 0 Then
        If (AttivaSelezioneSocioCertPerVarieta = 1) Then
            Me.Check1.Value = vbChecked
        End If
    End If
    
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
If (Me.Check1.Value = vbUnchecked) Then
    sSQL = sSQL & "FROM RV_POIEAnagraficaSocio "
Else
    If Me.cboVarieta.CurrentID > 0 Then
        sSQL = sSQL & "FROM RV_POIEAnagraficaSocioPerVarieta "
    Else
        sSQL = sSQL & "FROM RV_POIEAnagraficaSocio "
    End If
End If

sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
If (Me.txtDataRevoca.Value > 0) Then
    sSQL = sSQL & " AND ((DataUscita IS NULL) OR (DataUscita>" & fnNormDate(txtDataRevoca.Text) & "))"
End If
If (frmMain.CDSocioFatt.KeyFieldID > 0) Then
    sSQL = sSQL & " AND IDAnagraficaFatturazione=" & frmMain.CDSocioFatt.KeyFieldID
End If
If Len(Trim(txtCodice.Text)) > 0 Then
    sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + Trim(txtCodice.Text) + "%")
End If
If Len(Trim(txtDescrizione.Text)) > 0 Then
    sSQL = sSQL & " AND Anagrafica LIKE " + fnNormString("%" + Trim(txtDescrizione.Text) + "%")
End If
If Me.cboSocioDiretto.CurrentID > 0 Then
    If (Me.cboSocioDiretto.CurrentID = 1) Then
        sSQL = sSQL & " AND SocioDiretto = 1"
    End If
    If (Me.cboSocioDiretto.CurrentID = 2) Then
        sSQL = sSQL & " AND SocioDiretto = 0"
    End If
End If
If Me.cboNonCompLibroSoci.CurrentID > 0 Then
    If (Me.cboNonCompLibroSoci.CurrentID = 1) Then
        sSQL = sSQL & " AND NonCompilarePerLibroSoci = 1"
    End If
    If (Me.cboNonCompLibroSoci.CurrentID = 2) Then
        sSQL = sSQL & " AND NonCompilarePerLibroSoci = 0"
    End If
End If
If (Me.Check1.Value = vbChecked) Then
    If Me.cboVarieta.CurrentID > 0 Then
        sSQL = sSQL & " AND Provvisorio=0"
        sSQL = sSQL & " AND Chiuso=0"
        sSQL = sSQL & " AND IDRV_PO01_Varieta=" & Me.cboVarieta.CurrentID
    End If
End If
sSQL = sSQL & " ORDER BY Nome, Anagrafica"

Set rsGriglia = New ADODB.Recordset
rsGriglia.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockBatchOptimistic

With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    
    .ColumnsHeader.Clear
    
    .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAnagraficaFatturazione", "IDAnagraficaFatturazione", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignRight
    .ColumnsHeader.Add "Codice", "Codice", dgchar, True, 1500, dgAlignleft
    .ColumnsHeader.Add "Anagrafica", "Socio", dgchar, True, 3500, dgAlignleft
    .ColumnsHeader.Add "Nome", "Cooperativa", dgchar, True, 3500, dgAlignleft
    .ColumnsHeader.Add "SocioDiretto", "Diretto", dgBoolean, True, 1500, dgAligncenter
    .ColumnsHeader.Add "DataUscita", "DataUscita", dgDate, False, 1500, dgAlignleft
    .ColumnsHeader.Add "NonCompilarePerLibroSoci", "Non compilare nel libro soci", dgBoolean, False, 1500, dgAligncenter
    
    Set .Recordset = rsGriglia
    .LoadUserSettings
    .Refresh
End With

Cn.CursorLocation = OLDCursor
Exit Sub

ERR_GET_GRIGLIA_PROCESSI:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Griglia_DblClick()
On Error GoTo ERR_Griglia_DblClick
    LINK_ANA_SOCIO_SEL = fnNotNullN(Me.Griglia.AllColumns("IDAnagrafica").Value)
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
