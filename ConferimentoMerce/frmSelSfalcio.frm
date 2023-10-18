VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelSfalcio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleziona sfalcio"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12645
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelSfalcio.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   6376
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
Attribute VB_Name = "frmSelSfalcio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsGriglia As ADODB.Recordset

Private Sub Form_Load()
    CREA_RECORDSET
End Sub
Private Sub CREA_RECORDSET()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

rsGriglia.Fields.Append "IDRV_PO01_LottoCampagnaSfalcio", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "IDRV_PO01_Sfalcio", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "Descrizione", adVarChar, 250, adFldIsNullable
rsGriglia.Fields.Append "DataPresunta", adDBDate, , adFldIsNullable
rsGriglia.Fields.Append "DataEffettiva", adDBDate, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

If ATTIVA_SEQUENZA_SFALCIO = 1 Then
    sSQL = "SELECT * FROM RV_PO01_IELottoCampagnaSfalcio "
    sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagna=" & LINK_LOTTO_PRODUZIONE_SEL
    sSQL = sSQL & " AND NOT (DataEffettivaInizio IS NULL ) "
    sSQL = sSQL & " ORDER BY Sequenza"
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        rsGriglia.AddNew
            rsGriglia!IDRV_PO01_LottoCampagnaSfalcio = fnNotNullN(rs!IDRV_PO01_LottoCampagnaSfalcio)
            rsGriglia!IDRV_PO01_Sfalcio = fnNotNullN(rs!IDRV_PO01_Sfalcio)
            rsGriglia!Descrizione = fnNotNull(rs!Sfalcio)
            rsGriglia!DataPresunta = rs!DataPresuntaInizio
            rsGriglia!DataEffettiva = rs!DataEffettivaInizio
        rsGriglia.Update
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
        
    sSQL = "SELECT  TOP (1) * FROM RV_PO01_IELottoCampagnaSfalcio "
    sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagna=" & LINK_LOTTO_PRODUZIONE_SEL
    sSQL = sSQL & " AND DataEffettivaInizio IS NULL "
    sSQL = sSQL & " ORDER BY Sequenza"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        rsGriglia.AddNew
            rsGriglia!IDRV_PO01_LottoCampagnaSfalcio = fnNotNullN(rs!IDRV_PO01_LottoCampagnaSfalcio)
            rsGriglia!IDRV_PO01_Sfalcio = fnNotNullN(rs!IDRV_PO01_Sfalcio)
            rsGriglia!Descrizione = fnNotNull(rs!Sfalcio)
            rsGriglia!DataPresunta = rs!DataPresuntaInizio
            rsGriglia!DataEffettiva = rs!DataEffettivaInizio
        rsGriglia.Update
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    
Else
    sSQL = "SELECT * FROM RV_PO01_IELottoCampagnaSfalcio "
    sSQL = sSQL & " WHERE IDRV_PO01_LottoCampagna=" & LINK_LOTTO_PRODUZIONE_SEL
    sSQL = sSQL & " ORDER BY Sequenza"
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        rsGriglia.AddNew
            rsGriglia!IDRV_PO01_LottoCampagnaSfalcio = fnNotNullN(rs!IDRV_PO01_LottoCampagnaSfalcio)
            rsGriglia!IDRV_PO01_Sfalcio = fnNotNullN(rs!IDRV_PO01_Sfalcio)
            rsGriglia!Descrizione = fnNotNull(rs!Sfalcio)
            rsGriglia!DataPresunta = rs!DataPresuntaInizio
            rsGriglia!DataEffettiva = rs!DataEffettivaInizio
        rsGriglia.Update
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
End If

SettaggioGriglia
End Sub
Private Sub SettaggioGriglia()
On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
    With Me.Griglia
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_PO01_LottoCampagnaSfalcio", "IDRV_PO01_LottoCampagnaSfalcio", dgNumeric, False, 1000, dgAlignRight
            .ColumnsHeader.Add "IDRV_PO01_Sfalcio", "IDRV_PO01_Sfalcio", dgNumeric, False, 1000, dgAlignRight
            .ColumnsHeader.Add "Descrizione", "Descrizione sfalcio", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "DataPresunta", "Data presunta", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DataEffettiva", "Data effettiva", dgDate, True, 2000, dgAlignleft
        Set .Recordset = rsGriglia
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia"
End Sub



Private Sub Griglia_DblClick()
    If Not ((Me.Griglia.Recordset.BOF) And (Me.Griglia.Recordset.EOF)) Then
        frmMain.txtIDSfalcioLotto.Value = Me.Griglia.AllColumns("IDRV_PO01_LottoCampagnaSfalcio").Value
    End If
    
    Unload Me
End Sub
