VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmContrattoDettaglio 
   Caption         =   "SELEZIONA DETTAGLIO CONTRATTO"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContrattoDettaglio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   12390
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   5953
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
Attribute VB_Name = "frmContrattoDettaglio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private ColonnaSelezionata As String



Private Sub Form_Activate()
    ColonnaSelezionata = ""
    CREA_RECORDSET
End Sub

Private Sub CREA_RECORDSET()
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim I As Long

If Not (rsGriglia Is Nothing) Then
    If rsGriglia.State > 0 Then
        rsGriglia.Close
    End If
    Set rsGriglia = Nothing
End If
Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

If Not (rsContrattoDettaglioSel Is Nothing) Then
    If rsContrattoDettaglioSel.State > 0 Then
        rsContrattoDettaglioSel.Close
    End If
    Set rsContrattoDettaglioSel = Nothing
End If
Set rsContrattoDettaglioSel = New ADODB.Recordset
rsContrattoDettaglioSel.CursorLocation = adUseClient

sSQL = "SELECT * FROM RV_POIEContrattoDettaglioSel "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

For I = 0 To rs.Fields.Count - 1
    Select Case rs.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
        Case adInteger
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsGriglia.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsGriglia.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
            rsContrattoDettaglioSel.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
    End Select
Next

rsGriglia.Fields.Append "Selezionato", adSmallInt, , adFldIsNullable
rsContrattoDettaglioSel.Fields.Append "Selezionato", adSmallInt, , adFldIsNullable

rs.Close
Set rs = Nothing

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic
rsContrattoDettaglioSel.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_POIEContrattoDettaglioSel "
sSQL = sSQL & " WHERE IDOggetto=" & LINK_CONTRATTO
sSQL = sSQL & " AND RV_POTipoRiga=1"


Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

While Not rs.EOF
    rsGriglia.AddNew
    For I = 0 To rs.Fields.Count - 1
        rsGriglia.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
    Next
    rsGriglia!Selezionato = 0
    rsGriglia.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

GET_GRIGLIA

End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_CURSOR As Long

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
    .LoadUserSettings

    .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
    .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
    .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft

    .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
    .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1500, dgAlignleft
    .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, True, 3000, dgAlignleft

    Set cl = .ColumnsHeader.Add("Art_numero_colli", "Totale colli", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbGreen
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 2
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_quantita_totale", "Q.tà U.M.", dgDouble, True, 1300, dgAlignRight)
        cl.BackColor = vbGreen
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("Art_pre_uni_net_sco_net_IVA", "Prezzo", dgDouble, True, 1600, dgAlignRight)
        cl.BackColor = vbGreen
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("RV_POImportoUnitarioMin", "Prezzo minimo", dgDouble, True, 1600, dgAlignRight)
        cl.BackColor = vbGreen
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set cl = .ColumnsHeader.Add("RV_POImportoUnitarioMax", "Prezzo massimo", dgDouble, True, 1600, dgAlignRight)
        cl.BackColor = vbGreen
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 5
        cl.FormatOptions.FormatNumericThousandSep = "."
    Set .Recordset = rsGriglia
    .LoadUserSettings
    .Refresh
    
End With

Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyEscape) Then Unload Me
    If KeyCode = vbKeyReturn Then
        Griglia_DblClick
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_Form_Unload
    rsGriglia.Close
    Set rsGriglia = Nothing
Exit Sub
ERR_Form_Unload:
    MsgBox Err.Description, vbCritical, "Form_Unload"
End Sub

Private Sub Griglia_DblClick()
    frmMain.txtIDContrattoRiga.Value = fnNotNullN(Me.Griglia.AllColumns("IDValoriOggettoDettaglio").Value)
    Unload Me
End Sub


