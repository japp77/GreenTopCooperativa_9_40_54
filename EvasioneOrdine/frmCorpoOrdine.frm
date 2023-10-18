VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmCorpoOrdine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CORPO ORDINE COLLEGATO"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCorpoOrdine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   19110
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpoOrdine 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   9128
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
Attribute VB_Name = "frmCorpoOrdine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

'RECUPERO CAMPI
sSQL = "SELECT * FROM RV_POIEOrdineSelCorpo "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

For I = 0 To rs.Fields.Count - 1
    Select Case rs.Fields(I).Type
        Case adChar, adVarChar, adVarWChar, adWChar, 201
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
        Case adInteger
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsGriglia.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsGriglia.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
    End Select
Next

rsGriglia.Fields.Append "ImportoUnitarioImballo", adDouble, , adFldIsNullable


rs.Close
Set rs = Nothing

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_POIEOrdineSelCorpo "
sSQL = sSQL & "WHERE IDOggetto=" & LINK_ORDINE_PER_PREZZO
sSQL = sSQL & " AND RV_POTipoRiga=1 "
If LINK_ARTICOLO_ORDINE > 0 Then
    sSQL = sSQL & " AND Link_art_articolo=" & LINK_ARTICOLO_ORDINE
End If

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    rsGriglia.AddNew
        For I = 0 To rs.Fields.Count - 1
            rsGriglia.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
        Next
        rsGriglia!ImportoUnitarioImballo = GET_PREZZO_IMBALLO_DA_ORDINE(fnNotNullN(rsGriglia!RV_POIDImballo), fnNotNullN(rsGriglia!RV_POLinkRiga), fnNotNullN(rsGriglia!IDOggetto))
    rsGriglia.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

GET_GRIGLIA

Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long

OLDCursor = CnDMT.CursorLocation
CnDMT.CursorLocation = 3

With Me.GrigliaCorpoOrdine
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
           
        .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1100, dgAlignleft
        .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, True, 3000, dgAlignleft
        
        .ColumnsHeader.Add "RV_POIDTipoPedana", "RV_POIDTipoPedana", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceTipoPedana", "Codice tipo pedana", dgchar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneTipoPedana", "Descrizione codice pedana", dgchar, False, 3000, dgAlignleft
        
        .ColumnsHeader.Add "RV_PONotaRigaOrdRaggr", "Raggr. ord.", dgchar, True, 2500, dgAlignleft
        
        Set cl = .ColumnsHeader.Add("RV_POQuantitaPedana", "Q.t� pedana ord.", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("RV_POQuantitaPedanaEffettiva", "Q.t� pedana", dgDouble, True, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_numero_colli", "Colli", dgDouble, True, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_quantita_pezzi", "Pezzi", dgDouble, True, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
            
        .ColumnsHeader.Add "RV_POIDTipoUMOrdine", "RV_POIDTipoUMOrdine", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POTipoUMOrdine", "Tipo U.M. riga ordine", dgchar, True, 2500, dgAlignleft

        Set cl = .ColumnsHeader.Add("Art_peso", "Peso lordo", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("Art_tara", "Tara", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        Set cl = .ColumnsHeader.Add("PesoNetto", "Peso netto", dgDouble, False, 1300, dgAlignRight)
            'cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."

        Set cl = .ColumnsHeader.Add("Art_prezzo_unitario_neutro", "Imp. uni. merce", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 5
            cl.FormatOptions.FormatNumericThousandSep = "."

        Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_1", "% Sc1", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."

         Set cl = .ColumnsHeader.Add("Art_sco_in_percentuale_2", "% Sc2", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."
        .ColumnsHeader.Add "RV_POIDImballo", "RV_POIDImballo", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "RV_POCodiceImballo", "Codice imballo", dgchar, True, 1100, dgAlignleft
        .ColumnsHeader.Add "RV_PODescrizioneImballo", "Descrizione imballo", dgchar, True, 3000, dgAlignleft
         Set cl = .ColumnsHeader.Add("ImportoUnitarioImballo", "Imp. uni. imb.", dgDouble, False, 1300, dgAlignRight)
            cl.BackColor = vbYellow
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = ""
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."

   
        Set cl = .ColumnsHeader.Add("RV_POImportoImballoInArticolo", "Incluso Imballo", dgBoolean, False, 1300, dgAligncenter)
            cl.BackColor = vbYellow
   
        .ColumnsHeader.Add "RV_POIDTipoCategoria", "RV_POIDTipoCategoria", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoCategoria", "Tipo categoria", dgchar, True, 1500, dgAlignleft
        
        .ColumnsHeader.Add "RV_POIDCalibro", "RV_POIDCalibro", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "Calibro", "Calibro", dgchar, True, 1500, dgAlignleft
        
        .ColumnsHeader.Add "RV_POIDTipoLavorazione", "RV_POIDTipoLavorazione", dgNumeric, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoLavorazione", "Tipo lavorazione", dgchar, True, 1500, dgAlignleft

        .ColumnsHeader.Add "RV_POAnnotazioniRigaOrdine", "Note riga ord.", dgchar, False, 3000, dgAlignleft
        
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With

CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Function GET_PREZZO_IMBALLO_DA_ORDINE(IDImballo As Long, linkRiga As Long, IDOggettoOrdine As Long) As Double
On Error GoTo ERR_GET_PREZZO_IMBALLO_DA_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_PREZZO_IMBALLO_DA_ORDINE = 0

sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=2 "
sSQL = sSQL & " AND RV_POLinkRiga=" & linkRiga
sSQL = sSQL & " AND Link_Art_articolo=" & IDImballo

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
   GET_PREZZO_IMBALLO_DA_ORDINE = fnNotNullN(rs!Art_prezzo_unitario_netto_IVA)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_PREZZO_IMBALLO_DA_ORDINE:
    GET_PREZZO_IMBALLO_DA_ORDINE = 0
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CONFERMA_SEL_PREZZO_DA_ORD = 0
    CREA_RECORDSET
End Sub
Private Sub GrigliaCorpoOrdine_DblClick()
On Error GoTo ERR_GrigliaCorpoOrdine_DblClick
    
    If MODALITA_RECUPERO_RIGA_ORD = 0 Then
        RECORDSET_RETURN_PER_PREZZO!ImportoUnitarioArticolo = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("Art_prezzo_unitario_neutro").Value)
        RECORDSET_RETURN_PER_PREZZO!Sconto1 = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("Art_sco_in_percentuale_1").Value)
        RECORDSET_RETURN_PER_PREZZO!Sconto2 = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("Art_sco_in_percentuale_2").Value)

        RECORDSET_RETURN_PER_PREZZO!NotaRigaOrdRaggr = fnNotNull(Me.GrigliaCorpoOrdine.AllColumns("RV_PONotaRigaOrdRaggr").Value)
        
        If RECORDSET_RETURN_PER_PREZZO!IDImballoVendita = fnNotNullN(rsGriglia!RV_POIDImballo) Then
            RECORDSET_RETURN_PER_PREZZO!ImportoUnitarioImballo = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("ImportoUnitarioImballo").Value)
            RECORDSET_RETURN_PER_PREZZO!MerceInclusoImballo = Abs(fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("RV_POImportoImballoInArticolo").Value))
            RETURN_SEL_PREZZO_IMB_DA_ORD = 1
        End If
        
        CONFERMA_SEL_PREZZO_DA_ORD = 1
        
    End If
    If MODALITA_RECUPERO_RIGA_ORD = 1 Then
        RETURN_SEL_PREZZO_IMB_DA_ORD = 1
        CONFERMA_SEL_PREZZO_DA_ORD = 1
        frmAssegnazioneMerce.txtIDRigaOrdine.Value = fnNotNullN(Me.GrigliaCorpoOrdine.AllColumns("IDValoriOggettoDettaglio").Value)
    End If
    
    Unload Me
    
Exit Sub
ERR_GrigliaCorpoOrdine_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCorpoOrdine_DblClick"
    
End Sub

