VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelezionaLottoDiCampagna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SELEZIONE LOTTO DI PRODUZIONE"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   18885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelezionaLottoDiCampagna.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   18885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLottoChiuso 
      Caption         =   "Visualizza anche i lotti di produzione chiusi"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4095
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   11668
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
Attribute VB_Name = "frmSelezionaLottoDiCampagna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset

Private Sub chkLottoChiuso_Click()
    SettaggioGriglia
End Sub

Private Sub Form_Activate()
On Error GoTo ERR_Form_Activate
    
    SettaggioGriglia
    Me.Griglia.Recordset.Requery
    
Exit Sub
ERR_Form_Activate:
    MsgBox Err.Description, vbCritical, "Form_Activate"
    Unload Me
End Sub


Private Sub SettaggioGriglia()
On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT dbo.RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna, dbo.RV_PO01_LottoCampagna.IDAzienda, dbo.RV_PO01_LottoCampagna.IDFiliale, dbo.RV_PO01_LottoCampagna.IDSocio, "
    sSQL = sSQL & "dbo.RV_PO01_LottoCampagna.IDRV_PO01_Varieta, dbo.RV_PO01_LottoCampagna.IDRV_PO01_FamigliaProdotti, dbo.RV_PO01_LottoCampagna.IDRV_PO01_PeriodoCampagna, dbo.RV_PO01_LottoCampagna.CodiceLotto,"
    sSQL = sSQL & "dbo.RV_PO01_LottoCampagna.DescrizioneLotto, dbo.RV_PO01_LottoCampagna.Chiuso, dbo.RV_PO01_LottoCampagna.IDRV_PO01_ClassificazioneLottoProd01,"
    sSQL = sSQL & "dbo.RV_PO01_LottoCampagna.IDRV_PO01_ClassificazioneLottoProd02, dbo.RV_PO01_FamigliaProdotti.FamigliaProdotti, dbo.RV_PO01_FamigliaProdotti.ResaMinima, dbo.RV_PO01_FamigliaProdotti.ResaMassima,"
    sSQL = sSQL & "dbo.RV_PO01_FamigliaProdotti.UtilizzaNelCertificato, dbo.RV_PO01_Varieta.Varieta, dbo.RV_PO01_ClassificazioneLottoProd01.ClassificazioneLottoProd01,"
    sSQL = sSQL & "dbo.RV_PO01_ClassificazioneLottoProd02.ClassificazioneLottoProd02, dbo.RV_PO01_LottoCampagna.Acquistato, dbo.RV_PO01_LottoCampagna.Provvisorio, "
    sSQL = sSQL & "dbo.RV_PO01_StatoLotto.StatoLotto, CASE WHEN dbo.RV_PO01_Varieta.ResaMinima IS NOT NULL THEN (DimensioneMQ / 10000) "
    sSQL = sSQL & "* dbo.RV_PO01_Varieta.ResaMinima ELSE 0 END AS ResaTotale,"
    sSQL = sSQL & "(SELECT SUM(PesoNettoCalcolato) AS PesoTotale"
    sSQL = sSQL & " From dbo.RV_POCertificato"
    sSQL = sSQL & " WHERE (IDLottoProduzione = dbo.RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna)) AS PesoUtilizzato "
    sSQL = sSQL & "FROM RV_PO01_LottoCampagna INNER JOIN "
    sSQL = sSQL & "RV_PO01_FamigliaProdotti ON RV_PO01_LottoCampagna.IDRV_PO01_FamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti INNER JOIN "
    sSQL = sSQL & "RV_PO01_Varieta ON RV_PO01_LottoCampagna.IDRV_PO01_Varieta = RV_PO01_Varieta.IDRV_PO01_Varieta LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_ClassificazioneLottoProd01 ON RV_PO01_LottoCampagna.IDRV_PO01_ClassificazioneLottoProd01 = RV_PO01_ClassificazioneLottoProd01.IDRV_PO01_ClassificazioneLottoProd01 LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_ClassificazioneLottoProd02 ON RV_PO01_LottoCampagna.IDRV_PO01_ClassificazioneLottoProd02 = RV_PO01_ClassificazioneLottoProd02.IDRV_PO01_ClassificazioneLottoProd02 LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_StatoLotto ON RV_PO01_LottoCampagna.IDRV_PO01_StatoLotto = RV_PO01_StatoLotto.IDRV_PO01_StatoLotto "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDSocio=" & frmMain.CDSocio.KeyFieldID
    sSQL = sSQL & " AND UtilizzaNelCertificato=1 "
    sSQL = sSQL & " AND ((Provvisorio=0 OR Provvisorio IS NULL))"
    sSQL = sSQL & " AND ((VirtualDelete=0 OR VirtualDelete IS NULL))"
    If Me.chkLottoChiuso.Value = vbUnchecked Then
        sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(Me.chkLottoChiuso.Value)
    End If
    If Len(STRINGA_RICERCA_LOTTO) > 0 Then
        sSQL = sSQL & " AND CodiceLotto LIKE %" & fnNormString(STRINGA_RICERCA_LOTTO) & "%"
    End If
    
    Set rsArt = Cn.OpenResultset(sSQL)
        Set rsEvent = rsArt.Data
    
    With Me.Griglia
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_PO01_LottoCampagna", "ID", dgNumeric, False, 1000, dgAlignleft
            .ColumnsHeader.Add "CodiceLotto", "Codice", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "StatoLotto", "Stato", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "FamigliaProdotti", "Famiglia prodotti", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "Varieta", "Varietà", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Acquistato", "Acquistato", dbBoolean, True, 1000, dgAligncenter
            Set cl = .ColumnsHeader.Add("ResaTotale", "Resa", dgDouble, True, 1600, dgAlignRight)
                cl.BackColor = vbGreen
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PesoUtilizzato", "Utilizzato", dgDouble, True, 1600, dgAlignRight)
                cl.BackColor = vbGreen
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ResaMinima", "Resa per HA", dgDouble, False, 1600, dgAlignRight)
                cl.BackColor = vbGreen
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ResaMassina", "Resa massima per HA", dgDouble, False, 1600, dgAlignRight)
                cl.BackColor = vbGreen
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
        Set .Recordset = rsArt.Data
        .LoadUserSettings
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia Articoli"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
If KeyCode = vbKeyReturn Then
    Griglia_DblClick
End If

End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_Form_Unload
    rsArt.CloseResultset
    Set rsArt = Nothing
Exit Sub
ERR_Form_Unload:
    MsgBox Err.Description, vbCritical, "Form_Unload"
End Sub

Private Sub Griglia_DblClick()
    frmMain.txtIDLottoCampagna.Value = fnNotNullN(Me.Griglia.AllColumns("IDRV_PO01_LottoCampagna").Value)
    Unload Me
End Sub
