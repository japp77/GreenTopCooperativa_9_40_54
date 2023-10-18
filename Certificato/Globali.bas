Attribute VB_Name = "Globali"
Option Explicit

'Declares
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Sub sbOpenURL Lib "Diamante.dll" (ByVal hwnd As Long, ByVal sURL As String)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOP = 0
Public Const WM_SETREDRAW = &HB

'Costanti globali
Public Const TOTAL_CONTROLS_NUMBER = 10
Public Const SPLITLIMIT = 1000
Public Const SRCNEXT = 1
Public Const SRCPREVIOUS = 2
Public Const HELP_FINDER = &HB
Public Const HELP_CONTEXT = &H1
Public Const URL_DIAMANTE = "http://www.diamante.it"

'*** Costanti per la gestione della Attivazione-Disattivazione Menu e ToolBar
Public Const BTN_NEW = 1
Public Const BTN_SAVE = 2
Public Const BTN_PRINT = 4
Public Const BTN_PREVIEW = 8
Public Const BTN_CUT = 16
Public Const BTN_COPY = 32
Public Const BTN_PASTE = 64
Public Const BTN_DELETE = 128
Public Const BTN_CLEAR = 256
Public Const BTN_FIND = 512
Public Const BTN_SEARCH = 1024
Public Const BTN_VIEWMODE = 2048
Public Const BTN_PREVIOUS = 4096
Public Const BTN_NEXT = 8192
Public Const BTN_WORD = 16384
Public Const BTN_EXCEL = 32768
Public Const BTN_HTML = 65536
Public Const BTN_SEARCHFORM = 131072
Public Const BTN_SEARCHTABLE = 262144
Public Const BTN_FILTER = 262144 * 2
Public Const BTN_TOOLS = BTN_FILTER * 2
Public Const BTN_PDF = BTN_TOOLS * 2
Public Const BTN_EXPORT = BTN_PDF * 2
Public Const BTN_ALL = BTN_EXPORT * 2 - 1

'Il nome della ToolBar dell'Anteprima di stampa
Public Const BAND_CLOSE_PREVIEW = "Band_ClosePreview"

'Elenco errori
Public Const ERR_TABLE_STRUCT = vbObjectError + 10000
Public Const ERR_NO_DEFAULT_TABLEVIEW = vbObjectError + 10001
Public Const ERR_NO_PROCESSES = vbObjectError + 10002
Public Const ERR_NDELFILTER = vbObjectError + 2500



'La variabile globale TheApp mantiene un riferimento all'oggetto
'applicazione che viene utilizzato per eseguire le funzionalità
'ed i relativi processi del gestore.
Public TheApp As Application

'La variabile globale gResource mantiene un riferimento all'oggetto
'utilizzato per l'accesso alle risorse stringa, icon e bitmap di Diamante
Public gResource As Resource

Public Cn As DmtOleDbLib.adoConnection
Public Db As DMTDataLayer.Database


Public REGISTRY_KEY As String

Public LINK_CLIENTE_CONTRATTO As Long
Public LINK_CONTRATTO As Long
Public rsContrattoDettaglioSel As ADODB.Recordset
Public DATI_DA_CONTRATTO As Boolean


Public STRINGA_RICERCA_LOTTO As String

Public LINK_TIPO_ARROTONDAMENTO As Long
Public ATTIVA_SEZIONALE_DA_SOCIO As Long
Public Link_TipoSocio As Long
Public Link_TipoImballo As Long
Public UtilizzaDataENumSocioPerDDT As Long
Public NumeroColliPerAutomezzoCert As Long
Public LINK_MAGAZZINO_DOCUMENTO As Long
Public NUMERO_ZERI_DOC_RIF As Long

Public LINK_DOCUMENTO_COLLEGATO As Long

Public TIPO_SALVATAGGIO As Long

Public IDAnagrafica_PREC As Long
Public IDDestinazione_PREC As Long
Public IDVettore_PREC As Long
Public IDContratto_PREC As Long
Public IDContrattoRiga_PREC As Long
Public IDCooperativa_PREC As Long
Public IDAnagraficaSocio_PREC As Long

'Oggetto utilizzato per gestire l'inserimento / variazione del documento (DmtDocs.Dll)
Public oDoc As DmtDocs.cDocument
'Variabile utilizzata per ottenere il nome della tabella di testata del documento
Public sTabellaTestata As String
'Variabile utilizzata per ottenere il nome della tabella di dettaglio del documento
Public sTabellaDettaglio As String
'Variabile utilizzata per ottenere il nome della tabella delle scadenze del documento
Public sTabellaScadenze As String
'Variabile utilizzata per ottenere il nome della tabella del castelletto IVA del documento
Public sTabellaIVA As String

Public Rip_InXMLRifLetteraIntento As Long
Public Rip_InXMLRifNoteIva As Long
Public Rip_InXMLRifNota01Doc As Long
Public Rip_InXMLRifNota02Doc As Long
Public Rip_InXMLRifNota03Doc As Long
Public Rip_InXMLRifNotaDoc As Long
Public Rip_InXMLRifIstrMitt As Long
Public Rip_InXMLRifVettSucc As Long
Public Rip_InXMLRifAgenziaTrasp As Long
Public Rip_InXMLRifTargaAutoMezzo As Long

Public IDClassLottoProdPerFuoriQuota As Long
Public MsgInDocSeRigaMerceSenzaImballo As Long
Public IDAnagraficaDestSociDiretti  As Long

Public IDCategoriaAnagraficaSocioDiretto As Long
Public IDCategoriaAnagraficaProdAcq As Long
Public IDArticoloScartoPerCertificato As Long
Public IDCategoriaAnagraficaNoProd As Long

Public RiportaDestinazioneDaContrattoCertificato As Long
Public RiportaVettoreDaContrattoCertificato As Long
Public ForzaDestinazioneDaContrattoCertificato As Long
Public ForzaVettoreDaContrattoCertificato As Long
Public LINK_ANA_COOP_SEL As Long
Public LINK_ANA_SOCIO_SEL As Long
Public AttivaSelezioneSocioCertPerVarieta As Long
Public AttivaSelezioneAnaVeloceInCert As Long
Public NumeroMesiPerDataRevocaCertificato As Long

Public LINK_SOCIO_LOTTO_SEL As Long
Public NonRiportaInXMLRifVsNumOrd As Long
Public NonRiportareRifCerticatoInDDT As Long



