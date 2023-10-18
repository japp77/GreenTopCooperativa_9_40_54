Attribute VB_Name = "ModAnagraficaPLUS"
Option Explicit
Public oDb As DMTDataLayer.Database
Public cnDMT As DmtOleDbLib.adoConnection

Public IDArticolo As Long
Public IDEvent As Integer
Public IDFunzione As Long
Public IDTipoOggetto As Long
Public CallerName As String
Public HwndCaller As Long
'**********************VARIABILI GLOBALI AZIENDA**************************
    Public VarIDAzienda As Long
    Public VarIDAttivitaAzienda As Long
    Public VarIDFiliale As Long
    Public VarIDEsercizio As Long
    Public VarIDUtente As Long
    
'*************************************************************************
Public Utente As String
Public Password As String

Public gResource As Resource
Public REGISTRY_KEY As String
Public strConnectionString As String

Sub Main()

    Set gResource = New Resource
    REGISTRY_KEY = Trim(gResource.GetMessage(LBL_REGISTRY_KEY))
    'Legge i dati relativi all'anagrafica corrente dal file di registro
    IDArticolo = GetSetting(REGISTRY_KEY, "Links\RV_POArticoliPLUS.exe", "IDField")
    IDEvent = GetSetting(REGISTRY_KEY, "Links\RV_POArticoliPLUS.exe", "IDEvent", 0)
    IDFunzione = GetSetting(REGISTRY_KEY, "Links\RV_POArticoliPLUS.exe", "IDFunzione", 0)
    IDTipoOggetto = GetSetting(REGISTRY_KEY, "Links\RV_POArticoliPLUS.exe", "IDTipoOggetto", 0)
    CallerName = GetSetting(REGISTRY_KEY, "Links\RV_POArticoliPLUS.exe", "ApplicationNameCaller", "")
    HwndCaller = GetSetting(REGISTRY_KEY, "Links\RV_POArticoliPLUS.exe", "HwndCaller", 0)
    
    strConnectionString = GetSetting(REGISTRY_KEY, "MenuSettings", "ConnectionString")
    Utente = GetSetting(REGISTRY_KEY, "MenuSettings", "LASTUSER")
    Password = fnCryptString(GetSetting(REGISTRY_KEY, "MenuSettings", "LASTUSERPWD"))
    
    'Se nell'anagrafica di Diamante non era stata selezionata
    'nessuna anagrafica viene dato un messaggio di avvertimento
    If IDArticolo = 0 Then
        MsgBox "Impossibile procedere!" & vbCrLf & "Non è stata selezionato nessun articolo!", vbInformation, "Articoli PLUS"
    Else
        Select Case IDEvent
            Case 5   'OnSave
            Case 6   'OnDelete
                'Non si deve cancellare nulla poichè i campi aggiuntivi sono
                'nella tabella articoli

            Case 7   'OnLink
                frmMain.Show
        End Select
    End If
End Sub



Public Sub CloseConnection()
    'Set oDb = Nothing
    If Not (cnDMT Is Nothing) Then
        cnDMT.CloseConnection
        Set cnDMT = Nothing
    End If
End Sub

Public Sub ConnessioneADODBLib()
Dim StringaDiConnessione As String

    If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
        StringaDiConnessione = MenuOptions.ConnectionString
    Else
        StringaDiConnessione = MenuOptions.ConnectionString & ";"
    End If
    
    Set cnDMT = DmtOleDbLib.adoEnvironments(0).OpenConnection((StringaDiConnessione & "User Id=" & Utente & ";Password=" & Password))

End Sub
Public Sub OpenConnection()
Dim StringaDiConnessione As String

    If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
        StringaDiConnessione = MenuOptions.ConnectionString
    Else
        StringaDiConnessione = MenuOptions.ConnectionString & ";"
    End If
    'Apre una connessione DBLib
    Set oDb = New DMTDataLayer.Database
    oDb.OpenConnection StringaDiConnessione & "User Id=" & Utente & ";Password=" & Password
        
End Sub
Public Sub PrelevaAzienda()

    Dim TmpFiliale As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    'TmpFiliale = GetSetting("Diamante", "MenuSettings", "LASTBRANCH")
    
    sSQL = "SELECT Azienda.IDAzienda, Anagrafica.Anagrafica, AttivitaAzienda.IDAttivitaAzienda, AttivitaAzienda.AttivitaAzienda, Filiale.IDFiliale, Filiale.Filiale"
    sSQL = sSQL & " FROM (Anagrafica INNER JOIN Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica) INNER JOIN (Filiale INNER JOIN AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda) ON Azienda.IDAzienda = AttivitaAzienda.IDAzienda"
    sSQL = sSQL & " WHERE (((Filiale.IDFiliale)=" & MenuOptions.LastBranch & "))"
    
    
    Set rs = cnDMT.OpenResultset(sSQL)
        VarIDAzienda = rs!IDAzienda
        VarIDAttivitaAzienda = rs!IDAttivitaAzienda
        VarIDFiliale = rs!IDFiliale
        VarIDUtente = MenuOptions.LastUserID
        
    rs.CloseResultset
    Set rs = Nothing
    
    
End Sub
