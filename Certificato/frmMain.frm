VERSION 5.00
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{7A1D73E4-F461-11D0-8F01-004033A00AF2}#1.0#0"; "DmtWheel.ocx"
Object = "{5C67DC8E-40E7-11D3-AF44-00105A2FBE61}#3.0#0"; "DmtPrnDlgCtl.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{9385BB2E-6637-11D1-850D-002018802E11}#3.1#0"; "Dmtsplit.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F95AA20B-3F80-11D3-A741-00105A2E9BAF}#2.1#0"; "DmtSearchAccount2.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   12090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17385
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   12090
   ScaleWidth      =   17385
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar stbStatusbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   65
      Top             =   11745
      Width           =   17385
      _ExtentX        =   30665
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin ActiveBar3LibraryCtl.ActiveBar3 BarMenu 
      Height          =   11745
      Left            =   0
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   0
      Width           =   17385
      _LayoutVersion  =   2
      _ExtentX        =   30665
      _ExtentY        =   20717
      _DataPath       =   ""
      Bands           =   "frmMain.frx":4781A
      Begin DMTSPLIT.DMTSplitBar DMTSplitBar1 
         Height          =   510
         Left            =   3000
         TabIndex        =   72
         Top             =   360
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
      End
      Begin VB.PictureBox PicForm 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   11475
         Left            =   0
         ScaleHeight     =   11445
         ScaleWidth      =   17265
         TabIndex        =   69
         Top             =   0
         Width           =   17295
         Begin VB.PictureBox PicForm2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   11160
            Left            =   90
            ScaleHeight     =   11130
            ScaleWidth      =   17010
            TabIndex        =   70
            Top             =   135
            Width           =   17040
            Begin VB.Frame Frame1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   10860
               Left            =   120
               TabIndex        =   71
               Top             =   120
               Width           =   16755
               Begin VB.Frame Frame3 
                  Caption         =   "COOPERATIVA"
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
                  Height          =   2295
                  Left            =   120
                  TabIndex        =   80
                  Top             =   1680
                  Width           =   7215
                  Begin VB.TextBox txtAnaSocio 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   20
                     Top             =   1155
                     Width           =   4455
                  End
                  Begin VB.CommandButton Command14 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":479EA
                     Style           =   1  'Graphical
                     TabIndex        =   191
                     TabStop         =   0   'False
                     ToolTipText     =   "Elimina riferimento "
                     Top             =   1155
                     Width           =   375
                  End
                  Begin VB.CommandButton Command13 
                     Height          =   315
                     Left            =   480
                     Picture         =   "frmMain.frx":47F74
                     Style           =   1  'Graphical
                     TabIndex        =   17
                     ToolTipText     =   "Ricerca"
                     Top             =   1155
                     Width           =   375
                  End
                  Begin VB.TextBox txtCodiceAnaSocio 
                     Height          =   315
                     Left            =   1800
                     TabIndex        =   19
                     Top             =   1155
                     Width           =   855
                  End
                  Begin VB.TextBox txtAnaCoop 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   16
                     Top             =   480
                     Width           =   4455
                  End
                  Begin VB.CommandButton Command12 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":484FE
                     Style           =   1  'Graphical
                     TabIndex        =   187
                     TabStop         =   0   'False
                     ToolTipText     =   "Elimina riferimento"
                     Top             =   480
                     Width           =   375
                  End
                  Begin VB.CommandButton Command11 
                     Height          =   315
                     Left            =   480
                     Picture         =   "frmMain.frx":48A88
                     Style           =   1  'Graphical
                     TabIndex        =   13
                     ToolTipText     =   "Ricerca"
                     Top             =   480
                     Width           =   375
                  End
                  Begin VB.TextBox txtCodiceAnaCoop 
                     Height          =   315
                     Left            =   1800
                     TabIndex        =   15
                     Top             =   480
                     Width           =   855
                  End
                  Begin VB.CheckBox Check1 
                     Caption         =   "Acquisto"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   23
                     TabStop         =   0   'False
                     Top             =   1800
                     Width           =   1455
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDSocioFatt 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   21
                     TabStop         =   0   'False
                     Top             =   2280
                     Width           =   6855
                     _ExtentX        =   12091
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":49012
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":49060
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":490B6
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDSocio 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   22
                     TabStop         =   0   'False
                     Top             =   2880
                     Width           =   6855
                     _ExtentX        =   12091
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":49110
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4915E
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":491AE
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIDAnagraficaCoop 
                     Height          =   315
                     Left            =   840
                     TabIndex        =   14
                     TabStop         =   0   'False
                     Top             =   480
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIDSocio 
                     Height          =   315
                     Left            =   840
                     TabIndex        =   18
                     TabStop         =   0   'False
                     Top             =   1155
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Socio"
                     Height          =   255
                     Index           =   35
                     Left            =   2640
                     TabIndex        =   194
                     Top             =   960
                     Width           =   4455
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Codice"
                     Height          =   255
                     Index           =   34
                     Left            =   1800
                     TabIndex        =   193
                     Top             =   960
                     Width           =   975
                  End
                  Begin VB.Label Label2 
                     Caption         =   "ID"
                     Height          =   255
                     Index           =   33
                     Left            =   840
                     TabIndex        =   192
                     Top             =   960
                     Width           =   975
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Cooperativa"
                     Height          =   255
                     Index           =   32
                     Left            =   2640
                     TabIndex        =   190
                     Top             =   285
                     Width           =   4455
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Codice"
                     Height          =   255
                     Index           =   31
                     Left            =   1800
                     TabIndex        =   189
                     Top             =   285
                     Width           =   975
                  End
                  Begin VB.Label Label2 
                     Caption         =   "ID"
                     Height          =   255
                     Index           =   30
                     Left            =   840
                     TabIndex        =   188
                     Top             =   280
                     Width           =   975
                  End
               End
               Begin VB.Frame Frame6 
                  Caption         =   "LOTTO DI PRODUZIONE"
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
                  Height          =   1455
                  Left            =   120
                  TabIndex        =   160
                  Top             =   3960
                  Width           =   16575
                  Begin VB.TextBox txtLottoDiConferimento 
                     Height          =   315
                     Left            =   840
                     TabIndex        =   162
                     Top             =   480
                     Width           =   3885
                  End
                  Begin VB.CommandButton Command5 
                     Height          =   315
                     Left            =   480
                     Picture         =   "frmMain.frx":49208
                     Style           =   1  'Graphical
                     TabIndex        =   37
                     ToolTipText     =   "Trova conferimento/Acquisto merce"
                     Top             =   480
                     Width           =   375
                  End
                  Begin VB.CommandButton Command6 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":49792
                     Style           =   1  'Graphical
                     TabIndex        =   36
                     TabStop         =   0   'False
                     ToolTipText     =   "Elimina riferimento al conferimento"
                     Top             =   480
                     Width           =   375
                  End
                  Begin VB.TextBox txtRegioneLotto 
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   12360
                     TabIndex        =   161
                     TabStop         =   0   'False
                     Top             =   480
                     Width           =   1935
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIDLottoCampagna 
                     Height          =   315
                     Left            =   4800
                     TabIndex        =   163
                     TabStop         =   0   'False
                     Top             =   480
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDataCmb.DMTCombo cboRegione 
                     Height          =   315
                     Left            =   14400
                     TabIndex        =   164
                     TabStop         =   0   'False
                     Top             =   480
                     Width           =   2055
                     _ExtentX        =   3625
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
                  Begin DMTDataCmb.DMTCombo cboFamigliaLotto 
                     Height          =   315
                     Left            =   8400
                     TabIndex        =   165
                     Top             =   480
                     Width           =   3855
                     _ExtentX        =   6800
                     _ExtentY        =   556
                     Enabled         =   0   'False
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin DMTDataCmb.DMTCombo cboVarietaLotto 
                     Height          =   315
                     Left            =   6120
                     TabIndex        =   166
                     Top             =   480
                     Width           =   2175
                     _ExtentX        =   3836
                     _ExtentY        =   556
                     Enabled         =   0   'False
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTotaleEttariLotto 
                     Height          =   315
                     Left            =   840
                     TabIndex        =   173
                     Top             =   1080
                     Width           =   1695
                     _Version        =   65536
                     _ExtentX        =   2990
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtResaMinPerHa 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   175
                     Top             =   1080
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtResaMaxPerHa 
                     Height          =   315
                     Left            =   4680
                     TabIndex        =   177
                     Top             =   1080
                     Width           =   2055
                     _Version        =   65536
                     _ExtentX        =   3625
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtResaMaxTotale 
                     Height          =   315
                     Left            =   9240
                     TabIndex        =   179
                     Top             =   1080
                     Width           =   2295
                     _Version        =   65536
                     _ExtentX        =   4048
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtResaMinTotale 
                     Height          =   315
                     Left            =   6840
                     TabIndex        =   181
                     Top             =   1080
                     Width           =   2295
                     _Version        =   65536
                     _ExtentX        =   4048
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQtaUtilizzataLotto 
                     Height          =   315
                     Left            =   11640
                     TabIndex        =   183
                     Top             =   1080
                     Width           =   2295
                     _Version        =   65536
                     _ExtentX        =   4048
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Quantità utilizzata"
                     Height          =   255
                     Index           =   28
                     Left            =   11640
                     TabIndex        =   184
                     Top             =   840
                     Width           =   2295
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Resa massima totale lotto"
                     Height          =   255
                     Index           =   27
                     Left            =   9240
                     TabIndex        =   182
                     Top             =   840
                     Width           =   2295
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Resa minima totale lotto"
                     Height          =   255
                     Index           =   26
                     Left            =   6840
                     TabIndex        =   180
                     Top             =   840
                     Width           =   2415
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Resa massima per Ha"
                     Height          =   255
                     Index           =   25
                     Left            =   4680
                     TabIndex        =   178
                     Top             =   840
                     Width           =   2055
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Resa minima per Ha"
                     Height          =   255
                     Index           =   24
                     Left            =   2640
                     TabIndex        =   176
                     Top             =   840
                     Width           =   1935
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Ha"
                     Height          =   255
                     Index           =   23
                     Left            =   840
                     TabIndex        =   174
                     Top             =   840
                     Width           =   1695
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Varietà/Tipologia"
                     Height          =   255
                     Index           =   0
                     Left            =   6120
                     TabIndex        =   172
                     Top             =   240
                     Width           =   2175
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Famiglia/Tipo coltura"
                     Height          =   255
                     Index           =   0
                     Left            =   8400
                     TabIndex        =   171
                     Top             =   240
                     Width           =   3015
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Identificativo"
                     Height          =   255
                     Index           =   2
                     Left            =   4800
                     TabIndex        =   170
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Codice"
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Index           =   17
                     Left            =   840
                     TabIndex        =   169
                     Top             =   240
                     Width           =   2775
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Regione"
                     Height          =   255
                     Index           =   3
                     Left            =   14400
                     TabIndex        =   168
                     Top             =   240
                     Width           =   2055
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Regione lotto"
                     Height          =   255
                     Index           =   4
                     Left            =   12360
                     TabIndex        =   167
                     Top             =   240
                     Width           =   1455
                  End
               End
               Begin VB.Frame FraTab 
                  Caption         =   "PARAMETRI QUALITATIVI"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   1035
                  Index           =   8
                  Left            =   120
                  TabIndex        =   118
                  Top             =   8280
                  Width           =   16605
                  Begin VB.CommandButton Command8 
                     Height          =   375
                     Left            =   15960
                     Picture         =   "frmMain.frx":49D1C
                     Style           =   1  'Graphical
                     TabIndex        =   156
                     TabStop         =   0   'False
                     ToolTipText     =   "Recupera i prezzi da contratto"
                     Top             =   0
                     Width           =   495
                  End
                  Begin VB.CommandButton Command4 
                     Height          =   375
                     Left            =   15360
                     Picture         =   "frmMain.frx":4A2A6
                     Style           =   1  'Graphical
                     TabIndex        =   153
                     TabStop         =   0   'False
                     ToolTipText     =   "Aggiorna parametri qualitativi"
                     Top             =   0
                     Width           =   495
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtQual01 
                     Height          =   285
                     Left            =   120
                     TabIndex        =   119
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtQual02 
                     Height          =   285
                     Left            =   960
                     TabIndex        =   120
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual03 
                     Height          =   285
                     Left            =   1800
                     TabIndex        =   121
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtQual04 
                     Height          =   285
                     Left            =   2640
                     TabIndex        =   122
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual15 
                     Height          =   285
                     Left            =   11880
                     TabIndex        =   123
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual14 
                     Height          =   285
                     Left            =   11040
                     TabIndex        =   124
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual13 
                     Height          =   285
                     Left            =   10200
                     TabIndex        =   125
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual12 
                     Height          =   285
                     Left            =   9360
                     TabIndex        =   126
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual16 
                     Height          =   285
                     Left            =   12720
                     TabIndex        =   127
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual09 
                     Height          =   285
                     Left            =   6840
                     TabIndex        =   128
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual10 
                     Height          =   285
                     Left            =   7680
                     TabIndex        =   129
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual11 
                     Height          =   285
                     Left            =   8520
                     TabIndex        =   130
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual08 
                     Height          =   285
                     Left            =   6000
                     TabIndex        =   131
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual07 
                     Height          =   285
                     Left            =   5160
                     TabIndex        =   132
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual06 
                     Height          =   285
                     Left            =   4320
                     TabIndex        =   133
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQual05 
                     Height          =   285
                     Left            =   3480
                     TabIndex        =   134
                     Top             =   480
                     Width           =   795
                     _Version        =   65536
                     _ExtentX        =   1402
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "15% +"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   47
                     Left            =   12810
                     TabIndex        =   150
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "5%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   46
                     Left            =   3720
                     TabIndex        =   149
                     Top             =   240
                     Width           =   315
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "6%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   45
                     Left            =   4560
                     TabIndex        =   148
                     Top             =   240
                     Width           =   315
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "7%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   44
                     Left            =   5400
                     TabIndex        =   147
                     Top             =   240
                     Width           =   315
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "8%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   43
                     Left            =   6240
                     TabIndex        =   146
                     Top             =   240
                     Width           =   315
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "9%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   42
                     Left            =   7080
                     TabIndex        =   145
                     Top             =   240
                     Width           =   315
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "10%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   41
                     Left            =   7860
                     TabIndex        =   144
                     Top             =   240
                     Width           =   435
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "11%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   40
                     Left            =   8700
                     TabIndex        =   143
                     Top             =   240
                     Width           =   435
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "12%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   34
                     Left            =   9540
                     TabIndex        =   142
                     Top             =   240
                     Width           =   435
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "13%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   29
                     Left            =   10380
                     TabIndex        =   141
                     Top             =   240
                     Width           =   435
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "14%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   12
                     Left            =   11220
                     TabIndex        =   140
                     Top             =   240
                     Width           =   435
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "15%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   11
                     Left            =   12060
                     TabIndex        =   139
                     Top             =   240
                     Width           =   435
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "1%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   10
                     Left            =   120
                     TabIndex        =   138
                     Top             =   240
                     Width           =   780
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "2%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   9
                     Left            =   960
                     TabIndex        =   137
                     Top             =   240
                     Width           =   780
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "3%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   8
                     Left            =   2010
                     TabIndex        =   136
                     Top             =   240
                     Width           =   315
                  End
                  Begin VB.Label lblDocument 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "4%"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   7
                     Left            =   2850
                     TabIndex        =   135
                     Top             =   240
                     Width           =   315
                  End
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "Pulisci"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   975
                  Left            =   12480
                  Picture         =   "frmMain.frx":4A830
                  Style           =   1  'Graphical
                  TabIndex        =   64
                  Top             =   9600
                  Width           =   4095
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Salva e pulisci modulo"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   975
                  Left            =   6360
                  Picture         =   "frmMain.frx":4BD72
                  Style           =   1  'Graphical
                  TabIndex        =   63
                  Top             =   9600
                  Width           =   4095
               End
               Begin VB.Frame Frame5 
                  Caption         =   "SCARICO"
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
                  Height          =   2895
                  Left            =   120
                  TabIndex        =   87
                  Top             =   5400
                  Width           =   16575
                  Begin VB.CommandButton Command10 
                     Height          =   315
                     Left            =   14520
                     Picture         =   "frmMain.frx":4D2B4
                     Style           =   1  'Graphical
                     TabIndex        =   186
                     TabStop         =   0   'False
                     ToolTipText     =   "Attiva modifica"
                     Top             =   1800
                     Width           =   375
                  End
                  Begin VB.CommandButton Command7 
                     Height          =   315
                     Left            =   16080
                     Picture         =   "frmMain.frx":4D83E
                     Style           =   1  'Graphical
                     TabIndex        =   154
                     TabStop         =   0   'False
                     ToolTipText     =   "Refresh descrizione estesa articolo"
                     Top             =   600
                     Width           =   375
                  End
                  Begin VB.TextBox txtDescrizioneArticolo 
                     Height          =   315
                     Left            =   8400
                     MaxLength       =   255
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   39
                     Top             =   600
                     Width           =   7695
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   38
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   5895
                     _ExtentX        =   10398
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":4DDC8
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4DE17
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":4DE77
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDImballo 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   40
                     TabStop         =   0   'False
                     Top             =   960
                     Width           =   5895
                     _ExtentX        =   10398
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":4DED1
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4DF20
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":4DF7F
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtColliEntrata 
                     Height          =   315
                     Left            =   6120
                     TabIndex        =   41
                     Top             =   1200
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtColliUscita 
                     Height          =   315
                     Left            =   7560
                     TabIndex        =   42
                     TabStop         =   0   'False
                     Top             =   1200
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTaraUnitaria 
                     Height          =   315
                     Left            =   9000
                     TabIndex        =   43
                     Top             =   1200
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   5
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTaraTotale 
                     Height          =   315
                     Left            =   3600
                     TabIndex        =   48
                     Top             =   1800
                     Width           =   1815
                     _Version        =   65536
                     _ExtentX        =   3201
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTaraCamion 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   45
                     Top             =   1800
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTaraTotaleImballo 
                     Height          =   315
                     Left            =   10440
                     TabIndex        =   44
                     Top             =   1200
                     Width           =   1695
                     _Version        =   65536
                     _ExtentX        =   2990
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   5
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtPesoLordo 
                     Height          =   315
                     Left            =   1680
                     TabIndex        =   46
                     Top             =   1800
                     Width           =   1815
                     _Version        =   65536
                     _ExtentX        =   3201
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtScarto 
                     Height          =   315
                     Left            =   7440
                     TabIndex        =   47
                     Top             =   1800
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtPesoNetto 
                     Height          =   315
                     Left            =   5520
                     TabIndex        =   49
                     Top             =   1800
                     Width           =   1815
                     _Version        =   65536
                     _ExtentX        =   3201
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtPercRidPesoNetto 
                     Height          =   315
                     Left            =   9000
                     TabIndex        =   50
                     Top             =   1800
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQtaFatturazione 
                     Height          =   315
                     Left            =   10080
                     TabIndex        =   51
                     TabStop         =   0   'False
                     Top             =   1800
                     Width           =   2055
                     _Version        =   65536
                     _ExtentX        =   3625
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtPrezzoDaContratto 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   53
                     Top             =   2400
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   5
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtPrezzoDiFatturazione 
                     Height          =   315
                     Left            =   6240
                     TabIndex        =   56
                     Top             =   2400
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   5
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTotaleRiga 
                     Height          =   315
                     Left            =   8280
                     TabIndex        =   57
                     Top             =   2400
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIndiceDiVariazione 
                     Height          =   315
                     Left            =   12240
                     TabIndex        =   59
                     Top             =   2400
                     Width           =   1575
                     _Version        =   65536
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   5
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDataCmb.DMTCombo cboIvaArticolo 
                     Height          =   315
                     Left            =   12240
                     TabIndex        =   52
                     TabStop         =   0   'False
                     Top             =   1800
                     Width           =   2295
                     _ExtentX        =   4048
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
                  Begin DMTEDITNUMLib.dmtNumber txtPrezzoContrattoMin 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   54
                     Top             =   2400
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   5
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtPrezzoContrattoMax 
                     Height          =   315
                     Left            =   4200
                     TabIndex        =   55
                     Top             =   2400
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648384
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   5
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIndiceDiVariazioneEff 
                     Height          =   315
                     Left            =   13920
                     TabIndex        =   60
                     TabStop         =   0   'False
                     Top             =   2400
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   0
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIndice 
                     Height          =   315
                     Left            =   15480
                     TabIndex        =   61
                     TabStop         =   0   'False
                     Top             =   2400
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   0
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDataCmb.DMTCombo cboVarietaArticolo 
                     Height          =   315
                     Left            =   6120
                     TabIndex        =   157
                     Top             =   600
                     Width           =   2175
                     _ExtentX        =   3836
                     _ExtentY        =   556
                     Enabled         =   0   'False
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIndiceDiVariazione100 
                     Height          =   315
                     Left            =   10320
                     TabIndex        =   58
                     Top             =   2400
                     Width           =   1815
                     _Version        =   65536
                     _ExtentX        =   3201
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label Label2 
                     Caption         =   "% di var. base 100"
                     Height          =   255
                     Index           =   29
                     Left            =   10320
                     TabIndex        =   185
                     ToolTipText     =   "Indice di variazione effettivo"
                     Top             =   2160
                     Width           =   1815
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Varietà/Tipologia"
                     Height          =   255
                     Index           =   1
                     Left            =   6120
                     TabIndex        =   158
                     Top             =   360
                     Width           =   2295
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Indice"
                     Height          =   255
                     Index           =   22
                     Left            =   15480
                     TabIndex        =   117
                     ToolTipText     =   "Indice di variazione effettivo"
                     Top             =   2160
                     Width           =   855
                  End
                  Begin VB.Label Label2 
                     Caption         =   "% di var. eff."
                     Height          =   255
                     Index           =   21
                     Left            =   13920
                     TabIndex        =   115
                     ToolTipText     =   "Indice di variazione effettivo"
                     Top             =   2160
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Prezzo massimo"
                     Height          =   255
                     Index           =   20
                     Left            =   4200
                     TabIndex        =   114
                     Top             =   2160
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Prezzo minimo"
                     Height          =   255
                     Index           =   19
                     Left            =   2160
                     TabIndex        =   113
                     Top             =   2160
                     Width           =   1695
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Aliquota I.V.A."
                     Height          =   255
                     Index           =   1
                     Left            =   12240
                     TabIndex        =   112
                     Top             =   1560
                     Width           =   2295
                  End
                  Begin VB.Label Label2 
                     Caption         =   "% di var."
                     Height          =   255
                     Index           =   15
                     Left            =   12240
                     TabIndex        =   103
                     ToolTipText     =   "Indice di variazione effettivo"
                     Top             =   2160
                     Width           =   1095
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Totale riga"
                     Height          =   255
                     Index           =   14
                     Left            =   8280
                     TabIndex        =   102
                     Top             =   2160
                     Width           =   1935
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Prezzo di fatturazione"
                     Height          =   255
                     Index           =   13
                     Left            =   6240
                     TabIndex        =   101
                     Top             =   2160
                     Width           =   2055
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Prezzo da contratto"
                     Height          =   255
                     Index           =   12
                     Left            =   120
                     TabIndex        =   100
                     Top             =   2160
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Peso di fatturazione"
                     Height          =   255
                     Index           =   11
                     Left            =   10080
                     TabIndex        =   99
                     Top             =   1560
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Caption         =   "% Rid."
                     Height          =   255
                     Index           =   9
                     Left            =   9000
                     TabIndex        =   98
                     Top             =   1560
                     Width           =   735
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Peso netto"
                     Height          =   255
                     Index           =   8
                     Left            =   5520
                     TabIndex        =   97
                     Top             =   1560
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Scarto"
                     Height          =   255
                     Index           =   7
                     Left            =   7440
                     TabIndex        =   96
                     Top             =   1560
                     Width           =   1455
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Peso lordo"
                     Height          =   255
                     Index           =   6
                     Left            =   1680
                     TabIndex        =   95
                     Top             =   1560
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara totale imballo"
                     Height          =   255
                     Index           =   5
                     Left            =   10440
                     TabIndex        =   94
                     Top             =   960
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara camion"
                     Height          =   255
                     Index           =   4
                     Left            =   120
                     TabIndex        =   93
                     Top             =   1560
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara complessiva"
                     Height          =   255
                     Index           =   3
                     Left            =   3600
                     TabIndex        =   92
                     Top             =   1560
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara unitaria"
                     Height          =   255
                     Index           =   2
                     Left            =   9000
                     TabIndex        =   91
                     Top             =   960
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Colli in uscita"
                     Height          =   255
                     Index           =   1
                     Left            =   7560
                     TabIndex        =   90
                     Top             =   960
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Colli in entrata"
                     Height          =   255
                     Index           =   0
                     Left            =   6120
                     TabIndex        =   89
                     Top             =   960
                     Width           =   1575
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Descrizione estesa"
                     Height          =   255
                     Index           =   0
                     Left            =   8400
                     TabIndex        =   88
                     Top             =   360
                     Width           =   4455
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "RIFERIMENTI DOCUMENTO DI TRASPORTO"
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
                  Height          =   2295
                  Left            =   7440
                  TabIndex        =   81
                  Top             =   1680
                  Width           =   9255
                  Begin VB.CommandButton Command9 
                     Height          =   315
                     Left            =   8760
                     Picture         =   "frmMain.frx":4DFD9
                     Style           =   1  'Graphical
                     TabIndex        =   159
                     TabStop         =   0   'False
                     ToolTipText     =   "Refresh descrizione estesa articolo"
                     Top             =   1080
                     Visible         =   0   'False
                     Width           =   375
                  End
                  Begin VB.TextBox txtDescrizioneDocumento 
                     Height          =   315
                     Left            =   4320
                     Locked          =   -1  'True
                     TabIndex        =   151
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   4845
                  End
                  Begin VB.CommandButton cmdEliminaRifLetInt 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":4E563
                     Style           =   1  'Graphical
                     TabIndex        =   31
                     TabStop         =   0   'False
                     ToolTipText     =   "Elimina riferimento lettera intento"
                     Top             =   1635
                     Width           =   375
                  End
                  Begin VB.CommandButton cmdLetteraIntento 
                     Height          =   315
                     Left            =   480
                     Picture         =   "frmMain.frx":4EAED
                     Style           =   1  'Graphical
                     TabIndex        =   32
                     TabStop         =   0   'False
                     ToolTipText     =   "Lettere di intento del cliente"
                     Top             =   1635
                     Width           =   375
                  End
                  Begin VB.TextBox txtNLetteraIntento 
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   840
                     TabIndex        =   33
                     TabStop         =   0   'False
                     Top             =   1635
                     Width           =   1455
                  End
                  Begin VB.TextBox txtNumeroDDT 
                     Height          =   315
                     Left            =   2880
                     TabIndex        =   26
                     Top             =   480
                     Width           =   1335
                  End
                  Begin VB.TextBox txtNumeroCertificato 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   24
                     Top             =   480
                     Width           =   1335
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataCertificato 
                     Height          =   315
                     Left            =   1560
                     TabIndex        =   25
                     Top             =   480
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDSezionale 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   28
                     TabStop         =   0   'False
                     Top             =   840
                     Width           =   4215
                     _ExtentX        =   7435
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":4F077
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4F0C5
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":4F119
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataDocumento 
                     Height          =   315
                     Left            =   5640
                     TabIndex        =   29
                     TabStop         =   0   'False
                     Top             =   480
                     Visible         =   0   'False
                     Width           =   1425
                     _Version        =   65536
                     _ExtentX        =   2514
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtNumeroDocumento 
                     Height          =   315
                     Left            =   7200
                     TabIndex        =   30
                     TabStop         =   0   'False
                     Top             =   480
                     Visible         =   0   'False
                     Width           =   1890
                     _Version        =   65536
                     _ExtentX        =   3334
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataDDT 
                     Height          =   315
                     Left            =   4320
                     TabIndex        =   27
                     Top             =   480
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboIvaCliente 
                     Height          =   315
                     Left            =   4320
                     TabIndex        =   35
                     TabStop         =   0   'False
                     Top             =   1635
                     Width           =   3615
                     _ExtentX        =   6376
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
                  Begin DMTEDITNUMLib.dmtNumber txtIDLetteraIntento 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   109
                     TabStop         =   0   'False
                     Top             =   1440
                     Visible         =   0   'False
                     Width           =   375
                     _Version        =   65536
                     _ExtentX        =   661
                     _ExtentY        =   450
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataLetteraIntento 
                     Height          =   315
                     Left            =   2280
                     TabIndex        =   34
                     TabStop         =   0   'False
                     Top             =   1635
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Documento di trasporto collegato"
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Index           =   18
                     Left            =   4320
                     TabIndex        =   152
                     Top             =   840
                     Width           =   4695
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Esenzione I.V.A."
                     Height          =   255
                     Index           =   2
                     Left            =   4320
                     TabIndex        =   111
                     Top             =   1440
                     Width           =   3615
                  End
                  Begin VB.Label lblLetteraIntento 
                     Caption         =   "Lettera d'intento"
                     Height          =   255
                     Left            =   840
                     TabIndex        =   110
                     Top             =   1455
                     Width           =   1575
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Data D.D.T."
                     Height          =   255
                     Index           =   6
                     Left            =   4320
                     TabIndex        =   105
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Label Label3 
                     Caption         =   "N° D.D.T."
                     Height          =   255
                     Index           =   5
                     Left            =   2880
                     TabIndex        =   104
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Data"
                     Height          =   255
                     Index           =   4
                     Left            =   5640
                     TabIndex        =   86
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1215
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Numero"
                     Height          =   255
                     Index           =   3
                     Left            =   7200
                     TabIndex        =   85
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1935
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Data"
                     Height          =   255
                     Index           =   2
                     Left            =   1560
                     TabIndex        =   84
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Label Label3 
                     Caption         =   "N° Certificato"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   82
                     Top             =   240
                     Width           =   1335
                  End
               End
               Begin VB.Frame Frame2 
                  Caption         =   "INDUSTRIA/DESTINATARIO"
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
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   74
                  Top             =   120
                  Width           =   16575
                  Begin VB.CommandButton cmdEliminaRifContrattoRiga 
                     Height          =   315
                     Left            =   9000
                     Picture         =   "frmMain.frx":4F173
                     Style           =   1  'Graphical
                     TabIndex        =   9
                     TabStop         =   0   'False
                     ToolTipText     =   "Elimina riferimento al conferimento"
                     Top             =   1050
                     Width           =   375
                  End
                  Begin VB.CommandButton cmdTrovaContrattoRiga 
                     Height          =   315
                     Left            =   9360
                     Picture         =   "frmMain.frx":4F6FD
                     Style           =   1  'Graphical
                     TabIndex        =   10
                     TabStop         =   0   'False
                     ToolTipText     =   "Trova conferimento/Acquisto merce"
                     Top             =   1050
                     Width           =   375
                  End
                  Begin VB.TextBox txtDescrizioneContrattoRiga 
                     Height          =   315
                     Left            =   11040
                     Locked          =   -1  'True
                     TabIndex        =   12
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   5295
                  End
                  Begin VB.TextBox txtDescrizioneContratto 
                     Height          =   315
                     Left            =   2160
                     Locked          =   -1  'True
                     TabIndex        =   8
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   6735
                  End
                  Begin VB.CommandButton cmdTrovaContratto 
                     Height          =   315
                     Left            =   480
                     Picture         =   "frmMain.frx":4FC87
                     Style           =   1  'Graphical
                     TabIndex        =   6
                     TabStop         =   0   'False
                     ToolTipText     =   "Trova conferimento/Acquisto merce"
                     Top             =   1050
                     Width           =   375
                  End
                  Begin VB.CommandButton cmdEliminaRifContratto 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":50211
                     Style           =   1  'Graphical
                     TabIndex        =   5
                     TabStop         =   0   'False
                     ToolTipText     =   "Elimina riferimento al conferimento"
                     Top             =   1050
                     Width           =   375
                  End
                  Begin DMTDataCmb.DMTCombo cboAltroSito 
                     Height          =   315
                     Left            =   7080
                     TabIndex        =   1
                     TabStop         =   0   'False
                     Top             =   450
                     Width           =   3735
                     _ExtentX        =   6588
                     _ExtentY        =   556
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin DMTDATETIMELib.dmtTime txtOraTrasporto 
                     Height          =   315
                     Left            =   12120
                     TabIndex        =   3
                     TabStop         =   0   'False
                     Top             =   450
                     Width           =   855
                     _Version        =   65536
                     _ExtentX        =   1508
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataTrasporto 
                     Height          =   315
                     Left            =   10920
                     TabIndex        =   2
                     TabStop         =   0   'False
                     Top             =   450
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIDContratto 
                     Height          =   315
                     Left            =   840
                     TabIndex        =   7
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIDContrattoRiga 
                     Height          =   315
                     Left            =   9720
                     TabIndex        =   11
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DmtCodDescCtl.DmtCodDesc cdAnagrafica 
                     Height          =   585
                     Left            =   240
                     TabIndex        =   108
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   5625
                     _ExtentX        =   9922
                     _ExtentY        =   1032
                     PropCodice      =   $"frmMain.frx":5079B
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":507FF
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":50850
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin DMTDataCmb.DMTCombo cboVettore 
                     Height          =   315
                     Left            =   13080
                     TabIndex        =   4
                     TabStop         =   0   'False
                     Top             =   480
                     Width           =   3255
                     _ExtentX        =   5741
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
                  Begin DmtSearchAccount2.DmtSearchACS2 ACSCliente 
                     Height          =   600
                     Left            =   120
                     TabIndex        =   0
                     Top             =   240
                     Width           =   6870
                     _ExtentX        =   12118
                     _ExtentY        =   1058
                     WidthCode       =   800
                     WidthDescription=   3450
                     WidthSecondDescription=   2500
                     Object.Visible         =   0   'False
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     HideLeaf        =   0   'False
                     BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     CaptionDescription=   "Cognome o ragione sociale"
                     CaptionCode     =   "Codice"
                     OnlyAccounts    =   -1  'True
                  End
                  Begin VB.Label Label8 
                     Caption         =   "Vettore"
                     Height          =   255
                     Index           =   0
                     Left            =   13080
                     TabIndex        =   116
                     Top             =   240
                     Width           =   3375
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Identificativo"
                     Height          =   255
                     Index           =   16
                     Left            =   9720
                     TabIndex        =   107
                     Top             =   840
                     Width           =   1215
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Descrizione riga contratto"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Left            =   11040
                     MouseIcon       =   "frmMain.frx":508AA
                     MousePointer    =   99  'Custom
                     TabIndex        =   106
                     Top             =   840
                     Width           =   5295
                  End
                  Begin VB.Label lblCollegamentoContratto 
                     Caption         =   "Riferimento contratto"
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Left            =   2160
                     MouseIcon       =   "frmMain.frx":50BB4
                     MousePointer    =   99  'Custom
                     TabIndex        =   79
                     Top             =   840
                     Width           =   6735
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Identificativo"
                     Height          =   255
                     Index           =   10
                     Left            =   840
                     TabIndex        =   78
                     Top             =   840
                     Width           =   1215
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Data trasp."
                     Height          =   255
                     Index           =   6
                     Left            =   10920
                     TabIndex        =   77
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Ora trasp."
                     Height          =   255
                     Index           =   7
                     Left            =   12120
                     TabIndex        =   76
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Destinazione "
                     Height          =   255
                     Index           =   0
                     Left            =   7080
                     TabIndex        =   75
                     Top             =   240
                     Width           =   3735
                  End
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Salva e rimani nella cooperativa"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   975
                  Left            =   120
                  Picture         =   "frmMain.frx":50EBE
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  Top             =   9600
                  Width           =   4215
               End
               Begin VB.Label lblInfo 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   155
                  Top             =   9360
                  Width           =   16575
               End
            End
         End
         Begin DmtGridCtl.DmtGrid BrwMain 
            Height          =   1095
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1931
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
            ColumnsHeaderHeight=   20
         End
      End
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   7875
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   13891
         BackColor       =   -2147483643
         ForeColor       =   -2147483630
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
      Begin DmtPrnDlgCtl.DMTDialog DmtPrnDlg 
         Left            =   0
         Top             =   1290
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   480
         ScaleHeight     =   4935
         ScaleWidth      =   60
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin DMTWheelCtrl.SpareWheel SpareWheel 
         Left            =   465
         Top             =   660
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   4695
         Left            =   2100
         MousePointer    =   9  'Size W E
         Top             =   0
         Width           =   60
      End
   End
   Begin VB.Label Label3 
      Caption         =   "N° Certificato"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'L'applicazione corrente
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1
'Il processo corrente
Private m_Process As DMTRunAppLib.Process
'Il tipo di documento corrente
Private m_DocType As DmtDocManLib.DBFormDocType
'Il documento corrente
Private WithEvents m_Document As DmtDocManLib.DBFormDocument
Attribute m_Document.VB_VarHelpID = -1
'La vista tabellare attiva
Private m_ActiveTableView As DmtDocManLib.TableView
'Il filtro attivo
Private m_ActiveFilter As DmtDocManLib.Filter
'Il report da stampare
Private m_Report As DmtDocManLib.Report
'La collezione dei campi del documento
'collegati ai controlli del form
Private m_FormFields As FormFields
'Il campo con la proprietà TabIndex uguale a 0
Private m_ControlTabIndex0 As Control
'La variabile  m_Semaphore mantiene un riferimento all'oggetto
'Semaphore che gestisce i conflitti di multiutenza
Private m_Semaphore As Semaforo.dmtSemaphore
'Indica se all'evento KeyPress del Form il tasto deve essere annullato
Private m_EatKey As Boolean
'Indica se l'utente ha modificato uno dei campi del documento
Private m_Changed As Boolean
'Indica se i valori dei campi del documento sono stati salvati
Private m_Saved As Boolean
'Indica se è in corso la definizione di una ricerca
Private m_Search As Boolean
'Indica se uno dei filtri è stato selezionato
Private m_FilterSelected As Boolean
'Indica lo stato di visibilità della vista tabellare
'prima dell'inizio della fase di esecuzione della
'anteprima di stampa
Private m_TabMode As Boolean
'Indica se si sta muovendo lo splitter
Private m_SplitterMoving As Boolean
'Nome dell'eventuale database esteso
Private m_ExtendedDatabase As String
'Processo "Shell su evento OnSave" nome del campo collegato
Private m_LinkedField As String
'Handle della finestra della anteprima di stampa
Private m_PreviewWindowHandle As Long
'Flag che permette l'esecuzione di Form_Activate soltanto all'avvio del programma
Private m_bOnFirstTime As Boolean
'Impedisce il Reposition della browse.
Private m_bAvoidReposition As Boolean
'Consente l'esecuzione del codice contenuto in BrwMain_OnChangeGuiMode()
Private bEnableGuiEvent As Boolean
'Indica se è stato attivato un link
Private m_LinkActive As Boolean

'cbcx
'Oggetto adibito alla gestione del processo On_Extend
'Private m_ExtendApplication As DmtExtendAppLib.ExtendApplication


'rif1
'L'oggetto per la gestione dei sottodocumenti
'RICAMBI
Private WithEvents m_DocumentsLink As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink.VB_VarHelpID = -1
'SERVIZI
Private WithEvents m_DocumentsLink1 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink1.VB_VarHelpID = -1
'FASI INTERVENTO
Private WithEvents m_DocumentsLink2 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink2.VB_VarHelpID = -1
'ALTRE SPESE
Private WithEvents m_DocumentsLink3 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink3.VB_VarHelpID = -1
'CONDIZIONI CLIENTE X INTERVENTO
Private WithEvents m_DocumentsLink4 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink4.VB_VarHelpID = -1

'Costanti che rappresentano le modalità di visualizzazione
Private Enum neVisualModality
    Insert          'Modalità INSERIMENTO
    Modify          'Modalità VARIAZIONE
    Find            'Modalità TROVA
    Browse          'Modalità ELENCO
    Preview         'Modalità ANTEPRIMA
End Enum

'Costanti usate da SetStatus4Modality per l'apertura/chiusura dell'anteprima di stampa
Private Enum nePreviewModality
    OpenPrw
    ClosePrw
End Enum

Private m_iNumeroCopieDefault As Integer
Private m_OrientamentoDefault As OrientationConsts


'----- Oggetti e variabili per la gestione del riquadro attività -----------
'***Reports                                                                -
Private WithEvents oReportsActivity As DmtActBoxLib.ReportsActivity       '-
Attribute oReportsActivity.VB_VarHelpID = -1
'***Filtri                                                                 -
Private WithEvents oFiltersActivity As DmtActBoxLib.FiltersActivity       '-
Attribute oFiltersActivity.VB_VarHelpID = -1
'***Viste tabellari                                                        -
Private WithEvents oTableViewsActivity As DmtActBoxLib.TableViewsActivity '-
Attribute oTableViewsActivity.VB_VarHelpID = -1
'***Esportazioni                                                           -
Private oExportActivity As DmtActBoxLib.ExportActivity                    '-
'***Supporto tecnico                                                       -
Private oSupportActivity As DmtActBoxLib.SupportActivity                  '-
'***Nome dell'attività predefinita del riquadro attività                   -
Private m_DefaultActivity As String                                       '-
'---------------------------------------------------------------------------

Public bNotReturnValue As Boolean

'///////////////////////////////////////////////////////////////////////////////////
' ATTENZIONE:
' Occorre impostare questa costante!
' (ed eventualmente personalizzare il codice della funzione Caption2Display
'///////////////////////////////////////////////////////////////////////////////////
' Costante che identifica il campo più significativo del documento, il cui valore
' verrà visualizzato nella Caption del Form ed in quei messaggi in cui è mostrato
' il contenuto del campo principale del documento attivo.
' La costante può essere una stringa tipo "NomeCampo" o un intero che funge da indice
' nella collection m_Document.Fields().
'(Se l'applicazione può essere chiamata da un link occorre impostare anche la variabile
'sMessage1 presente nel metodo FormUnload.)
Private Const CAMPO_PER_CAPTION = "Anagrafica"

Private rsGriglia As ADODB.Recordset

'Versione del controllo ActiveBar
Private Const BARMENUVERSION = "3.0"
'Variabile per la gestione degli shortcut del Menu
Private aryShortCut(1) As New ActiveBar3LibraryCtl.ShortCut

Private bloading As Boolean


Public LINK_VARIETA_ART_CONTRATTO As Long
Public LINK_FAMIGLIA_ART_CONTRATTO As Long

'Private NuovoRicambio As Integer
'Private NuovoServizio As Integer


Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property



'**+
'Nome: ChangeStringsLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge le stringe dal file di risorse per gestire l'opzione multilingue.
'Qui vanno inserite tutte le stringhe aggiunte in frmMain solo se si vuole
'gestire l'opzione multilingua
'**/
Public Sub ChangeStringsLanguage()
    '//////////////////////////////////////////////////////////////////////////////
    'ATTENZIONE
    'Inserire qui il codice per la lettura dal file di risorse di tutte le stringhe
    'per le quali si vuole gestire l'opzione multilingue.
    '//////////////////////////////////////////////////////////////////////////////
End Sub

'**+
'Nome: ChangeToolBarLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge dal file di risorse le stringe delle ToolTipText e dei suggerimenti da visualizzare
'sulla Statusbar per gestire l'opzione multilingue
'**/
Public Sub ChangeToolBarLanguage()

    'New
    BarMenu.Bands("Standard").Tools("New").ToolTipText = GetToolTipText4ToolBar("New")
    BarMenu.Bands("Standard").Tools("New").Description = GetDescription4StatusBar("New")
    
    'Save
    BarMenu.Bands("Standard").Tools("Save").ToolTipText = GetToolTipText4ToolBar("Save")
    BarMenu.Bands("Standard").Tools("Save").Description = GetDescription4StatusBar("Save")

    'Print
    BarMenu.Bands("Standard").Tools("Print").ToolTipText = GetToolTipText4ToolBar("Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")

    'PrePrint
    BarMenu.Bands("Standard").Tools("PrePrint").ToolTipText = GetToolTipText4ToolBar("PrePrint")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")

    'Cut
    BarMenu.Bands("Standard").Tools("Cut").ToolTipText = GetToolTipText4ToolBar("Cut")
    BarMenu.Bands("Standard").Tools("Cut").Description = GetDescription4StatusBar("Cut")

    'Copy
    BarMenu.Bands("Standard").Tools("Copy").ToolTipText = GetToolTipText4ToolBar("Copy")
    BarMenu.Bands("Standard").Tools("Copy").Description = GetDescription4StatusBar("Copy")

    'Paste
    BarMenu.Bands("Standard").Tools("Paste").ToolTipText = GetToolTipText4ToolBar("Paste")
    BarMenu.Bands("Standard").Tools("Paste").Description = GetDescription4StatusBar("Paste")

    'Delete
    BarMenu.Bands("Standard").Tools("Delete").ToolTipText = GetToolTipText4ToolBar("Delete")
    BarMenu.Bands("Standard").Tools("Delete").Description = GetDescription4StatusBar("Delete")

    'Clear
    BarMenu.Bands("Standard").Tools("Clear").ToolTipText = GetToolTipText4ToolBar("Clear")
    BarMenu.Bands("Standard").Tools("Clear").Description = GetDescription4StatusBar("Clear")

    'NewSearch
    BarMenu.Bands("Standard").Tools("NewSearch").ToolTipText = GetToolTipText4ToolBar("NewSearch")
    BarMenu.Bands("Standard").Tools("NewSearch").Description = GetDescription4StatusBar("NewSearch")

    'ExecuteSearch
    BarMenu.Bands("Standard").Tools("ExecuteSearch").ToolTipText = GetToolTipText4ToolBar("ExecuteSearch")
    BarMenu.Bands("Standard").Tools("ExecuteSearch").Description = GetDescription4StatusBar("ExecuteSearch")

    'ChangeView
    BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").ToolTipText = GetToolTipText4ToolBar("ChangeView")

    'SearchPrevious
    BarMenu.Bands("Standard").Tools("SearchPrevious").ToolTipText = GetToolTipText4ToolBar("SearchPrevious")
    BarMenu.Bands("Standard").Tools("SearchPrevious").Description = GetDescription4StatusBar("SearchPrevious")

    'SearchNext
    BarMenu.Bands("Standard").Tools("SearchNext").ToolTipText = GetToolTipText4ToolBar("SearchNext")
    BarMenu.Bands("Standard").Tools("SearchNext").Description = GetDescription4StatusBar("SearchNext")

    'ExportWord
    BarMenu.Bands("Band_Export").Tools("ExportWord").ToolTipText = GetToolTipText4ToolBar("ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")

    'ExportExcel
    BarMenu.Bands("Band_Export").Tools("ExportExcel").ToolTipText = GetToolTipText4ToolBar("ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")

    'ExportHtml
    BarMenu.Bands("Band_Export").Tools("ExportHtml").ToolTipText = GetToolTipText4ToolBar("ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")

    'ExportPDF
    BarMenu.Bands("Band_Export").Tools("ExportPDF").ToolTipText = GetToolTipText4ToolBar("ExportPDF")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Description = GetDescription4StatusBar("ExportPDF")

    BarMenu.RecalcLayout
End Sub


'**+
'Nome: ChangeMenuLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge dal file di risorse le stringe delle Caption e dei suggerimenti da visualizzare
'sulla Statusbar per gestire l'opzione multilingue
'**/
Public Sub ChangeMenuLanguage()

    '--- Menu PopUp del pulsante "ChangeView" della Toolbar ---
    'ChangeView - Form
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'ChangeView - Tabella
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'ChangeView - Filtro
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    '---                           ---                      ---
    

    'File
    BarMenu.Bands("Band_Menu").Tools("File").Caption = GetCaption4MenuBar("File")
    BarMenu.Bands("Band_Menu").Tools("File").Description = GetDescription4StatusBar("File")

    'File-New
    BarMenu.Bands("Band_File").Tools("Mnu_New").Caption = GetCaption4MenuBar("Mnu_New")
    BarMenu.Bands("Band_File").Tools("Mnu_New").Description = GetDescription4StatusBar("Mnu_New")
    
    'File-Save
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Description = GetDescription4StatusBar("Mnu_Save")
    
    'File-PrePrint
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Caption = GetCaption4MenuBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    
    'File-Print
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    
    'File-Exit
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Caption = GetCaption4MenuBar("Mnu_Exit")
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Description = GetDescription4StatusBar("Mnu_Exit")
    
    'Edit
    BarMenu.Bands("Band_Menu").Tools("Edit").Caption = GetCaption4MenuBar("Edit")
    BarMenu.Bands("Band_Menu").Tools("Edit").Description = GetDescription4StatusBar("Edit")
    
    'Edit-Delete
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = GetCaption4MenuBar("Mnu_Delete")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Description = GetDescription4StatusBar("Mnu_Delete")
    
    'Edit-Clear
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Caption = GetCaption4MenuBar("Mnu_Clear")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Description = GetDescription4StatusBar("Mnu_Clear")
    
    'Edit-Cut
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Caption = GetCaption4MenuBar("Mnu_Cut")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Description = GetDescription4StatusBar("Mnu_Cut")
    
    'Edit-Copy
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Caption = GetCaption4MenuBar("Mnu_Copy")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Description = GetDescription4StatusBar("Mnu_Copy")
    
    'Edit-Paste
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Caption = GetCaption4MenuBar("Mnu_Paste")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Description = GetDescription4StatusBar("Mnu_Paste")
    
    'Edit-NewSearch
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Caption = GetCaption4MenuBar("Mnu_NewSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Description = GetDescription4StatusBar("Mnu_NewSearch")
    
    'Edit-ExecuteSearch
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Caption = GetCaption4MenuBar("Mnu_ExecuteSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Description = GetDescription4StatusBar("Mnu_ExecuteSearch")
    
    'Edit-SearchPrevious
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Caption = GetCaption4MenuBar("Mnu_SearchPrevious")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Description = GetDescription4StatusBar("Mnu_SearchPrevious")
    
    'Edit-SearchNext
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Caption = GetCaption4MenuBar("Mnu_SearchNext")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Description = GetDescription4StatusBar("Mnu_SearchNext")
    
    'View
    BarMenu.Bands("Band_Menu").Tools("View").Caption = GetCaption4MenuBar("View")
    BarMenu.Bands("Band_Menu").Tools("View").Description = GetDescription4StatusBar("View")
    
    'View-FormView
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'View-TableView
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'View-SearchFilter
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    'View-Folders
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Caption = GetCaption4MenuBar("Mnu_Folders")
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Description = GetDescription4StatusBar("Mnu_Folders")
    
    'View-ToolBar
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Caption = GetCaption4MenuBar("Mnu_ToolBar")
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Description = GetDescription4StatusBar("Mnu_ToolBar")
    
    'Tools
    BarMenu.Bands("Band_Menu").Tools("Tools").Caption = GetCaption4MenuBar("Tools")
    BarMenu.Bands("Band_Menu").Tools("Tools").Description = GetDescription4StatusBar("Tools")
    
    'Tools-Export
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Caption = GetCaption4MenuBar("Mnu_Export")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Description = GetDescription4StatusBar("Mnu_Export")
    
    'Tools-Options
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Caption = GetCaption4MenuBar("Mnu_Options")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Description = GetDescription4StatusBar("Mnu_Options")
    
    'Tools-Export-ExportWord
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("Mnu_ExportWord")
    
    'Tools-Export-ExportExcel
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("Mnu_ExportExcel")
    
    'Tools-Export-ExportHtml
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("Mnu_ExportHtml")

    'Tools-Export-ExportPDF
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Caption = GetCaption4MenuBar("Mnu_ExportPDF")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Description = GetDescription4StatusBar("Mnu_ExportPDF")

    'Help-HelpOnLine
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Caption = GetCaption4MenuBar("Mnu_HelpOnLine")
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Description = GetDescription4StatusBar("Mnu_HelpOnLine")
    
    'Help-Arg
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Caption = GetCaption4MenuBar("Mnu_Arg")
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Description = GetDescription4StatusBar("Mnu_Arg")
    
    'Help-Web
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Caption = GetCaption4MenuBar("Mnu_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Description = GetDescription4StatusBar("Mnu_Web")
    
    
    'Help-Info
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Caption = GetCaption4MenuBar("Mnu_Info")
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Description = GetDescription4StatusBar("Mnu_Info")
    
    'Help-Agg_Web
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Caption = GetCaption4MenuBar("Mnu_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Description = GetDescription4StatusBar("Mnu_Agg_Web")
    
    'Help-Info
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Caption = GetCaption4MenuBar("Mnu_Info")
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Description = GetDescription4StatusBar("Mnu_Info")
    
    'PopUp-RunApplication
    BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").Caption = GetCaption4MenuBar("Mnu_RunApplication")
    
    'PopUp-SearchObject
    BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").Caption = GetCaption4MenuBar("Mnu_SearchObject")
    
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: SetStatusBarVisibility
'
'Parametri:Boolean che valorizzerà la proprietà Visible della StatusBar
'
'Valori di ritorno:
'
'Funzionalità:
'Su richiesta di frmOption, Mostra/Nasconde la Statusbar
'**/
Public Sub SetStatusBarVisibility(ByVal bVisible As Boolean)
    stbStatusbar.Visible = bVisible
End Sub

'**+
'Nome: SetToolBarIcons
'
'Parametri:
'LargeIcons - Il tipo di icona da usare per i bottoni,
'grandi o piccole
'
'Valori di ritorno:
'
'Funzionalità:
'Cambia il tipo di icona della ToolBar standard
'**/
Public Sub SetToolBarIcons(ByVal LargeIcons As Boolean)
    Dim iPicture As Integer

    BarMenu.LargeIcons = LargeIcons
    If LargeIcons Then
        BarMenu.Bands("Standard").Tools("New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Export").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB32), &HC0C0C0
        
        'cbc - L'icona del pulsante "ChangeView" dipende dalla modalità attuale
        iPicture = IIf(BrwMain.Visible, IDB_STD_FORM32, IDB_STD_GRID32)
        BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
        
        BarMenu.LargeIcons = False
    Else
        BarMenu.Bands("Standard").Tools("New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Export").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB16), &HC0C0C0
    
        'L'icona del pulsante "ChangeView" dipende dalla modalità attuale
        iPicture = IIf(BrwMain.Visible, IDB_STD_FORM16, IDB_STD_GRID16)
        BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
    End If
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: SetVisibilityIDFields
'
'Parametri: Optional IDVisible As Variant (Boolean)
'           Se IDVisible è presente (chiamata da frmOption) viene usato il suo valore
'           per settare la visibilità dei campi ID, altrimenti viene letta l'impostazione
'           del registry
'
'Valori di ritorno:
'
'Funzionalità: Mostra/Nasconde i campi ID della Browse
'**/
Public Sub SetVisibilityIDFields(Optional ByVal IDVisible As Variant)
    Dim Col As dmtgridctl.dgColumnHeader
    Dim bValue As Boolean

    'Legge le impostazioni dal registry
    bValue = IIf(IsMissing(IDVisible), AppOptions.IDFieldsVisibility, IDVisible)

    For Each Col In BrwMain.ColumnsHeader
        If Left(Col.FieldName, 2) = "ID" Then
            Col.Visible = bValue
        End If
    Next Col

    'L'aspetto della browse viene ridisegnato.
    BrwMain.Refresh
    
End Sub


'ATTENZIONE: Nella funzione GetDescription4StatusBar vanno impostati tutti i
'            suggerimenti dei pulsanti della Toolbar e delle voci di menu
'            da visualizzare sulla Statusbar.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.

'**+
'Nome:   GetDescription4StatusBar
'
'Parametri: sToolName è il nome del pulsante o della voce di menu per i quali
'           si vuole ottenere il messaggio sulla Statusbar
'
'Valori di ritorno: La stringa da visualizzare sulla StatusBar
'
'Funzionalità: Restituisce la stringa del suggerimento associato ad un bottone
'              della toolbar o ad una voce di menu
'**/
Private Function GetDescription4StatusBar(ByVal sToolName As String) As String
    Dim sApplicationName As String
    Dim sTipoOggetto As String
    Dim sStr As String
    Dim sTemp As String

    
    sApplicationName = m_App.FunctionName
    sTipoOggetto = m_DocType.Name
    
    Select Case sToolName
    
        Case "File"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_App.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_FILE)
        
        Case "New", "Mnu_New"
            sStr = "Crea un nuovo " & sTipoOggetto
            
        Case "Save", "Mnu_Save"
            sStr = "Memorizza il " & sTipoOggetto & " corrente"
        
        Case "Print", "Mnu_Print"
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sStr = "Stampa i " & sTipoOggetto & " correnti"
            Else
                'Si è in modalità form
                sStr = "Stampa il " & sTipoOggetto & " corrente"
            End If
            
        Case "PrePrint", "Mnu_PrePrint"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                sTemp = sTipoOggetto & " correnti"
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_SETPREVIEW)
            Else
                'Si è in modalità form
                sTemp = sTipoOggetto & " corrente"
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_SETPREVIEW)
            End If
    
        Case "Mnu_Exit"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_EXIT)
            
        Case "Edit"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_App.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_MODIFY)
    
        Case "Cut", "Mnu_Cut"
            sStr = gResource.GetCustomizedMessage(IDS_SB_CUT)
            
        Case "Copy", "Mnu_Copy"
            sStr = gResource.GetCustomizedMessage(IDS_SB_COPY)
            
        Case "Paste", "Mnu_Paste"
            sStr = gResource.GetCustomizedMessage(IDS_SB_PASTE)
            
        Case "Delete", "Mnu_Delete"
            sStr = "Elimina il " & sTipoOggetto & " corrente"
            
        Case "Clear", "Mnu_Clear"
            sStr = gResource.GetCustomizedMessage(IDS_SB_CLEAR)
            
        Case "NewSearch", "Mnu_NewSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHWINDOW)
            
        Case "ExecuteSearch", "Mnu_ExecuteSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHEXECUTE)
            
        Case "Mnu_FormView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_FORM)
            
        Case "Mnu_TableView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_TABLE)
            
        Case "Mnu_SearchFilter"
            sStr = "Espone " & m_DocType.Name & " in modo <filtri>."
            
        Case "ChangeView"
            If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
                'Si è in modalità tabellare
                gResource.CustomStrings.Clear
                gResource.CustomStrings.Add m_DocType.Name, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_FORM)
            Else
                'Si è in modalità form
                gResource.CustomStrings.Clear
                gResource.CustomStrings.Add m_DocType.Name, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_TABLE)
            End If
            
        Case "View"
            sStr = gResource.GetMessage(IDS_SB_DISPLAY)
            
        Case "SearchPrevious", "Mnu_SearchPrevious"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHPREVIOUS)
            
        Case "SearchNext", "Mnu_SearchNext"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHNEXT)
            
        Case "Mnu_Folders"
            sStr = "Riquadro attività"
            
        Case "Mnu_ToolBar"
            sStr = gResource.GetMessage(IDS_SB_TOOLBAR)
            
        Case "Tools"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_TOOLS)
            
        Case "Mnu_Export"
            sStr = gResource.GetMessage(IDS_SB_EXPORT)
            
        Case "Mnu_Options"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_OPTION)

            
        Case "ExportWord", "Mnu_ExportWord"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTWORD)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTWORD)
            End If
        
        Case "ExportExcel", "Mnu_ExportExcel"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTEXCEL)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTEXCEL)
            End If
        
        Case "ExportHtml", "Mnu_ExportHtml"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTHTML)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTHTML)
            End If
        
        Case "ExportPDF", "Mnu_ExportPDF"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTACROBAT)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTACROBAT)
            End If
        
        Case "Mnu_HelpOnLine"
            sStr = gResource.GetMessage(IDS_SB_SUMMARY)
            
        Case "Mnu_Arg"
            sStr = gResource.GetMessage(IDS_SB_ARG)
            
        Case "Mnu_Web"
            sStr = gResource.GetMessage(IDS_SB_WEB)
            
        Case "Mnu_Info"
            sStr = gResource.GetMessage(IDS_SB_INFO)
            
        Case "Mnu_Web", "Web"
            sStr = gResource.GetMessage(IDS_SB_WEB)
        
        Case "Mnu_Agg_Web", "Agg_Web"
            sStr = gResource.GetMessage(IDS_SB_AGG_WEB)
    End Select
    
    GetDescription4StatusBar = sStr
End Function
'//////////////////////////////////////////////////////////////////////////////////
'ATTENZIONE: Nella funzione GetToolTipText4ToolBar vanno impostate tutte le
'            stringhe dei ToolTipText dei pulsanti della Toolbar.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.
'//////////////////////////////////////////////////////////////////////////////////
'**+
'Nome:   GetToolTipText4ToolBar
'
'Parametri: sToolName è il nome del pulsante per il quale
'           si vuole ottenere la stringa per la proprietà ToolTipText
'
'Valori di ritorno: La stringa ToolTipText
'
'Funzionalità: Restituisce la stringa del suggerimento associato ad un bottone
'              della toolbar (ToolTipext)
'**/
Private Function GetToolTipText4ToolBar(ByVal sToolName As String) As String
    Dim sStr As String
    
    gResource.CustomStrings.Clear
    
    Select Case sToolName
    
        Case "New"
            sStr = gResource.GetMessage(TT_NEW)
            
        Case "Save"
            sStr = gResource.GetMessage(TT_SAVE)
        
        Case "Print"
            sStr = gResource.GetMessage(TT_PRINT)
            
        Case "PrePrint"
            sStr = gResource.GetMessage(TT_PREVIEW)
    
        Case "Cut"
            sStr = gResource.GetMessage(TT_CUT)
            
        Case "Copy"
            sStr = gResource.GetMessage(TT_COPY)
            
        Case "Paste"
            sStr = gResource.GetMessage(TT_PASTE)
            
        Case "Delete"
            sStr = gResource.GetMessage(TT_DELETE)
            
        Case "Clear"
            sStr = gResource.GetMessage(TT_CLEAR)
            
        Case "NewSearch"
            sStr = gResource.GetMessage(TT_SEARCH)
            
        Case "ExecuteSearch"
            sStr = gResource.GetMessage(TT_SEARCHEXECUTE)
            
        Case "ChangeView"
            If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
                'Si è in modalità tabellare
                sStr = gResource.GetMessage(TT_FORM)
            Else
                'Si è in modalità form
                sStr = gResource.GetMessage(TT_SEARCHRESULT)
            End If
            
        Case "SearchPrevious"
            sStr = gResource.GetMessage(TT_SEARCHPREVIOUS)
            
        Case "SearchNext"
            sStr = gResource.GetMessage(TT_SEARCHNEXT)
            
        Case "ExportWord"
            sStr = gResource.GetMessage(TT_WORD)
        
        Case "ExportExcel"
            sStr = gResource.GetMessage(TT_EXCEL)
        
        Case "ExportHtml"
            sStr = gResource.GetMessage(TT_HTML)
        
        Case "ViewAssistant" 'toolbar
            sStr = gResource.GetMessage(TT_SHOW_ASSISTANT)
            
        Case "Help" 'toolbar e menu
            sStr = gResource.GetMessage(TT_HELP)

    End Select
    
    GetToolTipText4ToolBar = sStr
End Function

'//////////////////////////////////////////////////////////////////////////////////
'ATTENZIONE: Nella funzione GetCaption4MenuBar vanno impostate tutte le
'            stringhe delle Caption delle voci di menu.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.
'//////////////////////////////////////////////////////////////////////////////////
'**+
'Nome:   GetCaption4MenuBar
'
'Parametri: sToolName è il nome della voce di menu per la quale
'           si vuole ottenere la stringa per la Caption
'
'Valori di ritorno: La stringa da visualizzare nella Caption del menu
'
'Funzionalità: Restituisce la stringa della Caption di una voce di menu
'**/
Private Function GetCaption4MenuBar(ByVal sToolName As String) As String
    Dim sStr As String
    
    gResource.CustomStrings.Clear
    
    Select Case sToolName
    
        Case "File"
            sStr = gResource.GetMessage(MNU_FILE)
        
        Case "Mnu_New"
            sStr = gResource.GetMessage(MNU_NEW)
            aryShortCut(1).Value = "Control+N"
            BarMenu.Bands("Band_File").Tools("Mnu_New").ShortCuts = aryShortCut
            
        Case "Mnu_Save"
            If m_App.Language <> 1 Then
                sStr = gResource.GetMessage(MNU_SAVE)
                aryShortCut(1).Value = "Control+S"
                BarMenu.Bands("Band_File").Tools("Mnu_Save").ShortCuts = aryShortCut
            Else
                sStr = gResource.GetMessage(MNU_SAVE)
                aryShortCut(1).Value = "Shift+F12"
                BarMenu.Bands("Band_File").Tools("Mnu_Save").ShortCuts = aryShortCut
            End If
        
        Case "Mnu_PrePrint"
            sStr = gResource.GetMessage(MNU_PREVIEW)
        
        Case "Mnu_Print"
            If m_App.Language <> 1 Then
                sStr = gResource.GetMessage(MNU_PRINT) & "..."
                aryShortCut(1).Value = "Control+P"
                BarMenu.Bands("Band_File").Tools("Mnu_Print").ShortCuts = aryShortCut
            Else
                sStr = gResource.GetMessage(MNU_PRINT) & "..."
                aryShortCut(1).Value = "Control+Shift+F12"
                BarMenu.Bands("Band_File").Tools("Mnu_Print").ShortCuts = aryShortCut
            End If
    
        Case "Mnu_Exit"
            sStr = gResource.GetMessage(MNU_EXIT)
            
        Case "Edit"
            sStr = gResource.GetMessage(MNU_MODIFY)
    
        Case "Mnu_Delete"
            sStr = gResource.GetMessage(MNU_DELETE)
            aryShortCut(1).Value = "Delete"
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
            
        Case "Mnu_Clear"
            sStr = gResource.GetMessage(MNU_CLEAR)
    
        Case "Mnu_Cut"
            sStr = gResource.GetMessage(MNU_CUT)
            aryShortCut(1).Value = "Control+X"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").ShortCuts = aryShortCut
            
        Case "Mnu_Copy"
            sStr = gResource.GetMessage(MNU_COPY)
            aryShortCut(1).Value = "Control+C"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").ShortCuts = aryShortCut
            
        Case "Mnu_Paste"
            sStr = gResource.GetMessage(MNU_PASTE)
            aryShortCut(1).Value = "Control+V"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").ShortCuts = aryShortCut
            
        Case "Mnu_NewSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_FIND)
            aryShortCut(1).Value = "Control+T"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").ShortCuts = aryShortCut
            
        Case "Mnu_ExecuteSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_EXECUTE_SEARCH)
            aryShortCut(1).Value = "Control+E"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").ShortCuts = aryShortCut
            
        Case "Mnu_SearchPrevious"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_PREVIOUS_SEARCH)
            aryShortCut(1).Value = "Control+P"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").ShortCuts = aryShortCut
            
        Case "Mnu_SearchNext"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_NEXT_SEARCH)
            aryShortCut(1).Value = "Control+S"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").ShortCuts = aryShortCut
            
        Case "View"
            sStr = gResource.GetMessage(MNU_DISPLAY)
            
        Case "Mnu_FormView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_FORM)
            aryShortCut(1).Value = "Control+F"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_FormView").ShortCuts = aryShortCut
            
        Case "Mnu_TableView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_TABLE)
            aryShortCut(1).Value = "Control+M"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_TableView").ShortCuts = aryShortCut
            
        Case "Mnu_SearchFilter"
            sStr = "Mo&dalità filtri"
            aryShortCut(1).Value = "Control+Shift+T"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").ShortCuts = aryShortCut
            
        Case "Mnu_Folders"
            sStr = "&Riquadro attività"
            
        Case "Mnu_ToolBar"
            sStr = gResource.GetMessage(MNU_TOOLBAR)
            
        Case "Tools"
            sStr = gResource.GetMessage(MNU_TOOL)
            
        Case "Mnu_Export"
            sStr = gResource.GetMessage(MNU_EXPORT)
            
        Case "Mnu_Options"
            sStr = gResource.GetMessage(MNU_OPTION)
            
        Case "Mnu_ExportWord"
                sStr = gResource.GetMessage(MNU_EXPORT_WORD)
        
        Case "Mnu_ExportExcel"
                sStr = gResource.GetMessage(MNU_EXPORT_EXCEL)
        
        Case "Mnu_ExportHtml"
                sStr = gResource.GetMessage(MNU_EXPORT_HTML)
        
        Case "Mnu_ExportPDF"
                sStr = gResource.GetMessage(MNU_EXPORT_ACROBAT)
        
        Case "Help" 'toolbar e menu
            sStr = "&?"

        Case "Mnu_HelpOnLine"
            sStr = gResource.GetMessage(MNU_HELP)
            aryShortCut(1).Value = "F1"
            frmMain.BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").ShortCuts = aryShortCut
            
        Case "Mnu_Arg"
            sStr = gResource.GetMessage(MNU_ARG)
            aryShortCut(1).Value = "Shift+F1"
            frmMain.BarMenu.Bands("Band_Help").Tools("Mnu_Arg").ShortCuts = aryShortCut
            
        Case "Mnu_Web"
            sStr = gResource.GetMessage(MNU_WEB)
            
        Case "Mnu_Agg_Web"
            sStr = gResource.GetMessage(MNU_AGG_WEB)
            
        Case "Mnu_Info"
            sStr = gResource.GetMessage(MNU_INFO)
            
        Case "Mnu_RunApplication"
            sStr = gResource.GetMessage(MNU_EXE_GEST)
            aryShortCut(1).Value = "Control+G"
            frmMain.BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").ShortCuts = aryShortCut
        
        Case "Mnu_SearchObject"
            sStr = gResource.GetMessage(MNU_SEARCH)
            aryShortCut(1).Value = "Control+R"
            frmMain.BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").ShortCuts = aryShortCut
            
    End Select
    
    GetCaption4MenuBar = sStr
End Function


'**+
'Nome:   RefreshDescriptions4StatusBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Reimposta i messaggi da visualizzare sulla StatusBar per quelle
'              voci che dipendono dalla modalità di visualizzazione (Form/Tabella).
'
'**/
Private Sub RefreshDescriptions4StatusBar()
    'ATTENZIONE:
    'Inserire qui tutte le voci di menu ed i pulsanti della toolbar per i quali si
    'vuole cambiare il suggerimento sulla StatusBar in funzione della modalità di
    'visualizzazione. Ad esempio è possibile avere dei messaggi al SINGOLARE per
    'la modalità form e PLURALE per la modalità tabellare.
    'La funzione GetDescription4StatusBar si occupa di determinare la frase esatta.
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands("Band_Export").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Description = GetDescription4StatusBar("ExportPDF")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("ExportWord")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Description = GetDescription4StatusBar("ExportPDF")

End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 25/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: Caption2Display
'
'Parametri:
'  Boolean ReadFromGrid - determina se le stringhe per la costruzione della caption devono essere lette direttamente
'  dai campi del documento o dalla collection AllColumns della BrwMain.
'
'Valori di ritorno: String
'
'Funzionalità:
'                  ///////////////////////////////////////////////////////////////////////////////////////////////////////
'                  In questa funzione va inserito il codice per la determinazione della caption del form principale
'                  per le modalità Modify e Browse in base alle esigenze specifiche.
'                  Di default viene usato esclusivamente il campo del documento individuato dalla
'                  costante CAMPO_PER_CAPTION.
'                  ///////////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Function Caption2Display(Optional ByVal ReadFromGrid As Boolean) As String
    If Not m_Document.EOF And Not m_Document.BOF Then
        If Not ReadFromGrid Then
            
            Caption2Display = m_App.Caption & ": " & fnNotNull(m_Document.Fields(CAMPO_PER_CAPTION).Value) & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
        Else
            
            Caption2Display = m_App.Caption & ": " & fnNotNull(BrwMain.AllColumns(CAMPO_PER_CAPTION).Value) & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
        End If
    Else
        Caption2Display = m_App.FunctionName & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    End If
End Function



'**+
'Nome :SetStatus4Modality
'
'Parametri:NewModality rappresenta la modalità di visualizzazione
'          che si vuole ottenere.
'          ModePreview è uno switch per apertura o chiusura anteprima di stampa.
'
'Valori di ritorno:
'
'Funzionalità: Abilita i pulsanti della Toolbar e le voci di menu in funzione
'              di una determinata modalità di visualizzazione.
'              (disabilita tutti i rimanenti pulsanti e voci di menu)
'              Imposta la Caption del form in funzione della modalità di visualizzazione
'**/
Private Sub SetStatus4Modality(ByVal NewModality As neVisualModality, _
                                Optional ByVal ModePreview As nePreviewModality)
    Dim KeyON As Currency
    Dim KeyOFF As Currency
    Dim iPicture As Integer
   
    
    'Indica lo stato di visibilità della ToolBar standard
    'prima della visualizzazione della ToolBar della anteprima
    'di stampa
    Static bToolBarStandardVisible As Boolean
    
    'Indica lo stato di attivazione dei bottoni della ToolBar
    'standard prima della visualizzazione della ToolBar della
    'anteprima di stampa
    Static curToolBarStandardStatus As Currency
    
    
    'Elimina l'acceleratore CUT
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = gResource.GetMessage(MNU_DELETE)
    'Rimuove lo shortcut "Delete"
    aryShortCut(1).Clear
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
    
    'Imposta i pulsanti e le voci di menu
    Select Case NewModality
    
        Case Insert
            KeyOFF = BTN_SAVE + BTN_PRINT + BTN_PREVIEW + BTN_DELETE + BTN_SEARCH
            KeyOFF = KeyOFF + BTN_SEARCHTABLE + BTN_SEARCHFORM + BTN_VIEWMODE
            KeyOFF = KeyOFF + BTN_FILTER
            KeyOFF = KeyOFF + BTN_PREVIOUS + BTN_NEXT
            KeyOFF = KeyOFF + BTN_WORD + BTN_EXCEL + BTN_HTML + BTN_PDF
            KeyON = BTN_ALL - KeyOFF
            Me.Caption = m_App.Caption
            oFiltersActivity.AbortNewFilter
            
            If BrwMain.GuiMode = dgFilterDefinition Then
                bEnableGuiEvent = False
                BrwMain.GuiMode = dgNormal
                bEnableGuiEvent = True
            End If
            
            m_Search = False
            
        Case Modify
            KeyOFF = BTN_SAVE + BTN_CLEAR + BTN_SEARCH + BTN_SEARCHFORM
            KeyON = BTN_ALL - KeyOFF
            'in modalità variazione si è necessariamente in modalità form
            'pertanto il pulsante ChangeView della toolbar deve visualizzare
            'l'icona della griglia
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_GRID32, IDB_STD_GRID16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
                'Monta la caption del form principale
                Me.Caption = Caption2Display(False)
            End If
            oFiltersActivity.AbortNewFilter
                        
            m_Search = False
            
        Case Find
        
            'Solo se esiste almeno un elemento nel data manager.
            If Not (m_Document.EOF = True And m_Document.BOF = True) Then
                KeyON = BTN_VIEWMODE + BTN_SEARCHTABLE + BTN_SEARCHFORM
            End If
            KeyON = KeyON + BTN_NEW + BTN_CUT + BTN_COPY + BTN_PASTE
            KeyON = KeyON + BTN_CLEAR + BTN_SEARCH
            KeyOFF = BTN_ALL - KeyON
            'In modalità Find verrà proposto il pulsante per andare in modalità tabella
            'pertanto il pulsante ChangeView della toolbar deve visualizzare
            'l'icona della griglia
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_GRID32, IDB_STD_GRID16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
            Me.Caption = gResource.GetMessage(TT_SEARCH) & " - " & m_App.Caption
            
            oFiltersActivity.AbortNewFilter
                
            'Cancella eventuali blocchi su qualsiasi azione.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            
        Case Browse
            KeyOFF = BTN_SAVE + BTN_CLEAR + BTN_SEARCH + BTN_PREVIOUS + BTN_NEXT
            KeyOFF = KeyOFF + BTN_SEARCHTABLE + BTN_CUT + BTN_COPY + BTN_PASTE
            KeyON = BTN_ALL - KeyOFF
            'Seleziona l'icona grande o piccola in base alle impostazioni correnti
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_FORM32, IDB_STD_FORM16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
                'Monta la caption del form principale
                Me.Caption = Caption2Display(False)
            End If
            'Inserisce l'acceleratore CUT
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = gResource.GetMessage(MNU_DELETE)
            'Inserisce lo shortcut "Delete"
            aryShortCut(1).Value = "Delete"
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
            
            'Questo controllo si è reso necessario per evitare un loop infinito
            'con la gestione dell'evento BrwMain_OnChangeGuiMode() quando dal
            'Menu della browse si va in modalità tabellare.
            If BrwMain.GuiMode <> dgNormal Then
                BrwMain.GuiMode = dgNormal
            End If
            
            'Se il filtro attivo è un filtro temporaneo viene abilitato il pulsante
            'Salva Filtro del DocTypeExplorer per poterlo rendere permanente.
            If m_ActiveFilter.ID = -1 Then
                oFiltersActivity.NewFilterBegin   'Abilita il pulsante Salva Filtro
            Else
                oFiltersActivity.AbortNewFilter   'Disabilita il pulsante Salva Filtro
            End If
            ActivityBox.Redraw = True
            
            m_Search = False
            
            'Cancella eventuali blocchi su qualsiasi azione.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            
            
        Case Preview
            If ModePreview = OpenPrw Then
                bToolBarStandardVisible = BarMenu.Bands("Standard").Visible
                curToolBarStandardStatus = GetStatusToolBar(True)
                KeyON = BTN_PRINT + BTN_EXCEL + BTN_WORD + BTN_HTML + BTN_PDF
                KeyOFF = BTN_ALL - KeyON
                BarMenu.Bands("Band_View").Tools("Mnu_Folders").Enabled = False
                BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = False
                BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Enabled = False
                BarMenu.Bands("Standard").Visible = False
                BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = True
                BarMenu.RecalcLayout
            Else
                BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = False
                BarMenu.Bands("Standard").Visible = bToolBarStandardVisible
                ActivateBarButtons curToolBarStandardStatus, True
                ActivateBarButtons BTN_ALL - curToolBarStandardStatus, False
                BarMenu.Bands("Band_View").Tools("Mnu_Folders").Enabled = True
                BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = True
                BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Enabled = True
                BarMenu.RecalcLayout
            End If
            
    End Select
    
    'Attiva/disattiva i pulsanti e le voci di menu
    ActivateBarButtons KeyON, True
    ActivateBarButtons KeyOFF, False
End Sub

'**+
'Autore                 : Diamante S.p.a
'
'Nome                   : PermissionToSave
'
'Parametri:
'
'Valori di ritorno: True se il documento può essere salvato, False altrimenti.
'
'Funzionalità: Controlli da effettuare PRIMA di salvare il documento corrente
'
'**/
Private Function PermissionToSave() As Boolean
Dim Testo As String

    PermissionToSave = True
    
    If (Me.ACSCliente.IDAnagrafica = 0) Then
        MsgBox "Inserire il cliente!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    If (Me.CDSocio.KeyFieldID = 0) Then
        MsgBox "Inserire il socio!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    
    If (LINK_DOCUMENTO_COLLEGATO > 0) Then
        If (CONTROLLO_STATO_DOCUMENTO) Then
            MsgBox "Documento di trasporto collegato risulta bloccato!", vbCritical, "Controllo dati"
            PermissionToSave = False
            Exit Function
        End If
    End If
    If Len(Trim(Me.txtNumeroCertificato.Text)) = 0 Then
        MsgBox "Inserire il numero del certificato!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    If Me.txtDataCertificato.Value = 0 Then
        MsgBox "Inserire la data del certificato del socio!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    If Len(Trim(Me.txtNumeroDDT.Text)) = 0 Then
        MsgBox "Inserire il numero DDT del socio!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    If Me.txtDataDDT.Value = 0 Then
        MsgBox "Inserire la data del DDT del socio!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    If CDSezionale.KeyFieldID = 0 Then
        MsgBox "Inserire il sezionale!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    If CDArticolo.KeyFieldID = 0 Then
        MsgBox "Inserire l'articolo!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    If CDImballo.KeyFieldID = 0 Then
        MsgBox "Inserire l'imballo!", vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
'    If Me.txtColliEntrata.Value = 0 Then
'        MsgBox "Inserire i colli in entrata!", vbCritical, "Controllo dati"
'        PermissionToSave = False
'        Exit Function
'    End If
    If txtQtaFatturazione.Value <= 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Non è stata indicata la quantità di fatturazione!"
        MsgBox Testo, vbCritical, "Controllo dati"
        PermissionToSave = False
        Exit Function
    End If
    If (Me.txtIDLottoCampagna.Value > 0) Then
        If (LINK_SOCIO_LOTTO_SEL <> Me.txtIDSocio.Value) Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Il socio del documento non è uguale al socio del lotto selezionato!" & vbCrLf
            MsgBox Testo, vbCritical, "Controllo dati"
            PermissionToSave = False
        End If
    End If
    If txtIDLottoCampagna.Value = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Il lotto di produzione non è stato indicato!" & vbCrLf
        Testo = Testo & "Sei sicuro di continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
            PermissionToSave = False
            Exit Function
        End If
    End If
    If txtIDContratto.Value = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Non è stato indicato nessun riferimento ad un contratto!" & vbCrLf
        Testo = Testo & "Sei sicuro di continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
            PermissionToSave = False
            Exit Function
        End If
    End If
    If txtIDContrattoRiga.Value > 0 Then
        If txtIDContrattoRiga.Value = 0 Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Non è stato indicato nessun riferimento ad una riga del contratto selezionato!" & vbCrLf
            Testo = Testo & "Sei sicuro di continuare?"
            If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
                PermissionToSave = False
                Exit Function
            End If
        End If
    End If
    If txtPrezzoDiFatturazione.Value = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Non è stato indicato il prezzo di fatturazione!" & vbCrLf
        Testo = Testo & "Sei sicuro di continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
            PermissionToSave = False
            Exit Function
        End If
    End If
    If (Me.txtIDLottoCampagna.Value > 0) Then
        If Me.cboVarietaArticolo.CurrentID <> Me.cboVarietaLotto.CurrentID Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "La varietà del lotto selezionato non coincide con quella dell'articolo!" & vbCrLf
            Testo = Testo & "Sei sicuro di continuare?"
            If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
                PermissionToSave = False
                Exit Function
            End If
        End If
    End If
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value <= 0) Then
        If (Me.txtResaMaxTotale.Value > 0) Then
            If ((Me.txtQtaFatturazione.Value + Me.txtQtaUtilizzataLotto.Value) > Me.txtResaMaxTotale.Value) Then
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "Il totale utilizzato supera la resa massima del lotto selezionato!" & vbCrLf
                Testo = Testo & "Sei sicuro di continuare?"
                If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
                    PermissionToSave = False
                    Exit Function
                End If
            End If
        End If
    End If

    If (CONTROLLO_NUM_CERT(Me.cdAnagrafica.KeyFieldID, Me.cboAltroSito.CurrentID, Me.txtDataCertificato.Text, Me.txtNumeroCertificato.Text, fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = False) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Il numero del certificato risulta già essere inserito!" & vbCrLf
        Testo = Testo & "Sei sicuro di continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
            PermissionToSave = False
            Exit Function
        End If
    End If
    If (DateDiff("d", Me.txtDataCertificato.Text, Date) < 0) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La data del certificato è maggiore della data odierna!" & vbCrLf
        Testo = Testo & "Sei sicuro di continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
            PermissionToSave = False
            Exit Function
        End If
    End If
    If (CONTROLLO_DATA_CERT_ESERCIZIO_IN_CORSO(Me.txtDataCertificato.Text) = False) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La data del certificato non rientra nel periodo dell'esercizio in corso!" & vbCrLf
        Testo = Testo & "Sei sicuro di continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
            PermissionToSave = False
            Exit Function
        End If
    End If
    If (Me.txtDataDDT.Value > Me.txtDataCertificato.Value) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La data del D.d.T. è maggiore della data del certificato!" & vbCrLf
        Testo = Testo & "Sei sicuro di continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
            PermissionToSave = False
            Exit Function
        End If
    End If
    If (CONTROLLO_DATA_CERT_ESERCIZIO_IN_CORSO(Me.txtDataDDT.Text) = False) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La data del D.d.T. non rientra nel periodo dell'esercizio in corso!" & vbCrLf
        Testo = Testo & "Sei sicuro di continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
            PermissionToSave = False
            Exit Function
        End If
    End If
    
'    If (fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0) Then
'        If (MsgInDocSeRigaMerceSenzaImballo = 1) Then
'            If (Me.CDImballo.KeyFieldID = 0) Then
'                Testo = "ATTENZIONE!!!!"
'                Testo = Testo & "L'imballo secondario non è stato inserito!" & vbCrLf
'                If MsgBox(Testo, vbQuestion + vbYesNo, "Vuoi continuare?") = vbNo Then
'                    PermissionToSave = False
'                    Exit Function
'                End If
'            End If
'        End If
'    End If
End Function


'**+
'Nome: SearchNext
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Posizionamento al record successivo
'**/
Private Sub SearchNext()
    
    m_Document.MoveNext
    
    If m_Document.EOF Then
        'Si era già sull'ultimo record (prima di MoveNext).
        
        'Si annulla l'operazione
        m_Document.MovePrevious
        sbMsgInfo gResource.GetMessage(MESS_NO_NEXT_ELEMENTS), m_App.FunctionName
        Exit Sub
    Else
        'Controlla la presenza di eventuali conflitti nel caso di multiutenza.
        
        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            m_Document.MovePrevious
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
        End If
    End If
    
End Sub

'**+
'Nome: SearchPrevious
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Posizionamento al record precedente
'**/
Private Sub SearchPrevious()

    m_Document.MovePrevious
    
    If m_Document.BOF Then
        'Si era già sul primo record (prima di MovePrevious).
        
        'Si annulla l'operazione
        m_Document.MoveNext
        sbMsgInfo gResource.GetMessage(MESS_NO_PREVIOUS_ELEMENTS), m_App.FunctionName
        Exit Sub
    Else
        'Controlla la presenza di eventuali conflitti nel caso di multiutenza.
        
        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            m_Document.MoveNext
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
        End If
    End If
End Sub

'**+
'Nome: BrowseReposition
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni da compiere al riposizionamento del record corrente
'**/
Private Sub BrowseReposition()

    'Dopo un Save del documento avviene un Refresh della Browse ma in tal caso
    'è inutile effettuare il refresh del form.
    If Not m_bAvoidReposition Then
    
        'Refresh dei campi del form
        RefreshFormFields
        
        'Refresh della caption del Form
        If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
            'Monta la caption del form principale
            Me.Caption = Caption2Display(False)
        End If
        
    End If
 
    'Refresh delle variabili di stato
    m_Changed = False
    m_Saved = False
    m_Search = False
    
    'Annullamento di un eventuale inizio di inserimento di un nuovo record
    m_Document.AbortNew
    
End Sub



'**+
'Nome: NewRecord
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni su richiesta nuovo record
'**/
Private Sub NewRecord()


'--------------------------------------------------------------------------------------------
'NOTA:
'Il gruppo di istruzioni sottostanti e la riga  'Imposta il blocco su inserimento'
'sono state commentate per far si che la manutenzione NON imposti alcun blocco per
'l'azione Inserimento.
'Pertanto 2 o più utenti potranno effettuare contemporaneamente la suddetta azione.
'Se si intende impedire questa possibilità sarà sufficiente ripristinare le righe commentate.
'--------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'    'Controllo se ho il permesso di salvare ( nel caso di conflitti di multiutenza )
'    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, SemAllObjects, SemInsertAction) Then
'        'C'è un altro utente in modalità inserimento che blocca la medesima azione per
'        'tutti gli altri utenti. Pertanto annullo l'operazione di inserimento ed esco.
'        Exit Sub
'    End If

    'Ho il permesso per l'azione inserimento.
    '
    'Cancella il blocco precedente
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    
    '--------------------------------
    'Imposta il blocco su inserimento
'    m_Semaphore.SetObjectAction m_DocType.ID, SemAllObjects, SemInsertAction
    '
    'A questo punto nessun altro utente potrà effettuare una operazione di inserimento
    'finchè non verrà cancellato il blocco su inserimento.
    

    'Annulla una eventuale operazione precedente.
    If m_Document.TableNew Then
        m_Document.AbortNew
    End If
     
    bloading = True
    'Creazione buffers vuoti
    
    m_Document.NewDoc
    
    bloading = False
    
    ACSCliente.Description = ""
    ACSCliente.SecondDescription = ""
    ACSCliente.IDAnagrafica = 0
    Me.ACSCliente.sbLoadCFByIDAnagrafica 0, 0
    
    
    If (TIPO_SALVATAGGIO = 0) Then
        GET_SEZIONALE_DEFAULT
    End If
    
    If (TIPO_SALVATAGGIO = 1) Then
        Me.cdAnagrafica.Load IDAnagrafica_PREC
        
        Me.cboAltroSito.WriteOn IDDestinazione_PREC
        Me.cboVettore.WriteOn IDVettore_PREC
        Me.txtIDContratto.Value = IDContratto_PREC
        LINK_CONTRATTO = Me.txtIDContratto.Value
        Me.txtIDContrattoRiga.Value = IDContrattoRiga_PREC
        If (IDCooperativa_PREC = 0) Then
            Me.CDSocioFatt.Load IDCooperativa_PREC
            Me.CDSocio.Load IDAnagraficaSocio_PREC
        Else
            Me.CDSocioFatt.Load IDCooperativa_PREC
            Me.txtCodiceAnaSocio.SetFocus
        End If
    End If
    
    Me.txtColliEntrata.Value = NumeroColliPerAutomezzoCert
    Me.txtColliUscita.Value = NumeroColliPerAutomezzoCert
    'Refresh delle variabili di stato
    m_Search = False
    m_Changed = False
    m_Saved = False
    
    'Refresh della toolbar in modalità inserimento
    SetStatus4Modality Insert
    
    'Ripristina la vista del Form
    BrwMain.Visible = False
    
    
    'Il primo campo del Form riceve l'input focus
    If (TIPO_SALVATAGGIO = 0) Then
        SetFocusTabIndex0
    End If
    If (TIPO_SALVATAGGIO = 1) Then
        If (IDCooperativa_PREC = 0) Then
            If (AttivaSelezioneAnaVeloceInCert = 1) Then
                Command5_Click
            Else
                Me.Command5.SetFocus
            End If
        Else
            If (AttivaSelezioneAnaVeloceInCert = 1) Then
                frmSelAnagraficaSocio.Show vbModal
                If LINK_ANA_SOCIO_SEL > 0 Then
                    Me.CDSocio.Load LINK_ANA_SOCIO_SEL
                    If (Me.txtIDLottoCampagna.Value = 0) Then
                        Command5_Click
                    End If
                    Me.txtNumeroCertificato.SetFocus
                    
                End If
            Else
                Me.txtCodiceAnaSocio.SetFocus
            End If
        End If
    End If
    
    TIPO_SALVATAGGIO = 0
    
End Sub

'**+
'Nome                   : ClearControl
'
'Parametri              : ctrControl As Control - controllo da pulire
'
'Valori di ritorno      :
'
'Funzionalità           : Pulisce un controllo sulla base del tipo del controllo stesso
'
'**/
Private Sub ClearControl(ByVal ctrControl As Control)
    Dim sType As String

    sType = TypeName(ctrControl)
    
    If sType = "fpDateTime" Or sType = "TextBox" Or sType = "fpText" Or sType = "fpLongInteger" Or sType = "fpCurrency" Or sType = "fpDoubleSingle" Or sType = "dmtDate" Or sType = "dmtTime" Then
        ctrControl.Text = ""
    ElseIf sType = "CheckBox" Then
        ctrControl.Value = 0
    ElseIf sType = "fpBoolean" Then
        ctrControl.Value = 0
    ElseIf sType = "ComboBox" Then
        ctrControl.ListIndex = -1
    ElseIf sType = "DMTCombo" Then
        ctrControl.WriteOn 0
    ElseIf sType = "ListBox" Then
        ctrControl.ListIndex = -1
    ElseIf sType = "ListView" Then
        ctrControl.ListItems.Clear
    ElseIf sType = "TreeView" Then
        ctrControl.Nodes.Clear
    ElseIf sType = "Town" Then
        ctrControl.Reset
    ElseIf sType = "dmtCurrency" Or sType = "dmtNumber" Then
        ctrControl.Value = 0
        ctrControl.Text = ""
    ElseIf sType = "DmtSearchACS2" Then
        ctrControl.IDAnagrafica = 0
        ctrControl.Code = ""
        ctrControl.Description = ""
        ctrControl.SecondDescription = ""
    ElseIf sType = "DmtFirmGerarchy" Then
        ctrControl.LoadActivity 0
    ElseIf sType = "DMTProgControl" Then
        'Queste istruzioni forzano il refresh
        'e il reset del componente
        ctrControl.IDArticolo = 0
        ctrControl.Show
    ElseIf sType = "DmtCodDesc" Then
        ctrControl.Load 0
        
    End If
End Sub


'**+
'Nome: ClearFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Pulisce il contenuto dei campi di input del Form
'**/
Private Sub ClearFormFields()
    Dim cField As FormField
    
    For Each cField In m_FormFields
        'Viene ripulito il campo di immissione.
        ClearControl cField.Control
    Next
End Sub

'**+
'Nome: ExecuteMenuCommand
'
'Parametri:
'sToolName - Nome del comando selezionato
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione dei comandi generati dal controllo ActiveBar
'**/
Private Sub ExecuteMenuCommand(ByVal sToolName As String)
    Dim iAnswer As Integer

    'cbcxn
    'Notifica alla (eventuale) applicazione che gestisce il processo On_Extend la
    'pressione di un Tool. Se l'applicazione chiamata restituisce True viene annullata l'operazione.
    'If m_ExtendApplication.BeforeCommandClick(sToolName) Then Exit Sub

    Select Case sToolName
        Case "Cut", "Mnu_Cut"
            SendKeys ("+{DEL}")
            
        Case "Copy", "Mnu_Copy"
            SendKeys ("^{INSERT}")
            
        Case "Paste", "Mnu_Paste"
            SendKeys ("+{INSERT}")
            
        Case "Mnu_Folders"
            OnFolders
            
        Case "ClosePreview"
            ClosePreview
            
        Case "Save", "Mnu_Save"
            OnSave
            
        Case "Mnu_Exit"
            Unload frmMain
            
        Case "Delete", "Mnu_Delete"
            OnDelete
            
        Case "Clear", "Mnu_Clear"
            OnClear
            
        Case "ExecuteSearch", "Mnu_ExecuteSearch"
            OnExecuteSearch
        
        Case "SearchNext", "Mnu_SearchNext"
            
            OnMoveCurrentRecord SRCNEXT, sToolName
        
        Case "SearchPrevious", "Mnu_SearchPrevious"
            
            OnMoveCurrentRecord SRCPREVIOUS, sToolName
        
        Case "ChangeView", "Mnu_FormView", "Mnu_TableView"
            OnChangeView sToolName
            
        Case "Mnu_ToolBar"
            OnToolBarOptions
            
        Case "Mnu_Options"
            OnOptions
            
        Case "Mnu_Info"
            OnInfo
            
        Case "PrePrint", "Mnu_PrePrint", "Print", "Mnu_Print", "ExportPDF", "Mnu_ExportPDF", "MailPDF", "ExportWord", "Mnu_ExportWord", "MailWord", "ExportExcel", "Mnu_ExportExcel", "MailExcel", "ExportHtml", "Mnu_ExportHtml", "MailMHTL"
            OnPrint sToolName
            
        Case "NewSearch", "Mnu_NewSearch", "Mnu_SearchFilter"
            OnNewSearch
            
        Case "New", "Mnu_New"
            
            OnNew sToolName
           
        Case "Mnu_RunApplication", "Mnu_SearchObject"
            OnRunApplication sToolName
        Case "Mnu_Summary"
            OnSummary
        Case "Mnu_FastHelp", "Help"
            OnFastHelp
        Case "Mnu_HelpOnLine"
            OnHelpOnLine
        Case "Mnu_Arg"
             OnArg
        Case "Mnu_Web1"
             sbOpenURL hwnd, URL_DIAMANTE
    End Select
    
    
    'cbcxn
    'Notifica alla (eventuale) applicazione che gestisce il processo On_Extend la
    'pressione di un Tool DOPO avere eseguito l'operazione ad esso associata
    'm_ExtendApplication.AfterCommandClick sToolName
    
End Sub

'**+
'Nome: RefreshFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Riempie i valori dei campi del Form con i valori
'del documento.
'**/
Private Sub RefreshFormFields()

    'rif3 start
    
    Dim Fields As DmtDocManLib.Fields
    Dim Control As Control
    Dim Field As FormField
    
    On Error Resume Next
    
    'In questi casi non si deve far nulla
    If Not (m_Document.EOF = True Or m_Document.BOF = True) Then

       'Passa alla collezione Fields dell'oggetto
        'Document i valori da salvare
        For Each Field In m_FormFields
           Select Case TypeName(Field.Control)
                Case "TextBox"
                    Field.Control.Text = fnNotNull(m_Document.Fields(Field.Name).Value)
                Case "DMTCombo"
                    Field.Control.WriteOn fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "Town"
                    If Field.Name = "IDComune" Then
                        Field.Control.TownID = fnNotNullN(m_Document.Fields(Field.Name).Value)
                    ElseIf Field.Name = "Cap" Then
                        Field.Control.Zip = fnNotNull(m_Document.Fields(Field.Name).Value)
                    End If
                Case "dmtDate"
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtNumber"
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                    
                Case "dmtCurrency"
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtTime"
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "DmtSearchACS2"
                        Field.Control.Description = ""
                        Field.Control.SecondDescription = ""
                        Field.Control.IDAnagrafica = 0
                        
                        If fnNotNullN(m_Document.Fields(Field.Name).Value) > 0 Then
                            Field.Control.sbLoadCFByIDAnagrafica 0, fnNotNullN(m_Document.Fields(Field.Name).Value)
                        End If
                Case "CheckBox"
                    Field.Control.Value = Abs(fnNotNullN((m_Document.Fields(Field.Name).Value)))
                Case "DmtCodDesc"
                    Field.Control.Load fnNotNullN(m_Document.Fields(Field.Name).Value)
            End Select
        Next
    End If

    'rif3 end
    
End Sub

'**+
'Nome: ClearFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Azzera i valori dei campi del documento
'**/
Private Sub ClearFields()
    Dim Field As DmtDocManLib.Field
    
    For Each Field In m_Document.Fields
        Field.Value = Empty
    Next
End Sub

'**+
'Nome: Change
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni su variazione di un campo del Form
'**/
Private Sub Change()
    'Se si è in modalità tabellare non deve essere eseguita perchè
    'altrimenti al Click della Browse si attiverebbe il pulsante Salva
    If Not m_Search And Not BrwMain.Visible Then
        ActivateBarButtons BTN_SAVE, True
    
        m_Changed = True
        m_Saved = False
        m_Search = False
    End If
End Sub

'**+
'Nome: CreateFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Crea la collezione FormFields che associa i campi del
'documento con i controlli di input del Form. Vengono
'anche creati i controlli del Form necessari e calcolato
'il layout del Form.
'**/
Private Sub CreateFormFields()
    Dim Field As FormField
        
        
    'rif2 start
    
    'Se non esiste il documento aperto non si può creare la collezione
    If m_Document Is Nothing Then Exit Sub
    
    'Se la collezione è già stata creata esce
    If Not m_FormFields Is Nothing Then Exit Sub
    
    'Istanzia la collezione.  Il codice sottostante viene eseguito soltanto la prima volta
    Set m_FormFields = New FormFields
    
    'rif2   End
    
    
    'IDAnagrafica
    Set Field = New FormField
    Set Field.Control = Me.cdAnagrafica
    Field.Name = "IDAnagrafica"
    Field.Visible = True
    Me.cdAnagrafica.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDDestinazioneDiversa
    Set Field = New FormField
    Set Field.Control = Me.cboAltroSito
    Field.Name = "IDDestinazioneDiversa"
    Field.Visible = True
    Me.cboAltroSito.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDAnagraficaCooperativa
    Set Field = New FormField
    Set Field.Control = Me.CDSocioFatt
    Field.Name = "IDAnagraficaCooperativa"
    Field.Visible = True
    Me.CDSocioFatt.Tag = Field.Name
    m_FormFields.Add Field

    'IDAnagraficaSocio
    Set Field = New FormField
    Set Field.Control = Me.CDSocio
    Field.Name = "IDAnagraficaSocio"
    Field.Visible = True
    Me.CDSocio.Tag = Field.Name
    m_FormFields.Add Field
    
    'Acquistato
    Set Field = New FormField
    Set Field.Control = Me.Check1
    Field.Name = "Acquistato"
    Field.Visible = True
    Me.Check1.Tag = Field.Name
    m_FormFields.Add Field
    
    'NumeroCertificato
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroCertificato
    Field.Name = "NumeroCertificato"
    Field.Visible = True
    Me.txtNumeroCertificato.Tag = Field.Name
    m_FormFields.Add Field
    
    'DataCertificato
    Set Field = New FormField
    Set Field.Control = Me.txtDataCertificato
    Field.Name = "DataCertificato"
    Field.Visible = True
    Me.txtDataCertificato.Tag = Field.Name
    m_FormFields.Add Field
    
    'NumeroDocumentoSocio
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroDDT
    Field.Name = "NumeroDocumentoSocio"
    Field.Visible = True
    Me.txtNumeroDDT.Tag = Field.Name
    m_FormFields.Add Field
    
    'DataDocumentoSocio
    Set Field = New FormField
    Set Field.Control = Me.txtDataDDT
    Field.Name = "DataDocumentoSocio"
    Field.Visible = True
    Me.txtDataDDT.Tag = Field.Name
    m_FormFields.Add Field

    'IDSezionale
    Set Field = New FormField
    Set Field.Control = Me.CDSezionale
    Field.Name = "IDSezionale"
    Field.Visible = True
    Me.CDSezionale.Tag = Field.Name
    m_FormFields.Add Field

    'NumeroDocumento
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroDocumento
    Field.Name = "NumeroDocumento"
    Field.Visible = True
    Me.txtNumeroDocumento.Tag = Field.Name
    m_FormFields.Add Field
    
    'DataDocumento
    Set Field = New FormField
    Set Field.Control = Me.txtDataDocumento
    Field.Name = "DataDocumento"
    Field.Visible = True
    Me.txtDataDocumento.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDArticolo
    Set Field = New FormField
    Set Field.Control = Me.CDArticolo
    Field.Name = "IDArticolo"
    Field.Visible = True
    Me.CDArticolo.Tag = Field.Name
    m_FormFields.Add Field

    'DescrizioneArticolo
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneArticolo
    Field.Name = "DescrizioneArticolo"
    Field.Visible = True
    Me.txtDescrizioneArticolo.Tag = Field.Name
    m_FormFields.Add Field

    'IDArticoloImballo
    Set Field = New FormField
    Set Field.Control = Me.CDImballo
    Field.Name = "IDArticoloImballo"
    Field.Visible = True
    Me.CDImballo.Tag = Field.Name
    m_FormFields.Add Field

    'TaraImballo
    Set Field = New FormField
    Set Field.Control = Me.txtTaraUnitaria
    Field.Name = "TaraImballo"
    Field.Visible = True
    Me.txtTaraUnitaria.Tag = Field.Name
    m_FormFields.Add Field

    'TaraAutomezzo
    Set Field = New FormField
    Set Field.Control = Me.txtTaraCamion
    Field.Name = "TaraAutomezzo"
    Field.Visible = True
    Me.txtTaraCamion.Tag = Field.Name
    m_FormFields.Add Field

    'ColliEntrata
    Set Field = New FormField
    Set Field.Control = Me.txtColliEntrata
    Field.Name = "ColliEntrata"
    Field.Visible = True
    Me.txtColliEntrata.Tag = Field.Name
    m_FormFields.Add Field

    'ColliUscita
    Set Field = New FormField
    Set Field.Control = Me.txtColliUscita
    Field.Name = "ColliUscita"
    Field.Visible = True
    Me.txtColliUscita.Tag = Field.Name
    m_FormFields.Add Field

    'TaraTotaleColli
    Set Field = New FormField
    Set Field.Control = Me.txtTaraTotaleImballo
    Field.Name = "TaraTotaleColli"
    Field.Visible = True
    Me.txtTaraTotaleImballo.Tag = Field.Name
    m_FormFields.Add Field

    'Tara
    Set Field = New FormField
    Set Field.Control = Me.txtTaraTotale
    Field.Name = "Tara"
    Field.Visible = True
    Me.txtTaraTotale.Tag = Field.Name
    m_FormFields.Add Field

    'PesoNetto
    Set Field = New FormField
    Set Field.Control = Me.txtPesoNetto
    Field.Name = "PesoNetto"
    Field.Visible = True
    Me.txtPesoNetto.Tag = Field.Name
    m_FormFields.Add Field

    'PercRiduzionePesoNetto
    Set Field = New FormField
    Set Field.Control = Me.txtPercRidPesoNetto
    Field.Name = "PercRiduzionePesoNetto"
    Field.Visible = True
    Me.txtPercRidPesoNetto.Tag = Field.Name
    m_FormFields.Add Field

    'PesoNettoCalcolato
    Set Field = New FormField
    Set Field.Control = Me.txtQtaFatturazione
    Field.Name = "PesoNettoCalcolato"
    Field.Visible = True
    Me.txtQtaFatturazione.Tag = Field.Name
    m_FormFields.Add Field

    'ImportoUnitarioContratto
    Set Field = New FormField
    Set Field.Control = Me.txtPrezzoDaContratto
    Field.Name = "ImportoUnitarioContratto"
    Field.Visible = True
    Me.txtPrezzoDaContratto.Tag = Field.Name
    m_FormFields.Add Field

    'ImportoUnitario
    Set Field = New FormField
    Set Field.Control = Me.txtPrezzoDiFatturazione
    Field.Name = "ImportoUnitario"
    Field.Visible = True
    Me.txtPrezzoDiFatturazione.Tag = Field.Name
    m_FormFields.Add Field

    'TotaleRiga
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleRiga
    Field.Name = "TotaleRiga"
    Field.Visible = True
    Me.txtTotaleRiga.Tag = Field.Name
    m_FormFields.Add Field

    'IndiceVariazione
    Set Field = New FormField
    Set Field.Control = Me.txtIndiceDiVariazione
    Field.Name = "IndiceVariazione"
    Field.Visible = True
    Me.txtIndiceDiVariazione.Tag = Field.Name
    m_FormFields.Add Field

    'IndiceVariazioneEffettivo
    Set Field = New FormField
    Set Field.Control = Me.txtIndiceDiVariazioneEff
    Field.Name = "IndiceVariazioneEffettivo"
    Field.Visible = True
    Me.txtIndiceDiVariazioneEff.Tag = Field.Name
    m_FormFields.Add Field

    'IDRegione
    Set Field = New FormField
    Set Field.Control = Me.cboRegione
    Field.Name = "IDRegione"
    Field.Visible = True
    Me.cboRegione.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDContratto
    Set Field = New FormField
    Set Field.Control = Me.txtIDContratto
    Field.Name = "IDContratto"
    Field.Visible = True
    Me.txtIDContratto.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDContrattoRiga
    Set Field = New FormField
    Set Field.Control = Me.txtIDContrattoRiga
    Field.Name = "IDContrattoRiga"
    Field.Visible = True
    Me.txtIDContrattoRiga.Tag = Field.Name
    m_FormFields.Add Field

    'IDLottoProduzione
    Set Field = New FormField
    Set Field.Control = Me.txtIDLottoCampagna
    Field.Name = "IDLottoProduzione"
    Field.Visible = True
    Me.txtIDLottoCampagna.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indice
    Set Field = New FormField
    Set Field.Control = Me.txtIndice
    Field.Name = "Indice"
    Field.Visible = True
    Me.txtIndice.Tag = Field.Name
    m_FormFields.Add Field

    'DataInizioTrasporto
    Set Field = New FormField
    Set Field.Control = Me.txtDataTrasporto
    Field.Name = "DataInizioTrasporto"
    Field.Visible = True
    Me.txtDataTrasporto.Tag = Field.Name
    m_FormFields.Add Field
    
    'OraInizioTrasporto
    Set Field = New FormField
    Set Field.Control = Me.txtOraTrasporto
    Field.Name = "OraInizioTrasporto"
    Field.Visible = True
    Me.txtOraTrasporto.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDLetteraIntento
    Set Field = New FormField
    Set Field.Control = Me.txtIDLetteraIntento
    Field.Name = "IDLetteraIntento"
    Field.Visible = True
    Me.txtIDLetteraIntento.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDLetteraIntento
    Set Field = New FormField
    Set Field.Control = Me.cboIvaCliente
    Field.Name = "IDIvaEsente"
    Field.Visible = True
    Me.cboIvaCliente.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDIva
    Set Field = New FormField
    Set Field.Control = Me.cboIvaArticolo
    Field.Name = "IDIva"
    Field.Visible = True
    Me.cboIvaArticolo.Tag = Field.Name
    m_FormFields.Add Field
    
    'PesoLordo
    Set Field = New FormField
    Set Field.Control = Me.txtPesoLordo
    Field.Name = "PesoLordo"
    Field.Visible = True
    Me.txtPesoLordo.Tag = Field.Name
    m_FormFields.Add Field
    
    'ScartoPesoLordo
    Set Field = New FormField
    Set Field.Control = Me.txtScarto
    Field.Name = "ScartoPesoLordo"
    Field.Visible = True
    Me.txtScarto.Tag = Field.Name
    m_FormFields.Add Field
    
    'ImportoUnitarioContrattoMinimo
    Set Field = New FormField
    Set Field.Control = Me.txtPrezzoContrattoMin
    Field.Name = "ImportoUnitarioContrattoMinimo"
    Field.Visible = True
    Me.txtPrezzoContrattoMin.Tag = Field.Name
    m_FormFields.Add Field
    
    'ImportoUnitarioContrattoMassimo
    Set Field = New FormField
    Set Field.Control = Me.txtPrezzoContrattoMax
    Field.Name = "ImportoUnitarioContrattoMassimo"
    Field.Visible = True
    Me.txtPrezzoContrattoMax.Tag = Field.Name
    m_FormFields.Add Field
    
    'IndiceVariazione100
    Set Field = New FormField
    Set Field.Control = Me.txtIndiceDiVariazione100
    Field.Name = "IndiceVariazione100"
    Field.Visible = True
    Me.txtIndiceDiVariazione100.Tag = Field.Name
    m_FormFields.Add Field
End Sub


'**+
'Nome: ClosePreview
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Chiude la finestra della anteprima di stampa
'**/
Private Sub ClosePreview()
    Dim myDate
    
    On Error GoTo errHandler
        
    If m_Report.ClosePreview Then
        m_PreviewWindowHandle = 0
        PicForm.Visible = True
        BrwMain.Visible = m_TabMode
        ActivityBox.Visible = BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked
        FormRecalcLayout
        Set m_Report = Nothing
        SetStatus4Modality Preview, ClosePrw
    End If
    Exit Sub
errHandler:
    'Se si verifica un errore "SQL server in use"
    'la subroutine entra in un ciclo di attesa per
    '3 secondi prima di tentare nuovamente la chiusura
    myDate = Now
    If Err.Description = "SQL server in use" Then
        While Not (Now = DateAdd("s", 3, myDate))
        Wend
        Resume
    End If
    Err.Raise Err.Number, , Err.Description
End Sub




'**+
'Nome: ShortCut
'
'Parametri:
'KeyCode - Codice del tasto
'Shift - Stato del tasto Shift
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione degli accelleratori da tastiera
'**/
'**+
'Nome: ShortCut
'
'Parametri:
'KeyCode - Codice del tasto
'Shift - Stato del tasto Shift
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione degli accelleratori da tastiera
'**/
Private Function ShortCut(KeyCode As Integer, Shift As Integer) As Boolean
    Dim bCtrlDown As Boolean
    Dim bShiftDown As Boolean
    Dim bAltDown As Boolean
    
    bShiftDown = (Shift And vbShiftMask) > 0
    bCtrlDown = (Shift And vbCtrlMask) > 0
    bAltDown = (Shift And vbAltMask) > 0
    
    Select Case KeyCode
         Case vbKeyF12
            If bShiftDown Then
                If bCtrlDown Then
                    If BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled Then
                        ExecuteMenuCommand ("Mnu_Print")
                        ShortCut = True
                    End If
                Else
                    If BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled Then
                    
                        'Forza il lostfocus ed attende l'esecuzione di eventuali eventi associati
                        AutoLostFocus
                        
                        ExecuteMenuCommand ("Mnu_Save")
                        ShortCut = True
                    End If
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyF1
            SendMessage hwnd, WM_SETREDRAW, 0, 0
            'SendKeys ("{ESC}")
            DoEvents
            SendMessage hwnd, WM_SETREDRAW, 1, 0
            If bShiftDown Then
                'case shift F1
                ExecuteMenuCommand ("Mnu_Arg")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            Else
                ExecuteMenuCommand ("Mnu_HelpOnLine")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyN
            If bCtrlDown Then
                If BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled Then
                    ExecuteMenuCommand ("Mnu_New")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyX
'            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Cut")
'                    ShortCut = True
'                End If
''                KeyCode = 0
''                Shift = 0
'            End If
            
        Case vbKeyC
            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Copy")
'                    ShortCut = True
'                End If
                KeyCode = 0
                Shift = 0
            End If
            If bAltDown And BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible Then
                ClosePreview
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyV
            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Paste")
'                    ShortCut = True
'                End If
                'KeyCode = 0
                'Shift = 0
            End If
            
        Case vbKeyT
            If bCtrlDown And bShiftDown = False Then   'CTRL + T
                If BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled Then
                    ExecuteMenuCommand ("Mnu_NewSearch")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            If bCtrlDown And bShiftDown = True Then     'CTRL + MAIUSC + T
                If BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchFilter")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyE
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled Then
                    ExecuteMenuCommand "Mnu_ExecuteSearch"
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyP
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchPrevious")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyS
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchNext")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyM
            If bCtrlDown Then
                If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled Then
                    ExecuteMenuCommand ("Mnu_TableView")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyF
            If bCtrlDown Then
                If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled Then
                    ExecuteMenuCommand ("Mnu_FormView")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyDelete
            'Il tasto Canc ha effetto solo se il controllo attivo è la browse principale.
            If ActiveControl.Name = "BrwMain" And Not bShiftDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled Then
                    If BrwMain.Visible Then
                        ExecuteMenuCommand ("Mnu_Delete")
                        ShortCut = True
                        KeyCode = 0
                        Shift = 0
                    End If
                End If
            End If
            
        Case vbKeyR
            If bCtrlDown Then
                ExecuteMenuCommand ("Mnu_SearchObject")
                ShortCut = True
                'La condizione sottostante è necessaria per attivare l'acceleratore CTRL+R dalla modalità
                'filtri della DmrGrid
                If Not BrwMain.Visible Or (BrwMain.Visible And BrwMain.GuiMode = dgNormal) Then
                    KeyCode = 0
                    Shift = 0
                End If
            End If
            
        Case vbKeyG
            If bCtrlDown Then
                ExecuteMenuCommand ("Mnu_RunApplication")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
    
        Case vbKeyEscape
             If Not ActiveControl Is Nothing Then
                    If TypeName(ActiveControl) = "DmtGrid" Then
                        If BrwMain.GuiMode = dgFilterDefinition Then
                            If Not (m_Document.EOF = True And m_Document.BOF = True) Then
                                BrwMain.GuiMode = dgNormal
                                ExecuteMenuCommand "Mnu_TableView"
                                ShortCut = True
                            Else
                                'Ripulisce il contenuto delle condizioni.
                                BrwMain.Conditions.ClearValues
                                'Imposta la modalità FilterDefinition
                                BrwMain.GuiMode = dgFilterDefinition
                                ShortCut = True
                            End If
                        End If
                    End If
                KeyCode = 0
                Shift = 0
            End If
    
    End Select

End Function


'**+
'Nome: ShowErrorLog
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Mostra il dialogo di informazioni su l'ultimo errore
'bloccante verificatosi durante l'esecuzione del programma
'**/
Private Sub ShowErrorLog()
    Load frmErrorLog
    frmErrorLog.DMTErrorContol.MainProgram.Comments = App.Comments
    frmErrorLog.DMTErrorContol.MainProgram.Company = App.CompanyName
    frmErrorLog.DMTErrorContol.MainProgram.Copyright = App.LegalCopyright
    frmErrorLog.DMTErrorContol.MainProgram.Description = App.FileDescription
    frmErrorLog.DMTErrorContol.MainProgram.FileName = App.EXEName
    frmErrorLog.DMTErrorContol.MainProgram.Version = App.Major & "." & App.Minor & "." & App.Revision
    frmErrorLog.DMTErrorContol.ErrorNumber = Err.Number
    frmErrorLog.DMTErrorContol.ErrorDescription = Err.Description
    frmErrorLog.DMTErrorContol.Show
    frmErrorLog.Show vbModal
    End
End Sub


'**+
'Nome: OnBeforeOpenDoc
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazioni da effettuare prima dell'apertura del documento.
'**/
Private Sub OnBeforeOpenDoc()


End Sub


'**+
'Autore: Carlo B. Collovà
'Data creazione: 20/11/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: InitExtensions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Inizializza la componente adibita alla gestione dell'evento On_Extend
'
'**/
'Private Sub InitExtensions()
    'cbcx
    
    'Istanzia l'oggetto
    'Set m_ExtendApplication = New DmtExtendApp.ExtendApplication
    'Set m_ExtendApplication = New DmtExtendAppLib.ExtendApplication
    
    'Assegna un riferimento all'oggetto Application.
    'In questo modo la maggior parte dei parametri di inizializzazione vengono
    'letti da quest'ultimo
    'Set m_ExtendApplication.Application = m_App
    
    'Assegna un riferimento al controllo ActiveBar affinchè la classe
    'che gestisce i dati aggiuntivi possa interagire con la user interface
    'della manutenzione.
    'Set m_ExtendApplication.MenuBar = BarMenu
        
    'Se la funzione correntemente in esecuzione prevede l'evento On_Extend
    'vengono effettuate tutte le inizializzazioni del caso (come l'aggiunta di bottoni
    ' e menu alla BarMenu, ecc.) altrimenti la classe ExtendApplication non effettua
    'alcuna operazione.
    'm_ExtendApplication.Initialize
    
    'NOTA:
    '-----------------------------------------------------------------------------------------------------
    'Tutte le proprietà di m_ExtendApplication presenti anche nell'interfaccia IExtendApplication ed impostate
    'dopo la chiamata al metodo Initialize saranno settate anche in cContactPlus
    '-----------------------------------------------------------------------------------------------------
    
    'Assegna un riferimento del documento corrente
    'Set m_ExtendApplication.CurrentDocument = m_Document

'End Sub


'**+
'Nome: Start
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione e procedura di avvio
'**/
Private Sub Start()
    Dim OLDCursor As Integer
    Dim ToolID As Integer
    Dim Field As DmtDocManLib.Field
    Dim oActivity As IActivity
    Dim o As Activity
    Dim oFilter As Filter

        
    
    'Apertura del documento
    If Len(m_ExtendedDatabase) > 0 Then
        'Apre un nuovo documento usando il database esteso
        Set m_Document = m_App.OpenDocument(m_DocType, m_ExtendedDatabase)
    Else
        'Apre un nuovo documento usando il database diamante
        Set m_Document = m_App.OpenDocument(m_DocType)
    End If
    
    
    'NOTA: Con la sottostante proprietà settata a TRUE i metodi OnXXXDocumentsLink()
    'non sono più necessari in quanto il modello ad oggetti si occupa della gestione
    'dei sottodocumenti.
    '
    'Abilita la gestione automatica degli eventuali DocumentsLink
    m_Document.EnableRefreshDocumentsLinks = True
    
    
    'Clessidra
    OLDCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Caption = m_App.Caption
    
    'Inizializzazione del controllo ActiveBar
    InitMenuBar ToolID
    InitToolBar ToolID
    ActivateBarButtons BTN_ALL, True
    
    'Inizializzazione del riquadro attività
    With ActivityBox
        .Activities.Clear
        
        'Aggiunge l'attività dei reports
        Set oActivity = .Activities.Add("DmtActBoxLib.ReportsActivity", "Reports")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID, TheApp.IDFirm
        Set o = oActivity
        Set oReportsActivity = o.InternalClass
        
        'Aggiunge l'attività dei filtri
        Set oActivity = .Activities.Add("DmtActBoxLib.FiltersActivity", "Filters")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType
        Set o = oActivity
        Set oFiltersActivity = o.InternalClass
        
        'Aggiunge l'attività delle viste tabellari
        Set oActivity = .Activities.Add("DmtActBoxLib.TableViewsActivity", "TableViews")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID
        Set o = oActivity
        Set oTableViewsActivity = o.InternalClass

        'Aggiunge l'attività delle esportazioni
        Set oActivity = .Activities.Add("DmtActBoxLib.ExportActivity", "Export")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID
        Set o = oActivity
        Set oExportActivity = o.InternalClass
        
        'Aggiunge l'attività del supporto tecnico
        Set oActivity = .Activities.Add("DmtActBoxLib.SupportActivity", "Support")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load
        Set o = oActivity
        Set oSupportActivity = o.InternalClass
        
        'attiva/disattiva la visualizzazione delle attività
        EnableDOMActivitiesItems
        
        'imposta quale attività deve essere attivata per default
        If m_DefaultActivity <> "" Then
            Set .CurrentActivity = .Activities(m_DefaultActivity)
        End If
        
        'ridisegna il controllo
        .Redraw = True
    End With


    'Lettura impostazioni dal registry
    ReadRegistrySettings
        
    'Aggiunge due filtri temporanei, uno per le ricerche temporanee
    'e uno per la stampa in modalità form
    m_DocType.AddFilter "Temp"
    m_DocType.AddFilter "Form"
    
    

    'Connessione di tipo DMTADODBLib
    ConnessioneDiamanteADO
    
    GET_PARAMATRI_FILIALE
    
    'Inizializzazioni da fare prima dell'apertura del documento
    OnBeforeOpenDoc
    
    
    'rif12
    'Altre inizializzazioni
    OnStart
    
    
    
    If Len(m_App.Caller) > 0 And m_App.CallerFieldValue > 0 Then
        '-------------------------------------------------
        '     Il programma è stato chiamato da un link.
        '-------------------------------------------------
        
        'In tal caso occorre mostrare in modalità variazione il record richiesto dal programma client.
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        For Each Field In m_DocType.Fields
            Field.Value = Empty
        Next
                
        'Imposta una condizione di ricerca basata sull'ID del record richiesto dal programma client.
        m_DocType.Fields("ID" & m_App.TableName).Value = m_App.CallerFieldValue
        
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"
        
        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")
        
        'Inidica, nel caso di esegui gestione, se riportare il valore corrente al chiamante
        bNotReturnValue = CBool(Val(GetSetting(REGISTRY_KEY, App.EXEName, "NoReturnValue", "0")))
    Else
        '---------------------------------------------------
        '     Il programma non è stato chiamato da un link.
        '---------------------------------------------------
    
        'Il filtro attivo alla partenza è quello predefinito
        For Each oFilter In m_DocType.Filters
            If oFilter.ID = oFiltersActivity.DefaultFilterID Then
                Set m_ActiveFilter = m_DocType.Filters(oFilter.Name)
                Exit For
            End If
        Next
    End If
    
        
    'Si comunica al documento quale filtro eseguire all'avvio.
    Set m_Document.ActiveFilter = m_ActiveFilter
    'm_Document.Dataset.Recordset.Sort = "DataDocumento DESC, NumeroDocumento DESC"
    'Set Me.BrwMain.Recordset = m_Document.Dataset.Recordset
    'Prima di aprire il documento occorre comunicargli qual'è il campo chiave primaria.
    m_Document.PrimaryKey = "ID" & m_Document.TableName
    'Apertura del documento.
    m_Document.OpenDoc
    
    'Questa impostazione serve per conservare le impostazioni grafiche
    BrwMain.IDUser = m_App.IDUser
    'Permette di gestire l'evento BrwMain_OnApplyFilter
    BrwMain.AutoFiltering = False
    'Con questa impostazione la dmtGrid NON effettua mai il Move sul documento.
    'Questo pertanto andrà forzato in BrwMain_DblClick e BrwMain_KeyDown.
    BrwMain.EnableMove = False
    'Inizializza le colonne da visualizzare nella griglia
    If m_DocType.DefaultTableView Is Nothing Then
        Err.Raise ERR_NO_DEFAULT_TABLEVIEW, , "Default TableView not found"
    Else
        LOAD_COLUMN 'BrwMain.LoadColumns m_DocType.DefaultTableView
        BrwMain.LoadUserSettings
        SetVisibilityIDFields
    End If
    
    
    'Crea i campi per la ricerca.
    CreateBrowserConditions
    'Assegnazione del riferimento alla fonte dati (binding sul recordset del documento)
    
    'rif14

    
    'Set BrwMain.Recordset = m_Document.Dataset.Recordset
    Set BrwMain.Recordset = m_Document.Data
    
    
            
     'Viene inizializzato il dialogo di stampa
    With DmtPrnDlg
        Set .Application = m_App
        Set .DocType = m_DocType
    End With
    
    
    
    'Ripulisco la tabella semaforo.
    'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    
    'Evita il blocco della toolbar
    'BarMenu.ResetHooks
    
    Screen.MousePointer = OLDCursor

End Sub




'**+
'Nome: ConditionType
'
'Parametri: DBType è il valore di DMTDocManLib.Field.DBType e rappresenta
'           il tipo di dato corrispondente all'oggetto Field in base dati.
'
'Valori di ritorno: una costante di tipo ConditionTypeConstants usata dalla Browse
'                   per costruire una condizione di ricerca.
'
'Funzionalità: Trasforma una costante DBType in una costante compatibile ConditionTypeConstants
'**/
Private Function ConditionType(ByVal DBType As Integer) As dmtgridctl.ConditionTypeConstants
    Select Case DBType

        'dbTypeCHAR, dbTypeVARCHAR, dbTypeWCHAR, dbTypeWVARCHAR
        Case 1, 12, -8, -9
            ConditionType = dgCondTypeText
       
        'dbTypeNUMERIC, dbTypeDECIMAL, dbTypeINTEGER, dbTypeSMALLINT, dbTypeFLOAT
        'dbTypeREAL, dbTypeDOUBLE, dbTypeBIGINT, dbTypeTINYINT
        Case 2, 3, 4, 5, 6, 7, 8, -5, -6
            ConditionType = dgCondTypeNumber
            
        'dbTypeTIMESTAMP  ////NOTA: Se si desidera un campo dmCondTypeTime occorre gestirlo ad Hoc.
        Case 135
            ConditionType = dgCondTypeDate
    
        'dbTypeBIT
        Case -7, 11
            ConditionType = dgCondTypeBoolean
            
    End Select
End Function

'**+
'Nome: CreateBrowserConditions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Crea automaticamente i campi per la ricerca (modalità DefineFilter)
'              a partire dai campi non ID del documento.
'**/
Private Sub CreateBrowserConditions()
    Dim Field As DmtDocManLib.Field
    Dim Cond As dmtgridctl.dgCondition
    
    'Vengono creati automaticamente i campi per la ricerca.
    'In una applicazione specifica questo codice andrà sostituito integralmente per definire
    'dei campi di ricerca ad hoc.
    
    'Non viene visualizzata la Check Intervallo perchè attualmente
    'il modello ad oggetti non prevede la gestione di filtri con
    'clausole BETWEEN.
    
    If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
        Me.BrwMain.ConnectionString = MenuOptions.ConnectionString & "User Id=" & m_App.User & ";Password=" & m_App.Password
    Else
        Me.BrwMain.ConnectionString = MenuOptions.ConnectionString & ";" & "User Id=" & m_App.User & ";Password=" & m_App.Password
    End If
    
    BrwMain.Conditions.Clear

    BrwMain.Conditions.WidthConditions = 350
    BrwMain.Conditions.WidthFields = 250
    BrwMain.Conditions.WidthIntervals = 100
    
    BrwMain.Title.BackColor = vb3DFace
    BrwMain.Title.ForeColor = vbBlack
    BrwMain.Title.Font.Bold = True
    
    BrwMain.Conditions.Add "Group1", "Dati generali documento", ""
    BrwMain.Conditions("Group1").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("NumeroCertificato", "Numero certificato", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("DataCertificato", "Data certificato", m_DocType.TableName, False, True, False, dgCondTypeDate)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("NumeroDocumentoSocio", "Numero D.D.T. socio", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("DataDocumentoSocio", "Data D.D.T. socio", m_DocType.TableName, False, True, False, dgCondTypeDate)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("IDSezionale", "Sezionale", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.Indentation = 20
        Cond.RecordSource = "SELECT * FROM Sezionale WHERE IDFiliale=" & TheApp.Branch & "  ORDER BY Sezionale"
        Cond.DisplayField = "Sezionale"
        Cond.KeyField = "IDSezionale"

    
    BrwMain.Conditions.Add "Group2", "Anagrafica cliente", ""
    BrwMain.Conditions("Group2").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("CodiceAnagraficaCliente", "Codice", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("Anagrafica", "Ragione sociale", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("Nome", "Ragione sociale 2", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("SitoPerAnagrafica", "Destinazione", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
        
    BrwMain.Conditions.Add "Group3", "Cooperativa", ""
    BrwMain.Conditions("Group3").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("CodiceAnagraficaCooperativa", "Codice", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("AnagraficaCooperativa", "Ragione sociale", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("NomeCooperativa", "Ragione sociale 2", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
        
    BrwMain.Conditions.Add "Group4", "Socio", ""
    BrwMain.Conditions("Group4").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("CodiceAnagraficaSocio", "Codice", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("AnagraficaSocio", "Ragione sociale", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("NomeSocio", "Ragione sociale 2", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("Acquistato", "Acquisto", m_DocType.TableName, False, False, , dgCondTypeBoolean)
       'Cond.FromValue = "NO"
       Cond.Indentation = 20
       
    BrwMain.Conditions.Add "Group5", "Articolo", ""
    BrwMain.Conditions("Group5").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("CodiceArticolo", "Codice", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("DescrizioneArticolo", "Descrizione", m_DocType.TableName, False, False, False, dgCondTypeText)
        Cond.Indentation = 20
        
    BrwMain.Conditions.Add "Group9", "Lotto di produzione", ""
    BrwMain.Conditions("Group9").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("CodiceLotto", "Codice", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("Varieta", "Varietà", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("FamigliaProdotti", "Famiglia", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("AnnoDiRiferimentoPeriodoDiCampagna", "Anno di riferimento", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("DataInizioPeriodoDiCampagna", "Data inizio periodo di campagna", m_DocType.TableName, False, True, , dgCondTypeDate)
        Cond.Indentation = 20
        Cond.RangeChecked = True
    Set Cond = BrwMain.Conditions.Add("DataFinePeriodoDiCampagna", "Data fine periodo di campagna", m_DocType.TableName, False, True, , dgCondTypeDate)
        Cond.Indentation = 20
        Cond.RangeChecked = True
        
    BrwMain.Conditions.Add "Group6", "Documento di trasporto", ""
    BrwMain.Conditions("Group6").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("DataDocumentoDDT", "Data documento", m_DocType.TableName, False, True, , dgCondTypeDate)
        Cond.Indentation = 20
        Cond.RangeChecked = True
    Set Cond = BrwMain.Conditions.Add("NumeroDocumentoDDT", "Numero", m_DocType.TableName, False, True, False, dgCondTypeNumber)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("PrefissoSezionaleDocumentoDDT", "Prefisso", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
        
    BrwMain.Conditions.Add "Group7", "Fattura differita", ""
    BrwMain.Conditions("Group7").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("DataDocumentoFD", "Data documento", m_DocType.TableName, False, True, , dgCondTypeDate)
        Cond.Indentation = 20
        Cond.RangeChecked = True
    Set Cond = BrwMain.Conditions.Add("NumeroDocumentoFD", "Numero", m_DocType.TableName, False, True, False, dgCondTypeNumber)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("PrefissoSezionaleDocumentoFD", "Prefisso", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
        
    BrwMain.Conditions.Add "Group8", "Contratto", ""
    BrwMain.Conditions("Group8").IsHeader = True
    Set Cond = BrwMain.Conditions.Add("DataDocumentoContratto", "Data documento", m_DocType.TableName, False, True, , dgCondTypeDate)
        Cond.Indentation = 20
        Cond.RangeChecked = True
    Set Cond = BrwMain.Conditions.Add("NumeroDocumentoContratto", "Numero", m_DocType.TableName, False, True, False, dgCondTypeNumber)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("NumeroNsDocumentoContratto", "Riferimento interno", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("NumeroVsDocumentoContratto", "Riferimento cliente", m_DocType.TableName, False, True, False, dgCondTypeText)
        Cond.Indentation = 20
        
End Sub

'**+
'Nome: Export
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue l'esportazione del documento con controllo di errore
'**/
Private Sub ExportDocument(ByVal Appl As Long)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    
    OLDCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    m_Document.Export m_Report, Appl
    Screen.MousePointer = OLDCursor
    Exit Sub
errHandler:
    Screen.MousePointer = OLDCursor
    
    If Err.Number = 20507 Then
        'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
        sbMsgInfo "File di report non trovato", m_App.FunctionName
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
End Sub

'**+
'Nome: PrintDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue la stampa del documento con controllo di errore per nessuna stampante
'definita
'**/
Private Sub PrintDocument(ByVal ToolName As String)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    
    '**+ Riferimento al cursore corrente
    OLDCursor = Screen.MousePointer
    
    '**+ Inizializzazione selezioni di stampa
    m_Report.Copies = 1
    m_Report.Orientation = ocPortrait
    m_Report.PrinterName = ""
    
    If ToolName = "Mnu_Print" Then
        '**+ stampa con dialogo
        Set DmtPrnDlg.Report = m_Report
        DmtPrnDlg.Show
        If Not DmtPrnDlg.Cancel Then
            Screen.MousePointer = vbHourglass
            m_Document.DoPrint m_Report
        End If
    Else
        'Stampa diretta
        Screen.MousePointer = vbHourglass
        m_Document.DoPrint m_Report
    End If
    
    Screen.MousePointer = OLDCursor
    Exit Sub

errHandler:
    Screen.MousePointer = OLDCursor
    If Err.Number = vbObjectError + 36 Then
        ' errore generato all'interno della DMTDocManLib per nessuna stampante
        sbMsgInfo "Non è possibile ottenere informazioni sulla stampante." & Chr(13) & "Controllare che sia installata correttamente", m_App.FunctionName
    ElseIf Err.Number = vbObjectError + 4 Then
        'Si è annullata la stampa.
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
    
End Sub

'**+
'Nome: DoNewDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Procedure per la richiesta di un nuovo documento
'**/
Private Function DoNewDocument() As Integer
    
    '------------------------------------------------
    'Inserire qui se occorre del codice specifico
    'per la manutenzione.
    '------------------------------------------------
    
    DoNewDocument = ChooseAboutSaving
End Function

'**+
'Nome: WriteStatusBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Scrive una stringa di testo nella StatusBar
'**/
Private Sub WriteStatusBar(ByVal sTesto As String)
    If stbStatusbar.Style = sbrSimple Then
        stbStatusbar.SimpleText = sTesto
    Else
        stbStatusbar.Panels(1).Text = sTesto
    End If
End Sub

'**+
'Nome: FormUnload
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue i controlli alla richiesta di abbandono del form
'**/
Private Function FormUnload() As Integer
    Dim sMessage As String
    Dim sMessage1 As String
    Dim lIDField As Long

    
    If m_Changed Then
        Select Case ChooseAboutSaving
            Case vbCancel
                FormUnload = 1
                Exit Function
            Case vbYes
                OnSave
                'Se la registrazione non è andata a buon fine
                'esce e non chiude il programma
                If Not m_Saved Then
                    FormUnload = 1
                    Exit Function
                End If
        End Select
    End If
        
    
    If m_PreviewWindowHandle > 0 Then
        ClosePreview
    End If
    
    SaveRegistrySettings
    
    'Se il programma è stato chiamato da un link occorre restituire l'ID del record attivo
    'all'applicazione chiamante.
    If Len(m_App.Caller) > 0 Then
        'Il programma è stato chiamato da un link.
        
        'Se non verrà correttamente selezionato un elemento sarà restituito il valore -1 all'applicazione client.
        lIDField = -1
        
        'Se il documento è vuoto non si deve far nulla.
        'Se la browse è in modalità Filter Definition non formula la domanda di riporto dei dati nel programma chiamante.
        If (Not (m_Document.EOF And m_Document.BOF)) And (BrwMain.GuiMode <> dgFilterDefinition) Then
        
            'ATTENZIONE: La stringa sMessage1 deve essere personalizzata a seconda dei casi!!!
            sMessage1 = " il " & m_DocType.Name
            sMessage = sMessage1 & " """ & m_Document.Fields(CAMPO_PER_CAPTION).Value & """"
            
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add sMessage, 1
                              
            'Viene chiesto se si intende riportare il record corrente al programma chiamante.
            If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYPASTE), m_App.FunctionName) = vbYes Then
                'Legge l'ID del record corrente affinchè venga riportato all'applicazione chiamante.
                lIDField = m_Document.Fields("ID" & m_App.TableName).Value
            End If
            
        End If
        
        'Scrive sul registry l'ID da passare all'aplicazione chiamante.
        SaveSetting "Diamante", m_App.Caller, "IDField", lIDField
                                
    End If
    
End Function

'**+
'Nome: FormRecalcLayout
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Ricalcolo del layout del form
'**/
Private Sub FormRecalcLayout()
    Dim Height As Single
    Dim Width As Single

    'Se il form è minimizzato non serve il ricalcolo del layout
    If WindowState <> vbMinimized Then
        ActivityBox.Top = BarMenu.ClientAreaTop
        ActivityBox.Left = BarMenu.ClientAreaLeft
        ActivityBox.Height = IIf(BarMenu.ClientAreaHeight > 0, BarMenu.ClientAreaHeight, 0)
        
        imgSplitter.Visible = ActivityBox.Visible
        imgSplitter.Top = ActivityBox.Top
        imgSplitter.Height = ActivityBox.Height
        
        If ActivityBox.Visible Then
            imgSplitter.Left = ActivityBox.Width + ActivityBox.Left
            picSplitter.Left = imgSplitter.Left
        End If
        
        PicForm.Top = BarMenu.ClientAreaTop
        
        If ActivityBox.Visible Then
            PicForm.Left = imgSplitter.Left + imgSplitter.Width
        Else
            PicForm.Left = BarMenu.ClientAreaLeft
        End If
        


        Width = BarMenu.ClientAreaWidth - IIf(ActivityBox.Visible, ActivityBox.Width + imgSplitter.Width, 0)
        Height = BarMenu.ClientAreaHeight
        
        PicForm.Width = IIf(Width < 100, 100, Width)
        PicForm.Height = IIf(Height < 100, 100, Height)
        
        'RIDIMENSIONA LA SPLIT BAR IN BASE ALLA DIMENSIONE DEL FORM
        DMTSplitBar1.Move PicForm.Left, PicForm.Top, PicForm.Width, PicForm.Height
        'INIZIALIZZA LA SPLIT BAR
        DMTSplitBar1.SetSplitBar Height, Width, Me.PicForm2.Height, Me.PicForm2.Width
        
        'PicForm.Top = BarMenu.ClientAreaTop
        
        'If ActivityBox.Visible Then
        '    PicForm.Left = imgSplitter.Left + imgSplitter.Width
        'Else
        '    PicForm.Left = BarMenu.ClientAreaLeft
        'End If
        
        BrwMain.Top = PicForm.ScaleTop
        BrwMain.Left = PicForm.ScaleLeft
        BrwMain.Width = PicForm.ScaleWidth
        BrwMain.Height = PicForm.ScaleHeight
        

    End If
End Sub

'**+
'Nome: GetStatusToolBar
'
'Parametri:
'Enabled - Stato di abilitazione da controllare
'
'Valori di ritorno:
'
'Funzionalità:
'Calcola lo stato dei bottoni della ToolBar standard
'**/
Private Function GetStatusToolBar(ByVal Enabled As Boolean) As Currency
    Dim Valore As Currency

    Valore = 0
    If BarMenu.Bands("Standard").Tools("New").Enabled = Enabled Then Valore = Valore Or BTN_NEW
    If BarMenu.Bands("Standard").Tools("Save").Enabled = Enabled Then Valore = Valore Or BTN_SAVE
    If BarMenu.Bands("Standard").Tools("Print").Enabled = Enabled Then Valore = Valore Or BTN_PRINT
    If BarMenu.Bands("Standard").Tools("PrePrint").Enabled = Enabled Then Valore = Valore Or BTN_PREVIEW
    If BarMenu.Bands("Standard").Tools("Cut").Enabled = Enabled Then Valore = Valore Or BTN_CUT
    If BarMenu.Bands("Standard").Tools("Copy").Enabled = Enabled Then Valore = Valore Or BTN_COPY
    If BarMenu.Bands("Standard").Tools("Paste").Enabled = Enabled Then Valore = Valore Or BTN_PASTE
    If BarMenu.Bands("Standard").Tools("Delete").Enabled = Enabled Then Valore = Valore Or BTN_DELETE
    If BarMenu.Bands("Standard").Tools("Clear").Enabled = Enabled Then Valore = Valore Or BTN_CLEAR
    If BarMenu.Bands("Standard").Tools("NewSearch").Enabled = Enabled Then Valore = Valore Or BTN_FIND
    If BarMenu.Bands("Standard").Tools("ExecuteSearch").Enabled = Enabled Then Valore = Valore Or BTN_SEARCH
    If BarMenu.Bands("Standard").Tools("ChangeView").Enabled = Enabled Then Valore = Valore Or BTN_VIEWMODE
    If BarMenu.Bands("Standard").Tools("SearchPrevious").Enabled = Enabled Then Valore = Valore Or BTN_PREVIOUS
    If BarMenu.Bands("Standard").Tools("SearchNext").Enabled = Enabled Then Valore = Valore Or BTN_NEXT
    If BarMenu.Bands("Standard").Tools("Export").Enabled = Enabled Then Valore = Valore Or BTN_EXPORT
    If BarMenu.Bands("Band_Export").Tools("ExportWord").Enabled = Enabled Then Valore = Valore Or BTN_WORD
    If BarMenu.Bands("Band_Export").Tools("ExportExcel").Enabled = Enabled Then Valore = Valore Or BTN_EXCEL
    If BarMenu.Bands("Band_Export").Tools("ExportHtml").Enabled = Enabled Then Valore = Valore Or BTN_HTML
    If BarMenu.Bands("Band_Export").Tools("ExportPDF").Enabled = Enabled Then Valore = Valore Or BTN_PDF
    
    If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHFORM
    If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHTABLE
    If BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled = Enabled Then Valore = Valore Or BTN_FILTER

    If BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled = Enabled Then Valore = Valore Or BTN_NEW
    If BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled = Enabled Then Valore = Valore Or BTN_SAVE
    If BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled = Enabled Then Valore = Valore Or BTN_PRINT
    If BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Enabled = Enabled Then Valore = Valore Or BTN_PREVIEW

    If BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled = Enabled Then Valore = Valore Or BTN_CUT
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled = Enabled Then Valore = Valore Or BTN_COPY
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled = Enabled Then Valore = Valore Or BTN_PASTE
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled = Enabled Then Valore = Valore Or BTN_DELETE
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Enabled = Enabled Then Valore = Valore Or BTN_CLEAR
    If BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled = Enabled Then Valore = Valore Or BTN_FIND
    If BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled = Enabled Then Valore = Valore Or BTN_SEARCH
    If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled = Enabled Then Valore = Valore Or BTN_PREVIOUS
    If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled = Enabled Then Valore = Valore Or BTN_NEXT

    If BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = Enabled Then Valore = Valore Or BTN_TOOLS
    If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHFORM
    If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHTABLE

    If BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Enabled = Enabled Then Valore = Valore Or BTN_EXPORT
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Enabled = Enabled Then Valore = Valore Or BTN_WORD
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Enabled = Enabled Then Valore = Valore Or BTN_EXCEL
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Enabled = Enabled Then Valore = Valore Or BTN_HTML
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Enabled = Enabled Then Valore = Valore Or BTN_PDF

    GetStatusToolBar = Valore
End Function

'**+
'Nome: ReadRegistrySettings
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge i valori registrati nel registry relativi allo stato
'dei controlli del Form
'
'
'**/
Private Sub ReadRegistrySettings()
    Dim Index As Integer
    Dim FormHeight As Single
    Dim FormWidth As Single
    Dim NomeBanda As String
    Dim lngIDLanguage As Long
    Dim bFoldersVisible As Boolean
    Dim lValue As Long
           
           
    'Lettura file di help
    App.HelpFile = MenuOptions.ProgramsPath & "\Diamante.chm"
           
    ' Legge dal Registry le impostazioni sulla lingua
    lngIDLanguage = AppOptions.IDLanguage
           
    ' Modifica tutte le stringhe nel linguaggio corrente ( se <> da default )
    If lngIDLanguage <> NATIVE_LANGUAGE Then
        gResource.IDCurrentLanguage = lngIDLanguage
        'Setta i nuovi ToolTipText della Toolbar
        'e le Caption dei menu
        ChangeMenuLanguage
        ChangeToolBarLanguage
        'Traduce tutte le stringhe presenti sul form
        '(Solo se ChangeStringsLanguage è gestita dal programmatore !!!)
        ChangeStringsLanguage
    End If
    
    'Settaggio per la statusbar
    stbStatusbar.Visible = AppOptions.StatusBarVisibility
        
        
    '**+ settaggi per la barra degli strumenti
    With BarMenu
    
        '**+ E' necessario verificare la versione dell'activebar xchè nella nuova vesione 3.0
        'sono stati cambiati i valori di impostazione della proprietà DockingArea
        If AppOptions.BARMENUVERSION = BARMENUVERSION Then
    
            For Index = 0 To .Bands.Count - 1
            
                'Settaggi sulle toolbar (ancoraggio e dimensioni)
                If .Bands(Index).Type <> ddBTPopup Then
                    With .Bands(Index)
                        If AppOptions.ToolbarDockingArea(Index) > -1 Then
                            .DockingArea = AppOptions.ToolbarDockingArea(Index)
                            .DockLine = AppOptions.ToolbarDockLine(Index)
                            lValue = AppOptions.ToolbarHeight(Index)
                            If lValue > 0 Then .Height = lValue
                            lValue = AppOptions.ToolbarWidth(Index)
                            If lValue > 0 Then .Width = lValue
                            '**+ Attenzione le impostazioni del Left e Top devono essere effettuate dopo
                            'quelle dell'Height e del Width xchè se siamo in presenza di valori superiori
                            'a quelli della ClientArea azzera il left e top impostati in precedenza **/
                            lValue = AppOptions.ToolbarLeft(Index)
                            If lValue > 0 Then .Left = lValue
                            lValue = AppOptions.ToolbarTop(Index)
                            If lValue > 0 Then .Top = lValue
                            .DockingOffset = AppOptions.ToolbarDockingOffset(Index)
                        End If
                    End With
                End If
            
                'Settaggi sulla visibilità delle toolbar.
                If .Bands(Index).Type = ddBTNormal And .Bands(Index).Name <> BAND_CLOSE_PREVIEW Then
                     NomeBanda = .Bands(Index).Name
                     .Bands(NomeBanda).Visible = AppOptions.ToolbarVisibility(NomeBanda)
                End If
        
            Next Index
            
        End If
        
        'Settaggio sulla visualizzazione dei tooltip.
        .DisplayToolTips = AppOptions.DisplayTooltip
    End With
        
    
    'Dimensione delle icone della ToolBar
    SetToolBarIcons AppOptions.LargeIcon
    
    BarMenu.RecalcLayout
   
    bFoldersVisible = AppOptions.FoldersVisibility
   
    'Settaggi del riquadro attività
    ActivityBox.Visible = bFoldersVisible
    ActivityBox.Width = AppOptions.FoldersWidth
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked = bFoldersVisible
    m_DefaultActivity = AppOptions.DefaultActivity
    
    '**+ settaggi per la finestra principale del programma
    WindowState = AppOptions.WindowState
    If WindowState = 0 Then
        FormHeight = AppOptions.FormHeight
        If FormHeight <> -1 Then
            Height = FormHeight
        End If
        FormWidth = AppOptions.FormWidth
        If FormWidth <> -1 Then
            Width = FormWidth
        End If
    End If
End Sub

'**+
'Nome: SaveRegistrySettings
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Salva i valori relativi allo stato dei controlli del Form
'nel registry
'
'**/
Private Sub SaveRegistrySettings()
    Dim I As Integer

    '**+ Salva le impostazioni relative alle toolbar
    With AppOptions
        
        For I = 0 To BarMenu.Bands.Count - 1
            If BarMenu.Bands(I).Type <> ddBTPopup Then
                    .ToolbarDockingArea(I) = BarMenu.Bands(I).DockingArea
                    .ToolbarDockLine(I) = BarMenu.Bands(I).DockLine
                    .ToolbarLeft(I) = BarMenu.Bands(I).Left
                    .ToolbarTop(I) = BarMenu.Bands(I).Top
                    .ToolbarHeight(I) = BarMenu.Bands(I).Height
                    .ToolbarWidth(I) = BarMenu.Bands(I).Width
                    .ToolbarDockingOffset(I) = BarMenu.Bands(I).DockingOffset
            End If
        Next I
        .BARMENUVERSION = BARMENUVERSION
        
        'Salva le impostazioni relative alla finestra principale.
        If WindowState <> vbMinimized Then
            .FormHeight = Height
            .FormWidth = Width
            .WindowState = WindowState
        End If
        
        'Salva le impostazioni del riquadro attività
        .FoldersWidth = ActivityBox.Width
        .FoldersVisibility = ActivityBox.Visible
        .DefaultActivity = ActivityBox.CurrentActivityKey
    End With
End Sub

'**+
'Nome: ChangeView
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Cambia modalità di visualizzazione dei dati tra Form e vista tabellare
'
'**/
Private Sub ChangeView(Optional ByVal sToolName As Variant)

    'Se non vi sono record presenti nel browser
    'la modalità di visualizzazione non cambia e si esce.
    If (m_Document.EOF = True And m_Document.BOF = True) Then Exit Sub

    'Se si proviene dalla modalità tabellare
    '( o dalla modalità filtro provenendo dalla modalità tabellare )
    'potrebbe essere necessario allineare il documento con l'ultima selezione fatta nella browse.
    If BrwMain.Visible = True Then
        If BrwMain.ListIndex > 0 Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    End If
    

    If IsMissing(sToolName) Then sToolName = "ChangeView"

    'Cambia la visibiltà del browser
    If sToolName = "ChangeView" Then
        BrwMain.Visible = IIf(BrwMain.Visible And BrwMain.GuiMode = dgNormal, False, True)
    Else
        BrwMain.Visible = IIf((sToolName = "Mnu_FormView"), False, True)
    End If
    
    'Se si va in modalità form ed il record è bloccato si torna in modalità tabellare
    'impedendo di effettuare modifiche su quel record.
    'Quando si va in modalità tabellare il controllo non è necessario.
    If Not BrwMain.Visible Then

        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            'Il record è bloccato - si va in modalità tabellare
            
            BrwMain.Visible = True

            'Input Focus al browser
            'BrwMain.SetFocus

            'Refresh dello stato dei bottoni della ToolBar standard e dei menu
            SetStatus4Modality Browse

            Exit Sub
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
        End If

    End If
    
    
    
    'Se si era in fase di immissione di un nuovo record viene annullata
    m_Document.AbortNew
    
    If BrwMain.Visible Then 'Modalità tabellare
        
        'Input Focus al browser
        'BrwMain.SetFocus
        
        'Refresh dello stato dei bottoni della ToolBar standard e dei menu
        SetStatus4Modality Browse
        
    Else 'Modalità form
        
        'Refresh dello stato dei bottoni della ToolBar standard e dei menu
        SetStatus4Modality Modify
        
        'Input Focus al primo campo del form
        SetFocusTabIndex0
    End If
       
    'Imposta i suggerimenti da visualizzare sulla Statusbar in funzione
    'della modalità di visualizzazione corrente.
    'Ad esempio in alcuni casi le frasi sono al Singolare/Plurare.
    'La funzione GetDescription4StatusBar si occupa di determinare la frase esatta.
    'La Sub RefreshDescriptions4StatusBar deve essere chiamata anche in Execute_Search()--> Vedi.
    RefreshDescriptions4StatusBar
End Sub



'**+
'Nome: InitMenuBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione della MenuBar
'
'**/
Private Sub InitMenuBar(ByRef ToolID As Integer)
    BarMenu.Bands.Add "Band_Menu"
    BarMenu.Bands("Band_Menu").WrapTools = True
    BarMenu.Bands("Band_Menu").Type = ddBTMenuBar
    BarMenu.Bands("Band_Menu").DockLine = 1
    BarMenu.Bands("Band_Menu").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
    BarMenu.Bands("Band_Menu").GrabHandleStyle = ddGSNormal

    'File
    BarMenu.Bands.Add "Band_File"
    BarMenu.Bands("Band_File").Type = ddBTPopup
    BarMenu.Bands("Band_File").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "File"
    BarMenu.Bands("Band_Menu").Tools("File").SubBand = "Band_File"
    BarMenu.Bands("Band_Menu").Tools("File").Caption = GetCaption4MenuBar("File")
    BarMenu.Bands("Band_Menu").Tools("File").Description = GetDescription4StatusBar("File")

    'File-New
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_New"
    BarMenu.Bands("Band_File").Tools("Mnu_New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW16), &HC0C0C0
    BarMenu.Bands("Band_File").Tools("Mnu_New").Caption = GetCaption4MenuBar("Mnu_New")
    BarMenu.Bands("Band_File").Tools("Mnu_New").Description = GetDescription4StatusBar("Mnu_New")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_Save"
    BarMenu.Bands("Band_File").Tools("SepMnu_Save").ControlType = ddTTSeparator
    
    'File-Save
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Save"
    BarMenu.Bands("Band_File").Tools("Mnu_Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE16), &HC0C0C0
    If m_App.Language <> 1 Then
        BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    Else
        BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    End If
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Description = GetDescription4StatusBar("Mnu_Save")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_PrePrint"
    BarMenu.Bands("Band_File").Tools("SepMnu_PrePrint").ControlType = ddTTSeparator
    
    'File-PrePrint
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_PrePrint"
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW16), &HC0C0C0
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Caption = GetCaption4MenuBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    
    'File-Print
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Print"
    BarMenu.Bands("Band_File").Tools("Mnu_Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT16), &HC0C0C0
    If m_App.Language <> 1 Then
        BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    Else
        BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    End If
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_Exit"
    BarMenu.Bands("Band_File").Tools("SepMnu_Exit").ControlType = ddTTSeparator
    
    'File-Exit
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Exit"
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Caption = GetCaption4MenuBar("Mnu_Exit")
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Description = GetDescription4StatusBar("Mnu_Exit")
    
    'Edit
    BarMenu.Bands.Add "Band_Edit"
    BarMenu.Bands("Band_Edit").Type = ddBTPopup
    BarMenu.Bands("Band_Edit").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Edit"
    BarMenu.Bands("Band_Menu").Tools("Edit").SubBand = "Band_Edit"
    BarMenu.Bands("Band_Menu").Tools("Edit").Caption = GetCaption4MenuBar("Edit")
    BarMenu.Bands("Band_Menu").Tools("Edit").Description = GetDescription4StatusBar("Edit")
    
    'Edit-Delete
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Delete"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = GetCaption4MenuBar("Mnu_Delete")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Description = GetDescription4StatusBar("Mnu_Delete")
    
    'Edit-Clear
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Clear"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Caption = GetCaption4MenuBar("Mnu_Clear")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Description = GetDescription4StatusBar("Mnu_Clear")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "SepMnu_Cut"
    BarMenu.Bands("Band_Edit").Tools("SepMnu_Cut").ControlType = ddTTSeparator
    
    'Edit-Cut
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Cut"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Caption = GetCaption4MenuBar("Mnu_Cut")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Description = GetDescription4StatusBar("Mnu_Cut")
    
    'Edit-Copy
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Copy"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Caption = GetCaption4MenuBar("Mnu_Copy")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Description = GetDescription4StatusBar("Mnu_Copy")
    
    'Edit-Paste
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Paste"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Caption = GetCaption4MenuBar("Mnu_Paste")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Description = GetDescription4StatusBar("Mnu_Paste")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "SepMnu_NewSearch"
    BarMenu.Bands("Band_Edit").Tools("SepMnu_NewSearch").ControlType = ddTTSeparator
    
    'Edit-NewSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_NewSearch"
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Caption = GetCaption4MenuBar("Mnu_NewSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Description = GetDescription4StatusBar("Mnu_NewSearch")
    
    'Edit-ExecuteSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_ExecuteSearch"
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Caption = GetCaption4MenuBar("Mnu_ExecuteSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Description = GetDescription4StatusBar("Mnu_ExecuteSearch")
    
    'Edit-SearchPrevious
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_SearchPrevious"
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Caption = GetCaption4MenuBar("Mnu_SearchPrevious")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Description = GetDescription4StatusBar("Mnu_SearchPrevious")
    
    'Edit-SearchNext
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_SearchNext"
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Caption = GetCaption4MenuBar("Mnu_SearchNext")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Description = GetDescription4StatusBar("Mnu_SearchNext")
    
    'View
    BarMenu.Bands.Add "Band_View"
    BarMenu.Bands("Band_View").Type = ddBTPopup
    BarMenu.Bands("Band_View").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "View"
    BarMenu.Bands("Band_Menu").Tools("View").SubBand = "Band_View"
    BarMenu.Bands("Band_Menu").Tools("View").Caption = GetCaption4MenuBar("View")
    BarMenu.Bands("Band_Menu").Tools("View").Description = GetDescription4StatusBar("View")
    
    'View-FormView
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_FormView"
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").SetPicture 0, gResource.GetBitmap(IDB_STD_FORM16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'View-TableView
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_TableView"
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").SetPicture 0, gResource.GetBitmap(IDB_STD_GRID16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'View - SearchFilter
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_SearchFilter"
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").SetPicture 0, gResource.GetBitmap(IDB_FILTRO16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "SepMnu_Folders"
    BarMenu.Bands("Band_View").Tools("SepMnu_Folders").ControlType = ddTTSeparator
    
    'View-Folders
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_Folders"
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Caption = GetCaption4MenuBar("Mnu_Folders")
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Description = GetDescription4StatusBar("Mnu_Folders")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "SepMnu_ToolBar"
    BarMenu.Bands("Band_View").Tools("SepMnu_ToolBar").ControlType = ddTTSeparator
    
    'View-ToolBar
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_ToolBar"
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Caption = GetCaption4MenuBar("Mnu_ToolBar")
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Description = GetDescription4StatusBar("Mnu_ToolBar")
    
    'Tools
    BarMenu.Bands.Add "Band_Tools"
    BarMenu.Bands("Band_Tools").Type = ddBTPopup
    BarMenu.Bands("Band_Tools").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Tools"
    BarMenu.Bands("Band_Menu").Tools("Tools").SubBand = "Band_Tools"
    BarMenu.Bands("Band_Menu").Tools("Tools").Caption = GetCaption4MenuBar("Tools")
    BarMenu.Bands("Band_Menu").Tools("Tools").Description = GetDescription4StatusBar("Tools")
    
    'Tools-Export
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Tools").Tools.Add ToolID, "Mnu_Export"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").ControlType = ddTTLabel
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").SubBand = "Mnu_Band_Export"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Caption = GetCaption4MenuBar("Mnu_Export")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Description = GetDescription4StatusBar("Mnu_Export")
    
    'Tools-Options
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Tools").Tools.Add ToolID, "Mnu_Options"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Caption = GetCaption4MenuBar("Mnu_Options")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Description = GetDescription4StatusBar("Mnu_Options")
    
    'Tools-Export
    BarMenu.Bands.Add "Mnu_Band_Export"
    BarMenu.Bands("Mnu_Band_Export").Type = ddBTPopup
    BarMenu.Bands("Mnu_Band_Export").DockingArea = ddDAPopup
    
    'Tools-Export-ExportPDF
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportPDF"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Caption = GetCaption4MenuBar("Mnu_ExportPDF")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Description = GetDescription4StatusBar("Mnu_ExportPDF")
    
    'Tools-Export-ExportWord
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportWord"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("Mnu_ExportWord")
    
    'Tools-Export-ExportExcel
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportExcel"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("Mnu_ExportExcel")
    
    'Tools-Export-ExportHtml
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportHtml"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("Mnu_ExportHtml")

    'Help
    BarMenu.Bands.Add "Band_Help"
    BarMenu.Bands("Band_Help").Type = ddBTPopup
    BarMenu.Bands("Band_Help").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Help"
    BarMenu.Bands("Band_Menu").Tools("Help").Caption = GetCaption4MenuBar("Help")
    BarMenu.Bands("Band_Menu").Tools("Help").Description = GetDescription4StatusBar("Help")
    BarMenu.Bands("Band_Menu").Tools("Help").SubBand = "Band_Help"
    
    'Help-HelpOnLine
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_HelpOnLine"
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Caption = GetCaption4MenuBar("Mnu_HelpOnLine")
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Description = GetDescription4StatusBar("Mnu_HelpOnLine")
    
    'Help-Arg
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Arg"
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Caption = GetCaption4MenuBar("Mnu_Arg")
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Description = GetDescription4StatusBar("Mnu_Arg")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "SepMnu_Web"
    BarMenu.Bands("Band_Help").Tools("SepMnu_Web").ControlType = ddTTSeparator
    
    'Help-Web
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Web"
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB16), &HC0C0C0
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Caption = GetCaption4MenuBar("Mnu_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Description = GetDescription4StatusBar("Mnu_Web")
    
    'Help-Blog
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Agg_Web"
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB16), &HC0C0C0
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Caption = GetCaption4MenuBar("Mnu_Agg_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Description = GetDescription4StatusBar("Mnu_Agg_Web")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "SepMnu_Info"
    BarMenu.Bands("Band_Help").Tools("SepMnu_Info").ControlType = ddTTSeparator
    
    'Help-Info
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Info"
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Caption = GetCaption4MenuBar("Mnu_Info")
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Description = GetDescription4StatusBar("Mnu_Info")
    
    'PopUp
    BarMenu.Bands.Add "Band_PopUp"
    BarMenu.Bands("Band_PopUp").Type = ddBTPopup
    BarMenu.Bands("Band_PopUp").DockingArea = ddDAPopup
    
    'PopUp-RunApplication
    ToolID = ToolID + 1
    BarMenu.Bands("Band_PopUp").Tools.Add ToolID, "Mnu_RunApplication"
    BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").Caption = GetCaption4MenuBar("Mnu_RunApplication")
    
    'PopUp-SearchObject
    ToolID = ToolID + 1
    BarMenu.Bands("Band_PopUp").Tools.Add ToolID, "Mnu_SearchObject"
    BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").Caption = GetCaption4MenuBar("Mnu_SearchObject")
    
    BarMenu.RecalcLayout
End Sub
'**+
'Nome: InitToolBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione della ToolBar
'
'**/
Private Sub InitToolBar(ByRef ToolID As Integer)


    BarMenu.Bands.Add "Standard"
    BarMenu.Bands("Standard").DockLine = 2
    BarMenu.Bands("Standard").Type = ddBTNormal
    BarMenu.Bands("Standard").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
    BarMenu.Bands("Standard").GrabHandleStyle = ddGSNormal
    BarMenu.Bands.Add BAND_CLOSE_PREVIEW
    BarMenu.Bands(BAND_CLOSE_PREVIEW).DockLine = 2
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Type = ddBTMenuBar
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Caption = "Chiudi"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).DockingArea = ddDATop
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = False

    'New
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "New"
    BarMenu.Bands("Standard").Tools("New").ToolTipText = GetToolTipText4ToolBar("New")
    BarMenu.Bands("Standard").Tools("New").Description = GetDescription4StatusBar("New")
    
    'Save
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Save"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep2"
    BarMenu.Bands("Standard").Tools("Sep2").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("Save").ToolTipText = GetToolTipText4ToolBar("Save")
    BarMenu.Bands("Standard").Tools("Save").Description = GetDescription4StatusBar("Save")
    
    'Print
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Print"
    BarMenu.Bands("Standard").Tools("Print").ToolTipText = GetToolTipText4ToolBar("Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")
    
    'PrePrint
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "PrePrint"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep3"
    BarMenu.Bands("Standard").Tools("Sep3").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("PrePrint").ToolTipText = GetToolTipText4ToolBar("PrePrint")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")
    
    'Cut
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Cut"
    BarMenu.Bands("Standard").Tools("Cut").ToolTipText = GetToolTipText4ToolBar("Cut")
    BarMenu.Bands("Standard").Tools("Cut").Description = GetDescription4StatusBar("Cut")
    
    'Copy
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Copy"
    BarMenu.Bands("Standard").Tools("Copy").ToolTipText = GetToolTipText4ToolBar("Copy")
    BarMenu.Bands("Standard").Tools("Copy").Description = GetDescription4StatusBar("Copy")
    
    'Paste
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Paste"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep"
    BarMenu.Bands("Standard").Tools("Sep").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("Paste").ToolTipText = GetToolTipText4ToolBar("Paste")
    BarMenu.Bands("Standard").Tools("Paste").Description = GetDescription4StatusBar("Paste")
    
    'Delete
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Delete"
    BarMenu.Bands("Standard").Tools("Delete").ToolTipText = GetToolTipText4ToolBar("Delete")
    BarMenu.Bands("Standard").Tools("Delete").Description = GetDescription4StatusBar("Delete")
    
    'Clear
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Clear"
    BarMenu.Bands("Standard").Tools("Clear").ToolTipText = GetToolTipText4ToolBar("Clear")
    BarMenu.Bands("Standard").Tools("Clear").Description = GetDescription4StatusBar("Clear")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SepNewSearch"
    BarMenu.Bands("Standard").Tools("SepNewSearch").ControlType = ddTTSeparator
    
    'NewSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "NewSearch"
    BarMenu.Bands("Standard").Tools("NewSearch").ToolTipText = GetToolTipText4ToolBar("NewSearch")
    BarMenu.Bands("Standard").Tools("NewSearch").Description = GetDescription4StatusBar("NewSearch")
    
    'ExecuteSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "ExecuteSearch"
    BarMenu.Bands("Standard").Tools("ExecuteSearch").ToolTipText = GetToolTipText4ToolBar("ExecuteSearch")
    BarMenu.Bands("Standard").Tools("ExecuteSearch").Description = GetDescription4StatusBar("ExecuteSearch")
    
    'ChangeView
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "ChangeView"
    BarMenu.Bands("Standard").Tools("ChangeView").ControlType = ddTTButtonDropDown
    BarMenu.Bands("Standard").Tools("ChangeView").SubBand = "Band_ChangeView"
    BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands.Add "Band_ChangeView"
    BarMenu.Bands("Band_ChangeView").Type = ddBTPopup
    BarMenu.Bands("Band_ChangeView").DockingArea = ddDATop
    
    'ChangeView - Form
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_FormView"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").SetPicture 0, gResource.GetBitmap(IDB_STD_FORM16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'ChangeView - Tabella
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_TableView"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").SetPicture 0, gResource.GetBitmap(IDB_STD_GRID16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
     'ChangeView - Filtro
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_SearchFilter"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").SetPicture 0, gResource.GetBitmap(IDB_FILTRO16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    'SearchPrevious
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SearchPrevious"
    BarMenu.Bands("Standard").Tools("SearchPrevious").ToolTipText = GetToolTipText4ToolBar("SearchPrevious")
    BarMenu.Bands("Standard").Tools("SearchPrevious").Description = GetDescription4StatusBar("SearchPrevious")
    
    'SearchNext
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SearchNext"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep4"
    BarMenu.Bands("Standard").Tools("Sep4").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("SearchNext").ToolTipText = GetToolTipText4ToolBar("SearchNext")
    BarMenu.Bands("Standard").Tools("SearchNext").Description = GetDescription4StatusBar("SearchNext")
        
    'Export
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Export"
    BarMenu.Bands("Standard").Tools("Export").ControlType = ddTTButtonDropDown
    BarMenu.Bands("Standard").Tools("Export").SubBand = "Band_Export"
    BarMenu.Bands("Standard").Tools("Export").ToolTipText = GetToolTipText4ToolBar("Export")
    BarMenu.Bands("Standard").Tools("Export").Description = GetDescription4StatusBar("Mnu_Export")
    BarMenu.Bands.Add "Band_Export"
    BarMenu.Bands("Band_Export").Type = ddBTPopup
    BarMenu.Bands("Band_Export").DockingArea = ddDATop
    
    'ExportPDF
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportPDF"
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Caption = GetCaption4MenuBar("Mnu_ExportPDF")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").ToolTipText = GetToolTipText4ToolBar("ExportPDF")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Description = GetDescription4StatusBar("ExportPDF")
    
    'ExportWord
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportWord"
    BarMenu.Bands("Band_Export").Tools("ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportWord").ToolTipText = GetToolTipText4ToolBar("ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")
    
    'ExportExcel
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportExcel"
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").ToolTipText = GetToolTipText4ToolBar("ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    
    'ExportHtml
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportHtml"
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").ToolTipText = GetToolTipText4ToolBar("ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    
    'Bottone chiusura anteprima
    ToolID = ToolID + 1
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools.Add ToolID, "ClosePreview"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Style = ddSText
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Caption = "&Chiudi"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").ToolTipText = "Chiudi anteprima"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Description = "Esci da modalità Anteprima di stampa"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").ControlType = ddTTButton
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Visible = True
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep5"
    BarMenu.Bands("Standard").Tools("Sep5").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Web"
    BarMenu.Bands("Standard").Tools("Web").ToolTipText = GetToolTipText4ToolBar("Web")
    BarMenu.Bands("Standard").Tools("Web").Description = GetDescription4StatusBar("Web")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Agg_Web"
    BarMenu.Bands("Standard").Tools("Agg_Web").ToolTipText = GetToolTipText4ToolBar("Agg_Web")
    BarMenu.Bands("Standard").Tools("Agg_Web").Description = GetDescription4StatusBar("Agg_Web")
    
    BarMenu.RecalcLayout
End Sub




'**+
'Nome: ChooseAboutSaving
'
'Parametri:
'Ritorna i valori vbYes, vbNo o vbCancel a seconda della risposta data
'
'Valori di ritorno:
'
'Funzionalità:
'Richiesta della registrazione di un record
'**/
Private Function ChooseAboutSaving() As Integer
    If m_Changed Then
        gResource.CustomStrings.Clear
        gResource.CustomStrings.Add Chr(34) & TheApp.FunctionName & Chr(34), 1

        ChooseAboutSaving = fnMsgQuestionWithCancel((gResource.GetCustomizedMessage(MESS_QUERYSAVE)), TheApp.FunctionName)
    End If
End Function

'**+
'Nome: ChooseAboutSavingOkCancel
'
'Parametri:
'
'Valori di ritorno:
'Ritorna i valori vbOK o vbCancel a seconda della risposta data
'
'Funzionalità:
'Come ChooseAboutSaving ma con pulsanti Ok e Annulla
'**/
Private Function ChooseAboutSavingOkCancel() As Integer
    Dim sRecord As String

    sRecord = IIf(m_Document.Fields(CAMPO_PER_CAPTION).Value <> Empty, m_Document.Fields(CAMPO_PER_CAPTION).Value, TheApp.FunctionName)
  
    gResource.CustomStrings.Clear
    gResource.CustomStrings.Add Chr(34) & sRecord & Chr(34), 1
    ChooseAboutSavingOkCancel = fnMsgQuestionOKCancel((gResource.GetCustomizedMessage(MESS_QUERYSAVE)), m_App.FunctionName)
    
End Function

'**+
'Nome: ActivateBarButtons
'
'Parametri:
'Buttons - Variabile lunga 8 byte con la combinazione
'della maschera di bit che indica di quali bottoni cambiare
'lo stato di abilitazione
'Enable - Valore booleano che indica lo stato di abilitazione
'da applicare
'
'Valori di ritorno:
'
'Funzionalità:
'Abilita o meno gruppi di bottoni e voci di menu
'**/
Private Sub ActivateBarButtons(ByVal Buttons As Currency, ByVal Enable As Boolean)

    'Pulsanti della Toolbar
    '----------------------
    If (Buttons And BTN_NEW) Then BarMenu.Bands("Standard").Tools("New").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_SAVE) Then BarMenu.Bands("Standard").Tools("Save").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PRINT) Then BarMenu.Bands("Standard").Tools("Print").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PREVIEW) Then BarMenu.Bands("Standard").Tools("PrePrint").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_CUT) Then BarMenu.Bands("Standard").Tools("Cut").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_COPY) Then BarMenu.Bands("Standard").Tools("Copy").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PASTE) Then BarMenu.Bands("Standard").Tools("Paste").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_DELETE) Then BarMenu.Bands("Standard").Tools("Delete").Enabled = CheckRights("Cancellazione", Enable)
    If (Buttons And BTN_CLEAR) Then BarMenu.Bands("Standard").Tools("Clear").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_FIND) Then BarMenu.Bands("Standard").Tools("NewSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCH) Then BarMenu.Bands("Standard").Tools("ExecuteSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_VIEWMODE) Then BarMenu.Bands("Standard").Tools("ChangeView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_PREVIOUS) Then BarMenu.Bands("Standard").Tools("SearchPrevious").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_NEXT) Then BarMenu.Bands("Standard").Tools("SearchNext").Enabled = CheckRights("Selezione", Enable)
    If Not oExportActivity Is Nothing Then
        If (Buttons And BTN_EXPORT) Then oExportActivity.EnableItems CheckRights("Stampa", Enable)
    End If
    If (Buttons And BTN_EXPORT) Then BarMenu.Bands("Standard").Tools("Export").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_WORD) Then BarMenu.Bands("Band_Export").Tools("ExportWord").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_EXCEL) Then BarMenu.Bands("Band_Export").Tools("ExportExcel").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_HTML) Then BarMenu.Bands("Band_Export").Tools("ExportHtml").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PDF) Then BarMenu.Bands("Band_Export").Tools("ExportPDF").Enabled = CheckRights("Stampa", Enable)
    
    If (Buttons And BTN_SEARCHFORM) Then
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Checked = Not Enable
    End If
    
    If (Buttons And BTN_SEARCHTABLE) Then
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Enabled = CheckRights("Selezione", Enable)
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Checked = Not Enable
    End If
    
    If (Buttons And BTN_FILTER) Then BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Enabled = CheckRights("Selezione", Enable)
    
    'VOCI DI MENU
    '------------
    
    'Menu File
    '---------
    If (Buttons And BTN_NEW) Then BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_SAVE) Then BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PRINT) Then BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PREVIEW) Then BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Enabled = CheckRights("Stampa", Enable)
    
    'Menu Edit
    '---------
    If (Buttons And BTN_CUT) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_COPY) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PASTE) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_DELETE) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled = CheckRights("Cancellazione", Enable)
    If (Buttons And BTN_CLEAR) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_FIND) Then BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCH) Then BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_PREVIOUS) Then BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_NEXT) Then BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled = CheckRights("Selezione", Enable)
    
    'Menu Visualizza
    '---------------
    If (Buttons And BTN_FILTER) Then BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCHFORM) Then BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCHTABLE) Then BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_VIEWMODE) Then
        BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
    End If

    'Menu Export
    '-----------
    If Not oExportActivity Is Nothing Then
        If (Buttons And BTN_EXPORT) Then oExportActivity.EnableItems CheckRights("Stampa", Enable)
    End If
    If (Buttons And BTN_EXPORT) Then BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_WORD) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_EXCEL) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_HTML) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PDF) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Enabled = CheckRights("Stampa", Enable)
End Sub

'**+
'Nome: NewSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni da compiere in caso di richiesta di una nuova ricerca
'**/
Private Sub NewSearch()

    'Refresh dello stato del Form
    m_Changed = False
    m_Saved = False
    m_Search = True
    
    'Annulla una eventuale operazione di inserimento di un nuovo record
    If m_Document.TableNew Then
        m_Document.AbortNew
        RefreshFormFields
    End If
    
    'Ripristina la vista del Form
    BrwMain.Visible = True
    
    'Predispone la modalità DefineFilter della Browse
    BrwMain.AbortFilterEdit = False
    BrwMain.GuiMode = dgFilterDefinition
    BrwMain.SetFocus
    
    'Refresh dello stato dei bottoni delle barre dei menu per la modalità ricerca
    SetStatus4Modality Find
    
End Sub

'**+
'Nome: ExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue la ricerca impostata.
'
'**/
Private Sub ExecuteSearch()
    Dim Cond As dmtgridctl.dgCondition
    Dim Field As DmtDocManLib.Field
    Dim OLDCursor As Integer
    Dim sWhere As String
    
    
    'Gestione della clessidra
    OLDCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    
    'Se non è stato selezionato nessun filtro dal controllo DocTypeExplorer
    'viene creato un filtro temporaneo in memoria e reso il filtro attivo
    If Not m_FilterSelected Then
        
        'Comunica all'oggetto DocType i valori da usare per la ricerca
        sWhere = fnFillDocTypeCondition
        
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"
        
        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")
        
        'Aggiunge al filtro eventuali condizioni aggiuntive restituite dalla funzione fnFillDocTypeCondition
        If sWhere <> "" Then m_ActiveFilter.AddCondition sWhere
        
    End If
    
    'Comunica al documento il nuovo filtro da usare
    Set m_Document.ActiveFilter = m_ActiveFilter
    
    'Viene effettuata la ricerca
    m_Document.OpenDoc
    
    
    'Assegnazione del riferimento alla fonte dati (binding sul recordset del documento)
    
    'rif13
    
    'Set BrwMain.Recordset = m_Document.Dataset.Recordset
    Set BrwMain.Recordset = m_Document.Data
    
    
    
    'Ripristina il cursore
    Screen.MousePointer = OLDCursor
    
    'Operazioni da effettuare dopo l'esecuzione della ricerca.
    AfterExecuteSearch
End Sub

'**+
'Nome: AfterExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Determina quali operazioni compiere dopo ExecuteSearch
'              in funzione dell'esito della ricerca.
'
'**/
Private Sub AfterExecuteSearch()

    If Not (m_Document.EOF = True And m_Document.BOF = True) Then
        'La ricerca ha avuto esito positivo
        'Attiva la vista tabellare
        BrwMain.Visible = True
        BrwMain.SetFocus

        'Imposta i menu e la toolbar per la modalità tabellare
        SetStatus4Modality Browse

        'Attiva le procedure di creazione di un nuovo filtro solo se l'ExecuteSearch
        'non è stata chiamata da una selezione del DocTypeExplorer

        'Se l'ExecuteSearch non è stata chiamata da un filtro del riquadro attività
        'si permette di salvare il nuovo filtro ed aggiungerlo nel ramo dei filtri.
        If Not m_FilterSelected Then
            oFiltersActivity.NewFilterBegin
        End If

        'Imposta i suggerimenti da visualizzare sulla Statusbar in funzione
        'della modalità di visualizzazione corrente.
        'Ad esempio in alcuni casi le frasi sono al Singolare/Plurare.
        'Le impostazioni sottostanti servono soltanto all'avvio del programma dopo la prima
        'ricerca. (in quanto ChangeView non è stata ancora eseguita)
        'La Sub RefreshDescriptions4StatusBar deve essere chiamata anche in ChangeView()--> Vedi.
        RefreshDescriptions4StatusBar

        m_Search = False
    Else
        'La ricerca ha avuto esito negativo. Viene mostrato un messaggio
        'e si torna in modalità ricerca.

        'Per questioni estetiche viene subito mostrata la modalità FilterDefinition
        'al posto della browse vuota e quindi viene mostrato il messaggio.
        BrwMain.GuiMode = dgFilterDefinition

        'Se si è selezionato il filtro "Nessun record" non occorre
        'visualizzare il messaggio
        If m_ActiveFilter.NothingSelected = False Or m_FilterSelected = False Then
            'Messaggio  "Nessun elemento trovato"
            sbMsgInfo gResource.GetMessage(MESS_NORECFOUND), m_App.FunctionName
        End If

        'Si torna in modalità form (modalità ricerca)
        OnNewSearch
    End If
    
End Sub



'**+
'Autore: Diamante s.p.a
'Data creazione: 26/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: fnFillDocTypeCondition
'
'Parametri:
'
'Valori di ritorno: String - in base alle esigenze specifiche di una manutenzione è possibile montare ad hoc
'                                 una clausola WHERE che potrà poi essere presa in considerazione nel filtro di selezione
'                                 con il metodo AddCondition dell'oggetto DmtDocManLib.Filter
'
'Funzionalità: Comunica all'oggetto DocType i valori da usare per la ricerca
'
'**/
Private Function fnFillDocTypeCondition() As String
    Dim Field As DmtDocManLib.Field
    Dim Cond As dmtgridctl.dgCondition
    Dim sWhere As String
    
    
    'NOTA per l'uso dei campi RANGE
    '--------------------------------------------------------------------------------------------------
    'E' consentito l'inserimento, nella modalità filtri e nel caso di campi di tipo range, del solo il valore iniziale
    '(in questo caso vengono filtrati tutti gli elementi maggiori o uguali a quello inserito)
    'o solo quello finale (in questo vengono filtrati tutti gli elementi minori o uguali a quello inserito).
    'Questo funzionamento vale per tutte le tipologie di campo.
    
    'Nel caso di condizione RANGE la sintassi da usare è del tipo della riga sotto:
    'm_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
    '--------------------------------------------------------------------------------------------------
    
    sWhere = vbNullString
    
    'Ripulisce la collezione Fields dell'oggetto DocType
    For Each Field In m_DocType.Fields
        Field.Value = Empty
    Next
    
    m_DocType.Fields("IDAzienda").Value = TheApp.IDFirm
    
    For Each Cond In BrwMain.Conditions
        If Cond.IsHeader = False Then
              Select Case Cond.ConditionType
                  
                  'Condizione boolean
                  Case dgCondTypeBoolean
                      m_DocType.Fields(Cond.FieldName).Value = IIf(IsEmpty(Cond.FromValue), Empty, Abs(CDbl(Cond.FromValue = "SI")))
                      
                  'Condizione associata ad una combo box
                  Case dgCondTypeComboDB
                      m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValueID
                  
                  'Condizione di tipo text, numeric, data, time
                  Case dgCondTypeText
                      If Cond.RangeChecked = True Then
                          m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                      Else
                          m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      End If
                  Case dgCondTypeNumber
                      If Cond.RangeChecked = True Then
                          m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                      Else
                          m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      End If
                  
                  Case dgCondTypeDate
                      If Cond.RangeChecked = True Then
                          m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                          
                          'sWhere = Cond.FieldName & ">=" & fnNormDate(Cond.FromValue)
                          'sWhere = sWhere & " AND " & Cond.FieldName & "<=" & fnNormDate(Cond.ToValue)
                      Else
                          m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      End If
                  
                  Case dgCondTypeTime
                      If Cond.RangeChecked = True Then
                          sWhere = Cond.FieldName & ">=" & fnNormTime(Cond.FromValue)
                          sWhere = sWhere & " AND " & Cond.FieldName & "<=" & fnNormTime(Cond.ToValue)
                      Else
                          m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      End If
                
                  'Altre condizioni
                  Case Else
                      m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      
        
              End Select
        End If
    Next Cond
        
    fnFillDocTypeCondition = sWhere
End Function



'**+
'Nome: CheckRights
'
'Parametri:
'ActionName - Nome della azione
'Enable - Valore da modificare o ritornare inalterato
'
'Valori di ritorno:
'Il valore in Enable o False se l'azione non è abilitata
'per il tipo di documento
'
'Funzionalità:
'Controlla se l'azione passata è abilitata per il tipo documento
'**/
Private Function CheckRights(ByVal ActionName As String, ByVal Enable As Boolean) As Boolean
    Dim Action As DmtDocManLib.Action
    Dim Dummy As String
    
    If m_DocType.Actions.Count = 0 Then
        CheckRights = Enable
        Exit Function
    End If
    For Each Action In m_DocType.Actions
        If Action.Name = "TUTTE LE AZIONI" Then
            CheckRights = Enable
            Exit Function
        End If
    Next
    On Error GoTo ActionNotFound
    Dummy = m_DocType.Actions(ActionName).Name
    CheckRights = Enable
    Exit Function
ActionNotFound:
    CheckRights = False
End Function

'**+
'Nome: SetFocusTabIndex0
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Da l'input focus al campo con TabIndex uguale a 0.
'**/
Private Sub SetFocusTabIndex0()
    On Error GoTo SetFocusTabIndex0_Error
    
    Dim ControlObject As Control
    Dim iIndex As Long
    Dim bError As Boolean
    
    If m_ControlTabIndex0 Is Nothing Then
        For Each ControlObject In frmMain.Controls
            iIndex = ControlObject.TabIndex
            If bError Then
                '**+ Controllo corrente non ha proprietà TabIndex,
                '    quindi va saltato.
                bError = False
            Else
                If ControlObject.TabIndex = 0 Then
                    Set m_ControlTabIndex0 = ControlObject
                    Exit For
                End If
            End If
        Next
    End If
    m_ControlTabIndex0.SetFocus

    Exit Sub
SetFocusTabIndex0_Error:
    bError = True
    Resume Next
End Sub

'**+
'Nome: IsFieldInput
'
'Parametri:
'Control - un oggetto Control da controllare
'
'Valori di ritorno:
'Se il controllo è abilitato all'input torna vero altrimenti falso
'
'Funzionalità:
'Controllo se un certo controllo è usabile come campo
'di input dei dati del Form
'**/
Private Function IsFieldInput(ByVal Control As Control) As Boolean
    'Controlla se il Controllo è di Immissione
    IsFieldInput = IsFieldInput Or TypeName(Control) = "TextBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "CheckBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "ComboBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "OptionButton"
    
    'rif5 begin
    
    IsFieldInput = IsFieldInput Or TypeName(Control) = "DMTCombo"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "Town"
    
    'rif5 end
    
    
End Function

'**+
'Nome: FieldPresent
'
'Parametri:
'Name - Nome del campo
'
'Valori di ritorno:
'Se il campo specificato nel parametro è presente nella
'collezione FormFields torna vero altrimenti torna falso
'
'Funzionalità:
'Controlla la presenza di un campo nella collezione FormFields
'**/
Private Function FieldPresent(ByVal Name As String) As Boolean
    Dim Field As FormField

    For Each Field In m_FormFields
        FieldPresent = (Name = Field.Name)
        If FieldPresent Then Exit For
    Next
End Function

'**+
'Nome: OnStart
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Altre inizializzazioni dopo quelle predefinite
'**/
Private Sub OnStart()
On Error GoTo ERR_OnStart
Dim sSQL As String
  
'SETTARE LE GRIGLIE DEI SOTTODOCUMENTI

    'Inizializzazione della griglia adibita alla visualizzazione tabellare dei sotto-documenti
    '-------------------------------------------------------------------------------
    'Articolo
    Set Me.ACSCliente.Connection = TheApp.Database.Connection
    ACSCliente.ApplicationName = App.Title
    ACSCliente.Client = App.EXEName
    ACSCliente.IDFirm = TheApp.IDFirm
    ACSCliente.IDUser = TheApp.IDUser
    ACSCliente.UserName = TheApp.User
    ACSCliente.SearchType = DmtSearchCustomers
    ACSCliente.HwndContainer = Me.hwnd
    
    'Inizializza il controllo Codice-Descrizione per la ricerca dei clienti
    With cdAnagrafica
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeCaption4Find = "Cognome / Ragione sociale"
        .CodeField = "Anagrafica"
        .CodeIsNumeric = False

        .DescriptionCaption4Find = "Nome"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .IDExecuteFunction = 29 'Anagrafica
    End With
    
    With Me.CDArticolo
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione "
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    'Imballo
    With Me.CDImballo
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto = " & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    'Cooperativa
    With Me.CDSocioFatt
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "DenominazioneCompleta"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaCooperativaDaLibroSoci"
        .Filter = "IDAzienda = " & m_App.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Denominazione Completa"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Denominazione completa"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Anagrafiche") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    'Socio
    With Me.CDSocio
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "DenominazioneCompleta"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaSocio"
        .Filter = "IDAzienda = " & m_App.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Denominazione Completa"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Denominazione completa"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Anagrafiche") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    'Sezionale
    With Me.CDSezionale
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Prefisso"
        .DescriptionField = "Sezionale"
        .KeyField = "IDSezionale"
        .TableName = "RV_POIESezionalePerTipoDocumento"
        .Filter = "IDFiliale = " & m_App.Branch & " AND IDTipoOggetto=2"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Prefisso"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Sezionale"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Prefisso"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Sezionale"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafiche") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    With Me.cboIvaCliente
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT IDIva, Iva FROM Iva"
        .SQL = .SQL & " ORDER BY Codice"
    End With

    With Me.cboIvaArticolo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT IDIva, Iva FROM Iva"
        .SQL = .SQL & " ORDER BY AliquotaIva"
    End With
    
    'Regione
    With Me.cboRegione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRegione"
        .DisplayField = "Regione"
        .SQL = "SELECT * FROM Regione ORDER BY Regione"
        .Fill
    End With
    

    With Me.cboVettore
        Set .Database = m_App.Database.Connection
        .DisplayField = "Vettore"
        .AddFieldKey "IDVettore"
        .SQL = "SELECT IDVettore, Vettore"
        .SQL = .SQL & " FROM Vettore"
        .SQL = .SQL & " ORDER BY Vettore"
        .Refresh
    End With
    
    'Famiglia prodotti
    With Me.cboFamigliaLotto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_FamigliaProdotti"
        .DisplayField = "FamigliaProdotti"
        .SQL = "SELECT * FROM RV_PO01_FamigliaProdotti ORDER BY FamigliaProdotti"
        .Fill
    End With
    
    'Varieta lotto
    With Me.cboVarietaLotto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_Varieta"
        .DisplayField = "Varieta"
        .SQL = "SELECT * FROM RV_PO01_Varieta ORDER BY Varieta"
        .Fill
    End With
    
    'Varieta articolo
    With Me.cboVarietaArticolo
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PO01_Varieta"
        .DisplayField = "Varieta"
        .SQL = "SELECT * FROM RV_PO01_Varieta ORDER BY Varieta"
        .Fill
    End With
Exit Sub
ERR_OnStart:
    MsgBox Err.Description, vbCritical, "OnStart"
End Sub

'**+
'Nome: OnSave
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Save
'**/
Private Function OnSave() As Boolean
On Error GoTo ERR_OnSave
Dim Field As DmtDocManLib.Field
Dim DocLink As DmtDocManLib.DocumentsLink
Dim sSQL As String
Dim creadocumentoOK As Boolean

    
    OnSave = False
    'Controlli preliminari sulla validità e consistenza dei dati da salvare
    If Not PermissionToSave Then
        Exit Function
    End If
    
    'CONTROLLO E CREAZIONE COOPERATIVA COME CLIENTE
    CONTROLLO_COOP_COME_CLIENTE Me.txtIDAnagraficaCoop.Value
    
    
    'Passa alla collezione Fields dell'oggetto
    'Document i valori da salvare
    For Each Field In m_Document.Fields
        'Sul campo chiave primaria non si deve far nulla
        If Not Field.PrimaryKey Then
            If FieldPresent(Field.Name) Then
            
                'rif4 begin

                Select Case TypeName(m_FormFields(Field.Name).Control)
                    Case "TextBox"
                        Field.Value = m_FormFields(Field.Name).Control.Text
                    Case "DmtCodDesc"
                        Field.Value = m_FormFields(Field.Name).Control.KeyFieldID
                    Case "DMTCombo"
                        Field.Value = m_FormFields(Field.Name).Control.CurrentID
                    Case "Town"
                        If Field.Name = "IDComune" Then
                            Field.Value = m_FormFields(Field.Name).Control.CityID
                        ElseIf Field.Name = "Cap" Then
                            Field.Value = m_FormFields(Field.Name).Control.Zip
                        End If
                    Case "dmtDate"
                        If (m_FormFields(Field.Name).Control.Text = "") Or (IsNull(m_FormFields(Field.Name).Control.Value)) Then
                            Field.Value = Null
                        Else
                            Field.Value = m_FormFields(Field.Name).Control.Value
                        End If
                    Case "dmtNumber"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    Case "dmtCurrency"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    
                    Case "dmtTime"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    Case "DmtSearchACS2"
                        Field.Value = m_FormFields(Field.Name).Control.IDAnagrafica
                    Case "CheckBox"
                        Field.Value = fnNormBoolean(m_FormFields(Field.Name).Control.Value)
                    Case "Label"
                        Field.Value = m_FormFields(Field.Name).Control.Caption
                
                End Select
                
                'rif4 end
                
            Else
                'Se il processo in corso è "Manutenzione da Shell"
                'la variabile m_LinkedField contiene il nome del
                'campo collegato alla applicazione chiamante
                'quindi il campo relativo deve essere valorizzato
                'con il valore ricevuto dalla applicazione chiamante
                'nCBC+
                If Field.Name = m_LinkedField Then
                    Field.Value = m_App.CallerFieldValue
                End If
                If Field.Name = "IDAzienda" Then
                    Field.Value = TheApp.IDFirm
                End If
                            
            End If
        End If
    Next
    
    m_Document.SaveDocument
    
    SALVA_PAR_QUAL fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    creadocumentoOK = CREA_DOCUMENTO
    
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
    
    
    
    'Refresh delle variabili di stato
    m_Changed = False
    m_Search = False
    m_Saved = True
    
    'Refresh dello stato della ToolBar standard in modalità variazione
    SetStatus4Modality Modify
        
    OnSave = creadocumentoOK
Exit Function
ERR_OnSave:
    MsgBox Err.Description, vbCritical, "OnSave"
End Function

'**+
'Nome: OnSaveDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando OnSaveDocumentsLink
'**/
Private Sub OnSaveDocumentsLink(ByVal DocumentLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnDelete
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Delete
'**/
Private Sub OnDelete()
    Dim sToRemove As String
    Dim DocLink As DmtDocManLib.DocumentsLink
    Dim IDOggettoCollegato As Long
    
    
    'Se si è in modalità tabellare potrebbe essere necessario sincronizzare
    'il documento con il record evidenziato nella browse
    If BrwMain.Visible = True Then
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    End If
    
    
    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemDeleteAction) Then
        Exit Sub
    End If
    
    'Se in fase di inserimento di un nuovo
    'record non c'è niente da fare
    If m_Document.TableNew Then
        Exit Sub
    End If
    
    'Conferma della cancellazione
    gResource.CustomStrings.Clear
    sToRemove = m_Document.Fields(CAMPO_PER_CAPTION).Value
    gResource.CustomStrings.Add Chr(34) & sToRemove & Chr(34), 1
    If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYREMOVE), m_App.FunctionName) = vbYes Then
    
    
    
        If (LINK_DOCUMENTO_COLLEGATO > 0) Then
            If (CONTROLLO_STATO_DOCUMENTO) Then
                MsgBox "Documento di trasporto collegato risulta bloccato!", vbCritical, "Controllo dati"
                Exit Sub
            End If
        End If
        
        
        If Not (m_Document.EOF Or m_Document.BOF) Then
            'Cancella l'eventuale blocco sul record da cancellare.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        End If
        'rif16
        
        If (ELIMINA_DOCUMENTO = False) Then Exit Sub
            
   
        
        'Cancellazione
        m_Document.DeleteDocument
        
        

        
        
        If (m_Document.EOF = True And m_Document.BOF = True) Then
            'Se è stato cancellato l'ultimo record si va in modalità inserimento
            NewRecord
        Else
            'Refresh dello stato della ToolBar standard e dei menu
            If BrwMain.Visible Then
                'Va in modalità tabellare
                SetStatus4Modality Browse
            Else
                'Essendo in modalità variazione occorre controllare se il record su cui
                'ci si è posizionati è bloccato.
                'Se non lo è lo si blocca e si procede altrimenti si andrà in modalità tabellare.
                If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
                    'Il record è bloccato.
                    'Va in modalità tabellare
                    BrwMain.Visible = True
                    SetStatus4Modality Browse
                Else
                    'Il record non è bloccato.
                    
                    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
                    m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
                    
                    'Va in modalità variazione
                    SetStatus4Modality Modify
                End If
            
                 RefreshDescriptions4StatusBar
            End If
        End If
        
        'Refresh delle variabili di stato
        m_Changed = False
        m_Saved = True
        m_Search = False
        
    End If
End Sub

'**+
'Nome: OnDeleteDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni relative ai documents link sul comando Delete
'**/
Private Sub OnDeleteDocumentsLink(ByVal DocumentLink As DmtDocManLib.DocumentsLink)
End Sub

'**+
'Nome: OnClear
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Clear
'**/
Private Sub OnClear()
'Se si è in modalità Filtro occorre ripulire i campi di immissione altrimenti,
'se si è in modalità Form, si cancella il contenuto di tutti i controlli
    
    
    If BrwMain.Visible And BrwMain.GuiMode = dgFilterDefinition Then
        '---Modalità Filtro---
        'Ripulisce i campi di immissione delle condizioni di ricerca.
        BrwMain.Conditions.ClearValues
    Else
        '---Modalità Form---
        'Ripulisce i campi del form
        ClearFormFields
        SetFocusTabIndex0
        
        'Se si era in modalità Nuovo viene disabilitato il pulsante Salva
        'e si ripristina la modalità stessa.
        If m_Document.TableNew Then
            ActivateBarButtons BTN_SAVE, False
            m_Changed = False
            m_Saved = True
        End If
    End If
End Sub

'**+
'Nome: OnExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ExecuteSearch
'**/
Private Sub OnExecuteSearch()
    
    'Nota: utilizzo la chiamata al metodo ApplyFilter della dmtGrid piuttosto
    'che la chiamata diretta di ExecuteSearch perchè in questo modo la dmtGrid
    'può gestire internamente le conditions di ricerca.
    'Verrà generato l'evento BrwMain_OnApplyFilter()
    '
    'ExecuteSearch
    '
    BrwMain.ApplyFilter
    
        
End Sub

'**+
'Nome: OnMoveCurrentRecord
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando di riposizionamento del record corrente
'**/
Private Sub OnMoveCurrentRecord(ByVal tipo As Integer, ByVal sToolName As String)
    Dim iResponse As Integer
    
    iResponse = ChooseAboutSaving
    If iResponse = vbYes Then
        OnSave
        'Se la registrazione non è andata a buon fine esce
        If Not m_Saved Then
            Exit Sub
        End If
    End If
    If iResponse <> vbCancel Then
       Select Case tipo
           Case SRCNEXT
               SearchNext
           Case SRCPREVIOUS
               SearchPrevious
       End Select
       m_Changed = False
       ActivateBarButtons BTN_SAVE, False
    End If
End Sub

'**+
'Nome: OnRepositionDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando di riposizionamento del record corrente
'per i DocumentsLink
'**/
Private Sub OnRepositionDocumentsLink(ByVal DocumentsLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnChangeView
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ChangeView
'**/
Private Sub OnChangeView(ByVal sToolName As String)
    Dim iResponse   As Integer
    
    If Not BrwMain.Visible And m_Changed Then
        iResponse = ChooseAboutSaving
        
        If iResponse = vbYes Then
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
        End If
        
        If iResponse <> vbCancel Then
            'cbc 20/04/1999
            'se si è scelto NO ripulisce i campi e va in modalità tabellare annullando
            'le ultime modifiche
            RefreshFormFields
            ChangeView sToolName
            m_Changed = False
        End If
    Else
        ChangeView sToolName
    End If
    
End Sub

'**+
'Nome: OnToolBarOptions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ToolBar
'
'
'**/
Private Sub OnToolBarOptions()
    Dim dlgToolBars As frmToolBars
    Dim bVisible As Boolean
    
    On Error Resume Next
    
    Set dlgToolBars = New frmToolBars
    'Imposta un riferimento al form chiamante
    Set dlgToolBars.FormClient = Me
    dlgToolBars.Show vbModal, Me
    Set dlgToolBars = Nothing
    
    'All'uscita dal form di dialogo la visibilità della toolbar dei filtri dipende dalla
    'visibilità del Riquadro attività e dall'impostazione fatta nel dialogo.
    bVisible = GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "Riquadro attività", True)
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: OnOptions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Option
'**/
Private Sub OnOptions()
    Dim dlgOption As frmOption
    
    Set dlgOption = New frmOption
    Set dlgOption.FormClient = Me
    dlgOption.Show vbModal, Me
    
    
    Set dlgOption = Nothing
    
    'Impedisce il 'blocco' della Toolbar alla chiusura di un form di dialogo.
    
End Sub

'**+
'Nome: OnInfo
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Info
'**/
Private Sub OnInfo()
    Dim dlgInfo As frmInformazioni
    
    Set dlgInfo = New frmInformazioni
    dlgInfo.Show vbModal, Me
    
    'Impedisce il 'blocco' della Toolbar alla chiusura di un form di dialogo.
    
End Sub

'**+
'Nome: OnPrint
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Print
'**/
Private Sub OnPrint(ByVal ToolName As String)
    Dim lFlags As Long
    Dim OLDCursor As Integer
    Dim sStr As String
    Dim Field As DmtDocManLib.Field
    
    
    OLDCursor = Screen.MousePointer
    
    'Se il filtro attivo è "Nessun record" è possibile eseguire una stampa/esportazione soltanto se
    'si è in modalità form. In tal caso, infatti, verrà passato al Crystals Reports un filtro
    'creato ad hoc sull'ID del record attuale.
    If m_ActiveFilter.NothingSelected And BrwMain.Visible Then
        sStr = "Impossibile effettuare l'operazione richiesta." & vbCrLf
        sStr = sStr & "Prima di procedere occorre eseguire un filtro."
        sbMsgInfo sStr, m_App.FunctionName
        Screen.MousePointer = OLDCursor
        Exit Sub
    End If
    
    'Se non esiste un report attivo occorre annullare l'operazione.
    If Len(oReportsActivity.SelectedReportName) > 0 Then
        Set m_Report = m_DocType.Reports.Item(oReportsActivity.SelectedReportName)
    End If
    If m_Report Is Nothing Then
        sbMsgError "Impossibile eseguire - Nessun report predefinito.", m_App.FunctionName
        GoTo OnPrint_Exit
    End If
    m_iNumeroCopieDefault = m_Report.Copies
    m_OrientamentoDefault = m_Report.Orientation
    
    
    'Se è attivo il pulsante Salva deve essere visualizzato un messaggio di avviso
    'con i pulsanti OK e annulla (occorre salvare PRIMA della stampa)
    'Se è attivo il pulsante Salva deve essere visualizzato un messaggio di avviso
    'con i pulsanti OK e annulla (occorre salvare PRIMA della stampa)
    If m_Changed Then
        Select Case ChooseAboutSavingOkCancel
            Case vbOK
                OnSave
                'Se la registrazione non è andata a buon fine esce
                If Not m_Saved Then
                    GoTo OnPrint_Exit
                End If
                
            Case vbCancel
                GoTo OnPrint_Exit
        End Select
    End If
    
    If Not BrwMain.Visible Then
        'Modalità Form - deve stampare solo il record corrente
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        For Each Field In m_DocType.Fields
            Field.Value = Empty
        Next
        
        'Viene inserita la condizione di ricerca basata sull'ID del record corrente.
        m_DocType.Fields("ID" & m_App.TableName).Value = m_Document.Fields("ID" & m_App.TableName).Value
        
        'Viene creato un filtro temporaneo per il Crystals Reports.
        m_DocType.RemoveFilter "Form"
        Set m_Report.Filter = m_DocType.AddFilterWithConditions("Form")
    Else
        'Modalità vista tabellare
        
        'Viene passato il filtro corrente al Crystals Reports.
        Set m_Report.Filter = m_ActiveFilter
    End If
            
    
    Select Case ToolName
    
        Case "PrePrint", "Mnu_PrePrint"
            On Error GoTo ErrorHandler
            
            Screen.MousePointer = vbHourglass
            
            m_TabMode = BrwMain.Visible
            PicForm.Visible = False
            BrwMain.Visible = False
            ActivityBox.Visible = False
            
            SetStatus4Modality Preview, OpenPrw
            Refresh
            
            m_PreviewWindowHandle = m_Document.Preview(m_Report, "", hwnd, CInt(BarMenu.ClientAreaLeft / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaTop / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaWidth / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaHeight / Screen.TwipsPerPixelY), False)
            lFlags = SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOMOVE
            SetWindowPos m_PreviewWindowHandle, HWND_TOP, 0, 0, 0, 0, lFlags
            
        Case "Print", "Mnu_Print"
            PrintDocument ToolName
            
        Case "ExportWord", "Mnu_ExportWord"
            ExportDocument ecWord
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Word, TheApp.Name

        Case "ExportExcel", "Mnu_ExportExcel"
            ExportDocument ecExcel
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Excel, TheApp.Name
            
        Case "ExportHtml", "Mnu_ExportHtml"
            ExportDocument ecHtml
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), HTML, TheApp.Name
        
        Case "ExportPDF", "Mnu_ExportPDF"
            ExportDocument ecPdf
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), PDF, TheApp.Name
        
        Case "MailWord"
            SendDocument ecWord
            
        Case "MailExcel"
            SendDocument ecExcel
            
        Case "MailHtml"
            SendDocument ecHtml
        
        Case "MailPDF"
            SendDocument ecPdf
    End Select
    
   
OnPrint_Exit:
    Set Field = Nothing
    Screen.MousePointer = OLDCursor
    Exit Sub
    
ErrorHandler:
    Const ERROR_PRINTING_ABORTED = 3
    Const ERROR_PRINTING_CANCELLED = 4
    Select Case Err.Number
        Case 20507
            'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
            sbMsgInfo "File di report non trovato", m_App.FunctionName
        Case ERROR_PRINTING_ABORTED, ERROR_PRINTING_CANCELLED
            'non deve far niente, è stato già segnalato da CrystalReport
        Case Else
            If Len(Trim(Err.Description)) > 0 Then
                sbMsgInfo Err.Description, m_App.FunctionName
            End If
    End Select

    'Si è verificato un errore durante la procedura di anteprima.
    Screen.MousePointer = OLDCursor
    
    'Ripristina la situazione del form
    m_PreviewWindowHandle = 0
    PicForm.Visible = True

    BrwMain.Visible = m_TabMode
    ActivityBox.Visible = BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked
    FormRecalcLayout
    SetStatus4Modality Preview, ClosePrw
        
    Set Field = Nothing
End Sub

'**+
'Nome: OnNewSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando NewSearch
'**/
Private Sub OnNewSearch()
    Dim iResponse As Integer

    m_FilterSelected = False
    
    If Not m_Changed Then
        NewSearch
    Else
        'cbc 20/04/1999
        'deve mostrare il messaggio con Si, No, Annulla
        iResponse = ChooseAboutSaving
        If iResponse = vbYes Then
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
        End If
        If iResponse <> vbCancel Then
            'se si è scelto NO ripristina i dati precedenti annullando le ultime modifiche
            'e predispone la modalità ricerca.
            RefreshFormFields
            NewSearch
            m_Changed = False
        End If
    
    End If
    
    
    
End Sub

'**+
'Nome: OnNew
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando New
'**/
Private Sub OnNew(ByVal sToolName As String)
    
    Select Case DoNewDocument
        Case vbYes
            'Si è risposto affermativamente alla
            'richiesta di Update delle modifiche apportate
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
            NewRecord
            
        Case vbCancel
            'Si è risposto Annulla alla richiesta di Update
            Exit Sub
            
        Case Else
            'Si è premuto il tasto <No> alla richiesta di Update
            NewRecord
    End Select
End Sub

'**+
'Nome: OnNewDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando New per i documentslink
'**/
Private Sub OnNewDocumentsLink(ByVal DocumentsLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnSummary
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Summary
'**/
Private Sub OnSummary()
    Dim lRes As Long
    
    lRes = WinHelp(hwnd, App.HelpFile, HELP_FINDER, 0)
End Sub

'**+
'Nome: OnFastHelp
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando FastHelp
'**/
Private Sub OnFastHelp()
    frmMain.WhatsThisMode
End Sub

'**+
'Nome: OnHelpOnLine
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando HelpOnLine
'**/
Private Sub OnHelpOnLine()
    Dim lRes As Long
    
    If Not ActiveControl Is Nothing Then
        If ActiveControl.HelpContextID <> 0 Then
            lRes = WinHelp(hwnd, App.HelpFile, HELP_CONTEXT, ActiveControl.HelpContextID)
        Else
            ExecuteMenuCommand "Mnu_Arg"
        End If
    End If
End Sub

'**+
'Nome: OnArg
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Arg
'**/
Private Sub OnArg()
    Dim lRes As Long
    
    If m_App.ContextHelpID <> 0 Then
        lRes = WinHelp(hwnd, App.HelpFile, HELP_CONTEXT, m_App.ContextHelpID)
    Else
        ExecuteMenuCommand "Mnu_Summary"
    End If
End Sub

'**+
'Nome: OnViewAssistant
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ViewAssistant
'**/
Private Sub OnViewAssistant()
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked = Not BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked
    If BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked Then
   Else
    End If
End Sub

'**/
'Autore                 : Diamante S.p.a
'
'Nome                   : OnFolders
'
'Parametri:
'
'
'Valori di ritorno:
'
'Funzionalità:
'Permette la visualizzazione o meno del DocTypeExplorer e della relativa toolbar.
'**/
Private Sub OnFolders()
    ActivityBox.Visible = Not ActivityBox.Visible
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked = ActivityBox.Visible
    FormRecalcLayout
    
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: OnRunApplication
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando RunApplication
'**/
Private Sub OnRunApplication(ByVal sToolName As String)
End Sub



Private Sub ACSCliente_ChangedElement()
    If Me.ACSCliente.IDAnagrafica > 0 Then
        If Me.ACSCliente.IDAnagrafica <> Me.cdAnagrafica.KeyFieldID Then
            Me.cdAnagrafica.Load Me.ACSCliente.IDAnagrafica
        End If
    End If
End Sub

'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : ActivityBox_CloseButtonPressed
'
'Parametri                  :
'
'Funzionalità               : Gestione della chiusura del Riquadro attività
'
'**/
Private Sub ActivityBox_CloseButtonPressed()
    OnFolders
End Sub

'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : ActivityBox_ItemSelected
'
'Parametri                  :
'
'Funzionalità               : Gestione della selezione delle voci del Riquadro attività
'
'**/
Private Sub ActivityBox_ItemSelected(ByVal Item As DmtActBoxTlb.Item, NeedRedraw As Boolean)
    Dim oFilter As Filter
    Dim oTableView As TableView
    
    If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
        Select Case ActivityBox.CurrentActivity.Caption
            Case "Filtri"
                For Each oFilter In m_DocType.Filters
                    If oFilter.ID = Val(Item.Tag) Then
                        Set m_ActiveFilter = m_DocType.Filters(oFilter.Name)
                        Exit For
                    End If
                Next
                'Flag usato per specificare che deve essere eseguito un filtro permanente.
                m_FilterSelected = True
                
                '---Modalità Filtro---
                'Ripulisce i campi di immissione delle condizioni di ricerca.
                BrwMain.Conditions.ClearValues
                
                'Se attivo, viene disabilitato il pulsante Salva Filtro del DocTypeExplorer.
                oFiltersActivity.AbortNewFilter
                ActivityBox.Redraw = True
                
                'Viene eseguita la ricerca basata sul nuovo filtro.
                ExecuteSearch

            Case "Viste tabellari"
                For Each oTableView In m_DocType.TableViews
                    If oTableView.ID = Val(Item.Tag) Then
                        Set m_ActiveTableView = m_DocType.TableViews(oTableView.Name)
                        Exit For
                    End If
                Next
                BrwMain.LoadColumns m_ActiveTableView
                SetVisibilityIDFields
        End Select
    End If
    If ActivityBox.CurrentActivity.Caption = "Esportazioni" Then
        If Item.Hyperlink Then
            Select Case Item.Name
                Case "E" & ExportConstants.PDF
                    ExecuteMenuCommand "ExportPDF"
                Case "E" & ExportConstants.Word
                    ExecuteMenuCommand "ExportWord"
                Case "E" & ExportConstants.Excel
                    ExecuteMenuCommand "ExportExcel"
                Case "E" & ExportConstants.HTML
                    ExecuteMenuCommand "ExportHtml"
                Case "S" & ExportConstants.PDF
                    ExecuteMenuCommand "MailPDF"
                Case "S" & ExportConstants.Word
                    ExecuteMenuCommand "MailWord"
                Case "S" & ExportConstants.Excel
                    ExecuteMenuCommand "MailExcel"
                Case "S" & ExportConstants.HTML
                    ExecuteMenuCommand "MailHtml"
            End Select
        End If
    End If
End Sub





Private Sub cboAltroSito_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboIvaArticolo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboIvaCliente_Click()
    
If bloading Then Exit Sub

If Me.cboIvaCliente.CurrentID > 0 Then
    Me.cboIvaArticolo.WriteOn Me.cboIvaCliente.CurrentID
Else
    Me.cboIvaArticolo.WriteOn GET_LINK_IVA_ARTICOLO(Me.CDArticolo.KeyFieldID)
End If

If Not (BrwMain.Visible) Then Change
    
    
End Sub

Private Sub cboRegione_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboVettore_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cdAnagrafica_ChangeElement()
On Error GoTo ERR_cdAnagrafica_ChangeElement
Dim IDSezionaleCliente As Long
    
    
    If bloading = True Then Exit Sub
    
    AggiornaAltreDestinazioni
    
    Me.ACSCliente.IDAnagrafica = 0
    Me.ACSCliente.Description = ""
    Me.ACSCliente.Code = ""
    Me.ACSCliente.SecondDescription = ""

    Me.ACSCliente.sbLoadCFByIDAnagrafica 0, Me.cdAnagrafica.KeyFieldID
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        Me.txtIDContratto.Value = 0
        Me.txtIDContrattoRiga.Value = 0
        IDSezionaleCliente = GET_SEZ_PER_CLIENTE(Me.cdAnagrafica.KeyFieldID)
        Me.cboAltroSito.WriteOn GET_DESTINAZIONE_PER_CLIENTE(Me.cdAnagrafica.KeyFieldID)
        
        If (IDSezionaleCliente > 0) Then
            Me.CDSezionale.Load IDSezionaleCliente
        End If
        Me.txtIDLetteraIntento.Value = GET_LINK_LETTERA_INTENTO_PRED(Me.cdAnagrafica.KeyFieldID, 2, Date, TheApp.IDFirm)
        Me.cboIvaCliente.WriteOn GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento, 0)
        If (Me.cboIvaCliente.CurrentID = 0) Then
            Me.cboIvaCliente.WriteOn GET_LINK_IVA_CLIENTE(Me.cdAnagrafica.KeyFieldID)
        End If
        RECUPERO_PAR_QUAL Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm
        If (TIPO_SALVATAGGIO = 0) Then GET_CONTRATTO
    Else
        If (m_Document("IDAnagrafica").Value <> Me.cdAnagrafica.KeyFieldID) Then
            Me.txtIDContratto.Value = 0
            Me.txtIDContrattoRiga.Value = 0
            IDSezionaleCliente = GET_SEZ_PER_CLIENTE(Me.cdAnagrafica.KeyFieldID)
            Me.cboAltroSito.WriteOn GET_DESTINAZIONE_PER_CLIENTE(Me.cdAnagrafica.KeyFieldID)
            If (IDSezionaleCliente > 0) Then
                Me.CDSezionale.Load IDSezionaleCliente
            End If
            Me.txtIDLetteraIntento.Value = GET_LINK_LETTERA_INTENTO_PRED(Me.cdAnagrafica.KeyFieldID, 2, Date, TheApp.IDFirm)
            Me.cboIvaCliente.WriteOn GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento, 0)
            If (Me.cboIvaCliente.CurrentID = 0) Then
                Me.cboIvaCliente.WriteOn GET_LINK_IVA_CLIENTE(Me.cdAnagrafica.KeyFieldID)
            End If
            RECUPERO_PAR_QUAL Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm
            If (TIPO_SALVATAGGIO = 0) Then GET_CONTRATTO
        End If
    End If
    
If Not (BrwMain.Visible) Then Change
Exit Sub
ERR_cdAnagrafica_ChangeElement:
    MsgBox Err.Description, vbCritical, "cdAnagrafica_ChangeElement"
End Sub


Private Sub CDArticolo_ChangeElement()
On Error GoTo ERR_CDArticolo_ChangeElement
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim Testo As String
If bloading = True Then Exit Sub

sSQL = "SELECT IDIvaVendita, AliquotaIva, Articolo, IDUnitaDiMisuraVendita, RV_POIDImballoVendita, RV_PO01_IDVarieta "
sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
sSQL = sSQL & "WHERE IDArticolo = " & Me.CDArticolo.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN((m_Document(m_Document.PrimaryKey).Value <= 0)) Then
        Me.cboIvaArticolo.WriteOn Me.cboIvaCliente.CurrentID
        If (Me.cboIvaArticolo.CurrentID = 0) Then
            Me.cboIvaArticolo.WriteOn fnNotNullN(rs!IDIvaVendita)
        End If
    End If
    
    cboVarietaArticolo.WriteOn fnNotNullN(rs!RV_PO01_IDVarieta)
    
End If

rs.CloseResultset
Set rs = Nothing

If fnNotNullN((m_Document(m_Document.PrimaryKey).Value <= 0)) Then
    REFRESH_DESCR_ARTICOLO
    
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = "SELECT * FROM RV_POIEContrattoDettaglioSel "
        sSQL = sSQL & " WHERE IDOggetto=" & Me.txtIDContratto.Value
        sSQL = sSQL & " AND RV_POTipoRiga=1"
        sSQL = sSQL & " AND Link_art_articolo=" & Me.CDArticolo.KeyFieldID
            
        Set rs = Cn.OpenResultset(sSQL)
        If rs.EOF Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Nel contratto non è presente l'articolo selezionato!" & vbCrLf
            Testo = Testo & "Impossibile continuare!"
            MsgBox Testo, vbCritical, "Controllo dati"
            Me.CDArticolo.Load 0
            Me.txtIDContrattoRiga.Value = 0
        Else
            Me.txtIDContrattoRiga.Value = fnNotNullN(rs!IDValoriOggettoDettaglio)
        End If
    End If
End If

If Not (BrwMain.Visible) Then Change
Exit Sub

ERR_CDArticolo_ChangeElement:
    MsgBox Err.Description, vbCritical, "CDArticolo_ChangeElement"
End Sub

Private Sub REFRESH_DESCR_ARTICOLO()
    Me.txtDescrizioneArticolo.Text = "Cert. N. " & Me.txtNumeroCertificato.Text & " del " & Me.txtDataCertificato.Text
    Me.txtDescrizioneArticolo.Text = Me.txtDescrizioneArticolo.Text & " - DDT n. " & Me.txtNumeroDDT.Text & " del " & Me.txtDataDDT.Text
    If (Me.cboVarietaArticolo.CurrentID > 0) Then
        Me.txtDescrizioneArticolo.Text = Me.txtDescrizioneArticolo.Text & " - " & Me.cboVarietaArticolo.Text
    Else
        Me.txtDescrizioneArticolo.Text = Me.txtDescrizioneArticolo.Text & " - " & Me.CDArticolo.Description
    End If
End Sub


Private Sub CDImballo_ChangeElement()

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        Me.txtTaraUnitaria.Value = fnGetTaraImballo(Me.CDImballo.KeyFieldID)
    End If
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDSezionale_ChangeElement()
    If bloading = True Then Exit Sub
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDSocio_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Sezionale_Socio As Long
Dim IDLotto As Long

Me.txtIDSocio.Value = Me.CDSocio.KeyFieldID

If bloading = True Then Exit Sub
If fnNotNullN((m_Document(m_Document.PrimaryKey).Value > 0)) Then Exit Sub

If (Me.CDSocioFatt.KeyFieldID = 0) Then
    sSQL = "SELECT IDAnagraficaFatturazione FROM RV_PO01_ConfigurazioneSocio "
    sSQL = sSQL & "WHERE IDAnagrafica=" & Me.CDSocio.KeyFieldID
    
    Set rs = Cn.OpenResultset(sSQL)
    If Not rs.EOF Then
        Me.CDSocioFatt.Load fnNotNullN(rs!IDAnagraficaFatturazione)
    End If
    rs.CloseResultset
    Set rs = Nothing
End If

If fnNotNullN((m_Document(m_Document.PrimaryKey).Value <= 0)) Then
    If (ATTIVA_SEZIONALE_DA_SOCIO = 1) Then
        Link_Sezionale_Socio = 0
        
        sSQL = "SELECT * FROM RV_PO01_ConfigurazioneSocioSez "
        sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
        sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
        sSQL = sSQL & " AND IDTipoOggetto=" & 2
        sSQL = sSQL & " AND IDAnagrafica=" & Me.CDSocioFatt.KeyFieldID
        sSQL = sSQL & " AND IDTipoUtilizzoSezionale=2"
        Set rs = Cn.OpenResultset(sSQL)
        
        If Not rs.EOF Then
            Link_Sezionale_Socio = fnNotNullN(rs!IDSezionale)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        
        If (Link_Sezionale_Socio = 0) Then
            sSQL = "SELECT * FROM RV_PO01_ConfigurazioneSocioSez "
            sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
            sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
            sSQL = sSQL & " AND IDTipoOggetto=" & 2
            sSQL = sSQL & " AND IDAnagrafica=" & Me.CDSocioFatt.KeyFieldID
            sSQL = sSQL & " AND ((IDTipoUtilizzoSezionale=0 OR IDTipoUtilizzoSezionale IS NULL))"
            Set rs = Cn.OpenResultset(sSQL)
        
            If Not rs.EOF Then
                Link_Sezionale_Socio = fnNotNullN(rs!IDSezionale)
            End If
            
            rs.CloseResultset
            Set rs = Nothing
        End If
        
        If Link_Sezionale_Socio > 0 Then
            Me.CDSezionale.Load Link_Sezionale_Socio
            Me.txtDataDocumento.Value = Date
        End If
    Else
        If (Me.CDSocioFatt.KeyFieldID = 0) Then
            Link_Sezionale_Socio = 0
            
            sSQL = "SELECT * FROM RV_PO01_ConfigurazioneSocioSez "
            sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
            sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
            sSQL = sSQL & " AND IDTipoOggetto=" & 2
            sSQL = sSQL & " AND IDAnagrafica=" & Me.CDSocioFatt.KeyFieldID
            sSQL = sSQL & " AND IDTipoUtilizzoSezionale=2"
            Set rs = Cn.OpenResultset(sSQL)
            
            If Not rs.EOF Then
                Link_Sezionale_Socio = fnNotNullN(rs!IDSezionale)
                
            End If
            
            rs.CloseResultset
            Set rs = Nothing
            
            If (Link_Sezionale_Socio = 0) Then
                sSQL = "SELECT * FROM RV_PO01_ConfigurazioneSocioSez "
                sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
                sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
                sSQL = sSQL & " AND IDTipoOggetto=" & 2
                sSQL = sSQL & " AND IDAnagrafica=" & Me.CDSocioFatt.KeyFieldID
                sSQL = sSQL & " AND ((IDTipoUtilizzoSezionale=0 OR IDTipoUtilizzoSezionale IS NULL))"
                Set rs = Cn.OpenResultset(sSQL)
            
                If Not rs.EOF Then
                    Link_Sezionale_Socio = fnNotNullN(rs!IDSezionale)
                End If
                
                rs.CloseResultset
                Set rs = Nothing
            End If
            
            If Link_Sezionale_Socio > 0 Then
                Me.CDSezionale.Load Link_Sezionale_Socio
                Me.txtDataDocumento.Value = Date
            End If
        End If
    End If
    
    sSQL = "SELECT IDAnagrafica, IDCategoriaAnagrafica "
    sSQL = sSQL & " FROM Anagrafica "
    sSQL = sSQL & " WHERE IDAnagrafica=" & Me.CDSocio.KeyFieldID

    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        If (IDCategoriaAnagraficaProdAcq > 0) Then
            If fnNotNullN(rs!IDCategoriaAnagrafica) = IDCategoriaAnagraficaProdAcq Then
                Me.Check1.Value = vbChecked
            End If
        End If
    End If

    rs.CloseResultset
    Set rs = Nothing
    
    IDLotto = GET_LOTTO_PROD_SINGOLO
    If (IDLotto > 0) Then
        Me.txtIDLottoCampagna.Value = IDLotto
    End If
    
    Me.txtNumeroCertificato.SetFocus
End If
If Not (BrwMain.Visible) Then Change
End Sub


Private Sub CDSocioFatt_ChangeElement()
On Error GoTo ERR_CDSocioFatt_ChangeElement
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Sezionale_Socio As Long
Dim DataControlloRevoca As String

Me.txtIDAnagraficaCoop.Value = Me.CDSocioFatt.KeyFieldID

If bloading = True Then Exit Sub
If fnNotNullN((m_Document(m_Document.PrimaryKey).Value <= 0)) Then
    Link_Sezionale_Socio = 0
    
    sSQL = "SELECT * FROM RV_PO01_ConfigurazioneSocioSez "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND IDTipoOggetto=" & 2
    sSQL = sSQL & " AND IDAnagrafica=" & Me.CDSocioFatt.KeyFieldID
    sSQL = sSQL & " AND IDTipoUtilizzoSezionale=2"
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        Link_Sezionale_Socio = fnNotNullN(rs!IDSezionale)
        
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If (Link_Sezionale_Socio = 0) Then
        sSQL = "SELECT * FROM RV_PO01_ConfigurazioneSocioSez "
        sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
        sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
        sSQL = sSQL & " AND IDTipoOggetto=" & 2
        sSQL = sSQL & " AND IDAnagrafica=" & Me.CDSocioFatt.KeyFieldID
        sSQL = sSQL & " AND ((IDTipoUtilizzoSezionale=0 OR IDTipoUtilizzoSezionale IS NULL))"
        Set rs = Cn.OpenResultset(sSQL)
    
        If Not rs.EOF Then
            Link_Sezionale_Socio = fnNotNullN(rs!IDSezionale)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
    End If
    
    If Link_Sezionale_Socio > 0 Then
        Me.CDSezionale.Load Link_Sezionale_Socio
        Me.txtDataDocumento.Value = Date
    End If
End If

'DataControlloRevoca = DateAdd("m", -NumeroMesiPerDataRevocaCertificato, Date)

If Me.CDSocioFatt.KeyFieldID > 0 Then
    'Socio
    With Me.CDSocio
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "DenominazioneCompleta"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaSocio"
        .Filter = "IDAzienda = " & m_App.IDFirm '& " AND IDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID & " AND ((DataUscita IS NULL) OR (DataUscita>" & fnNormDate(DataControlloRevoca) & "))"
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Denominazione Completa"
        .CodeCaption4Find = "Codice"
        .DescriptionCaption4Find = "Denominazione completa"
        .IDExecuteFunction = fncTrovaIDFunzione("Anagrafiche") 'Articoli
        .CodeIsNumeric = False
    End With
Else
    With Me.CDSocio
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "DenominazioneCompleta"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaSocio"
        .Filter = "IDAzienda = " & m_App.IDFirm '& " AND ((DataUscita IS NULL) OR (DataUscita>" & fnNormDate(DataControlloRevoca) & "))"
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Denominazione Completa"
        .CodeCaption4Find = "Codice"
        .DescriptionCaption4Find = "Denominazione completa"
        .IDExecuteFunction = fncTrovaIDFunzione("Anagrafiche") 'Articoli
        .CodeIsNumeric = False
    End With
End If

If Not (BrwMain.Visible) Then Change
Exit Sub
ERR_CDSocioFatt_ChangeElement:
    MsgBox Err.Description, vbCritical, "CDSocioFatt_ChangeElement"
    
End Sub



Private Sub Check1_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdEliminaRifContratto_Click()
    If (MsgBox("Sei sicuro di voler eliminare il riferimento?", vbQuestion + vbYesNo, "Eliminazione riferimento contratto") = vbNo) Then Exit Sub
    Me.txtIDContratto.Value = 0
    Me.txtIDContrattoRiga.Value = 0
End Sub

Private Sub cmdEliminaRifContrattoRiga_Click()
    If (MsgBox("Sei sicuro di voler eliminare il riferimento?", vbQuestion + vbYesNo, "Eliminazione riferimento la riga del contratto") = vbNo) Then Exit Sub
    
    txtIDContrattoRiga.Value = 0
    
End Sub
Private Sub cmdEliminaRifLetInt_Click()
On Error GoTo ERR_cmdEliminaRifLetInt_Click
Dim Testo As String
Dim IDIvaArticolo As Long

If Me.txtIDLetteraIntento.Value = 0 Then Exit Sub
Testo = "Sei sicuro di voler eliminare il riferimento alla lettera d'intento?" & vbCrLf
If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento lettera d'intento") = vbNo Then Exit Sub

Me.txtIDLetteraIntento.Value = 0


Me.cboIvaCliente.WriteOn GET_LINK_IVA_CLIENTE(Me.cdAnagrafica.KeyFieldID)

Me.cboIvaCliente.WriteOn GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, Me.cboIvaCliente.CurrentID)

If Me.cboIvaCliente.CurrentID > 0 Then
    Me.cboIvaArticolo.WriteOn Me.cboIvaCliente.CurrentID
Else
    Me.cboIvaArticolo.WriteOn GET_LINK_IVA_ARTICOLO(Me.CDArticolo.KeyFieldID)
End If

Exit Sub

ERR_cmdEliminaRifLetInt_Click:
MsgBox Err.Description, vbCritical, "ERR_cmdEliminaRifLetInt_Click"
End Sub

Private Sub cmdLetteraIntento_Click()
On Error GoTo ERR_cmdLetteraIntento_Click
    If Me.cdAnagrafica.KeyFieldID = 0 Then Exit Sub
    frmLetteraIntento.Show vbModal
    
    If Me.txtIDLetteraIntento.Value > 0 Then
        Me.cboIvaCliente.WriteOn GET_LINK_IVA_CLIENTE(Me.cdAnagrafica.KeyFieldID)
        
        Me.cboIvaCliente.WriteOn GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, Me.cboIvaCliente.CurrentID)
        
        If Me.cboIvaCliente.CurrentID > 0 Then
            Me.cboIvaArticolo.WriteOn Me.cboIvaCliente.CurrentID
        Else
           Me.cboIvaArticolo.WriteOn GET_LINK_IVA_ARTICOLO(Me.CDArticolo.KeyFieldID)
        End If
    End If
    
Exit Sub
ERR_cmdLetteraIntento_Click:
    MsgBox Err.Description, vbCritical, "cmdLetteraIntento_Click"
End Sub

Private Sub cmdTrovaContratto_Click()
    GET_CONTRATTO
End Sub

Private Sub cmdTrovaContrattoRiga_Click()
    If Me.txtIDContratto.Value = 0 Then Exit Sub
    
    GET_CONTRATTO
End Sub

Private Sub Command1_Click()
    If (OnSave = False) Then Exit Sub
    
    TIPO_SALVATAGGIO = 1
    
    IDAnagrafica_PREC = ACSCliente.IDAnagrafica
    IDDestinazione_PREC = Me.cboAltroSito.CurrentID
    IDVettore_PREC = Me.cboVettore.CurrentID
    IDContratto_PREC = Me.txtIDContratto.Value
    IDContrattoRiga_PREC = Me.txtIDContrattoRiga.Value
    IDCooperativa_PREC = Me.CDSocioFatt.KeyFieldID
    IDAnagraficaSocio_PREC = Me.CDSocio.KeyFieldID
    
    NewRecord
    
End Sub

Private Sub Command10_Click()
    Me.cboIvaArticolo.Enabled = Not Me.cboIvaArticolo.Enabled
End Sub

Private Sub Command11_Click()
    If txtIDAnagraficaCoop.Value = 0 Then
        GET_ANAGRAFICA_COOPERATIVA "", 0, True, Me.txtCodiceAnaCoop.Text, Me.txtAnaCoop.Text
    End If
End Sub

Private Sub Command12_Click()
Dim Testo As String

If (txtIDAnagraficaCoop.Value > 0) Then
    Testo = "Sei sicuro di voler eliminare il riferimento?"
    
    If (MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento") = vbNo) Then Exit Sub

    Me.CDSocioFatt.Load 0
End If
End Sub

Private Sub Command13_Click()
    If txtIDSocio.Value = 0 Then
        GET_ANAGRAFICA_SOCIO "", 0, True, Me.txtCodiceAnaSocio.Text, Me.txtAnaSocio.Text
    End If
End Sub

Private Sub Command14_Click()
Dim Testo As String

If (txtIDSocio.Value > 0) Then
    Testo = "Sei sicuro di voler eliminare il riferimento?"
    
    If (MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento") = vbNo) Then Exit Sub

    Me.CDSocio.Load 0
End If
End Sub

Private Sub Command2_Click()
    TIPO_SALVATAGGIO = 0
    
    If (OnSave = False) Then Exit Sub
    
    NewRecord
    
End Sub

Private Sub Command3_Click()
    OnNew "New"
End Sub

Private Sub Command4_Click()
    If (MsgBox("Sei sicuro di voler procedere con questo comando?", vbCritical + vbYesNo, "Refresh parametri qualitativi") = vbNo) Then Exit Sub
    
    RECUPERO_PAR_QUAL Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm
    RECUPERO_PAR_QUAL_CONTR Me.txtIDContratto.Value
    RECUPERA_INDICE
    
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub Command5_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroTotaleLotti As Long

If Me.txtIDLottoCampagna.Value > 0 Then Exit Sub

STRINGA_RICERCA_LOTTO = Me.txtLottoDiConferimento.Text

frmSelezionaLottoDiCampagna.Show vbModal

CDArticolo.SetFocus
    
End Sub

Private Sub Command6_Click()
    If (MsgBox("Sei sicuro di voler eliminare il riferimento?", vbQuestion + vbYesNo, "Eliminazione riferimento lotto di produzione") = vbNo) Then Exit Sub
    Me.txtIDLottoCampagna.Value = 0
End Sub

Private Sub Command7_Click()
    REFRESH_DESCR_ARTICOLO
End Sub

Private Sub Command8_Click()
On Error GoTo ERR_GET_RIGA_DA_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If (MsgBox("Sei sicuro di voler procedere con questo comando?", vbCritical + vbYesNo, "Recupero prezzi da contratto") = vbNo) Then Exit Sub



sSQL = "SELECT * FROM RV_POIEContrattoDettaglioSel "
sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & Me.txtIDContrattoRiga.Value
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtPrezzoDaContratto.Value = fnNotNullN(rs!Art_prezzo_unitario_neutro)
    Me.txtPrezzoContrattoMin.Value = fnNotNullN(rs!RV_POImportoUnitarioMin)
    Me.txtPrezzoContrattoMax.Value = fnNotNullN(rs!RV_POImportoUnitarioMax)
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_RIGA_DA_CONTRATTO:
    MsgBox Err.Description, vbCritical, "GET_RIGA_DA_CONTRATTO"
End Sub

Private Sub DMTCombo2_Change()

End Sub

Private Sub Form_Activate()
    'Il codice di Form_Activate deve essere eseguito soltanto la prima volta,
    'all'avvio del programma.
    '
    'La variabile m_bOnFirstTime è usata per evitare di eseguire il codice seguente
    'quando si chiude un Form di dialogo e si riattiva frmMain.
    '
    'Queste inizializzazioni non sono state effettuate nella Sub Main() per evitare di
    'rendere visibili le variabili m_DocType, m_Document e m_Changed.
    If m_bOnFirstTime = True Then
    
        m_bOnFirstTime = False

        'Se il filtro di default restituisce dei record si va in modalità variazione
        'ma solo se il primo record non è bloccato altrimenti si va in modalità tabellare
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            'Il filtro ha restituito almeno un record
             
            'Controlla se il primo record su cui si dovrebbe andare in variazione è bloccato.
            If m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
                'Il primo record NON è bloccato
                'allora si effettua il blocco e si va in modalità Variazione
                
                m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
                    
                'La vista alla partenza deve essere quella del Form
                BrwMain.Visible = False

                'Imposta la modalità variazione
                SetStatus4Modality Modify
                
            Else
                'Il primo record è bloccato
                'allora si parte in modalità tabellare
                
                BrwMain.Visible = True
                
                SetStatus4Modality Browse
            End If

            RefreshDescriptions4StatusBar
        Else
            'Il filtro di default non ha restituito nessun record.
            'Si va in modalità inserimento nuovo record
            NewRecord
        
        End If
               
    End If
    
End Sub

Private Sub Form_Initialize()
    ActivityBox.Visible = True
    
    'Impostazione iniziale del flag
    m_bOnFirstTime = True
    
    bEnableGuiEvent = True
End Sub

Private Sub Form_Load()
    'La vista tabellare deve trovarsi sopra tutti gli altri controlli
    BrwMain.ZOrder
    
   'IMPOSTA IL CONTROLLO CHE CONTIENTE I TUTTI GLI ALTRI CONTROLLI
    DMTSplitBar1.ZOrder 0
    Set DMTSplitBar1.dContainer = Me.PicForm2
    'IMPOSTA L'UNITA' DI MISURA DEL FORM
    DMTSplitBar1.ScaleMode = DMTSplit_Twips
    'INIZIALIZZA LA SPLIT BAR
    DMTSplitBar1.SetSplitBar Me.ScaleHeight, Me.ScaleWidth, Me.PicForm.ScaleHeight, Me.PicForm.ScaleWidth

End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If ActivityBox.Visible Then
        imgSplitter.Left = ActivityBox.Width + ActivityBox.Left
    End If
    If m_PreviewWindowHandle > 0 Then
        MoveWindow m_PreviewWindowHandle, CInt(BarMenu.ClientAreaLeft / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaTop / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaWidth / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaHeight / Screen.TwipsPerPixelX), True
    Else
        FormRecalcLayout
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bHandled As Boolean
    
    m_EatKey = False
    
    ShortCut KeyCode, Shift
    
    If KeyCode = 0 And Shift = 0 Then
        m_EatKey = True
    Else
        m_EatKey = False
    End If
    
    
    Select Case KeyCode
        Case vbKeyPageDown
            DMTSplitBar1.ScrollDown
        Case vbKeyPageUp
            DMTSplitBar1.ScrollUp
    End Select
    
    If (KeyCode = vbKeyF3) Then
        If (Me.Frame3.Height = 2295) Then
            Me.Frame3.Height = 3735
        Else
            Me.Frame3.Height = 2295
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If m_EatKey Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' ATTENZIONE
    '-------------------------------------------------------------------------------------
    ' In questo metodo qualsiasi riferimento a proprietà o metodi di un oggetto dovrebbe
    ' essere 'protetto' dal test
    '
    '                                         If obj Is Nothing then .....
    '
    ' perchè il form potrebbe essere scaricato prima che l'oggetto stesso vengana istanziato.
    '-------------------------------------------------------------------------------------
    
    'chiude e distrugge il riferimento alle connessioni
        'CloseConnection
    'Distrugge il riferimento al recordset
    Set BrwMain.Recordset = Nothing
    
    Cancel = FormUnload
End Sub

Private Sub Form_Terminate()

    'Distrugge tutti gli oggetti allocati e provvede ad eliminare gli eventuali blocchi
    'effettuati dalla Semaforo.
    '(Inserire in DestroyObjects il codice per la distruzione degli oggetti allocati)
    DestroyObjects

End Sub

Private Sub brwMain_DblClick()
    
    
'-------------------------------------------------------------------
    'Il documento si sincronizza con la browse
'    If BrwMain.ListIndex > 0 Then
'        m_Document.Move BrwMain.ListIndex - 1
'    End If
'-------------------------------------------------------------------
'NOTA: La versione attuale della dmtGrid effettua automaticamente il
'      Move sul documento.
'-------------------------------------------------------------------
    
    'Se si è in modalità FilterDefinition il DblClick e la pressione
    'di Invio non devono avere alcun effetto
    If BrwMain.GuiMode <> dgFilterDefinition Then
    
        ChangeView
        BrowseReposition
        
        m_Document.AbortNew
        m_Changed = False
        ActivateBarButtons BTN_SAVE, False
        
        
        
    End If
End Sub

Private Sub brwMain_KeyDown(KeyCode As Integer, Shift As Integer)

    'Alla pressione del tasto INVIO dalla modalità tabellare si passa in modalità form.
    If KeyCode = vbKeyReturn And BrwMain.GuiMode = dgNormal And BrwMain.Visible Then
        brwMain_DblClick
    End If
    
    'Viene intercettata la pressione del tasto CANC
    'e la si comunica al form.
    If KeyCode = vbKeyDelete Then
    
        'Prima di cancellare sincronizzo il documento con la selezione
        'fatta nella browse
        If BrwMain.GuiMode = dgNormal And BrwMain.ListIndex > 0 Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    
        ShortCut KeyCode, Shift
    End If
    
End Sub


'Quando si selezionano i documenti dalla modalità tabellare la Caption del form
'va costruita leggendo i valori direttamente dalla riga selezionata nella griglia
'e non da un campo del documento perchè in modalità tabellare non viene eseguito
'il Move sul documento.
Private Sub BrwMain_Reposition(ByVal AllColumns As dmtgridctl.dgColumns)
    If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
        'Monta la caption del form principale
        Me.Caption = Caption2Display(True)
    End If
End Sub


Private Sub BrwMain_OnChangeGuiMode()
    'Se si cambia modalità tramite il menù presente nel controllo
    'dmtGrid occorre effettuare delle impostazioni preliminari nella UserInterface
    
    If bEnableGuiEvent Then
    
        'Modalità FilterDefinition
        If BrwMain.GuiMode = dgFilterDefinition Then
            'Annulla una eventuale operazione di inserimento di un nuovo record
            If m_Document.TableNew Then
                m_Document.AbortNew
            End If
            
            'Impostazioni per la modalità Ricerca
            SetStatus4Modality Find
        End If
        
        'Modalità tabellare
        If BrwMain.GuiMode = dgNormal Then
            'Se si è premuto il pulsante "Visualizzazione tabellare" dalla browse
            'in modalità FilterDefinition e con il recordset vuoto, non si deve andare in
            'modalità tabellare (browse vuota) ma si deve restare in modalità ricerca.
            If (m_Document.EOF = True And m_Document.BOF = True) Then
                BrwMain.GuiMode = dgFilterDefinition
            Else
                'Impostazioni per la modalità tabellare
                SetStatus4Modality Browse
            End If
        End If
    
    End If
End Sub

'Scatenato prima che venga visualizzata la Toolbar della DmtGrid
Private Sub BrwMain_BeforeShowActions()
    
    'Quando si è in modalità FilterDefinition si può andare in
    'modalità tabellare solo se il documento contiene almeno un record.
    If BrwMain.GuiMode = dgFilterDefinition Then
        'Abilita/disabilita il pulsante Modalità Tabellare della dmtGrid
        BrwMain.Actions("TableMode").Enabled = (m_Document.EOF <> True And m_Document.BOF <> True)
    End If
End Sub

'Scatenato quando dalla Browse ( in modalità FilterDefinition ) si clicca su esegui ricerca.
Private Sub BrwMain_OnApplyFilter(ByVal Filter As String)
    ExecuteSearch
End Sub

Private Sub BarMenu_BandClose(ByVal Band As ActiveBar3LibraryCtl.Band)
     'Se la banda è una Toolbar allora viene registrata la chiusura.
    If Band.Type = ddBTNormal And Band.Name <> BAND_CLOSE_PREVIEW Then
        
        'Salva nel registry l'impostazione sulla visibilità della toolbar
        AppOptions.ToolbarVisibility(Band.Name) = False
        
    End If
End Sub

Private Sub BarMenu_BandMove(ByVal Band As ActiveBar3LibraryCtl.Band)
    Form_Resize
End Sub

Private Sub BarMenu_BandOpen(ByVal Band As ActiveBar3LibraryCtl.Band, ByVal Cancel As ActiveBar3LibraryCtl.ReturnBool)
     'Se la banda è una Toolbar allora viene registrata l'apertura.
    If Band.Type = ddBTNormal And Band.Name <> BAND_CLOSE_PREVIEW Then
        AppOptions.ToolbarVisibility(Band.Name) = True
    End If
End Sub

Private Sub BarMenu_MenuItemEnter(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar Tool.Description
End Sub

Private Sub BarMenu_MenuItemExit(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar ""
End Sub

Private Sub BarMenu_MouseEnter(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar Tool.Description
End Sub

Private Sub BarMenu_MouseExit(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar ""
End Sub
Private Sub BarMenu_QueryUnload(Cancel As Integer)
    Cancel = True
End Sub

Private Sub BarMenu_Resize(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    Form_Resize
End Sub

Private Sub BarMenu_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    Dim iKeyCode As Integer
    Dim iShift As Integer
    Dim bContinue As Boolean
    
    On Error GoTo BarMenu_ClickError
        
    'Forza il lostfocus ed attende l'esecuzione di eventuali eventi associati
    AutoLostFocus
        
    bContinue = True
    iShift = GetShift(Tool)
    iKeyCode = GetKeyCode(Tool)
    
    If iKeyCode <> 0 Or iShift <> 0 Then
        bContinue = Not ShortCut(iKeyCode, iShift)
        If bContinue Then
            SendKeys GetSendKeys(Tool) & "(" & GetKey(Tool) & ")"      '"^(R)"
        End If
    Else
        ExecuteMenuCommand Tool.Name
    End If
    
    Exit Sub
    
BarMenu_ClickError:
    If Err.Number = ERR_NDELFILTER Then
        'In seguito a particolari sequenze di eventi può risultare abilitato il cancella filtro sul
        'filtro di default. Se si esegue la cancellazione viene sollevata una eccezione.
        sbMsgError "Non è possibile eliminare il filtro di default.", m_App.FunctionName
    Else
        sbMsgError Err.Description, m_App.FunctionName
    End If
    
    Resume Next
End Sub





Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width, ActivityBox.Height
        picSplitter.AutoRedraw = True
    End With
    picSplitter.Visible = True
    m_SplitterMoving = True
    picSplitter.ZOrder
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If m_SplitterMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < SPLITLIMIT Then
            picSplitter.Left = SPLITLIMIT
        ElseIf sglPos > BarMenu.ClientAreaWidth - SPLITLIMIT Then
            picSplitter.Left = BarMenu.ClientAreaWidth - SPLITLIMIT
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActivityBox.Width = picSplitter.Left - ActivityBox.Left
    FormRecalcLayout
    picSplitter.Visible = False
    m_SplitterMoving = False
End Sub






Private Sub m_App_OnRun(ByVal Proc As Process)
    Dim Parameter As DMTRunAppLib.Parameter

    On Error GoTo ErrorHandler
    
    Set m_Process = Proc
    Set m_DocType = m_Process.IDocType
    
    
    '.................................................................................................................................
    '.................................................................................................................................
    'Gestione preliminare della Semaforo per il controllo dei conflitti di multiutenza
    
    
    'Inizializza la Semaforo
    InitSemaphore
    
    ' Verifica se l'applicazione corrente è bloccata da altri gestori.
    ' (Il controllo avviene sul Tipo Oggetto correntemente trattato.)
    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, SemAllObjects, SemAllActions) Then
        '-------------------------------------------------------------
        'Il programma è bloccato da un'altra manutenzione in esecuzione.
        '-------------------------------------------------------------
        
        'Scarica il form
        Unload Me
       
        'Prima di terminare il programma è bene distruggere tutti gli oggetti allocati
        DestroyObjects
       
        'Termina il programma
        End
    End If
    
    '----------------------------------------------------
    'Il programma non è bloccato e prosegue normalmente.
    '----------------------------------------------------
    
    'Ripulisce la tabella semaforo.
    'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
    SemaphoreUnlock
    
    'Imposta gli eventuali blocchi (semaforo) su altre manutenzioni.
    SemaphoreLock
    '.................................................................................................................................
    '.................................................................................................................................
    
    
    
    Select Case Proc.Name
        '*
        'Inserire il codice per la gestione del processo
        '*
        Case "Manutenzione"
        '   For Each Parameter In Proc.Parameters
        '       Select Case Parameter.Name
        '       *
        '       Inserire il codice per la gestione del parametro
        '       *
        '       Case ParameterName??????
        '       End Select
        '   next
           Start 'di solito
    
    Case Else
    
        'cbcx
        'QUESTA PARTE DEVE ESSERE RIVISTA
        '-----------------------------------------------------------------
        
'''''        Dim ErrorMsg As String
'''''
'''''        ErrorMsg = "No processes to execute" & vbCrLf
'''''        ErrorMsg = ErrorMsg & "This application is able to execute these processes:" & vbCrLf
'''''        '*
'''''        'Inserire i processi che l'applicazione sa eseguire
'''''        '*
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE & vbCrLf
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_EXTENDED_DATABASE & vbCrLf
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_DA_SHELL & vbCrLf
'''''        Err.Raise ERR_NO_PROCESSES, , ErrorMsg


    End Select
    Exit Sub
ErrorHandler:
    SemaphoreUnlock
    ShowErrorLog
End Sub

Private Sub m_Document_OnReposition()
        
    'Viene creata (se non è già stato fatto) la collezione FormFields
    CreateFormFields
   
    If Not m_Document.TableNew Then
        'Se EOF = true o BOF = true vuol dire che si è andati oltre l'ultimo o
        'prima del primo record. In tal caso non si deve fare il refresh dei
        'controlli del form.
        If Not (m_Document.EOF Or m_Document.BOF) Then
            BrowseReposition
            
            'cbcx
            '---------------------------------------------
            'Gestione processo On_Extend
'            If Not m_ExtendApplication Is Nothing Then
'                'Notifica l'identificativo unico del documento corrente
'                m_ExtendApplication.PrimaryID = m_Document.Fields("ID" & m_App.TableName).Value
'            End If
            
            
        End If
    Else
        'Nel caso di inserimento nuovo record ripulisce i campi del form
        ClearFormFields
    End If
    On Error Resume Next
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) > 0 Then
        RECUPERO_PAR_QUAL_CERT fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    End If
    
    GET_DOCUMENTO_COLLEGATO fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    'rif11 begin
    
    'rif11 end
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 07/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: AutoLostFocus
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Forza un LostFocus del controllo attivo ed attende la gestione di eventuali eventi associati.
'                  Alla fine ripristina il fuoco sul controllo iniziale.
'                  Usata quando si clicca sulla toolbar e quando si utilizza l'acceleratore per il salvataggio SHIFT + F12
'                  (in tal caso infatti non viene scatenato l'evento BarMenu_Click)
'
'**/
Private Sub AutoLostFocus()
    Dim Ctr As Control

    
    'Se si è in modalità FilterDefinition non si deve spostare il fuoco
    'altrimenti Taglia, Copia e Incolla (dalla toolbar) non possono funzionare
    If BrwMain.GuiMode <> dgFilterDefinition And Not Me.ActiveControl Is Nothing Then
    
        'Memorizza il controllo che ha il fuoco
        Set Ctr = Me.ActiveControl
    
        'Forza il lost focus del controllo attivo
        Globali.SetFocus PicForm.hwnd
        
        'Vengono gestiti gli eventi LostFocus (se previsti)
        DoEvents
        
        'Ripristina il fuoco sul controllo.
        Ctr.SetFocus
        
    End If

End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 11/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: InitSemaphore
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Inizializzazione del semaforo per la gestione
'                  dei conflitti in caso di multiutenza

'
'**/
Private Sub InitSemaphore()

    Set m_Semaphore = New Semaforo.dmtSemaphore
    Set m_Semaphore.Database = m_App.Database.Connection
    Set m_Semaphore.objRes = gResource
    
    m_Semaphore.IDUser = m_App.IDUser
    m_Semaphore.IDBranch = m_App.Branch
    m_Semaphore.IDFunction = m_App.FunctionID
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreLock
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'                 ////////////////////////////////////////////////////////////////////////
'                     Impostare qui gli eventuali blocchi sulle altre manutenzioni
'                 ////////////////////////////////////////////////////////////////////////
'**/
Private Sub SemaphoreLock()
    If Not m_Semaphore Is Nothing Then
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions

    End If
End Sub

'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreUnlock
'
'Parametri:
'
'Valori di ritorno:

'Funzionalità:
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'                     Sbloccare qui le altre manutenzioni (bloccate precedentemente in SemaphoreLock)
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Sub SemaphoreUnlock()
    If Not m_Semaphore Is Nothing Then
    
        'Ripulisce la tabella semaforo per quanto riguarda il Tipo Oggetto e l'utente correnti
        m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        
        
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
        'Sblocca le manutenzioni bloccate precedentemente
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions
    
    End If
End Sub




'**+
'Autore: Diamante s.p.a
'Data creazione: 11/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: DestroyObjects
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'                  ////////////////////////////////////////////////////////////////////////////////////////////////////
'                  /         Inserire qui il codice per distruggere (prima che venga terminato il programma)     /
'                  /         tutti gli oggetti allocati                                                                              /
'                  ////////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Sub DestroyObjects()
    
    'Sblocca gli eventuali gestori bloccati da questa manutenzione
    SemaphoreUnlock

    Set m_FormFields = Nothing
    Set m_Report = Nothing
    Set m_ActiveFilter = Nothing
    Set m_Document = Nothing
    Set m_Process = Nothing
    Set m_App = Nothing
    Set m_Semaphore = Nothing
    
    'cbcx
    'Set m_ExtendApplication = Nothing
End Sub





Public Sub ConnessioneDiamanteADO()
On Error GoTo ERR_ConnessioneDiamanteADO
    '------------------------------
    'APERTURA DELLA CONNESSIONE
    '------------------------------
    
    'Leggiamo il tipo di database utilizzato (Access o SQL Server)
    'Apriamo la connessione in base al tipo di database rilevato
    '(MenuOptions.DBType restituisce il valore del DBType)
    'Select Case MenuOptions.DBType
    '    Case 0 'CONNESSIONE_SQL_SERVER            'Microsoft SQL Server
    '        Set Cn = adoEngine.adoEnvironments(0).OpenConnection("", , , "DSN=Diamante;UID=sa;PWD=")
    '    Case 1 'CONNESSIONE_ACCESS               'Microsoft ACCESS
    '        Set Cn = adoEngine.adoEnvironments(0).OpenConnection("", , , "DSN=Diamante;UID=admin;PWD=dmt192981046")
    '    Case -1
            'Se la voce DBType non viene trovata nel file di registro
            'vuol dire che Diamante non è stato installato correttamente
    '        MsgBox "Impossibile avviare il programma. Diamante non è stato installatto correttamente!", vbCritical, "Aggiornamento scadenze"
    '        End
    'End Select
    
    Set Cn = m_App.Database.Connection
    
Exit Sub
ERR_ConnessioneDiamanteADO:
    MsgBox Err.Description, vbCritical, "Connessione Diamante di tipo ADO"
End Sub


'**+
'Nome: SendDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue l'esportazione del documento con controllo di errore
'**/
Private Sub SendDocument(ByVal Appl As Long)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    
    OLDCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    m_Document.SendMail m_Report, Appl
    Screen.MousePointer = OLDCursor
    Exit Sub
errHandler:
    Screen.MousePointer = OLDCursor
    
    If Err.Number = 20507 Then
        'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
        sbMsgInfo "File di report non trovato", m_App.FunctionName
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
End Sub
'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : InitSemaphore
'
'Parametri                  :
'
'Funzionalità               : Attiva/disattiva le attività del Riquadro attività
'
'**/
Private Sub EnableDOMActivitiesItems()
    oFiltersActivity.EnableItems (BrwMain.GuiMode = dgNormal And BrwMain.Visible)
    oTableViewsActivity.EnableItems (BrwMain.GuiMode = dgNormal And BrwMain.Visible)
    
    ActivityBox.Redraw = True
End Sub

Private Sub GET_PARAMATRI_FILIALE()
On Error GoTo ERR_GET_PARAMATRI_FILIALE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Link_TipoImballo = fnNotNullN(rs!IDTipoImballo)
    Link_TipoSocio = fnNotNullN(rs!IDCategoriaAnagrafica)
    LINK_TIPO_ARROTONDAMENTO = fnNotNullN(rs!IDTipoArrotondIndiceVarCert)
    ATTIVA_SEZIONALE_DA_SOCIO = fnNotNullN(rs!PrelevaSezionaleDaSocioCert)
    UtilizzaDataENumSocioPerDDT = fnNotNullN(rs!UtilizzaDataENumSocioPerDDT)
    NumeroColliPerAutomezzoCert = fnNotNullN(rs!NumeroColliPerAutomezzoCert)
    LINK_MAGAZZINO_DOCUMENTO = fnNotNullN(rs!IDMagazzino_Vendita)
    IDClassLottoProdPerFuoriQuota = fnNotNullN(rs!IDClassificazioneLottoProdPerFuoriQuota)
    MsgInDocSeRigaMerceSenzaImballo = fnNotNullN(rs!MsgInDocSeRigaMerceSenzaImballo)
    IDAnagraficaDestSociDiretti = fnNotNullN(rs!IDAnagraficaDestinazionePerCertificato)
    IDCategoriaAnagraficaSocioDiretto = fnNotNullN(rs!IDCategoriaAnagraficaSocioDiretto)
    IDCategoriaAnagraficaProdAcq = fnNotNullN(rs!IDCategoriaAnagraficaProdAcq)
    IDCategoriaAnagraficaNoProd = fnNotNullN(rs!IDCategoriaAnagraficaNoProd)
    IDArticoloScartoPerCertificato = fnNotNullN(rs!IDArticoloScartoPerCertificato)
    RiportaDestinazioneDaContrattoCertificato = fnNotNullN(rs!RiportaDestinazioneDaContrattoCertificato)
    RiportaVettoreDaContrattoCertificato = fnNotNullN(rs!RiportaVettoreDaContrattoCertificato)
    ForzaDestinazioneDaContrattoCertificato = fnNotNullN(rs!ForzaDestinazioneDaContrattoCertificato)
    ForzaVettoreDaContrattoCertificato = fnNotNullN(rs!ForzaVettoreDaContrattoCertificato)
    AttivaSelezioneSocioCertPerVarieta = fnNotNullN(rs!AttivaSelezioneSocioCertPerVarieta)
    AttivaSelezioneAnaVeloceInCert = fnNotNullN(rs!AttivaSelezioneAnaVeloceInCert)
    NumeroMesiPerDataRevocaCertificato = fnNotNullN(rs!NumeroMesiPerDataRevocaCertificato)
    NonRiportaInXMLRifVsNumOrd = fnNotNullN(rs!NonRiportaInXMLRifVsNumOrd)
    NonRiportareRifCerticatoInDDT = fnNotNullN(rs!NonInviareRifCertificatoIdDDT)
    
Else
    Link_TipoImballo = 0
    Link_TipoSocio = 0
    LINK_TIPO_ARROTONDAMENTO = 0
    ATTIVA_SEZIONALE_DA_SOCIO = 0
    UtilizzaDataENumSocioPerDDT = 0
    NumeroColliPerAutomezzoCert = 1
    LINK_MAGAZZINO_DOCUMENTO = 0
    IDClassLottoProdPerFuoriQuota = 0
    MsgInDocSeRigaMerceSenzaImballo = 0
    IDAnagraficaDestSociDiretti = 0
    IDCategoriaAnagraficaSocioDiretto = 0
    IDCategoriaAnagraficaProdAcq = 0
    IDCategoriaAnagraficaNoProd = 0
    IDArticoloScartoPerCertificato = 0
    RiportaDestinazioneDaContrattoCertificato = 0
    RiportaVettoreDaContrattoCertificato = 0
    ForzaDestinazioneDaContrattoCertificato = 0
    ForzaVettoreDaContrattoCertificato = 0
    AttivaSelezioneSocioCertPerVarieta = 0
    AttivaSelezioneAnaVeloceInCert = 0
    NumeroMesiPerDataRevocaCertificato = 1
    NonRiportaInXMLRifVsNumOrd = 0
    NonRiportareRifCerticatoInDDT = 0
End If

rs.CloseResultset
Set rs = Nothing

RECUPERA_CONFIG_CAUS_XML

Exit Sub
ERR_GET_PARAMATRI_FILIALE:
    MsgBox Err.Description, vbCritical, "GET_PARAMATRI_FILIALE"
End Sub
Private Function fncTrovaIDFunzione(Gestore As String, Optional Funzione As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE Gestore.Gestore = " & fnNormString(Gestore)
sSQL = sSQL & " AND Funzione = " & fnNormString(Funzione)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaIDFunzione = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub AggiornaAltreDestinazioni()
    With Me.cboAltroSito
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica FROM SitoPerAnagrafica"
        .SQL = .SQL & " WHERE IDAnagrafica = " & Me.cdAnagrafica.KeyFieldID
        .SQL = .SQL & " ORDER BY SitoPerAnagrafica"
    End With
End Sub
Private Sub GET_CONTRATTO()
On Error GoTo ERR_GET_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroTotaleContratti As Long
Dim NumeroTotaleRigheContratto As Long

If bloading = True Then Exit Sub

LINK_CONTRATTO = Me.txtIDContratto.Value
LINK_CLIENTE_CONTRATTO = Me.ACSCliente.IDAnagrafica
NumeroTotaleContratti = 0
NumeroTotaleRigheContratto = 0

If LINK_CONTRATTO = 0 Then
    
    sSQL = "SELECT COUNT(IDOggetto) AS NumeroRecord FROM RV_POIEContrattoSel "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND Link_Nom_anagrafica=" & Me.cdAnagrafica.KeyFieldID
    sSQL = sSQL & " AND RV_POContrattoChiuso=0"
    sSQL = sSQL & " AND ((Doc_data_scadenza>=" & fnNormDate(Date) & ") OR (Doc_data_scadenza IS NULL))"
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        NumeroTotaleContratti = fnNotNullN(rs!NumeroRecord)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If (NumeroTotaleContratti > 1) Then
        frmContratto.Show vbModal
    Else
        If (NumeroTotaleContratti = 0) Then
            MsgBox "Non ci sono contratti disponibili per il cliente selezionato!", vbInformation, "Controllo dati contratto"
            Exit Sub
        End If
        If (NumeroTotaleContratti = 1) Then
            sSQL = "SELECT IDOggetto FROM RV_POIEContrattoSel "
            sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
            sSQL = sSQL & " AND Link_Nom_anagrafica=" & Me.cdAnagrafica.KeyFieldID
            sSQL = sSQL & " AND RV_POContrattoChiuso=0"
            sSQL = sSQL & " AND ((Doc_data_scadenza>=" & fnNormDate(Date) & ") OR (Doc_data_scadenza IS NULL))"
            Set rs = Cn.OpenResultset(sSQL)
            
            If Not rs.EOF Then
                LINK_CONTRATTO = fnNotNullN(rs!IDOggetto)
            End If
            
            rs.CloseResultset
            Set rs = Nothing
            
        End If
    End If
    If LINK_CONTRATTO > 0 Then
        GET_DATI_DA_CONTRATTO LINK_CONTRATTO
    End If
End If
If LINK_CONTRATTO > 0 Then
    If Me.txtIDContrattoRiga.Value = 0 Then
        sSQL = "SELECT COUNT(IDValoriOggettoDettaglio) AS NumeroRecord FROM RV_POIEContrattoDettaglioSel "
        sSQL = sSQL & "WHERE IDOggetto=" & LINK_CONTRATTO
        sSQL = sSQL & " AND RV_POTipoRiga=1"
        Set rs = Cn.OpenResultset(sSQL)
        
        If Not rs.EOF Then
            NumeroTotaleRigheContratto = fnNotNullN(rs!NumeroRecord)
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        
        If (NumeroTotaleRigheContratto > 1) Then
            frmContrattoDettaglio.Show vbModal
            Me.CDSocioFatt.SetFocus
        Else
            If (NumeroTotaleRigheContratto = 0) Then
                MsgBox "Non è presente nessun dettaglio per il contratto selezionato!", vbInformation, "Controllo dati contratto"
                Exit Sub
            End If
            If (NumeroTotaleRigheContratto = 1) Then
                sSQL = "SELECT IDValoriOggettoDettaglio FROM RV_POIEContrattoDettaglioSel "
                sSQL = sSQL & "WHERE IDOggetto=" & LINK_CONTRATTO
                sSQL = sSQL & " AND RV_POTipoRiga=1"
                Set rs = Cn.OpenResultset(sSQL)
                
                If Not rs.EOF Then
                    Me.txtIDContrattoRiga.Value = fnNotNullN(rs!IDValoriOggettoDettaglio)
                End If
                
                rs.CloseResultset
                Set rs = Nothing
                
                Me.Command11.SetFocus
                
            End If
        End If
        
        If AttivaSelezioneAnaVeloceInCert = 1 Then
            If (Me.CDSocioFatt.KeyFieldID = 0) Then
                frmSelAnagraficaCoop.Show vbModal
                If LINK_ANA_COOP_SEL > 0 Then
                    Me.CDSocioFatt.Load LINK_ANA_COOP_SEL
                    frmSelAnagraficaSocio.Show vbModal
                    If LINK_ANA_SOCIO_SEL > 0 Then
                        Me.CDSocio.Load LINK_ANA_SOCIO_SEL
                        If (Me.txtIDLottoCampagna.Value = 0) Then
                            Command5_Click
                        End If
                        Me.txtNumeroCertificato.SetFocus
                    End If
                End If
            End If
        End If
    End If
End If
Exit Sub
ERR_GET_CONTRATTO:
    Screen.MousePointer = 0
    DATI_DA_CONTRATTO = False
    MsgBox Err.Description, vbCritical, "GET_CONTRATTO"
End Sub

Private Sub GET_DATI_DA_CONTRATTO(IDContratto As Long)
On Error GoTo ERR_DATI_DA_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ValoriOggettoPerTipo000E "
sSQL = sSQL & "WHERE IDOggetto=" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.cdAnagrafica.Load fnNotNullN(rs!Link_Nom_anagrafica)
    
    If (RiportaDestinazioneDaContrattoCertificato = 1) Then
        If (fnNotNullN(rs!Link_Nom_ult_sito) > 0) Then
            Me.cboAltroSito.WriteOn fnNotNullN(rs!Link_Nom_ult_sito)
        Else
            If (ForzaDestinazioneDaContrattoCertificato = 1) Then
                Me.cboAltroSito.WriteOn fnNotNullN(rs!Link_Nom_ult_sito)
            End If
        End If
    End If
    If (RiportaVettoreDaContrattoCertificato = 1) Then
        If (fnNotNullN(rs!Link_Vet_vettore) > 0) Then
            Me.cboVettore.WriteOn fnNotNullN(rs!Link_Vet_vettore)
        Else
            If (ForzaVettoreDaContrattoCertificato = 1) Then
                Me.cboVettore.WriteOn fnNotNullN(rs!Link_Vet_vettore)
            End If
        End If
    End If
    
    Me.txtIDContratto.Value = fnNotNullN(rs!IDOggetto)

    LINK_CONTRATTO = fnNotNullN(rs!IDOggetto)
    
    RECUPERO_PAR_QUAL_CONTR LINK_CONTRATTO
End If
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_DATI_DA_CONTRATTO:
    MsgBox Err.Description, vbCritical, "DATI_DA_CONTRATTO"
End Sub
Private Sub GET_RIGA_DA_CONTRATTO(IDContrattoRiga As Long)
On Error GoTo ERR_GET_RIGA_DA_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIEContrattoDettaglioSel "
sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & IDContrattoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.CDArticolo.Load fnNotNullN(rs!Link_Art_articolo)
    Me.txtDescrizioneArticolo.Text = fnNotNull(rs!Art_descrizione)
    Me.CDImballo.Load fnNotNullN(rs!RV_POIDImballo)
    Me.txtTaraUnitaria.Value = fnNotNullN(rs!RV_POTaraUnitariaImballo)
    Me.txtIDContrattoRiga.Value = fnNotNullN(rs!IDValoriOggettoDettaglio)
    Me.txtPrezzoDaContratto.Value = fnNotNullN(rs!Art_prezzo_unitario_neutro)
    Me.txtPrezzoContrattoMin.Value = fnNotNullN(rs!RV_POImportoUnitarioMin)
    Me.txtPrezzoContrattoMax.Value = fnNotNullN(rs!RV_POImportoUnitarioMax)
    REFRESH_DESCR_ARTICOLO
    CALCOLA_PESI
    CALCOLA_TOTALE_RIGA
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_RIGA_DA_CONTRATTO:
    MsgBox Err.Description, vbCritical, "GET_RIGA_DA_CONTRATTO"
End Sub

Private Function GET_LINK_LETTERA_INTENTO_PRED(IDAnagrafica As Long, IDTipoAnagrafica As Long, DataDocumento As String, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_LETTERA_INTENTO_PRED
Dim sSQL As String
Dim IDLetteraIntento As Integer
Dim cmd As ADODB.Command
    
    IDLetteraIntento = 0
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SP_LetteraIntentoDefault"
    cmd.ActiveConnection = Cn.InternalConnection
    
    cmd.Parameters.Append cmd.CreateParameter("IDAzienda", adInteger, adParamInput, , IDAzienda)
    cmd.Parameters.Append cmd.CreateParameter("IDTipoAnagrafica", adInteger, adParamInput, , IDTipoAnagrafica)
    cmd.Parameters.Append cmd.CreateParameter("IDAnagrafica", adInteger, adParamInput, , IDAnagrafica)
    cmd.Parameters.Append cmd.CreateParameter("DataDocumento", adDate, adParamInput, , DataDocumento)
    cmd.Parameters.Append cmd.CreateParameter("IDLetteraIntento", adInteger, adParamOutput, , IDLetteraIntento)
    cmd.Execute
    
    IDLetteraIntento = cmd.Parameters(4).Value
    
    GET_LINK_LETTERA_INTENTO_PRED = IDLetteraIntento
    
Exit Function
ERR_GET_LINK_LETTERA_INTENTO_PRED:
    MsgBox Err.Description, vbCritical, "Recupero lettera d'intento"
    
End Function





Private Sub txtAnaCoop_LostFocus()
    If txtIDAnagraficaCoop.Value = 0 Then
        GET_ANAGRAFICA_COOPERATIVA txtAnaCoop.Text, 2
    End If
End Sub

Private Sub txtAnaSocio_LostFocus()
    If txtIDSocio.Value = 0 Then
        GET_ANAGRAFICA_SOCIO txtAnaSocio.Text, 2
    End If
End Sub

Private Sub txtCodiceAnaCoop_LostFocus()
    If txtIDAnagraficaCoop.Value = 0 Then
        GET_ANAGRAFICA_COOPERATIVA txtCodiceAnaCoop.Text, 1
    End If
End Sub

Private Sub txtCodiceAnaSocio_LostFocus()
    If txtIDSocio.Value = 0 Then
        GET_ANAGRAFICA_SOCIO txtCodiceAnaSocio.Text, 1
    End If
End Sub

Private Sub txtColliEntrata_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtColliEntrata_LostFocus()
    Me.txtColliUscita.Value = Me.txtColliEntrata.Value
    CALCOLA_PESI
    
End Sub

Private Sub txtColliUscita_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataCertificato_LostFocus()
If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    REFRESH_DESCR_ARTICOLO
End If
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataDDT_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataDDT_LostFocus()
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        REFRESH_DESCR_ARTICOLO
    End If
    If (Me.txtDataTrasporto.Value = 0) Then
        Me.txtDataTrasporto.Value = Me.txtDataDDT.Value
        Me.txtOraTrasporto.Value = 32400
    End If
    If (Me.txtIDLottoCampagna.Value > 0) Then
        Me.txtColliEntrata.SetFocus
    End If
End Sub

Private Sub txtDataTrasporto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDescrizioneArticolo_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtIDAnagraficaCoop_Change()
    Me.txtCodiceAnaCoop.Text = Me.CDSocioFatt.Code
    Me.txtAnaCoop.Text = Me.CDSocioFatt.Description
    
    Me.txtAnaCoop.Locked = Me.txtIDAnagraficaCoop > 0
    Me.txtCodiceAnaCoop.Locked = Me.txtIDAnagraficaCoop > 0
End Sub

Private Sub txtIDContratto_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.txtDescrizioneContratto.Text = ""

If (fnNotNullN(Me.txtIDContratto.Text) = 0) Then Exit Sub

sSQL = "SELECT IDOggetto, Doc_data, Doc_numero, Doc_numero_vs_ordine_di_rifer,"
sSQL = sSQL & "Link_Nom_ult_sito, Link_Vet_vettore "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000E "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDContratto.Value

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtDescrizioneContratto.Text = "Contratto n° " & fnNotNullN(rs!Doc_numero) & " del " & fnNotNull(rs!Doc_data)
    If (RiportaDestinazioneDaContrattoCertificato = 1) Then
        If (fnNotNullN(rs!Link_Nom_ult_sito) > 0) Then
            Me.cboAltroSito.WriteOn fnNotNullN(rs!Link_Nom_ult_sito)
        Else
            If (ForzaDestinazioneDaContrattoCertificato = 1) Then
                Me.cboAltroSito.WriteOn fnNotNullN(rs!Link_Nom_ult_sito)
            End If
        End If
    End If
    If (RiportaVettoreDaContrattoCertificato = 1) Then
        If (fnNotNullN(rs!Link_Vet_vettore) > 0) Then
            Me.cboVettore.WriteOn fnNotNullN(rs!Link_Vet_vettore)
        Else
            If (ForzaVettoreDaContrattoCertificato = 1) Then
                Me.cboVettore.WriteOn fnNotNullN(rs!Link_Vet_vettore)
            End If
        End If
    End If
    
    If (Len(fnNotNull(rs!Doc_numero_vs_ordine_di_rifer)) > 0) Then
        Me.txtDescrizioneContratto.Text = Me.txtDescrizioneContratto.Text + " RIF: " & fnNotNull(rs!Doc_numero_vs_ordine_di_rifer)
    End If
End If

rs.CloseResultset
Set rs = Nothing

If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtIDContrattoRiga_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.txtDescrizioneContrattoRiga.Text = ""
LINK_VARIETA_ART_CONTRATTO = 0
LINK_FAMIGLIA_ART_CONTRATTO = 0

If (fnNotNullN(Me.txtIDContrattoRiga.Text) = 0) Then Exit Sub

sSQL = "SELECT ValoriOggettoDettaglio0038.IDValoriOggettoDettaglio, ValoriOggettoDettaglio0038.IDOggetto, ValoriOggettoDettaglio0038.Art_codice, ValoriOggettoDettaglio0038.Art_descrizione, "
sSQL = sSQL & "ValoriOggettoDettaglio0038.RV_PODataInizioConsegna , ValoriOggettoDettaglio0038.RV_PODataFineConsegna, "
sSQL = sSQL & "ValoriOggettoDettaglio0038.Link_Art_articolo, Articolo.RV_PO01_IDVarieta, Articolo.RV_PO01_IDFamigliaProdotti "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0038 INNER JOIN "
sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0038.Link_Art_articolo = Articolo.IDArticolo "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & Me.txtIDContrattoRiga.Value

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtDescrizioneContrattoRiga.Text = fnNotNull(rs!Art_codice)
    If (Len(fnNotNull(rs!Art_descrizione) > 0)) Then
        Me.txtDescrizioneContrattoRiga.Text = txtDescrizioneContrattoRiga.Text + " " & fnNotNull(rs!Art_descrizione)
    End If
    LINK_VARIETA_ART_CONTRATTO = fnNotNullN(rs!RV_PO01_IDVarieta)
    LINK_FAMIGLIA_ART_CONTRATTO = fnNotNullN(rs!RV_PO01_IDFamigliaProdotti)
End If

rs.CloseResultset
Set rs = Nothing

If (fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0) Then
    If (Me.txtIDContrattoRiga.Value > 0) Then
        GET_RIGA_DA_CONTRATTO Me.txtIDContrattoRiga.Value

    End If
End If

If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtIDLetteraIntento_Change()
On Error GoTo ERR_txtIDLetteraIntento_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If IsEmpty(Me.txtIDLetteraIntento.Value) Then Me.txtIDLetteraIntento.Value = 0

sSQL = "SELECT * FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & Me.txtIDLetteraIntento.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtNLetteraIntento.Text = ""
    Me.txtDataLetteraIntento.Value = 0
    
    'Me.lblLetteraIntento.ToolTipText = ""
Else
    Me.txtNLetteraIntento.Text = fnNotNull(rs!Numero)
    Me.txtDataLetteraIntento.Value = fnNotNullN(rs!Data)

    'Me.lblLetteraIntento.ToolTipText = "Prot. N° " & fnNotNull(rs!NumeroCliFor) & " del " & fnNotNull(rs!DataEmissione)
End If

rs.CloseResultset
Set rs = Nothing

If Not (BrwMain.Visible) Then Change
Exit Sub
ERR_txtIDLetteraIntento_Change:
    MsgBox Err.Description, vbCritical, "txtIDLetteraIntento_Change"
End Sub
Private Function GET_LINK_IVA_LETTERA_INTENTO(IDLetteraIntento As Long, IDIvaCliente As Long) As Long
On Error GoTo ERR_GET_LINK_IVA_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & IDLetteraIntento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
Else
    If fnNotNullN(rs!IDIva) > 0 Then
        GET_LINK_IVA_LETTERA_INTENTO = fnNotNullN(rs!IDIva)
    Else
        GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_IVA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_IVA_LETTERA_INTENTO"
End Function

Private Function GET_LINK_IVA_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_CLIENTE = 0
Else

    GET_LINK_IVA_CLIENTE = fnNotNullN(rs!IDIva)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
End Function
Private Function GET_LINK_IVA_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIvaVendita "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_ARTICOLO = 0
Else

    GET_LINK_IVA_ARTICOLO = fnNotNullN(rs!IDIvaVendita)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
End Function
Private Sub CALCOLA_PESI()
    Me.txtTaraTotaleImballo.Value = Me.txtColliEntrata.Value * Me.txtTaraUnitaria.Value
    Me.txtTaraTotale.Value = Me.txtTaraCamion.Value + Me.txtTaraTotaleImballo.Value
    Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTaraTotale.Value
    If (Me.txtPesoNetto.Value > 0) Then
        Me.txtPercRidPesoNetto.Value = (Me.txtScarto.Value / Me.txtPesoNetto.Value) * 100
    End If
    Me.txtQtaFatturazione.Value = Me.txtPesoNetto.Value - Me.txtScarto.Value
End Sub

Private Sub CALCOLA_TOTALE_RIGA()
Dim Testo As String

    Me.txtTotaleRiga.Value = Me.txtQtaFatturazione.Value * Me.txtPrezzoDiFatturazione.Value
    Me.txtIndiceDiVariazione.Value = 0
    txtIndiceDiVariazione100.Value = 0
    
    If (Me.txtPrezzoDaContratto.Value > 0) Then
        Me.txtIndiceDiVariazione.Value = (((Me.txtPrezzoDiFatturazione.Value - Me.txtPrezzoDaContratto.Value) / Me.txtPrezzoDaContratto.Value) * 100)
        Me.txtIndiceDiVariazione100.Value = (Me.txtPrezzoDiFatturazione.Value / Me.txtPrezzoDaContratto.Value) * 100
    End If
    Me.txtIndiceDiVariazione.Value = Round(Me.txtIndiceDiVariazione.Value, 5)
    
    Me.txtIndiceDiVariazioneEff.Value = fnRoundChange(Me.txtIndiceDiVariazione.Value, 1, 3)
    
    Select Case LINK_TIPO_ARROTONDAMENTO
        Case 3 'Difetto
            Me.txtIndiceDiVariazioneEff.Value = fnRoundDown(Me.txtIndiceDiVariazione.Value)
        Case 4 'Eccesso
            Me.txtIndiceDiVariazioneEff.Value = fnRoundUp(Me.txtIndiceDiVariazione.Value)
    End Select
    
    RECUPERA_INDICE
    

End Sub

Private Sub txtIDLottoCampagna_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Long
Dim IDRegioneLotto As Long
Dim Testo As String

Me.txtLottoDiConferimento.Text = ""
Me.cboFamigliaLotto.WriteOn 0
Me.cboVarietaLotto.WriteOn 0
Me.txtLottoDiConferimento.Enabled = True
Me.txtRegioneLotto.Text = ""
Me.txtTotaleEttariLotto.Value = 0
Me.txtResaMinPerHa.Value = 0
Me.txtResaMaxPerHa.Value = 0
Me.txtResaMinTotale.Value = 0
Me.txtResaMaxTotale.Value = 0
Me.txtQtaUtilizzataLotto.Value = 0
LINK_SOCIO_LOTTO_SEL = 0

If (fnNotNullN(Me.txtIDLottoCampagna.Text) = 0) Then Exit Sub

'''RECUPERO GENERALE DELLE INFORMAZIONI DEL LOTTO
sSQL = "SELECT * FROM RV_PO01_IELottoDiCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & Me.txtIDLottoCampagna.Value

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    LINK_SOCIO_LOTTO_SEL = fnNotNullN(rs!IDSocio)
    Me.txtLottoDiConferimento.Text = fnNotNull(rs!CodiceLotto)
    Me.cboVarietaLotto.WriteOn fnNotNullN(rs!IDRV_PO01_Varieta)
    Me.cboFamigliaLotto.WriteOn fnNotNullN(rs!IDRV_PO01_FamigliaProdotti)
    Me.txtTotaleEttariLotto.Value = fnNotNullN(rs!DimensioneMQ) / 10000
    Me.Check1.Value = vbUnchecked
    If (fnNotNullN(rs!Acquistato) = 1) Then
        Me.Check1.Value = vbChecked
    End If
End If

rs.CloseResultset
Set rs = Nothing

'''REGIONI IN BASE AL CATASTO
I = 0
sSQL = "SELECT * FROM RV_PO01_IERegionePerLottoDiCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & Me.txtIDLottoCampagna.Value

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If (I > 0) Then
        Me.txtRegioneLotto.Text = Me.txtRegioneLotto.Text = " - "
    End If
    Me.txtRegioneLotto.Text = fnNotNull(rs!Regione)
    IDRegioneLotto = fnNotNullN(rs!IDRegione)
    I = I + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing


'''RESA PER FAMIGLIA DEL LOTTO
sSQL = "SELECT * FROM RV_PO01_Varieta "
sSQL = sSQL & "WHERE IDRV_PO01_Varieta=" & Me.cboVarietaLotto.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If Not (rs.EOF) Then
    Me.txtResaMinPerHa.Value = fnNotNullN(rs!ResaMinima)
    Me.txtResaMaxPerHa.Value = fnNotNullN(rs!ResaMassima)
    Me.txtResaMinTotale.Value = Me.txtResaMinPerHa.Value * Me.txtTotaleEttariLotto.Value
    Me.txtResaMaxTotale.Value = Me.txtResaMaxPerHa.Value * Me.txtTotaleEttariLotto.Value
End If

rs.CloseResultset
Set rs = Nothing

'''RECUPERO DELLA QUANTITA
sSQL = "SELECT SUM(PesoNettoCalcolato) AS TotalePesoUtilizzato "
sSQL = sSQL & "FROM RV_POCertificato "
sSQL = sSQL & "WHERE IDLottoProduzione=" & Me.txtIDLottoCampagna.Value

Set rs = Cn.OpenResultset(sSQL)

If Not (rs.EOF) Then
    Me.txtQtaUtilizzataLotto.Value = fnNotNullN(rs!TotalePesoUtilizzato)
End If

rs.CloseResultset
Set rs = Nothing

If Me.txtIDLottoCampagna.Value > 0 Then
    Me.txtLottoDiConferimento.Enabled = False
End If
If (fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0) Then
    If (I = 1) Then
        Me.cboRegione.WriteOn IDRegioneLotto
    End If
    If (Me.txtResaMaxTotale.Value > 0) Then
        If (Me.txtQtaUtilizzataLotto.Value > Me.txtResaMaxTotale.Value) Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "La quantità utilizzato per il lotto selezionato è maggiore della resa massima consentita!" & vbCrLf
            Testo = Testo & "Vuoi continuare?"
            If (MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo) Then
                Me.txtIDLottoCampagna.Value = 0
            End If
        End If
    End If
    If Me.txtIDLottoCampagna.Value > 0 Then
        If (Me.CDArticolo.KeyFieldID > 0) Then
            If Me.cboVarietaArticolo.CurrentID <> Me.cboVarietaLotto.CurrentID Then
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "La varietà del lotto di produzione selezionato non è uguale alla varietà dell'articolo selezionato!" & vbCrLf
                Testo = Testo & "Vuoi continuare?"
                If (MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo) Then
                    Me.txtIDLottoCampagna.Value = 0
                End If
            End If
        End If
    End If
End If
If Not (BrwMain.Visible) Then Change
End Sub





Private Sub txtIDSocio_Change()
    Me.txtCodiceAnaSocio.Text = Me.CDSocio.Code
    Me.txtAnaSocio.Text = Me.CDSocio.Description
    
    Me.txtAnaSocio.Locked = Me.txtIDSocio > 0
    Me.txtCodiceAnaSocio.Locked = Me.txtIDSocio > 0
End Sub

Private Sub txtIndiceDiVariazioneEff_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtIndiceDiVariazioneEff_LostFocus()
    RECUPERA_INDICE
    
End Sub

Private Sub txtNumeroCertificato_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumeroCertificato_LostFocus()
On Error Resume Next
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        REFRESH_DESCR_ARTICOLO
    End If
End Sub

Private Sub txtNumeroDDT_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumeroDDT_LostFocus()
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        REFRESH_DESCR_ARTICOLO
    End If
End Sub

Private Sub txtOraTrasporto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPercRidPesoNetto_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPercRidPesoNetto_LostFocus()
    CALCOLA_PESI
End Sub

Private Sub txtPesoLordo_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPesoLordo_LostFocus()
    CALCOLA_PESI
    
End Sub

Private Sub txtPrezzoContrattoMax_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPrezzoContrattoMax_LostFocus()
    CALCOLA_TOTALE_RIGA
    
End Sub

Private Sub txtPrezzoContrattoMin_Change()
     If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPrezzoContrattoMin_LostFocus()
    CALCOLA_TOTALE_RIGA
   
End Sub

Private Sub txtPrezzoDaContratto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPrezzoDaContratto_LostFocus()
    CALCOLA_TOTALE_RIGA
    
End Sub

Private Sub txtPrezzoDiFatturazione_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPrezzoDiFatturazione_LostFocus()
Dim Testo As String

    CALCOLA_TOTALE_RIGA
    
    If (Me.txtPrezzoDiFatturazione.Value < Me.txtPrezzoContrattoMin.Value) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "L'importo unitario di fatturazione è minore del prezzo da contratto minimo"
        MsgBox Testo, vbInformation, "Controllo prezzo unitario"
    End If
    If (Me.txtPrezzoDiFatturazione.Value > Me.txtPrezzoContrattoMax.Value) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "L'importo unitario di fatturazione è maggiore del prezzo da contratto massimo"
        MsgBox Testo, vbInformation, "Controllo prezzo unitario"
    End If
End Sub

Private Sub txtScarto_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtScarto_LostFocus()
    CALCOLA_PESI
    
End Sub

Private Sub txtTaraCamion_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtTaraCamion_LostFocus()
    CALCOLA_PESI
   
End Sub

Private Sub txtTaraUnitaria_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtTaraUnitaria_LostFocus()
CALCOLA_PESI

End Sub
Private Sub RECUPERO_PAR_QUAL(IDAnagrafica As Long, IDAzienda As Long)
On Error GoTo ERR_RECUPERO_PAR_QUAL
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

txtQual01.Value = 0
txtQual02.Value = 0
txtQual03.Value = 0
txtQual04.Value = 0
txtQual05.Value = 0
txtQual06.Value = 0
txtQual07.Value = 0
txtQual08.Value = 0
txtQual09.Value = 0
txtQual10.Value = 0
txtQual11.Value = 0
txtQual12.Value = 0
txtQual13.Value = 0
txtQual14.Value = 0
txtQual15.Value = 0
txtQual16.Value = 0


sSQL = "SELECT * FROM RV_POParametriQualitaAnagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    txtQual01.Value = fnNotNullN(rs!Qualita01)
    txtQual02.Value = fnNotNullN(rs!Qualita02)
    txtQual03.Value = fnNotNullN(rs!Qualita03)
    txtQual04.Value = fnNotNullN(rs!Qualita04)
    txtQual05.Value = fnNotNullN(rs!Qualita05)
    txtQual06.Value = fnNotNullN(rs!Qualita06)
    txtQual07.Value = fnNotNullN(rs!Qualita07)
    txtQual08.Value = fnNotNullN(rs!Qualita08)
    txtQual09.Value = fnNotNullN(rs!Qualita09)
    txtQual10.Value = fnNotNullN(rs!Qualita10)
    txtQual11.Value = fnNotNullN(rs!Qualita11)
    txtQual12.Value = fnNotNullN(rs!Qualita12)
    txtQual13.Value = fnNotNullN(rs!Qualita13)
    txtQual14.Value = fnNotNullN(rs!Qualita14)
    txtQual15.Value = fnNotNullN(rs!Qualita15)
    txtQual16.Value = fnNotNullN(rs!Qualita16)
    
Else
    RECUPERO_PAR_QUAL_AZ IDAzienda
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_RECUPERO_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "RECUPERO_PAR_QUAL"
End Sub
Private Sub RECUPERO_PAR_QUAL_CONTR(IDOggetto As Long)
On Error GoTo ERR_RECUPERO_PAR_QUAL
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

txtQual01.Value = 0
txtQual02.Value = 0
txtQual03.Value = 0
txtQual04.Value = 0
txtQual05.Value = 0
txtQual06.Value = 0
txtQual07.Value = 0
txtQual08.Value = 0
txtQual09.Value = 0
txtQual10.Value = 0
txtQual11.Value = 0
txtQual12.Value = 0
txtQual13.Value = 0
txtQual14.Value = 0
txtQual15.Value = 0
txtQual16.Value = 0

sSQL = "SELECT * FROM RV_POParametriQualitaContratto "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    txtQual01.Value = fnNotNullN(rs!Qualita01)
    txtQual02.Value = fnNotNullN(rs!Qualita02)
    txtQual03.Value = fnNotNullN(rs!Qualita03)
    txtQual04.Value = fnNotNullN(rs!Qualita04)
    txtQual05.Value = fnNotNullN(rs!Qualita05)
    txtQual06.Value = fnNotNullN(rs!Qualita06)
    txtQual07.Value = fnNotNullN(rs!Qualita07)
    txtQual08.Value = fnNotNullN(rs!Qualita08)
    txtQual09.Value = fnNotNullN(rs!Qualita09)
    txtQual10.Value = fnNotNullN(rs!Qualita10)
    txtQual11.Value = fnNotNullN(rs!Qualita11)
    txtQual12.Value = fnNotNullN(rs!Qualita12)
    txtQual13.Value = fnNotNullN(rs!Qualita13)
    txtQual14.Value = fnNotNullN(rs!Qualita14)
    txtQual15.Value = fnNotNullN(rs!Qualita15)
    txtQual16.Value = fnNotNullN(rs!Qualita16)
    
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_RECUPERO_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "RECUPERO_PAR_QUAL"
End Sub

Private Sub RECUPERO_PAR_QUAL_AZ(IDAzienda As Long)
On Error GoTo ERR_RECUPERO_PAR_QUAL
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

txtQual01.Value = 0
txtQual02.Value = 0
txtQual03.Value = 0
txtQual04.Value = 0
txtQual05.Value = 0
txtQual06.Value = 0
txtQual07.Value = 0
txtQual08.Value = 0
txtQual09.Value = 0
txtQual10.Value = 0
txtQual11.Value = 0
txtQual12.Value = 0
txtQual13.Value = 0
txtQual14.Value = 0
txtQual15.Value = 0
txtQual16.Value = 0


sSQL = "SELECT * FROM RV_POParametriQualitaAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    txtQual01.Value = fnNotNullN(rs!Qualita01)
    txtQual02.Value = fnNotNullN(rs!Qualita02)
    txtQual03.Value = fnNotNullN(rs!Qualita03)
    txtQual04.Value = fnNotNullN(rs!Qualita04)
    txtQual05.Value = fnNotNullN(rs!Qualita05)
    txtQual06.Value = fnNotNullN(rs!Qualita06)
    txtQual07.Value = fnNotNullN(rs!Qualita07)
    txtQual08.Value = fnNotNullN(rs!Qualita08)
    txtQual09.Value = fnNotNullN(rs!Qualita09)
    txtQual10.Value = fnNotNullN(rs!Qualita10)
    txtQual11.Value = fnNotNullN(rs!Qualita11)
    txtQual12.Value = fnNotNullN(rs!Qualita12)
    txtQual13.Value = fnNotNullN(rs!Qualita13)
    txtQual14.Value = fnNotNullN(rs!Qualita14)
    txtQual15.Value = fnNotNullN(rs!Qualita15)
    txtQual16.Value = fnNotNullN(rs!Qualita16)
    
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_RECUPERO_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "RECUPERO_PAR_QUAL_AZ"
End Sub
Private Sub RECUPERA_INDICE()
Dim DeltaAssoluto As Double

DeltaAssoluto = Abs(Me.txtIndiceDiVariazioneEff.Value)
Me.txtIndice.Value = 8

If Me.txtIndiceDiVariazioneEff.Value > 0 Then
    If (Me.txtQual09.Value > DeltaAssoluto) Then
        Me.txtIndice.Value = 8
        Exit Sub

    End If
    If (Me.txtQual10.Value > DeltaAssoluto) Then
        Me.txtIndice.Value = 9
        Exit Sub
    End If
    If (Me.txtQual11.Value > DeltaAssoluto) Then
        Me.txtIndice.Value = 10
        Exit Sub

    End If
    If (Me.txtQual12.Value > DeltaAssoluto) Then
        Me.txtIndice.Value = 11
        Exit Sub

    End If
    If (Me.txtQual13.Value > DeltaAssoluto) Then
        Me.txtIndice.Value = 12
        Exit Sub

    End If
    If (Me.txtQual14.Value > DeltaAssoluto) Then
        Me.txtIndice.Value = 13
        Exit Sub

    End If
    If (Me.txtQual15.Value > DeltaAssoluto) Then
        Me.txtIndice.Value = 14
        Exit Sub
    End If
    If (Me.txtQual16.Value > DeltaAssoluto) Then
        Me.txtIndice.Value = 15
        Exit Sub
    End If
    
    If (Me.txtQual09.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 9
        Exit Sub

    End If
    If (Me.txtQual10.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 10
        Exit Sub
    End If
    If (Me.txtQual11.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 11
        Exit Sub

    End If
    If (Me.txtQual12.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 12
        Exit Sub

    End If
    If (Me.txtQual13.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 13
        Exit Sub

    End If
    If (Me.txtQual14.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 14
        Exit Sub

    End If
    If (Me.txtQual15.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 15
        Exit Sub
    End If

    If (DeltaAssoluto > Me.txtQual16.Value) Then
        Me.txtIndice.Value = 16
        Exit Sub
    End If
    
End If

If Me.txtIndiceDiVariazioneEff.Value < 0 Then
    If (Me.txtQual01.Value < DeltaAssoluto) Then
        Me.txtIndice.Value = 0
        Exit Sub
    End If
    If (Me.txtQual02.Value < DeltaAssoluto) Then
        Me.txtIndice.Value = 1
        Exit Sub
    End If
    If (Me.txtQual03.Value < DeltaAssoluto) Then
        Me.txtIndice.Value = 2
        Exit Sub
    End If
    If (Me.txtQual04.Value < DeltaAssoluto) Then
        Me.txtIndice.Value = 3
        Exit Sub
    End If
    If (Me.txtQual05.Value < DeltaAssoluto) Then
        Me.txtIndice.Value = 4
        Exit Sub
    End If
    If (Me.txtQual06.Value < DeltaAssoluto) Then
        Me.txtIndice.Value = 5
        Exit Sub
    End If
    If (Me.txtQual07.Value < DeltaAssoluto) Then
        Me.txtIndice.Value = 6
        Exit Sub
    End If
    
    If (Me.txtQual01.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 1
        Exit Sub
    End If
    If (Me.txtQual02.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 2
        Exit Sub
    End If
    If (Me.txtQual03.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 3
        Exit Sub
    End If
    If (Me.txtQual04.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 4
        Exit Sub
    End If
    If (Me.txtQual05.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 5
        Exit Sub
    End If
    If (Me.txtQual06.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 6
        Exit Sub
    End If
    If (Me.txtQual07.Value = DeltaAssoluto) Then
        Me.txtIndice.Value = 7
        Exit Sub
    End If
    If (DeltaAssoluto > txtQual01.Value) Then
        Me.txtIndice.Value = 0
        Exit Sub
    End If
    
End If

End Sub
Private Sub GET_SEZIONALE_DEFAULT()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale "
sSQL = sSQL & "FROM DefaultFilialePerTipoOggetto "
sSQL = sSQL & "WHERE (IDTipoOggetto = " & 2 & ") And (IDSezionale > 0) And (IDFiliale = " & TheApp.Branch & ")"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Me.CDSezionale.Load fnNotNullN(rs!IDSezionale)
Else
    Me.CDSezionale.Load 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub SALVA_PAR_QUAL(IDCertificato As Long)
On Error GoTo ERR_SALVA_PAR_QUAL
Dim sSQL As String
Dim rs As ADODB.Recordset

If (ELIMINA_PAR_QUAL(IDCertificato) = False) Then Exit Sub

sSQL = "SELECT * FROM RV_POParametriQualitaCertificato "
sSQL = sSQL & "WHERE IDRV_POCertificato=" & IDCertificato


Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDRV_POCertificato = IDCertificato
    rs!Qualita01 = txtQual01.Value
    rs!Qualita02 = txtQual02.Value
    rs!Qualita03 = txtQual03.Value
    rs!Qualita04 = txtQual04.Value
    rs!Qualita05 = txtQual05.Value
    rs!Qualita06 = txtQual06.Value
    rs!Qualita07 = txtQual07.Value
    rs!Qualita08 = txtQual08.Value
    rs!Qualita09 = txtQual09.Value
    rs!Qualita10 = txtQual10.Value
    rs!Qualita11 = txtQual11.Value
    rs!Qualita12 = txtQual12.Value
    rs!Qualita13 = txtQual13.Value
    rs!Qualita14 = txtQual14.Value
    rs!Qualita15 = txtQual15.Value
    rs!Qualita16 = txtQual16.Value
    'rs!QualitaPrezzo16 = txtQualPrz16.Value
rs.Update

rs.Close
Set rs = Nothing
Exit Sub
ERR_SALVA_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "SALVA_PAR_QUAL"
    
End Sub

Private Function ELIMINA_PAR_QUAL(IDCertificato As Long) As Boolean
On Error GoTo ERR_ELIMINA_PAR_QUAL
Dim sSQL As String
Dim rs As ADODB.Recordset

ELIMINA_PAR_QUAL = False

sSQL = "DELETE FROM RV_POParametriQualitaCertificato "
sSQL = sSQL & "WHERE IDRV_POCertificato=" & IDCertificato

Cn.Execute sSQL

ELIMINA_PAR_QUAL = True

Exit Function
ERR_ELIMINA_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "ELIMINA_PAR_QUAL"
End Function
Private Sub RECUPERO_PAR_QUAL_CERT(IDCertificato As Long)
On Error GoTo ERR_RECUPERO_PAR_QUAL
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

txtQual01.Value = 0
txtQual02.Value = 0
txtQual03.Value = 0
txtQual04.Value = 0
txtQual05.Value = 0
txtQual06.Value = 0
txtQual07.Value = 0
txtQual08.Value = 0
txtQual09.Value = 0
txtQual10.Value = 0
txtQual11.Value = 0
txtQual12.Value = 0
txtQual13.Value = 0
txtQual14.Value = 0
txtQual15.Value = 0
txtQual16.Value = 0


sSQL = "SELECT * FROM RV_POParametriQualitaCertificato "
sSQL = sSQL & "WHERE IDRV_POCertificato=" & IDCertificato

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    txtQual01.Value = fnNotNullN(rs!Qualita01)
    txtQual02.Value = fnNotNullN(rs!Qualita02)
    txtQual03.Value = fnNotNullN(rs!Qualita03)
    txtQual04.Value = fnNotNullN(rs!Qualita04)
    txtQual05.Value = fnNotNullN(rs!Qualita05)
    txtQual06.Value = fnNotNullN(rs!Qualita06)
    txtQual07.Value = fnNotNullN(rs!Qualita07)
    txtQual08.Value = fnNotNullN(rs!Qualita08)
    txtQual09.Value = fnNotNullN(rs!Qualita09)
    txtQual10.Value = fnNotNullN(rs!Qualita10)
    txtQual11.Value = fnNotNullN(rs!Qualita11)
    txtQual12.Value = fnNotNullN(rs!Qualita12)
    txtQual13.Value = fnNotNullN(rs!Qualita13)
    txtQual14.Value = fnNotNullN(rs!Qualita14)
    txtQual15.Value = fnNotNullN(rs!Qualita15)
    txtQual16.Value = fnNotNullN(rs!Qualita16)
    
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_RECUPERO_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "RECUPERO_PAR_QUAL_CERT"
End Sub

Private Sub GET_DOCUMENTO_COLLEGATO(IDCertificato As Long)
On Error GoTo ERR_GET_DOCUMENTO_COLLEGATO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

LINK_DOCUMENTO_COLLEGATO = 0
Me.txtDescrizioneDocumento.Text = "Nessun documento collegato!"

sSQL = "SELECT ValoriOggettoPerTipo0002.IDOggetto, ValoriOggettoPerTipo0002.IDTipoOggetto, Oggetto.Oggetto, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, Sezionale.Prefisso, Sezionale.Sezionale, "
sSQL = sSQL & "ValoriOggettoPerTipo0002.RV_POIDCertificato "
sSQL = sSQL & "FROM ValoriOggettoPerTipo0002 INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo0002.IDOggetto = Oggetto.IDOggetto INNER JOIN "
sSQL = sSQL & "Sezionale ON Oggetto.IDSezionale = Sezionale.IDSezionale "
sSQL = sSQL & " WHERE RV_POIDCertificato=" & IDCertificato

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    LINK_DOCUMENTO_COLLEGATO = fnNotNullN(rs!IDOggetto)
    Me.txtDescrizioneDocumento.Text = fnNotNull(rs!Oggetto) & " n. " & fnNotNullN(rs!Doc_numero) & " del " & fnNotNull(rs!Doc_data)
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_DOCUMENTO_COLLEGATO:
    MsgBox Err.Description, vbCritical, "GET_DOCUMENTO_COLLEGATO"
End Sub

Private Function CREA_DOCUMENTO() As Boolean
On Error GoTo ERR_CREA_DOCUMENTO

    CREA_DOCUMENTO = False
    
    If LINK_DOCUMENTO_COLLEGATO > 0 Then
        If (ELIMINA_DOCUMENTO = False) Then Exit Function
    End If
    
    SettaggioDocumento
    
    fncTestata
    
    fncRighe
    
    If (InserimentoDMT) Then
        LINK_DOCUMENTO_COLLEGATO = oDoc.IDOggetto
        CREA_DOCUMENTO = True
    End If
    
    GET_DOCUMENTO_COLLEGATO fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    Set oDoc = Nothing
    
    
Exit Function
ERR_CREA_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "CREA_DOCUMENTO"
End Function

Private Sub SettaggioDocumento()
On Error GoTo ERR_SettaggioDocumento
    If Not (oDoc Is Nothing) Then
        Set oDoc = Nothing
    End If
    
    If oDoc Is Nothing Then
        Set oDoc = New DmtDocs.cDocument
        
        With oDoc
            Set .Connection = Cn
            .SetTipoOggetto 2
            .IDFunzione = GET_LINK_FUNZIONE_DA_TIPO_OGGETTO(oDoc.IDTipoOggetto)
            .TablesNames oDoc.IDTipoOggetto, sTabellaTestata, sTabellaDettaglio, sTabellaIVA, sTabellaScadenze
            .IDAzienda = TheApp.IDFirm
            .IDFiliale = TheApp.Branch
            .IDAttivitaAzienda = GetAttivitaAzienda(TheApp.IDFirm, TheApp.Branch)
            .IDTipoAnagrafica = 2 'Cliente
            .IDUtente = TheApp.IDUser
            .descrizione = GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto)
            .UseAutomation = True
            .IDEsercizio = fncEsercizio(Me.txtDataDDT.Text)
            .IDSezionale = Me.CDSezionale.KeyFieldID
            '.DataEmissione = Me.txtDataDDT.Text
            .DataEmissione = Me.txtDataCertificato.Text
            .Numero = 0
            
            If .Tables.Count = 0 Then
                .Clear
                .SetTipoOggetto 2
            Else
                .ClearValues
            End If
    
        End With
    End If
Exit Sub
ERR_SettaggioDocumento:
    MsgBox Err.Description, vbCritical, "SettaggioDocumento"
End Sub
Private Function GET_LINK_FUNZIONE_DA_TIPO_OGGETTO(IDTipoOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FUNZIONE_DA_TIPO_OGGETTO = 0
Else
    GET_LINK_FUNZIONE_DA_TIPO_OGGETTO = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GetAttivitaAzienda(IDAzienda As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivitaAzienda.IDAttivitaAzienda, Azienda.IDAzienda, Filiale.IDFiliale "
sSQL = sSQL & "FROM AttivitaAzienda INNER JOIN "
sSQL = sSQL & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda INNER JOIN "
sSQL = sSQL & "Filiale ON AttivitaAzienda.IDAttivitaAzienda = Filiale.IDAttivitaAzienda "
sSQL = sSQL & "WHERE (Azienda.IDAzienda =" & IDAzienda & ") And (Filiale.IDFiliale = " & IDFiliale & ")"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GetAttivitaAzienda = 0
Else
    GetAttivitaAzienda = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_DESCRIZIONE_TIPOOGGETTO(IDTipoOggetto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select Oggetto "
    sSQL = sSQL & "FROM TipoOggetto "
    sSQL = sSQL & "WHERE IDTipoOggetto = " & IDTipoOggetto
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_DESCRIZIONE_TIPOOGGETTO = fnNotNull(rs!Oggetto)
    Else
        GET_DESCRIZIONE_TIPOOGGETTO = ""
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Function CONTROLLO_STATO_DOCUMENTO() As Boolean
On Error GoTo ERR_CONTROLLO_STATO_DOCUMENTO
    
    CONTROLLO_STATO_DOCUMENTO = False
    
    SettaggioDocumento

    oDoc.ReadWithTO LINK_DOCUMENTO_COLLEGATO, 2
    
    CONTROLLO_STATO_DOCUMENTO = oDoc.IsLocked
    
Exit Function
ERR_CONTROLLO_STATO_DOCUMENTO:
        
End Function
Private Function ELIMINA_DOCUMENTO() As Boolean
On Error GoTo ERR_CONTROLLO_STATO_DOCUMENTO
    
    ELIMINA_DOCUMENTO = False
    
    SettaggioDocumento

    oDoc.ReadWithTO LINK_DOCUMENTO_COLLEGATO, 2
    
    oDoc.DeleteWithTO LINK_DOCUMENTO_COLLEGATO, 2
    
    ELIMINA_DOCUMENTO = True
    
Exit Function
ERR_CONTROLLO_STATO_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "ELIMINA_DOCUMENTO"
End Function
Private Function fncTestata() As Boolean
'On Error GoTo ERR_fncTestata
Dim IDListinoDefault As Long
Dim Link_Pagamento As Long
Dim Link_Valuta_Cliente As Long


         
         With oDoc.Tables
        
            'Imposta la riga attiva per la tabella di testata
            
            oDoc.Tables(sTabellaTestata).SetActiveRetail 1
            
            oDoc.ReadDataFromCliFo Me.ACSCliente.IDAnagrafica, sTabellaTestata
            
            oDoc.ReadDataFromCliFoSite Me.cboAltroSito.CurrentID, sTabellaTestata
            
            If Me.cboVettore.CurrentID > 0 Then
                oDoc.Field "Link_Doc_spedizione", 3, sTabellaTestata
                oDoc.ReadDataFromCarrier Me.cboVettore.CurrentID, MainCarrier, sTabellaTestata
            End If
                        
            .Field "Doc_causale_trasporto", fnSetCausaleDocumento, sTabellaTestata
            .Field "Link_Doc_Magazzino", LINK_MAGAZZINO_DOCUMENTO, sTabellaTestata
            .Field "Link_Doc_sezionale", Me.CDSezionale.KeyFieldID, sTabellaTestata
            .Field "Doc_prefisso", GET_PREFISSO_SEZ(Me.CDSezionale.KeyFieldID), sTabellaTestata
            .Field "Doc_data", oDoc.DataEmissione, sTabellaTestata
            .Field "Doc_data_inizio_trasporto", oDoc.DataEmissione, sTabellaTestata
            .Field "Doc_ora_inizio_trasporto", time, sTabellaTestata
            .Field "Doc_crea_scadenze", fnNormBoolean(1), sTabellaTestata
            .Field "RV_PODataCompetenzaLiq", oDoc.DataEmissione, sTabellaTestata
            
            'PAGAMENTO DOPO I DATI DELL'ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             If fnNotNullN(.Field("Link_Doc_pagamento", , sTabellaTestata)) = 0 Then
                
                oDoc.ReadDataFromPayment oDoc.DBDefaults.IDPagamentoDocDefault
                
                If fnNotNullN(.Field("Link_Doc_pagamento", , sTabellaTestata)) = 0 Then
                    oDoc.ReadDataFromPayment oDoc.DBDefaults.IDPagamentoDocDefault
                End If
             
             End If
            
            'VALUTA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Field("Link_Val_valuta", , sTabellaTestata) = 0 Then
               .Field "Link_Val_valuta", oDoc.DBDefaults.Link_Val_valuta_nazionale, sTabellaTestata
            End If
            

            
            .Field "Doc_data_plafond", oDoc.DataEmissione, sTabellaTestata
            .Field "Link_Spe_esenti_art_10_IVA", oDoc.DBDefaults.Link_Spe_esenti_art_10_IVA, sTabellaTestata
            .Field "Link_Spe_bolli_eff_art_15_IVA", oDoc.DBDefaults.Link_Spe_bolli_eff_art_15_IVA, sTabellaTestata
            
            .Field "RV_POIDCertificato", fnNotNullN(m_Document(m_Document.PrimaryKey).Value), sTabellaTestata
            If (Me.CDSocioFatt.KeyFieldID > 0) Then
                .Field "RV_POIDAnagraficaDestinazione", Me.CDSocioFatt.KeyFieldID, sTabellaTestata
            Else
                .Field "RV_POIDAnagraficaDestinazione", IDAnagraficaDestSociDiretti, sTabellaTestata
            End If
            .Field "RV_POIDAnagraficaSocio", Me.CDSocio.KeyFieldID, sTabellaTestata
            
            oDoc.Field "Link_nom_lettera_intento", Me.txtIDLetteraIntento.Value, sTabellaTestata
            oDoc.ReadIvaFromLetter Me.txtIDLetteraIntento.Value
            oDoc.Field "Link_Nom_IVA", Me.cboIvaCliente.CurrentID, sTabellaTestata
            
            oDoc.Field "Doc_data_ns_ordine_di_rifer", Me.txtDataDDT.Text, sTabellaTestata
            oDoc.Field "Doc_numero_ns_ordine_di_rifer", Me.txtNumeroDDT.Text, sTabellaTestata
            
            If (NonRiportareRifCerticatoInDDT = 0) Then
                oDoc.Field "Doc_data_vs_ordine_di_rifer", Me.txtDataCertificato.Text, sTabellaTestata
                oDoc.Field "Doc_numero_vs_ordine_di_rifer", Me.txtNumeroCertificato.Text, sTabellaTestata
            End If
            
            oDoc.Field "Link_Nom_raggrup_fatturato", GET_RAGGRUPPAMENTO_FATTURATO, sTabellaTestata
            
        End With
        
        fncTestata = True
     
Exit Function
ERR_fncTestata:
    fncTestata = False
    
    
End Function

Private Function GET_RAGGRUPPAMENTO_FATTURATO() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim DescrizioneFatturato As String
Dim IDNuovoRaggruppamentoFatturato As Long

GET_RAGGRUPPAMENTO_FATTURATO = 0
IDNuovoRaggruppamentoFatturato = 0

DescrizioneFatturato = "Conferito"
If (Me.Check1.Value = vbChecked) Then
    DescrizioneFatturato = "Acquistato"
End If

sSQL = "SELECT * FROM RaggruppamentoFatturato"
sSQL = sSQL & " WHERE RaggruppamentoFatturato=" & fnNormString(DescrizioneFatturato)

Set rs = Cn.OpenResultset(sSQL)
If Not rs.EOF Then
    IDNuovoRaggruppamentoFatturato = fnNotNullN(rs!IDRaggruppamentoFatturato)
End If

rs.CloseResultset
Set rs = Nothing

If (IDNuovoRaggruppamentoFatturato = 0) Then
    IDNuovoRaggruppamentoFatturato = fnGetNewKey("RaggruppamentoFatturato", "IDRaggruppamentoFatturato")
    
    sSQL = "INSERT INTO RaggruppamentoFatturato (IDRaggruppamentoFatturato, RaggruppamentoFatturato)"
    sSQL = sSQL & " VALUES ("
    sSQL = sSQL & IDNuovoRaggruppamentoFatturato & ", "
    sSQL = sSQL & fnNormString(DescrizioneFatturato)
    sSQL = sSQL & ")"
    
    Cn.Execute sSQL
End If

GET_RAGGRUPPAMENTO_FATTURATO = IDNuovoRaggruppamentoFatturato


End Function

Private Function fnSetCausaleDocumento()
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    fnSetCausaleDocumento = ""
    
    sSQL = "SELECT CausaleTrasporto FROM CausaleTrasportoPerFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & oDoc.IDFunzione
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        fnSetCausaleDocumento = ""
    Else
        fnSetCausaleDocumento = fnNotNull(rs!CausaleTrasporto)
    End If
    
    
    rs.CloseResultset
    Set rs = Nothing
    
End Function
Private Function GET_PREFISSO_SEZ(IDSezionale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Prefisso FROM Sezionale "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDSezionale=" & IDSezionale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREFISSO_SEZ = ""
Else
    GET_PREFISSO_SEZ = Trim(fnNotNull(rs!Prefisso))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Function fncEsercizio(DataDocumento As String) As Long

    fncEsercizio = 0
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select IDEsercizio, Esercizio "
    sSQL = sSQL & " FROM Esercizio "
    sSQL = sSQL & " WHERE (IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND DataInizio<=" & fnNormDate(DataDocumento)
    sSQL = sSQL & " AND DataFine>=" & fnNormDate(DataDocumento)
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fncEsercizio = fnNotNullN(rs!IDEsercizio)
    Else
        fncEsercizio = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function fncRighe() As Boolean
On Error GoTo ERR_fncRighe
Dim I As Integer
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ProgressivoArticolo As Long
Dim Link_Riga As Long
Dim ImportoPedana As Double
Dim IDListinoDefault As Long
Dim ImportoLiquidazione As Double
Dim Sconto1 As Double
Dim Sconto2 As Double
Dim ImportoImballo As Double
Dim DescrizioneRigaDaOrdine As String
Dim Link_IVA_Articolo_Riga As Long
Dim Aliquota_IVA_Articolo_Riga As Long
Dim ImballoARendere As Long
Dim rsImballiANoleggio As ADODB.Recordset
Dim ImportoUnitarioArticoloMerceNetta As Double
Dim ImportoUnitarioArticoloMerceImballo As Double
Dim LINK_UM_COOP As Long
Dim LINK_REGOLA_PROVV As Long
Dim LINK_UM_LIQ As Long

    I = 1
    Link_Riga = 1
    ProgressivoArticolo = 0
        
    oDoc.Tables(sTabellaDettaglio).SetActiveRetail I
        
    Aliquota_IVA_Articolo_Riga = GET_ALIQUOTA_IVA_ARTICOLO(Me.cboIvaArticolo.CurrentID)
    LINK_UM_LIQ = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(Me.CDArticolo.KeyFieldID)
    
    oDoc.Field "Link_Art_articolo", Me.CDArticolo.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Art_codice", Me.CDArticolo.Code, sTabellaDettaglio
    oDoc.Field "Art_descrizione", Me.txtDescrizioneArticolo.Text, sTabellaDettaglio
    oDoc.Field "Art_quantita_totale", Me.txtQtaFatturazione.Value, sTabellaDettaglio

    oDoc.Field "Art_sco_in_percentuale_1", 0, sTabellaDettaglio
    oDoc.Field "Art_sco_in_percentuale_2", 0, sTabellaDettaglio
    
    oDoc.Field "Art_importo_totale_lordo_IVA", (Me.txtPrezzoDiFatturazione.Value * Me.txtQtaFatturazione.Value) + ((((Me.txtPrezzoDiFatturazione.Value * Me.txtQtaFatturazione.Value)) / 100) * Aliquota_IVA_Articolo_Riga), sTabellaDettaglio
    oDoc.Field "Art_importo_totale_netto_IVA", Me.txtPrezzoDiFatturazione.Value * Me.txtQtaFatturazione.Value, sTabellaDettaglio
    oDoc.Field "Art_prezzo_unitario_netto_IVA", Me.txtPrezzoDiFatturazione.Value, sTabellaDettaglio
    oDoc.Field "Art_prezzo_unitario_lordo_IVA", Me.txtPrezzoDiFatturazione.Value + ((Me.txtPrezzoDiFatturazione.Value / 100) * Aliquota_IVA_Articolo_Riga), sTabellaDettaglio
    oDoc.Field "Art_pre_uni_net_sco_net_IVA", Me.txtPrezzoDiFatturazione.Value, sTabellaDettaglio
    oDoc.Field "Art_pre_uni_net_sco_lor_IVA", Me.txtPrezzoDiFatturazione.Value + ((Me.txtPrezzoDiFatturazione.Value / 100) * Aliquota_IVA_Articolo_Riga), sTabellaDettaglio
    oDoc.Field "Art_Importo_totale_neutro", Me.txtPrezzoDiFatturazione.Value * Me.txtQtaFatturazione.Value, sTabellaDettaglio
    oDoc.Field "Art_prezzo_unitario_neutro", Me.txtPrezzoDiFatturazione.Value, sTabellaDettaglio
    oDoc.Field "Art_Importo_netto_IVA", Me.txtPrezzoDiFatturazione.Value, sTabellaDettaglio
    oDoc.Field "Art_importo_net_sconto_lor_IVA", Me.txtPrezzoDiFatturazione.Value + ((Me.txtPrezzoDiFatturazione.Value / 100) * Aliquota_IVA_Articolo_Riga), sTabellaDettaglio
    oDoc.Field "Art_importo_net_sconto_net_IVA", Me.txtPrezzoDiFatturazione.Value * Me.txtQtaFatturazione.Value, sTabellaDettaglio
    
    oDoc.Field "Link_Art_Magazzino", LINK_MAGAZZINO_DOCUMENTO, sTabellaDettaglio
    oDoc.Field "Link_art_IVA", Me.cboIvaArticolo.CurrentID, sTabellaDettaglio
    oDoc.Field "Art_aliquota_IVA", Aliquota_IVA_Articolo_Riga, sTabellaDettaglio
    
    oDoc.Field "Art_numero_colli", Me.txtColliEntrata.Value, sTabellaDettaglio
    If (Me.txtColliEntrata.Value = 0) Then
        oDoc.Field "Art_numero_colli", 1, sTabellaDettaglio
    End If
    oDoc.Field "Art_Peso", Me.txtPesoLordo.Value - Me.txtScarto.Value, sTabellaDettaglio
    oDoc.Field "Art_tara", Me.txtTaraTotale, sTabellaDettaglio
    oDoc.Field "Art_quantita_pezzi", 0, sTabellaDettaglio
            
    oDoc.Field "Link_Art_unita_di_misura", GET_LINK_UM_ART(Me.CDArticolo.KeyFieldID), sTabellaDettaglio
    oDoc.Field "Art_sigla_unita_di_misura", GET_SIGLA_UM(fnNotNullN(oDoc.Field("Link_Art_unita_di_misura", , sTabellaDettaglio))), sTabellaDettaglio
    
    LINK_UM_COOP = fnGetUMCoop(oDoc.Field("Link_Art_unita_di_misura", , sTabellaDettaglio))
        
    oDoc.Field "RV_POLinkRiga", Link_Riga, sTabellaDettaglio
    oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
    
    oDoc.Field "RV_PODataConferimento", Me.txtDataDDT.Value, sTabellaDettaglio
    
    oDoc.Field "RV_POIDSocio", Me.CDSocio.KeyFieldID, sTabellaDettaglio
    If (Me.CDSocioFatt.KeyFieldID > 0) Then
        oDoc.Field "RV_POIDAnagraficaFatturazione", Me.CDSocioFatt.KeyFieldID, sTabellaDettaglio
    Else
        oDoc.Field "RV_POIDAnagraficaFatturazione", IDAnagraficaDestSociDiretti, sTabellaDettaglio
    End If
    oDoc.Field "RV_POCodiceSocio", Me.CDSocio.Code, sTabellaDettaglio
    oDoc.Field "RV_POSocio", Me.CDSocio.Description, sTabellaDettaglio
    oDoc.Field "RV_POLottoCampagna", Me.txtLottoDiConferimento.Text, sTabellaDettaglio
    oDoc.Field "RV_POCodiceLotto", Me.txtLottoDiConferimento.Text, sTabellaDettaglio
    oDoc.Field "RV_POImportoImballoInArticolo", 0, sTabellaDettaglio
    oDoc.Field "RV_PODataLavorazione", Me.txtDataDDT.Value, sTabellaDettaglio
    oDoc.Field "RV_POVariazionePrezzoManuale", 0, sTabellaDettaglio
    oDoc.Field "RV_POImportoUnitarioListino", 0, sTabellaDettaglio
    oDoc.Field "RV_POImportoImballoSel", 0, sTabellaDettaglio
    
    
    oDoc.Field "RV_PODataOrdineCliente", Me.txtDataCertificato.Value, sTabellaDettaglio
    oDoc.Field "RV_PONumeroOrdineCliente", Me.txtNumeroCertificato.Text, sTabellaDettaglio
    oDoc.Field "RV_PODataOrdineInterno", Me.txtDataDDT.Value, sTabellaDettaglio
    oDoc.Field "RV_PONumeroOrdineInterno", Me.txtNumeroDDT.Text, sTabellaDettaglio
        
    oDoc.Field "RV_POIDTipoImportoVenditaLiq", GET_FORZATURA_PREZZO_LIQ_CLIENTE(Me.ACSCliente.IDAnagrafica, Me.CDArticolo.KeyFieldID), sTabellaDettaglio
    If (Me.Check1.Value = vbUnchecked) Then
        oDoc.Field "RV_POIDTipoDocumentoCoop", 1, sTabellaDettaglio
    Else
        oDoc.Field "RV_POIDTipoDocumentoCoop", 2, sTabellaDettaglio
    End If
    oDoc.Field "RV_POQuantitaLiq", Me.txtQtaFatturazione.Value, sTabellaDettaglio
    If (LINK_UM_LIQ > 0) Then
        Select Case LINK_UM_LIQ
            Case 1
                oDoc.Field "RV_POQuantitaLiq", txtColliEntrata.Value, sTabellaDettaglio
            Case 2
                oDoc.Field "RV_POQuantitaLiq", txtPesoLordo.Value, sTabellaDettaglio
            Case 3
                oDoc.Field "RV_POQuantitaLiq", txtQtaFatturazione.Value, sTabellaDettaglio
            Case 4
                oDoc.Field "RV_POQuantitaLiq", txtTaraTotale.Value, sTabellaDettaglio
            Case 5
                oDoc.Field "RV_POQuantitaLiq", 0, sTabellaDettaglio
        End Select
    End If
    
    oDoc.Field "RV_POImportoDaLiq", 0, sTabellaDettaglio
    ImportoUnitarioArticoloMerceNetta = Me.txtPrezzoDiFatturazione.Value
    oDoc.Field "RV_POImportoMerceNetta", ImportoUnitarioArticoloMerceNetta, sTabellaDettaglio
    oDoc.Field "RV_POVariazionePrezzoImballo", 0, sTabellaDettaglio
    
        
    oDoc.Field "ID_Art_dettaglio_prog", oDoc.SetIDArtDettaglioProg, sTabellaDettaglio
    oDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(Me.CDArticolo.KeyFieldID, Me.ACSCliente.IDAnagrafica, Me.cboAltroSito.CurrentID), sTabellaDettaglio
    oDoc.Field "RV_POIDLottoCampagnaLavorazione", Me.txtIDLottoCampagna.Value, sTabellaDettaglio
        
    If (oDoc.IDTipoOggetto <> 8) Then
        sbLoadElectronicInvoiceData4Article fnNotNullN(oDoc.Field("ID_Art_dettaglio_prog", , sTabellaDettaglio)), Me.CDArticolo.KeyFieldID
    End If
        

    oDoc.Field "RV_PORigaCompleta", 0, sTabellaDettaglio


    Link_Riga = Link_Riga + 1
    I = I + 1
    
    
    If (IDArticoloScartoPerCertificato > 0 And Me.txtScarto.Value > 0) Then
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail I
            
        Aliquota_IVA_Articolo_Riga = GET_ALIQUOTA_IVA_ARTICOLO(Me.cboIvaArticolo.CurrentID)
        LINK_UM_LIQ = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(IDArticoloScartoPerCertificato)
        
        oDoc.Field "Link_Art_articolo", IDArticoloScartoPerCertificato, sTabellaDettaglio
        oDoc.Field "Art_codice", GET_PROPRIETA_ARTICOLO(IDArticoloScartoPerCertificato, "CodiceArticolo"), sTabellaDettaglio
        oDoc.Field "Art_descrizione", GET_PROPRIETA_ARTICOLO(IDArticoloScartoPerCertificato, "Articolo"), sTabellaDettaglio
        oDoc.Field "Art_quantita_totale", Me.txtScarto.Value, sTabellaDettaglio
    
        oDoc.Field "Art_sco_in_percentuale_1", 0, sTabellaDettaglio
        oDoc.Field "Art_sco_in_percentuale_2", 0, sTabellaDettaglio
        
        oDoc.Field "Art_importo_totale_lordo_IVA", (0 * Me.txtScarto.Value) + ((((0 * Me.txtScarto.Value)) / 100) * Aliquota_IVA_Articolo_Riga), sTabellaDettaglio
        oDoc.Field "Art_importo_totale_netto_IVA", 0 * Me.txtScarto.Value, sTabellaDettaglio
        oDoc.Field "Art_prezzo_unitario_netto_IVA", 0, sTabellaDettaglio
        oDoc.Field "Art_prezzo_unitario_lordo_IVA", 0 + ((0 / 100) * Aliquota_IVA_Articolo_Riga), sTabellaDettaglio
        oDoc.Field "Art_pre_uni_net_sco_net_IVA", 0, sTabellaDettaglio
        oDoc.Field "Art_pre_uni_net_sco_lor_IVA", 0 + ((0 / 100) * Aliquota_IVA_Articolo_Riga), sTabellaDettaglio
        oDoc.Field "Art_Importo_totale_neutro", 0 * Me.txtScarto.Value, sTabellaDettaglio
        oDoc.Field "Art_prezzo_unitario_neutro", 0, sTabellaDettaglio
        oDoc.Field "Art_Importo_netto_IVA", 0, sTabellaDettaglio
        oDoc.Field "Art_importo_net_sconto_lor_IVA", 0 + ((0 / 100) * Aliquota_IVA_Articolo_Riga), sTabellaDettaglio
        oDoc.Field "Art_importo_net_sconto_net_IVA", 0 * Me.txtScarto.Value, sTabellaDettaglio
        
        oDoc.Field "Link_Art_Magazzino", LINK_MAGAZZINO_DOCUMENTO, sTabellaDettaglio
        oDoc.Field "Link_art_IVA", Me.cboIvaArticolo.CurrentID, sTabellaDettaglio
        oDoc.Field "Art_aliquota_IVA", Aliquota_IVA_Articolo_Riga, sTabellaDettaglio
        
        oDoc.Field "Art_numero_colli", 1, sTabellaDettaglio
        oDoc.Field "Art_Peso", Me.txtScarto.Value, sTabellaDettaglio
        oDoc.Field "Art_tara", 0, sTabellaDettaglio
        oDoc.Field "Art_quantita_pezzi", 0, sTabellaDettaglio
                
        oDoc.Field "Link_Art_unita_di_misura", GET_LINK_UM_ART(IDArticoloScartoPerCertificato), sTabellaDettaglio
        oDoc.Field "Art_sigla_unita_di_misura", GET_SIGLA_UM(fnNotNullN(oDoc.Field("Link_Art_unita_di_misura", , sTabellaDettaglio))), sTabellaDettaglio
        
        LINK_UM_COOP = fnGetUMCoop(oDoc.Field("Link_Art_unita_di_misura", , sTabellaDettaglio))
            
        oDoc.Field "RV_POLinkRiga", Link_Riga, sTabellaDettaglio
        oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
        
        oDoc.Field "RV_PODataConferimento", Me.txtDataDDT.Value, sTabellaDettaglio
        
        oDoc.Field "RV_POIDSocio", Me.CDSocio.KeyFieldID, sTabellaDettaglio
        If (Me.CDSocioFatt.KeyFieldID > 0) Then
            oDoc.Field "RV_POIDAnagraficaFatturazione", Me.CDSocioFatt.KeyFieldID, sTabellaDettaglio
        Else
            oDoc.Field "RV_POIDAnagraficaFatturazione", IDAnagraficaDestSociDiretti, sTabellaDettaglio
        End If
        oDoc.Field "RV_POCodiceSocio", Me.CDSocio.Code, sTabellaDettaglio
        oDoc.Field "RV_POSocio", Me.CDSocio.Description, sTabellaDettaglio
        oDoc.Field "RV_POLottoCampagna", Me.txtLottoDiConferimento.Text, sTabellaDettaglio
        oDoc.Field "RV_POCodiceLotto", Me.txtLottoDiConferimento.Text, sTabellaDettaglio
        oDoc.Field "RV_POImportoImballoInArticolo", 0, sTabellaDettaglio
        oDoc.Field "RV_PODataLavorazione", Me.txtDataDDT.Value, sTabellaDettaglio
        oDoc.Field "RV_POVariazionePrezzoManuale", 0, sTabellaDettaglio
        oDoc.Field "RV_POImportoUnitarioListino", 0, sTabellaDettaglio
        oDoc.Field "RV_POImportoImballoSel", 0, sTabellaDettaglio
        
        oDoc.Field "RV_PODataOrdineCliente", Me.txtDataCertificato.Value, sTabellaDettaglio
        oDoc.Field "RV_PONumeroOrdineCliente", Me.txtNumeroCertificato.Text, sTabellaDettaglio
        oDoc.Field "RV_PODataOrdineInterno", Me.txtDataDDT.Value, sTabellaDettaglio
        oDoc.Field "RV_PONumeroOrdineInterno", Me.txtNumeroDDT.Text, sTabellaDettaglio
            
        oDoc.Field "RV_POIDTipoImportoVenditaLiq", GET_FORZATURA_PREZZO_LIQ_CLIENTE(Me.ACSCliente.IDAnagrafica, Me.CDArticolo.KeyFieldID), sTabellaDettaglio
        If (Me.Check1.Value = vbUnchecked) Then
            oDoc.Field "RV_POIDTipoDocumentoCoop", 1, sTabellaDettaglio
        Else
            oDoc.Field "RV_POIDTipoDocumentoCoop", 2, sTabellaDettaglio
        End If
        oDoc.Field "RV_POQuantitaLiq", Me.txtScarto.Value, sTabellaDettaglio
        oDoc.Field "RV_POImportoDaLiq", 0, sTabellaDettaglio
        ImportoUnitarioArticoloMerceNetta = 0
        oDoc.Field "RV_POImportoMerceNetta", ImportoUnitarioArticoloMerceNetta, sTabellaDettaglio
        oDoc.Field "RV_POVariazionePrezzoImballo", 0, sTabellaDettaglio
            
        oDoc.Field "ID_Art_dettaglio_prog", oDoc.SetIDArtDettaglioProg, sTabellaDettaglio
        oDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(IDArticoloScartoPerCertificato, Me.ACSCliente.IDAnagrafica, Me.cboAltroSito.CurrentID), sTabellaDettaglio
        oDoc.Field "RV_POIDLottoCampagnaLavorazione", Me.txtIDLottoCampagna.Value, sTabellaDettaglio
            
        If (oDoc.IDTipoOggetto <> 8) Then
            sbLoadElectronicInvoiceData4Article fnNotNullN(oDoc.Field("ID_Art_dettaglio_prog", , sTabellaDettaglio)), IDArticoloScartoPerCertificato
        End If
    
        oDoc.Field "RV_PORigaCompleta", 0, sTabellaDettaglio
    End If
    
fncRighe = True
Exit Function
ERR_fncRighe:
    fncRighe = False


End Function
Private Function GET_PROPRIETA_ARTICOLO(IDArticolo As Long, NomeProprieta As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " + NomeProprieta + " FROM Articolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo


Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_PROPRIETA_ARTICOLO = ""
Else
    GET_PROPRIETA_ARTICOLO = fnNotNull(rs.adoColumns(NomeProprieta).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ALIQUOTA_IVA_ARTICOLO(IDIva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AliquotaIva FROM Iva "
sSQL = sSQL & " WHERE IDIva=" & IDIva


Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_ALIQUOTA_IVA_ARTICOLO = 0
Else
    GET_ALIQUOTA_IVA_ARTICOLO = fnNotNullN(rs!AliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PREZZO_MEDIO_CLIENTE(IDAnagrafica As Long, IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim EsisteDettaglio As Boolean

EsisteDettaglio = False

sSQL = "SELECT IDArticolo, RV_PONonPartecipaPrezzoMedio "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!RV_PONonPartecipaPrezzoMedio) = 1 Then
        GET_PREZZO_MEDIO_CLIENTE = 0
    Else
        GET_PREZZO_MEDIO_CLIENTE = 1
    End If
    
    EsisteDettaglio = True
End If

rs.CloseResultset
Set rs = Nothing


If EsisteDettaglio = True Then Exit Function

sSQL = "SELECT NonCalcolarePrezzoMedio "
sSQL = sSQL & " FROM RV_POConfigurazioneClienteArtVend"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!NonCalcolarePrezzoMedio) = 1 Then
        GET_PREZZO_MEDIO_CLIENTE = 0
    Else
        GET_PREZZO_MEDIO_CLIENTE = 1
    End If
    
    EsisteDettaglio = True
End If

rs.CloseResultset
Set rs = Nothing

If EsisteDettaglio = True Then Exit Function

sSQL = "SELECT NonCalcolarePrezzoMedio "
sSQL = sSQL & " FROM RV_POConfigurazioneCliente"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_MEDIO_CLIENTE = 1
Else
    If fnNotNullN(rs!NonCalcolarePrezzoMedio) = 1 Then
        GET_PREZZO_MEDIO_CLIENTE = 0
    Else
        GET_PREZZO_MEDIO_CLIENTE = 1
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_FORZATURA_PREZZO_LIQ_CLIENTE(IDAnagrafica As Long, IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim EsisteDettaglio As Boolean

EsisteDettaglio = False


sSQL = "SELECT IDRV_POTipoImportoVenditaLiq "
sSQL = sSQL & " FROM RV_POConfigurazioneClienteArtVend"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_FORZATURA_PREZZO_LIQ_CLIENTE = fnNotNullN(rs!IDRV_POTipoImportoVenditaLiq)
    
    EsisteDettaglio = True
End If

rs.CloseResultset
Set rs = Nothing

If EsisteDettaglio = True Then Exit Function


sSQL = "SELECT IDRV_POTipoImportoVenditaLiq "
sSQL = sSQL & " FROM RV_POConfigurazioneCliente"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_FORZATURA_PREZZO_LIQ_CLIENTE = 0
Else
    GET_FORZATURA_PREZZO_LIQ_CLIENTE = fnNotNullN(rs!IDRV_POTipoImportoVenditaLiq)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_RIF_PA_ARTICOLO(IDArticolo As Long, IDCliente As Long, IDDestinazione As Long) As String
On Error GoTo ERR_GET_RIF_PA_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_RIF_PA_ARTICOLO = ""

'Articolo - Cliente
sSQL = "SELECT RiferimentoPACliente FROM ClientePerArticolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAnagrafica=" & IDCliente

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPACliente)
End If
rs.CloseResultset
Set rs = Nothing

If Len(Trim(GET_RIF_PA_ARTICOLO)) > 0 Then Exit Function

'Destinazione
sSQL = "SELECT RiferimentoPAArticolo FROM SitoPerAnagrafica "
sSQL = sSQL & " WHERE IDSitoPerAnagrafica=" & IDDestinazione

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPAArticolo)
End If
rs.CloseResultset
Set rs = Nothing


If Len(Trim(GET_RIF_PA_ARTICOLO)) > 0 Then Exit Function
'Cliente
sSQL = "SELECT RiferimentoPAArticolo FROM Cliente "
sSQL = sSQL & " WHERE IDAnagrafica=" & IDCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPAArticolo)
End If
rs.CloseResultset
Set rs = Nothing

If Len(Trim(GET_RIF_PA_ARTICOLO)) > 0 Then Exit Function

sSQL = "SELECT RiferimentoPA FROM Articolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPA)
End If
rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_RIF_PA_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_RIF_PA_ARTICOLO"
End Function
Private Function GET_LINK_UM_ART(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_UM_ART = 0

sSQL = "SELECT IDUnitaDiMisuraVendita FROM Articolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_ART = fnNotNullN(rs!IDUnitaDiMisuraVendita)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_SIGLA_UM(IDUM As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoListinoImballo As Double

GET_SIGLA_UM = ""


sSQL = "SELECT * FROM UnitaDiMisura "
sSQL = sSQL & " WHERE IDUnitaDiMisura = " & IDUM

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SIGLA_UM = fnNotNull(rs!DescrizioneFattura)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
sSQL = sSQL & "IDUnitaDiMisura = " & Link_UMAcq

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF = False Then
    fnGetUMCoop = rs!RV_POIDUnitaDiMisuraCoop
Else
    fnGetUMCoop = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function InserimentoDMT() As Boolean
On Error GoTo ERR_InserimentoDMT
Dim VarNumeroDoc As String
Dim Link_Oggetto As Long
Dim sSQL As String
Dim TestoMessaggio As String
    
    InserimentoDMT = False

    Screen.MousePointer = vbHourglass
    
    If oDoc.IDTipoOggetto <> 8 Then
        ConsolidaDettaglioFatturaElettronica
    End If
   
    Set oDoc.Scadenze = Nothing
    oDoc.PerformDocument Nothing
    
    oDoc.AllowCreateMovements = True
    
    'CONTROLLO PLAFOND'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim sMsgPlafond As String
    If oDoc.PlafondExceed Then
        sMsgPlafond = oDoc.PlafondLastMessage
        If Len(sMsgPlafond) > 0 Then
            If oDoc.PlafondLastMessageStyle = vbCritical Then
                sbMsgError sMsgPlafond, TheApp.FunctionName
                Screen.MousePointer = 0
                Exit Function
            Else
                If fnMsgAlertQuestionOKCancel(sMsgPlafond, TheApp.FunctionName) = vbCancel Then
                    Screen.MousePointer = 0
                    'bSaving = 0
                    Exit Function
                End If
            End If
        End If
    End If
    
    VarNumeroDoc = oDoc.Insert
    
    If VarNumeroDoc > 0 Then
        If (NonRiportaInXMLRifVsNumOrd = 0) Then
            If Len(fnNotNull(oDoc.Tables.Field("Doc_numero_vs_ordine_di_rifer", , sTabellaTestata))) > 0 Then
                If oDoc.IDTipoOggetto <> 8 Then
                    SCRIVI_ORD_CLI_RIF_XML fnNotNull(oDoc.Tables.Field("Doc_data_vs_ordine_di_rifer", , sTabellaTestata)), fnNotNull(oDoc.Tables.Field("Doc_numero_vs_ordine_di_rifer", , sTabellaTestata))
                End If
            End If
        End If
        
        fnAggiornaDescrizioneDocumento oDoc.Tables.Field("Doc_prefisso", , sTabellaDettaglio), sTabellaDettaglio
        fnAggiornaDescrizioneOrdineCliente sTabellaDettaglio, oDoc.Tables.Field("Doc_data_vs_ordine_di_rifer", , sTabellaTestata), oDoc.Tables.Field("Doc_numero_vs_ordine_di_rifer", , sTabellaTestata)
        fnAggiornaDescrizioneOrdineInterno sTabellaDettaglio, oDoc.Tables.Field("Doc_data_ns_ordine_di_rifer", , sTabellaTestata), oDoc.Tables.Field("Doc_numero_ns_ordine_di_rifer", , sTabellaTestata)
        fnAggiornaDescrizioneDocumentoOrd oDoc.Tables.Field("Doc_prefisso", , sTabellaDettaglio), sTabellaDettaglio
        
        AGGIORNA_RIGHE_DOCUMENTO sTabellaDettaglio

        AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI 0, oDoc.IDOggetto, Me.cboAltroSito.CurrentID
        
        SCRIVI_CAUSALI_DOC oDoc.IDOggetto
        
    End If

    Screen.MousePointer = vbDefault
    
    InserimentoDMT = True
Exit Function

ERR_InserimentoDMT:
    InserimentoDMT = False
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "InserimentoDMT"
    
    
End Function
Private Sub ConsolidaDettaglioFatturaElettronica()
On Error GoTo ERR_ConsolidaDettaglioFatturaElettronica
    Dim oRs As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFilter As String
    
    'DATI
    'leggo i dati legati al documento
    With oDoc.ElectronicInvoiceAdditionalData.AdditionalData
        'conservo il filtro per rimetterlo dopo
        If Len(.Filter) > 0 Then sFilter = .Filter
                            
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            While Not .EOF
                If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                    .Fields("Temporaneo").Value = False
                End If
                .MoveNext
            Wend
        End If
        .Filter = sFilter
    End With
    'DATI
    'leggo i dati legati al documento
    With oDoc.ElectronicInvoiceAdditionalData.AdditionalCodes
        'conservo il filtro per rimetterlo dopo
        If Len(.Filter) > 0 Then sFilter = .Filter
                            
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            While Not .EOF
                If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                    .Fields("Temporaneo").Value = False
                End If
                .MoveNext
            Wend
        End If
        .Filter = sFilter
    End With
    
    oDoc.ElectronicInvoiceAdditionalData.Changed = True

Exit Sub
ERR_ConsolidaDettaglioFatturaElettronica:
    MsgBox Err.Description, vbCritical, "ConsolidaDettaglioFatturaElettronica"
End Sub
Private Sub sbLoadElectronicInvoiceData4Article(ByVal lID_Art_dettaglio_prog As Long, ByVal lIDArticle As Long)
On Error GoTo ERR_sbLoadElectronicInvoiceData4Article
    Dim oRs As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFilter As String

    'DATI
    'leggo i dati legati all'articolo con IDArticolo richiesto
    Set oRs = oDoc.ElectronicInvoiceAdditionalData.GetDataFromArticle(lIDArticle)
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then
            If oRs.RecordCount > 0 Then
                With oDoc.ElectronicInvoiceAdditionalData.AdditionalData
                    'conservo il filtro per rimetterlo dopo
                    If Len(.Filter) > 0 Then sFilter = .Filter
                    
                    'metto gli elementi precedenti in naftalina ;)
                    .Filter = "ID_Art_dettaglio_prog = " & lID_Art_dettaglio_prog
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        
                        While Not .EOF
                            If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                                .Fields("ID_Art_dettaglio_prog").Value = -1 * fnNotNullL(.Fields("ID_Art_dettaglio_prog").Value)
                            End If
                            .MoveNext
                        Wend
                    End If
                    .Filter = sFilter
                    
                    'riverso i dati aggiuntivi dell'articolo sul dettaglio
                    oRs.MoveFirst
                    
                    While Not oRs.EOF
                        .AddNew
                        
                        For Each oField In oRs.Fields
                            If oField.Name <> "IDDatoFatturaPAPerArticolo" Then
                                .Fields(oField.Name).Value = oField.Value
                            End If
                        Next
                        Set oField = Nothing
                        
                        .Fields("IDOggetto").Value = oDoc.IDOggetto
                        .Fields("IDTipoOggetto").Value = oDoc.IDTipoOggetto
                        .Fields("ID_Art_dettaglio_prog").Value = lID_Art_dettaglio_prog
                        .Fields("Eliminato").Value = False
                        'Impostare "Temporaneo" a False se i codici vengono immediatamente legati al dettaglio,
                        'a True se questi codici rimangono sospesi in attesa di un ulteriore conferma
                        '(Salva di dettaglio, ad es., i cui Temporaneo verrà finalmente posto a False)
                        .Fields("Temporaneo").Value = False
                        
                        oRs.MoveNext
                    Wend
                    oDoc.ElectronicInvoiceAdditionalData.Changed = True
                End With
            End If
            oRs.Close
        End If
    End If
    Set oRs = Nothing
    
    'CODICI
    'leggo i codici legati all'articolo con IDArticolo richiesto
    Set oRs = oDoc.ElectronicInvoiceAdditionalData.GetCodesFromArticle(lIDArticle)
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then
            If oRs.RecordCount > 0 Then
                With oDoc.ElectronicInvoiceAdditionalData.AdditionalCodes
                    'conservo il filtro per rimetterlo dopo
                    If Len(.Filter) > 0 Then sFilter = .Filter
                    'metto gli elementi precedenti in naftalina ;)
                    .Filter = "ID_Art_dettaglio_prog = " & lID_Art_dettaglio_prog
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        
                        While Not .EOF
                            If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                                .Fields("ID_Art_dettaglio_prog").Value = -1 * fnNotNullL(.Fields("ID_Art_dettaglio_prog").Value)
                            End If
                            .MoveNext
                        Wend
                    End If
                    .Filter = sFilter
                    
                    'riverso i codici aggiuntivi dell'articolo sul dettaglio
                    oRs.MoveFirst
                    
                    While Not oRs.EOF
                        .AddNew
                        
                        For Each oField In oRs.Fields
                            If oField.Name <> "IDCodiceFatturaPAPerArticolo" Then
                                .Fields(oField.Name).Value = oField.Value
                            End If
                        Next
                        Set oField = Nothing
                        
                        .Fields("IDOggetto").Value = oDoc.IDOggetto
                        .Fields("IDTipoOggetto").Value = oDoc.IDTipoOggetto
                        .Fields("ID_Art_dettaglio_prog").Value = lID_Art_dettaglio_prog
                        .Fields("Eliminato").Value = False
                        'Impostare "Temporaneo" a False se i codici vengono immediatamente legati al dettaglio,
                        'a True se questi codici rimangono sospesi in attesa di un ulteriore conferma
                        '(Salva di dettaglio, ad es., i cui Temporaneo verrà finalmente posto a False)
                        .Fields("Temporaneo").Value = False
                        
                        oRs.MoveNext
                    Wend
                    oDoc.ElectronicInvoiceAdditionalData.Changed = True
                End With
            End If
            oRs.Close
        End If
    End If
    Set oRs = Nothing
Exit Sub
ERR_sbLoadElectronicInvoiceData4Article:
    MsgBox Err.Description, vbCritical, "sbLoadElectronicInvoiceData4Article"
End Sub
Private Sub SCRIVI_ORD_CLI_RIF_XML(DataOrdine As String, NumeroOrdine As String)
On Error GoTo ERR_SCRIVI_ORD_CLI_RIF_XML
Dim sSQL As String
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM DatoFatturaPATestataDoc "
sSQL = sSQL & "WHERE IDDatoFatturaPATestataDoc=0"

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
rsNew.AddNew
    rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
    rsNew!IDBloccoXML = 1
    rsNew!IDOggetto = oDoc.IDOggetto
    rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
    rsNew!IDDocumento = NumeroOrdine
    If Len(DataOrdine) > 0 Then
        rsNew!Data = DataOrdine
    End If
    rsNew!NumItem = NumeroOrdine
rsNew.Update

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_SCRIVI_ORD_CLI_RIF_XML:
    MsgBox Err.Description, vbCritical, "SCRIVI_ORD_CLI_RIF_XML"

End Sub
Private Sub fnAggiornaDescrizioneDocumento(LetteraSezionale As String, NomeTabellaDettaglio As String)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_PODescrizioneDocumento="

Select Case oDoc.IDTipoOggetto
    Case 2
        If NUMERO_ZERI_DOC_RIF = 0 Then
            sSQL = sSQL & fnNormString("Rif. D.d.t. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(oDoc.Numero) & " del " & oDoc.DataEmissione)
        Else
            sSQL = sSQL & fnNormString("Rif. D.d.t. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & oDoc.Numero & " del " & oDoc.DataEmissione)
        End If
    Case 114
        If NUMERO_ZERI_DOC_RIF = 0 Then
            sSQL = sSQL & fnNormString("Rif. f.a. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(oDoc.Numero) & " del " & oDoc.DataEmissione)
        Else
            sSQL = sSQL & fnNormString("Rif. f.a. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & oDoc.Numero & " del " & oDoc.DataEmissione)
        End If
    Case 8
        If NUMERO_ZERI_DOC_RIF = 0 Then
            sSQL = sSQL & fnNormString("Rif. s.n.f. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(oDoc.Numero) & " del " & oDoc.DataEmissione)
        Else
            sSQL = sSQL & fnNormString("Rif. s.n.f. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & oDoc.Numero & " del " & oDoc.DataEmissione)
        End If
End Select

sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto

Cn.Execute sSQL
End Sub
Private Sub fnAggiornaDescrizioneDocumentoOrd(LetteraSezionale As String, NomeTabellaDettaglio As String)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_PODescrizioneDocumentoOrdinamento="

Select Case oDoc.IDTipoOggetto
    Case 2
        sSQL = sSQL & fnNormString("Rif. D.d.t. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(oDoc.Numero) & " del " & oDoc.DataEmissione)
    Case 114
        sSQL = sSQL & fnNormString("Rif. f.a. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(oDoc.Numero) & " del " & oDoc.DataEmissione)
    Case 8
        sSQL = sSQL & fnNormString("Rif. s.n.f. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(oDoc.Numero) & " del " & oDoc.DataEmissione)
End Select

sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto

Cn.Execute sSQL

End Sub
Private Sub fnAggiornaDescrizioneOrdineCliente(NomeTabellaDettaglio As String, DataOrdine As String, NumeroOrdine As String)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_POOggettoOrdineCliente="
sSQL = sSQL & fnNormString("Rif. certificato: n. " & NumeroOrdine & " del " & DataOrdine)
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto

Cn.Execute sSQL


End Sub
Private Sub fnAggiornaDescrizioneOrdineInterno(NomeTabellaDettaglio As String, DataOrdine As String, NumeroOrdine As String)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_POOggettoOrdineInterno="
sSQL = sSQL & fnNormString("Rif. D.D.T. Socio: n. " & NumeroOrdine & " del " & DataOrdine)
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto

Cn.Execute sSQL

End Sub

Private Function GetNumeroDocumentoModificato(NumeroDocumento As String) As String
Const Totale As Integer = 6
Dim I As Integer
Dim Count As Integer

GetNumeroDocumentoModificato = ""
For I = 1 To (Totale - Len(NumeroDocumento))
    GetNumeroDocumentoModificato = GetNumeroDocumentoModificato & "0"
Next

GetNumeroDocumentoModificato = GetNumeroDocumentoModificato & NumeroDocumento

End Function
Private Sub AGGIORNA_RIGHE_DOCUMENTO(NomeTabella As String)
On Error GoTo ERR_AGGIORNA_RIGHE_DOCUMENTO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsCount As DmtOleDbLib.adoResultset
Dim Unita_Progresso As Double
Dim NumeroRecord As Long
Dim Moltiplicatore As Double
Dim Prezzo As Double
Dim MerceNettaPerLiquidazione As Double
Dim IDUMCoop_ArtVenduto As Long
Dim PrezzoScontato As Double

sSQL = "SELECT IDOggetto, RV_POPrezzoMedioInLiq, RV_POVariazionePrezzoImballo, RV_POImportoMerceNetta, RV_POImportoLiq, "
sSQL = sSQL & "RV_POPrezzoUnitarioOrigine, RV_POIDIvaImballo, RV_POIDTipoVariazione, Art_pre_uni_net_sco_net_IVA ,"
sSQL = sSQL & "RV_POImportoDaLiq, RV_POLinkRiga, RV_POImportoImballoInArticolo, "
sSQL = sSQL & "Art_numero_colli, Art_quantita_Totale, Link_art_articolo, "
sSQL = sSQL & "RV_POVariazionePrezzoManuale, RV_POImportoRigaCommissioni,  "
sSQL = sSQL & "RV_POIDConferimentoRighe, RV_POIDAssegnazioneMerce, RV_POIDProcessoIVGamma, RV_PODataLavorazione, "
sSQL = sSQL & "RV_POAnnotazioniAggiuntiveLav, RV_PONotaRigaOrdRaggr, RV_PODataOrdineCliente, RV_PONumeroOrdineCliente, "
sSQL = sSQL & "RV_PODataOrdineInterno, RV_PONumeroOrdineInterno,  "
sSQL = sSQL & "RV_POIDImballoPrim, RV_PONumeroConfezioniPerImballo, RV_POCostoConfezioneImballo, RV_POCostoKitLiq, "
sSQL = sSQL & "RV_POCostoConfezioneImballoLiq, RV_POQuantitaLiq,  "
sSQL = sSQL & "Art_numero_colli,Art_Peso, Art_Tara, Art_quantita_pezzi, RV_POImportoImballoSel, RV_POImportoLiqDoc, Art_prezzo_unitario_neutro, "
sSQL = sSQL & "Art_sco_in_percentuale_1,Art_sco_in_percentuale_2 "
sSQL = sSQL & " FROM " & NomeTabella
sSQL = sSQL & " WHERE RV_POTipoRiga=1 "
sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    
    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_articolo))
    rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
    IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_articolo))
    
    Select Case IDUMCoop_ArtVenduto
        Case 1
            rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_numero_colli) * Moltiplicatore
        Case 2
            rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
        Case 3
            rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
        Case 4
            rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
        Case 5
            rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
    End Select

    rs!RV_POImportoDaLiq = 0
    If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1 Then
        rs!RV_POImportoDaLiq = -(fnNotNullN(rs!RV_POImportoImballoSel) * fnNotNullN(rs!Art_numero_colli)) / fnNotNullN(rs!RV_POQuantitaLiq)
    End If
    
    
    If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 0 Then
        rs!RV_POVariazionePrezzoImballo = 0
        rs!RV_POImportoMerceNetta = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
    Else
        rs!RV_POVariazionePrezzoImballo = ((Abs(fnNotNullN(rs!RV_POImportoDaLiq)) * fnNotNullN(rs!RV_POQuantitaLiq)) / fnNotNullN(rs!Art_quantita_totale))
        rs!RV_POImportoMerceNetta = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA) - rs!RV_POVariazionePrezzoImballo
    End If

    
    Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
    Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
    
    MerceNettaPerLiquidazione = Prezzo
    PrezzoScontato = Prezzo
    
    If fnNotNullN(rs!RV_POImportoImballoInArticolo) > 0 Then
        Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoDaLiq)
    Else
        Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoDaLiq))
    End If
    
    Prezzo = Prezzo
    
    If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
        Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
    Else
        Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
    End If
    
    rs!RV_POCostoConfezioneImballoLiq = 0
    rs!RV_POCostoKitLiq = 0

    
    rs!RV_POImportoLiq = Prezzo
    rs!RV_POImportoLiqDoc = Prezzo
    rs.Update
    
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
Exit Sub
ERR_AGGIORNA_RIGHE_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "AGGIORNA_RIGHE_DOCUMENTO"
End Sub
Private Function GET_MOLTIPLICATORE_ARTICOLO(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POMoltiplicatore FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_MOLTIPLICATORE_ARTICOLO = 1
Else
    If fnNotNullN(rs!RV_POMoltiplicatore) = 0 Then
        GET_MOLTIPLICATORE_ARTICOLO = 1
    Else
        GET_MOLTIPLICATORE_ARTICOLO = fnNotNullN(rs!RV_POMoltiplicatore)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = 0

sSQL = "SELECT RV_POIDUnitaDiMisuraLiq "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo
        
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI(IDTipoOggetto As Long, IDOggetto As Long, IDDestinazione As Long)
On Error GoTo ERR_AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim QuantitaMovimento As Double
Dim Link_Unita_di_Misura_Conferimeto As Long


sSQL = "SELECT " & sTabellaDettaglio & ".IDValoriOggettoDettaglio, "
sSQL = sSQL & sTabellaDettaglio & ".Art_pre_uni_net_sco_net_IVA, "
sSQL = sSQL & sTabellaDettaglio & ".Link_art_articolo, "
sSQL = sSQL & sTabellaDettaglio & ".Art_quantita_pezzi, "
sSQL = sSQL & sTabellaDettaglio & ".Art_numero_colli, "
sSQL = sSQL & sTabellaDettaglio & ".Art_tara, "
sSQL = sSQL & sTabellaDettaglio & ".Art_peso, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PODataConferimento, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDConferimentoRighe, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POCodiceLotto, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDSocio, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POLottoCampagna, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoDaLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POQuantitaLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDAssegnazioneMerce, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDProcessoIVGamma, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POPrezzoMedioInLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoUnitarioImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoMerceNetta, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDIvaImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POVariazionePrezzoImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoImballoInArticolo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDAnagraficaFatturazione, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoImportoVenditaLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POVariazionePrezzoManuale, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoDocumentoCoop, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoRigaCommissioni, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PODataLavorazione, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoLavorazione, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoCategoria, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDCalibro, "

sSQL = sSQL & sTabellaDettaglio & ".RV_POIDPedana, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POCodicePedana, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoPedana, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POPesoPedana, "

sSQL = sSQL & sTabellaDettaglio & ".RV_PORigaRiscontroPeso, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POAnnotazioniAggiuntiveLav, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PONotaRigaOrdRaggr, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PODataOrdineCliente, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PONumeroOrdineCliente, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PODataOrdineInterno, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PONumeroOrdineInterno, "

sSQL = sSQL & sTabellaDettaglio & ".RV_POIDImballoPrim, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PONumeroConfezioniPerImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POTaraConfezioneImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POCodiceImballoPrim, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PODescrizioneImballoPrim, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POCostoConfezioneImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POCostoKitLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POCostoConfezioneImballoLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDLottoCampagnaLavorazione, "

sSQL = sSQL & "RV_POCaricoMerceRighe.IDArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceTesta.NumeroDocumento, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDUnitaDiMisura, RV_POCaricoMerceRighe.CodiceLotto, "
sSQL = sSQL & "RV_POCaricoMerceTesta.IDMagazzinoConferimento, RV_POCaricoMerceRighe.IDUnitaDiMisuraDiamante, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDRV_POTipoLavorazione AS IDTipoLavorazioneConf, RV_POCaricoMerceRighe.PrezzoMedio AS PrezzoMedioConf "
sSQL = sSQL & "FROM " & sTabellaDettaglio & " LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDConferimentoRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE " & sTabellaDettaglio & ".IDOggetto=" & IDOggetto
sSQL = sSQL & " AND " & sTabellaDettaglio & ".RV_POTipoRiga=1"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection
    
While Not rs.EOF
    Select Case fnNotNullN(rs!IDUnitaDiMisura)
        Case 1
            QuantitaMovimento = fnNotNullN(rs!Art_numero_colli)
        Case 2
            QuantitaMovimento = fnNotNullN(rs!Art_peso)
        Case 3
            QuantitaMovimento = fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)
        Case 4
            QuantitaMovimento = fnNotNullN(rs!Art_tara)
        Case 5
            QuantitaMovimento = fnNotNullN(rs!Art_quantita_pezzi)
    End Select
    
    Aggiorna_Movimento_Documento fnNotNullN(rs!Link_Art_articolo), fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), rs!IDValoriOggettoDettaglio, fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma), fnNotNullN(rs!RV_POIDSocio), fnNotNull(rs!RV_PODataConferimento), fnNotNullN(rs!NumeroDocumento), _
    fnNotNull(rs!CodiceLotto), fnNotNull(rs!RV_POLottoCampagna), fnNotNull(rs!RV_POCodiceLotto), _
    fnNotNullN(rs!RV_POQuantitaLiq), fnNotNullN(rs!RV_POImportoDaLiq), fnNotNullN(rs!RV_POImportoLiq), QuantitaMovimento, _
    fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)), _
    fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi), fnNotNullN(rs!RV_POPrezzoMedioInLiq), fnNotNullN(rs!RV_POIDImballo), fnNotNullN(rs!RV_POImportoUnitarioImballo), _
    fnNotNullN(rs!RV_POIDIvaImballo), fnNotNullN(rs!RV_POVariazionePrezzoImballo), fnNotNullN(rs!RV_POImportoMerceNetta), fnNotNullN(rs!RV_POImportoImballoInArticolo), fnNotNullN(rs!RV_POIDAnagraficaFatturazione), IDDestinazione, fnNotNullN(rs!RV_POIDTipoImportoVenditaLiq), _
    fnNotNullN(rs!RV_POIDTipoDocumentoCoop), fnNotNullN(rs!RV_POVariazionePrezzoManuale), fnNotNull(rs!RV_PODataLavorazione), _
    fnNotNullN(rs!RV_POIDTipoLavorazione), fnNotNullN(rs!RV_POIDTipoCategoria), fnNotNullN(rs!RV_POIDCalibro), fnNotNullN(rs!IDTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), _
    fnNotNullN(rs!RV_POIDPedana), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNull(rs!RV_POCodicePedana), fnNotNullN(rs!RV_POPesoPedana), fnNotNullN(rs!RV_POImportoRigaCommissioni), _
    fnNotNull(rs!RV_POAnnotazioniAggiuntiveLav), fnNotNull(rs!RV_PONotaRigaOrdRaggr), fnNotNull(rs!RV_PODataOrdineCliente), fnNotNull(rs!RV_PONumeroOrdineCliente), fnNotNull(rs!RV_PODataOrdineInterno), fnNotNull(rs!RV_PONumeroOrdineInterno), _
    fnNotNullN(rs!RV_POIDImballoPrim), fnNotNull(rs!RV_POCodiceImballoPrim), fnNotNull(rs!RV_PODescrizioneImballoPrim), fnNotNullN(rs!RV_PONumeroConfezioniPerImballo), fnNotNullN(rs!RV_POTaraConfezioneImballo), _
    fnNotNullN(rs!RV_POCostoConfezioneImballo), fnNotNullN(rs!RV_POCostoConfezioneImballoLiq), fnNotNullN(rs!RV_POCostoKitLiq), fnNotNullN(rs!RV_POIDLottoCampagnaLavorazione)
    
rs.MoveNext
Wend

rs.Close
Set rs = Nothing
Exit Sub
ERR_AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI:
    MsgBox Err.Description, vbCritical, "ERR_AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI"
End Sub
Private Function Aggiorna_Movimento_Documento(IDArticolo As Long, ImportoUnitarioArticolo As Double, IDRiga As Long, IDRigaConferimento, IDAssegnazione As Long, IDProcessoIVGamma As Long, IDSocio As Long, _
DataConferimento As String, NumeroConferimento As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, CodiceLottoVendita As String, _
QuantitaLiquidazione As Double, ImportoInclusoImballo As Double, ImportoLiquidazione As Double, QuantitaMovimentata As Double, Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, _
PrezzoMedioLiq As Double, IDArticoloImballo As Long, ImportoUnitarioImballo, IDIvaImballo As Long, VariazionePrezzoImballo As Double, PrezzoMerceNetta As Double, MerceInclusaImballo As Long, IDAnagraficaFatturazione As Long, _
IDSitoPerAnagrafica As Long, IDTipoImportoLiq As Long, IDTipoDocumentoCoop As Long, VarImpLiqMan As Double, DataLavorazione As String, IDTipoLavorazione As Long, IDTipoCategoria As Long, IDCalibro As Long, IDTipoLavorazioneConf As Long, PrezzoMedioConf As Long, _
IDPedana As Long, IDTipoPedana As Long, CodicePedana As String, PesoPedana As Double, ImportoRigaCommissioni As Double, _
AnnotazioniAggiuntive As String, RaggrOrdine As String, DataOrdineCliente As String, NumeroOrdineCliente As String, DataOrdineInterno As String, NumeroOrdineInterno As String, _
IDImballoPrimario As Long, CodiceImballoPrimario As String, DescrizioneImballoPrimario As String, NConfezioniImballo As Double, TaraConfezione As Double, _
CostoConfezioneImballo As Double, CostoConfezioneImballoLiq As Double, CostoKitLiq As Double, IDLottoProduzioneLavorazione As Long) As Long

Dim Prezzo As Double
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim Moltiplicatore As Double
Dim MerceNettaPerLiquidazione As Double

    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(IDArticolo)
    
    sSQL = "SELECT * FROM Movimento "
    sSQL = sSQL & "WHERE IDTipoOggetto=" & oDoc.IDTipoOggetto
    sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDRiga
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    If Not rs.EOF Then
        
        rs("RV_POTipoRiga").Value = 1
        rs("RV_POIDCaricoMerceRighe").Value = IDRigaConferimento
        rs("RV_POIDAssegnazioneMerce").Value = IDAssegnazione
        rs("RV_POIDProcessoIVGamma") = IDProcessoIVGamma
        rs("RV_POIDAnagraficaSocio") = IDSocio
        If Len(DataConferimento) > 0 Then
            rs("RV_PODataConferimento") = DataConferimento
        End If
        rs("RV_PONumeroConferimento") = NumeroConferimento
        rs("RV_POCodiceLotto") = CodiceLottoEntrata
        rs("RV_POCodiceLottoCampagna") = CodiceLottoCampagna
        rs("RV_POCodiceLottoVendita") = CodiceLottoVendita
        rs("RV_POQuantitaLiquidazione") = QuantitaLiquidazione
        rs("RV_POImportoInclusoImballo") = ImportoInclusoImballo
        rs("RV_POPrezzoMerceNetta").Value = PrezzoMerceNetta
        
        rs("RV_POImportoRigaCommissioni") = ImportoRigaCommissioni
        rs("RV_POImportoLiquidazione") = ImportoLiquidazione
        rs("RV_POQuantitaMovimentata") = QuantitaMovimentata
        rs("RV_PONumeroColli") = Colli
        rs("RV_POPesoLordo") = PesoLordo
        rs("RV_POPesoNetto") = PesoLordo - Tara
        rs("RV_POTara") = Tara
        rs("RV_POQuantitaPezzi") = Pezzi
        
        rs("RV_POPrezzoMedioInLiq").Value = PrezzoMedioLiq
        rs("RV_POIDImballo").Value = IDArticoloImballo
        rs("RV_POImportoUnitarioImballo").Value = ImportoUnitarioImballo
        rs("RV_POIDIvaImballo").Value = IDIvaImballo
        rs("RV_POVariazionePrezzoImballo").Value = VariazionePrezzoImballo
        rs("RV_POQuantitaLiqPerPrezzoMedio").Value = QuantitaLiquidazione
        rs("RV_POMerceInclusaImballo").Value = MerceInclusaImballo
        rs("RV_POTipoRigaCollegata").Value = 0
        rs("RV_POIDAnagraficaFatturazione").Value = IDAnagraficaFatturazione
        rs("RV_POIDSitoPerAnagrafica").Value = IDSitoPerAnagrafica
        rs("RV_POIDTipoImportoVenditaLiq").Value = IDTipoImportoLiq
        rs("Oggetto").Value = GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto)
        
        rs("RV_POVariazionePrezzoManuale").Value = VarImpLiqMan
        rs("RV_POIDTipoDocumentoCoop").Value = IDTipoDocumentoCoop
        
        rs("RV_PODataLavorazione").Value = DataLavorazione
        rs("RV_POIDTipoLavorazione").Value = IDTipoLavorazione
        rs("RV_POIDCalibro").Value = IDCalibro
        rs("RV_POIDTipoCategoria").Value = IDTipoCategoria
        rs("RV_POIDTipoLavorazioneConf").Value = IDTipoLavorazioneConf
        rs("RV_POPrezzoMedioConf").Value = PrezzoMedioConf
        
        rs("RV_POIDPedana").Value = IDPedana
        rs("RV_POIDTipoPedana").Value = IDTipoPedana
        rs("RV_POCodicePedana").Value = CodicePedana
        rs("RV_POPesoPedana").Value = PesoPedana

        rs("RV_PORigaRiscontroPeso") = 0
        rs("RV_POAnnotazioniAggiuntiveLav").Value = AnnotazioniAggiuntive
        rs("RV_PONotaRigaOrdRaggr").Value = RaggrOrdine
        
        If Len(DataOrdineCliente) > 0 Then
            rs("RV_PODataOrdineCliente").Value = DataOrdineCliente
        End If
        rs("RV_PONumeroOrdineCliente").Value = NumeroOrdineCliente
        
        If Len(DataOrdineInterno) > 0 Then
            rs("RV_PODataOrdineInterno").Value = DataOrdineInterno
        End If
        
        rs("RV_PONumeroOrdineInterno").Value = NumeroOrdineInterno
        rs("RV_POIDImballoPrim").Value = IDImballoPrimario
        rs("RV_POCodiceImballoPrim").Value = CodiceImballoPrimario
        rs("RV_PODescrizioneImballoPrim").Value = DescrizioneImballoPrimario
        rs("RV_PONumeroConfezioniPerImballo").Value = NConfezioniImballo
        rs("RV_POTaraConfezioneImballo").Value = TaraConfezione
        rs("RV_POCostoConfezioneImballo").Value = CostoConfezioneImballo
        rs("RV_POCostoConfezioneImballoLiq").Value = CostoConfezioneImballoLiq
        rs("RV_POCostoKitLiq").Value = CostoKitLiq
        rs("RV_POImportoLiqDoc") = ImportoLiquidazione
        rs("RV_POIDLottoCampagnaLavorazione").Value = IDLottoProduzioneLavorazione
        rs("RV_PODataCompetenzaLiq").Value = oDoc.DataEmissione
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
End Function
Public Function fnGetNewKey(tabella As String, CampoKey As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim VarData As String
'Monta la query SQL per trovare il massimo valore della chiave primaria
sSQL = "SELECT MAX (" & CampoKey & ") AS MaxID FROM " & tabella ' & " WHERE " & >=" & VarData

'Apertura del recordset
Set rs = Cn.OpenResultset(fnAnsi2Jet(sSQL))

'Determina il primo progressivo disponibile
fnGetNewKey = fnNotNullN(rs.adoColumns("MaxID")) + 1
If fnGetNewKey <= 0 Then fnGetNewKey = 1

'Chiude il recordset e distrugge l'oggetto.
rs.CloseResultset
Set rs = Nothing

End Function
Private Sub SCRIVI_CAUSALI_DOC(IDOggetto As Long)
On Error GoTo ERR_SCRIVI_CAUSALI_DOC
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rsDoc As DmtOleDbLib.adoResultset
Dim tipo As Long
Dim Ordinamento As Long
Dim TestoVettoreSuccessivo As String
Dim TestoAgenziaTraporto As String

If oDoc.IDTipoOggetto = 8 Then Exit Sub

tipo = 0


If (oDoc.Field("RV_PODocumentoCRM", , sTabellaTestata) = 1) Then
    tipo = 1
Else
    If fnNotNullN(oDoc.Field("RV_POIDAnagraficaDestinazione", , sTabellaTestata)) > 0 Then
        tipo = 2
    End If
End If

sSQL = "SELECT * FROM DatoFatturaPATestataDoc "
sSQL = sSQL & "WHERE IDDatoFatturaPATestataDoc=0"

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

Ordinamento = ADD_NOTE_TIPO_OGGETTO(oDoc.IDTipoOggetto, oDoc.IDOggetto, tipo, rsNew)


If Rip_InXMLRifLetteraIntento = 1 Then
    ADD_NOTA_LETTERA_INTENTO oDoc.IDTipoOggetto, oDoc.IDOggetto, rsNew, fnNotNullN(oDoc.Field("Link_nom_lettera_intento", , sTabellaTestata)), Ordinamento
End If

If Rip_InXMLRifNoteIva = 1 Then
    ADD_NOTE_IVA oDoc.IDTipoOggetto, oDoc.IDOggetto, rsNew, Ordinamento
End If

'ANNOTAZIONE 1 DOCUMENTO
If Rip_InXMLRifNota01Doc = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni1", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni1", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ANNOTAZIONE 2 DOCUMENTO
If Rip_InXMLRifNota02Doc = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni2", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni2", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ANNOTAZIONE 3 DOCUMENTO
If Rip_InXMLRifNota03Doc = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni3", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni3", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If
'ANNOTAZIONE STANDARD DEL DOCUMENTO
If Rip_InXMLRifNotaDoc = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("Doc_annotazioni_variazio", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("Doc_annotazioni_variazio", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ISTRUZIONI DEL MITTENTE
If Rip_InXMLRifIstrMitt = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POIstruzioniMittente", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("RV_POIstruzioniMittente", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If




'TARGA AUTOMEZZO
If Rip_InXMLRifTargaAutoMezzo = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POTargaAutomezzo", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = "Targa automezzo: " & Mid(Trim(fnNotNull(oDoc.Field("RV_POTargaAutomezzo", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If
rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_SCRIVI_CAUSALI_DOC:
    MsgBox Err.Description, vbCritical, "SCRIVI_CAUSALI_DOC"
End Sub
Private Function ADD_NOTE_TIPO_OGGETTO(IDTipoOggetto As Long, IDOggetto As Long, tipo As Long, rsAdd As ADODB.Recordset) As Long
On Error GoTo ERR_ADD_NOTE_TIPO_OGGETTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Ordinamento As Long

sSQL = "SELECT * FROM RV_PONoteDocumentiCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

Ordinamento = 1

If Not rs.EOF Then
    Select Case tipo
        Case 0
            If Len(Trim(fnNotNull(rs!Annotazioni1))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione01) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni1)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni3))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione03) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni3)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni4))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione04) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni4)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni5))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione05) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
        Case 1
            If Len(Trim(fnNotNull(rs!Annotazioni6))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione06) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni6)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni7))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione07) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni7)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni8))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione08) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni8)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni9))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione09) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni9)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni10))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione10) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni10)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni5))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione05) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
        Case 2
            If Len(Trim(fnNotNull(rs!Annotazioni11))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione11) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni11)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni12))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione12) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni12)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni13))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione13) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni13)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni14))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione14) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni14)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni15))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione15) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni15)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni5))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione05) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 200)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
    End Select
End If

rs.CloseResultset
Set rs = Nothing

ADD_NOTE_TIPO_OGGETTO = Ordinamento
Exit Function
ERR_ADD_NOTE_TIPO_OGGETTO:
    MsgBox Err.Description, vbCritical, "ADD_NOTE_TIPO_OGGETTO"
End Function

Private Function ADD_NOTA_LETTERA_INTENTO(IDTipoOggetto As Long, IDOggetto As Long, rsAdd As ADODB.Recordset, IDLetteraIntento As Long, Ordinamento As Long) As Long
On Error GoTo ERR_ADD_NOTA_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Testo As String
Dim rsNota As DmtOleDbLib.adoResultset
Dim Nota2 As String

If IDLetteraIntento = 0 Then Exit Function

Nota2 = ""
Testo = ""

sSQL = "SELECT Annotazioni2 FROM RV_PONoteDocumentiCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto

Set rsNota = Cn.OpenResultset(sSQL)

If Not rsNota.EOF Then
    Nota2 = fnNotNull(rsNota!Annotazioni2)
End If

rsNota.CloseResultset
Set rsNota = Nothing

sSQL = "SELECT IDLetteraIntento, IDTipoLetteraIntento, IDAzienda, Data, Numero, Anno, NumeroCliFor, AnnoCliFor, "
sSQL = sSQL & "IDAnagrafica_CF, IDTipoAnagrafica_CF, IDAzienda_CF, ProgressivoDichiarazione, ProtocolloDichiarazione, DataEmissione "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & IDLetteraIntento

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Testo = Nota2 & " " & fnNotNull(rs!ProtocolloDichiarazione) & "/" & fnNotNull(rs!ProgressivoDichiarazione)
    Testo = Testo & " del " & fnNotNull(rs!Data)
    Testo = Testo & " - nr. c/o cliente " & fnNotNull(rs!NumeroCliFor)
    Testo = Testo & " del " & fnNotNull(rs!DataEmissione)
End If

rs.CloseResultset
Set rs = Nothing

If Len(Testo) > 0 Then
    rsAdd.AddNew
        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
        rsAdd!IDBloccoXML = 8
        rsAdd!IDOggetto = IDOggetto
        rsAdd!IDTipoOggetto = IDTipoOggetto
        rsAdd!Annotazioni = Mid(Trim(Testo), 1, 200)
        rsAdd!Ordinamento = Ordinamento
        Ordinamento = Ordinamento + 1
    rsAdd.Update
End If
Exit Function
ERR_ADD_NOTA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "ERR_ADD_NOTA_LETTERA_INTENTO"
End Function

Private Sub ADD_NOTE_IVA(IDTipoOggetto As Long, IDOggetto As Long, rsAdd As ADODB.Recordset, Ordinamento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NomeTabella As String

NomeTabella = sTabellaIVA

sSQL = "SELECT " & sTabellaIVA & ".IDValoriOggettoDettaglio, " & sTabellaIVA & ".IDOggetto, " & sTabellaIVA & ".IDTipoOggetto, " & sTabellaIVA & ".Link_Cst_IVA, Iva.Annotazioni"
sSQL = sSQL & " FROM " & sTabellaIVA & " INNER JOIN "
sSQL = sSQL & " Iva ON " & sTabellaIVA & ".Link_Cst_IVA = Iva.IDIva "
sSQL = sSQL & " WHERE (" & sTabellaIVA & ".IDOggetto = " & IDOggetto & ")"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If Len(Trim(rs!Annotazioni)) > 0 Then
        rsAdd.AddNew
            rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsAdd!IDBloccoXML = 8
            rsAdd!IDOggetto = IDOggetto
            rsAdd!IDTipoOggetto = IDTipoOggetto
            rsAdd!Annotazioni = Mid(Trim(rs!Annotazioni), 1, 200)
            rsAdd!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsAdd.Update
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_ADD_NOTA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "ADD_NOTA_LETTERA_INTENTO"

End Sub
Private Sub LOAD_COLUMN()
On Error GoTo ERR_LOAD_COLUMN
Dim cl As dmtgridctl.dgColumnHeader
    If Me.BrwMain.ColumnsHeader.Count = 0 Then
        With Me.BrwMain.ColumnsHeader
            .Add "IDRV_POCertificato", "IDRV_POCertificato", dgNumeric, False
            .Add "IDAzienda", "IDAzienda", dgNumeric, False
            .Add "IDContratto", "IDContratto", dgNumeric, False
            .Add "IDContrattoRiga", "IDContrattoRiga", dgNumeric, False
            
            .Add "IDAnagrafica", "IDAnagrafica", dgNumeric, False
            .Add "CodiceAnagraficaCliente", "Codice cliente", dgchar, True, 1800
            .Add "Anagrafica", "Cliente/Industria", dgchar, True, 3000
            .Add "Nome", "Nome cliente/Industria", dgchar, False, 1500
            .Add "PartitaIva", "Partita I.V.A. cliente/Industria", dgchar, False, 1500
            .Add "CodiceFiscale", "Codice fiscale cliente/Industria", dgchar, False, 1500
            .Add "IDDestinazioneDiversa", "IDDestinazioneDiversa", dgNumeric, False
            .Add "SitoPerAnagrafica", "Destinazione diversa", dgchar, True, 1500
            
            .Add "IDAnagraficaCooperativa", "IDAnagraficaCooperativa", dgNumeric, False
            .Add "CodiceAnagraficaCooperativa", "Codice cooperativa", dgchar, True, 1800
            .Add "AnagraficaCooperativa", "Cooperativa", dgchar, True, 3500
            .Add "NomeCooperativa", "Nome Cooperativa", dgchar, False, 1500

            .Add "IDAnagraficaSocio", "IDAnagraficaSocio", dgNumeric, False
            .Add "CodiceAnagraficaSocio", "Codice socio", dgchar, True, 1800
            .Add "AnagraficaSocio", "Socio", dgchar, True, 3500
            .Add "NomeSocio", "Nome socio", dgchar, False, 1500
            .Add "Acquistato", "Acquisto", dgBoolean, True, 1500
            
            .Add "NumeroCertificato", "Numero certificato", dgchar, True, 2000
            .Add "DataCertificato", "Data certificato", dgDate, True, 2000
        
            .Add "NumeroDocumentoSocio", "Numero DDT socio", dgchar, True, 2000
            .Add "DataDocumentoSocio", "Data DDT socio", dgDate, True, 2000
            
            'RIFERIMENTO DOCUMENTO DI TRASPORTO
            .Add "IDOggettoDDT", "IDOggettoDDT", dgNumeric, False
            .Add "IDTipoOggettoDDT", "IDTipoOggettoDDT", dgNumeric, False
            .Add "DataDocumentoDDT", "Data DDT collegato", dgDate, False, 2000
            .Add "NumeroDocumentoDDT", "Numero DDT collegato", dgNumeric, False, 2000
            .Add "PrefissoSezionaleDocumentoDDT", "Sezionale DDT Collegato", dgchar, False, 2000
            
            'RIFERIMENTO FATTURA DIFFERITA
            .Add "IDOggettoFD", "IDOggettoFD", dgNumeric, False
            .Add "IDTipoOggettoFD", "IDTipoOggettoFD", dgNumeric, False
            .Add "DataDocumentoFD", "Data FD collegato", dgDate, False, 2000
            .Add "NumeroDocumentoFD", "Numero FD collegato", dgNumeric, False, 2000
            .Add "PrefissoSezionaleDocumentoFD", "Sezionale FD Collegato", dgchar, False, 2000
            
            'RIFERIMENTO CONTRATTO
            .Add "DataDocumentoContratto", "Data contratto collegato", dgDate, False, 2000
            .Add "NumeroDocumentoContratto", "Numero contratto collegato", dgNumeric, False, 2000
            .Add "NumeroNsDocumentoContratto", "Numero ns contratto collegato", dgchar, False, 2000
            .Add "NumeroVsDocumentoContratto", "Numero vs contratto collegato", dgchar, False, 2000
            
            'LOTTO DI PRODUZIONE
            .Add "IDRV_PO01_PeriodoCampagna", "IDRV_PO01_PeriodoCampagna", dgNumeric, False
            .Add "IDRV_PO01_Varieta", "IDRV_PO01_Varieta", dgNumeric, False
            .Add "IDRV_PO01_FamigliaProdotti", "IDRV_PO01_FamigliaProdotti", dgNumeric, False
            .Add "CodiceLotto", "Codice lotto di produzione", dgchar, False, 2000
            .Add "Varieta", "Varietà lotto di produzione", dgchar, False, 2000
            .Add "FamigliaProdotti", "Famiglia prodotti lotto di produzione", dgchar, False, 2000
            .Add "PeriodoCampagna", "Periodo di campagna lotto di produzione", dgchar, False, 2000
            .Add "AnnoDiRiferimentoPeriodoDiCampagna", "Anno periodo di campagna", dgNumeric, False, 2000
            .Add "DataInizioPeriodoDiCampagna", "Data inizio periodo di campagna", dgDate, False, 2000
            .Add "DataFinePeriodoDiCampagna", "Data fine periodo di campagna", dgDate, False, 2000
            
            'ALTRI DATI CERTIFICATO
            .Add "DescrizioneArticolo", "Descrizione articolo", dgchar, False, 2000
            .Add "ScartoPesoLordo", "Scarto", dgDouble, False, 2000
            .Add "PesoNettoCalcolato", "Peso di fatturazione", dgDouble, False, 2000
            .Add "ImportoUnitario", "Importo unitario", dgDouble, False, 2000
            .Add "TotaleRiga", "Totale riga netto I.V.A.", dgDouble, False, 2000
            
        End With
    End If
    Me.BrwMain.LoadUserSettings
    Me.BrwMain.Refresh
    
Exit Sub
ERR_LOAD_COLUMN:
    MsgBox Err.Description, vbCritical, "LOAD_COLUMN"
End Sub

Private Function GET_SEZ_PER_CLIENTE(IDCliente As Long) As Long
On Error GoTo ERR_GET_SEZ_PER_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_SEZ_PER_CLIENTE = 0

sSQL = "SELECT IDRV_POConfigurazioneCliente, IDSezionalePerDDT "
sSQL = sSQL & " FROM RV_POConfigurazioneCliente "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm

sSQL = sSQL & " AND IDAnagrafica=" & IDCliente

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerDDT)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_SEZ_PER_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_SEZ_PER_CLIENTE"

End Function
Private Function GET_DESTINAZIONE_PER_CLIENTE(IDCliente As Long) As Long
On Error GoTo ERR_GET_DESTINAZIONE_PER_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_DESTINAZIONE_PER_CLIENTE = 0

sSQL = "SELECT IDAnagrafica, IDAzienda, IDSitoPerAnagrafica "
sSQL = sSQL & " FROM Cliente "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDCliente

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_DESTINAZIONE_PER_CLIENTE = fnNotNullN(rs!IDSitoPerAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_DESTINAZIONE_PER_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_DESTINAZIONE_PER_CLIENTE"
End Function


Private Function fnGetTaraImballo(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Tara FROM Articolo WHERE "
sSQL = sSQL & "IDArticolo = " & IDArticolo

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF = False Then
    If IsNull(rs!Tara) Then
        fnGetTaraImballo = 0
    Else
        fnGetTaraImballo = rs!Tara
    End If
    
Else
    
        fnGetTaraImballo = 0
    
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function CONTROLLO_NUM_CERT(IDCliente As Long, IDDestinazione As Long, DataCertificato As String, NumeroCertificato As String, ID As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim DataInizio As String
Dim DataFine As String

CONTROLLO_NUM_CERT = False

DataInizio = "01/01/" & Year(DataCertificato)
DataFine = "31/12/" & Year(DataCertificato)

sSQL = "SELECT IDRV_POCertificato FROM RV_POCertificato "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDCliente
sSQL = sSQL & " AND IDDestinazioneDiversa=" & IDDestinazione
sSQL = sSQL & " AND NumeroCertificato=" & fnNormString(NumeroCertificato)
sSQL = sSQL & " AND DataCertificato>=" & fnNormDate(DataInizio)
sSQL = sSQL & " AND DataCertificato<=" & fnNormDate(DataFine)
If (ID > 0) Then
    sSQL = sSQL & " AND IDRV_POCertificato<>" & ID
End If
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    CONTROLLO_NUM_CERT = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function CONTROLLO_DATA_CERT_ESERCIZIO_IN_CORSO(DataCertificato As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim DataInizio As String
Dim DataFine As String

CONTROLLO_DATA_CERT_ESERCIZIO_IN_CORSO = False

sSQL = "SELECT * FROM Esercizio "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDTipoEsercizio=1"
sSQL = sSQL & " AND DataInizio<=" & fnNormDate(DataCertificato)
sSQL = sSQL & " AND DataFine>=" & fnNormDate(DataCertificato)
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    CONTROLLO_DATA_CERT_ESERCIZIO_IN_CORSO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LOTTO_PROD_SINGOLO() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroElencoLotti As Long

NumeroElencoLotti = 0

sSQL = "SELECT RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna, RV_PO01_LottoCampagna.IDAzienda, RV_PO01_LottoCampagna.IDSocio, RV_PO01_FamigliaProdotti.UtilizzaNelCertificato, RV_PO01_LottoCampagna.Chiuso, "
sSQL = sSQL & "RV_PO01_LottoCampagna.Provvisorio, RV_PO01_LottoCampagna.IDRV_PO01_Varieta "
sSQL = sSQL & " FROM RV_PO01_LottoCampagna INNER JOIN "
sSQL = sSQL & "RV_PO01_FamigliaProdotti ON RV_PO01_LottoCampagna.IDRV_PO01_FamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDSocio=" & frmMain.CDSocio.KeyFieldID
sSQL = sSQL & " AND UtilizzaNelCertificato=1 "
sSQL = sSQL & " AND ((Provvisorio=0 OR Provvisorio IS NULL))"
sSQL = sSQL & " AND ((VirtualDelete=0 OR VirtualDelete IS NULL))"
sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(0)
sSQL = sSQL & " AND IDRV_PO01_Varieta=" & LINK_VARIETA_ART_CONTRATTO

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    NumeroElencoLotti = NumeroElencoLotti + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

If (NumeroElencoLotti = 1) Then
    sSQL = "SELECT RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna, RV_PO01_LottoCampagna.IDAzienda, RV_PO01_LottoCampagna.IDSocio, RV_PO01_FamigliaProdotti.UtilizzaNelCertificato, RV_PO01_LottoCampagna.Chiuso, "
    sSQL = sSQL & "RV_PO01_LottoCampagna.Provvisorio, RV_PO01_LottoCampagna.IDRV_PO01_Varieta "
    sSQL = sSQL & " FROM RV_PO01_LottoCampagna INNER JOIN "
    sSQL = sSQL & "RV_PO01_FamigliaProdotti ON RV_PO01_LottoCampagna.IDRV_PO01_FamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDSocio=" & frmMain.CDSocio.KeyFieldID
    sSQL = sSQL & " AND UtilizzaNelCertificato=1 "
    sSQL = sSQL & " AND ((Provvisorio=0 OR Provvisorio IS NULL))"
    sSQL = sSQL & " AND ((VirtualDelete=0 OR VirtualDelete IS NULL))"
    sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(0)
    sSQL = sSQL & " AND IDRV_PO01_Varieta=" & LINK_VARIETA_ART_CONTRATTO
    
    Set rs = Cn.OpenResultset(sSQL)

    If Not rs.EOF Then
        GET_LOTTO_PROD_SINGOLO = fnNotNullN(rs!IDRV_PO01_LottoCampagna)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If

End Function
Private Sub GET_ANAGRAFICA_COOPERATIVA(stringa As String, tipo As Long, Optional forza As Boolean = False, Optional codice As String = "", Optional descrizione As String = "")
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long

If (forza = False) Then
    If (Len(stringa) = 0) Then Exit Sub
End If

NumeroRecord = 0

sSQL = "SELECT COUNT(IDAnagraficaFatturazione) as Numero "
sSQL = sSQL & "FROM RV_POIEAnagraficaCooperativaDaLibroSoci "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
If (forza = False) Then
    If Len(stringa) > 0 Then
        If (tipo = 1) Then
            sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + stringa + "%")
        End If
        If (tipo = 2) Then
            sSQL = sSQL & " AND DenominazioneCompleta LIKE " + fnNormString("%" + stringa + "%")
        End If
    End If
Else
    If (Len(codice) > 0) Then
        sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + codice + "%")
    End If
    If (Len(descrizione) > 0) Then
        sSQL = sSQL & " AND DenominazioneCompleta LIKE " + fnNormString("%" + descrizione + "%")
    End If
End If

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    NumeroRecord = fnNotNullN(rs!Numero)
End If

rs.CloseResultset
Set rs = Nothing

If (NumeroRecord = 1) Then
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM RV_POIEAnagraficaCooperativaDaLibroSoci "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    If (forza = False) Then
        If Len(stringa) > 0 Then
            If (tipo = 1) Then
                sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + stringa + "%")
            End If
            If (tipo = 2) Then
                sSQL = sSQL & " AND DenominazioneCompleta LIKE " + fnNormString("%" + stringa + "%")
            End If
        End If
    Else
        If (Len(codice) > 0) Then
            sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + codice + "%")
        End If
        If (Len(descrizione) > 0) Then
            sSQL = sSQL & " AND DenominazioneCompleta LIKE " + fnNormString("%" + descrizione + "%")
        End If
    End If
    Set rs = Cn.OpenResultset(sSQL)

    If Not rs.EOF Then
        Me.CDSocioFatt.Load fnNotNullN(rs!IDAnagraficaFatturazione)
    End If

    rs.CloseResultset
    Set rs = Nothing
Else
    frmSelAnagraficaCoop.Show vbModal
    If LINK_ANA_COOP_SEL > 0 Then
        Me.CDSocioFatt.Load LINK_ANA_COOP_SEL
    End If
End If

End Sub
Private Sub GET_ANAGRAFICA_SOCIO(stringa As String, tipo As Long, Optional forza As Boolean = False, Optional codice As String = "", Optional descrizione As String = "")
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long
Dim DataControlloRevoca As String

If (forza = False) Then
    If (Len(stringa) = 0) Then Exit Sub
End If

NumeroRecord = 0

DataControlloRevoca = DateAdd("m", -NumeroMesiPerDataRevocaCertificato, Date)

sSQL = "SELECT COUNT(IDAnagrafica) as Numero "
If (AttivaSelezioneSocioCertPerVarieta = 0) Then
    sSQL = sSQL & "FROM RV_POIEAnagraficaSocio "
Else
    If (LINK_VARIETA_ART_CONTRATTO > 0) Then
        sSQL = sSQL & "FROM RV_POIEAnagraficaSocioPerVarieta "
    Else
        sSQL = sSQL & "FROM RV_POIEAnagraficaSocio "
    End If
End If
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND ((DataUscita IS NULL) OR (DataUscita>" & fnNormDate(DataControlloRevoca) & "))"
If (Me.CDSocioFatt.KeyFieldID > 0) Then
    sSQL = sSQL & " AND IDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
End If
If (forza = False) Then
    If Len(stringa) > 0 Then
        If (tipo = 1) Then
            sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + stringa + "%")
        End If
        If (tipo = 2) Then
            sSQL = sSQL & " AND Anagrafica LIKE " + fnNormString("%" + stringa + "%")
        End If
    End If
Else
    If (Len(codice) > 0) Then
        sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + codice + "%")
    End If
    If (Len(descrizione) > 0) Then
        sSQL = sSQL & " AND Anagrafica LIKE " + fnNormString("%" + descrizione + "%")
    End If
End If
If (AttivaSelezioneSocioCertPerVarieta = 1) Then
    If (LINK_VARIETA_ART_CONTRATTO > 0) Then
        sSQL = sSQL & " AND Provvisorio=0"
        sSQL = sSQL & " AND Chiuso=0"
        sSQL = sSQL & " AND IDRV_PO01_Varieta=" & LINK_VARIETA_ART_CONTRATTO
    End If
End If
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    NumeroRecord = fnNotNullN(rs!Numero)
End If

rs.CloseResultset
Set rs = Nothing

If (NumeroRecord = 1) Then
    sSQL = "SELECT * "
    If (AttivaSelezioneSocioCertPerVarieta = 0) Then
        sSQL = sSQL & "FROM RV_POIEAnagraficaSocio "
    Else
        If (LINK_VARIETA_ART_CONTRATTO > 0) Then
            sSQL = sSQL & "FROM RV_POIEAnagraficaSocioPerVarieta "
        Else
            sSQL = sSQL & "FROM RV_POIEAnagraficaSocio "
        End If
    End If
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND ((DataUscita IS NULL) OR (DataUscita>" & fnNormDate(DataControlloRevoca) & "))"
    If (Me.CDSocioFatt.KeyFieldID > 0) Then
        sSQL = sSQL & " AND IDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
    End If
    If (forza = False) Then
        If Len(stringa) > 0 Then
            If (tipo = 1) Then
                sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + stringa + "%")
            End If
            If (tipo = 2) Then
                sSQL = sSQL & " AND Anagrafica LIKE " + fnNormString("%" + stringa + "%")
            End If
        End If
    Else
        If (Len(codice) > 0) Then
            sSQL = sSQL & " AND Codice LIKE " + fnNormString("%" + codice + "%")
        End If
        If (Len(descrizione) > 0) Then
            sSQL = sSQL & " AND Anagrafica LIKE " + fnNormString("%" + descrizione + "%")
        End If
    End If
    If (AttivaSelezioneSocioCertPerVarieta = 1) Then
        If (LINK_VARIETA_ART_CONTRATTO > 0) Then
            sSQL = sSQL & " AND Provvisorio=0"
            sSQL = sSQL & " AND Chiuso=0"
            sSQL = sSQL & " AND IDRV_PO01_Varieta=" & LINK_VARIETA_ART_CONTRATTO
        End If
    End If
    Set rs = Cn.OpenResultset(sSQL)

    If Not rs.EOF Then
        Me.CDSocio.Load fnNotNullN(rs!IDAnagrafica)
    End If

    rs.CloseResultset
    Set rs = Nothing
Else
    frmSelAnagraficaSocio.Show vbModal
    If LINK_ANA_SOCIO_SEL > 0 Then
        Me.CDSocio.Load LINK_ANA_SOCIO_SEL
    End If
End If

End Sub
Private Sub CONTROLLO_COOP_COME_CLIENTE(idcoop As Long)
On Error GoTo ERR_CONTROLLO_COOP_COME_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim CreaAnagraficaCliente As Boolean
Dim objAna As dmtRegAna.CRegAnagrafica

If idcoop = 0 Then Exit Sub

CreaAnagraficaCliente = False

sSQL = "SELECT IDAnagrafica FROM Cliente "
sSQL = sSQL & " WHERE IDAnagrafica=" & idcoop
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    CreaAnagraficaCliente = True
End If

rs.CloseResultset
Set rs = Nothing


If (CreaAnagraficaCliente = True) Then
    
    Set objAna = New dmtRegAna.CRegAnagrafica
    
    objAna.Connection = TheApp.Database.Connection
    
    objAna.Read idcoop
    
    'If (objAna.IDAnagrafica > 0) Then
        
        objAna.Field "IDAzienda", TheApp.IDFirm, "Cliente" 'Richiesto
        objAna.Field "IDTipoAnagrafica", 2, "Cliente" 'Richiesto
        objAna.Field "IDPDCContabile", GET_LINK_PDC_CLIENTE, "Cliente"
        
        objAna.Field "DataUltimaVariazione", Date, "Cliente" 'Richiesto
        objAna.Field "IDUtenteUltimaVariazione", TheApp.IDUser, "Cliente" 'Richiesto
        objAna.Field "VirtualDelete", 0, "Cliente" 'Richiesto
        
        objAna.Field "DataUltimaVariazione", Date, "Cliente"
        objAna.Field "IDUtenteUltimaVariazione", TheApp.IDUser, "Cliente"
        objAna.Field "VirtualDelete", 0, "Cliente"
        
        objAna.Update
        
    'End If
    
    Set objAna = Nothing
End If
Exit Sub
ERR_CONTROLLO_COOP_COME_CLIENTE:
    MsgBox Err.Description, vbCritical, "CONTROLLO_COOP_COME_CLIENTE"
End Sub
Private Function GET_LINK_PDC_CLIENTE() As Long
On Error GoTo ERR_GET_LINK_PDC_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDEsercizio As Long
Dim IDPianoDeiConti As Long

GET_LINK_PDC_CLIENTE = 0

IDEsercizio = fnGetEsercizio(Date)

If IDEsercizio = 0 Then Exit Function

sSQL = "SELECT IDPianoDeiConti FROM PianoDeiConti"
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND TipoPDC=1"
sSQL = sSQL & " AND IDEsercizio=" & IDEsercizio

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    IDPianoDeiConti = fnNotNullN(rs!IDPianoDeiConti)
End If

rs.CloseResultset
Set rs = Nothing

If (IDPianoDeiConti > 0) Then
    sSQL = "SELECT IDElementoPDC FROM ElementoPDC "
    sSQL = sSQL & "WHERE IDTipoConto=6 "
    sSQL = sSQL & " AND IDPianoDeiConti=" & IDPianoDeiConti
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        GET_LINK_PDC_CLIENTE = fnNotNullN(rs!IDElementoPDC)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If
Exit Function
ERR_GET_LINK_PDC_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_LINK_PDC_CLIENTE"

End Function
Private Sub RECUPERA_CONFIG_CAUS_XML()
On Error GoTo ERR_RECUPERA_CONFIG_CAUS_XML
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Rip_InXMLRifLetteraIntento = 0
Rip_InXMLRifNoteIva = 0
Rip_InXMLRifNota01Doc = 0
Rip_InXMLRifNota02Doc = 0
Rip_InXMLRifNota03Doc = 0
Rip_InXMLRifNotaDoc = 0
Rip_InXMLRifIstrMitt = 0
Rip_InXMLRifVettSucc = 0
Rip_InXMLRifAgenziaTrasp = 0
Rip_InXMLRifTargaAutoMezzo = 0

sSQL = "SELECT RiportaInXMLRifLetteraIntento, RiportaInXMLRifNoteIva, RiportaInXMLRifNota01Doc, "
sSQL = sSQL & "RiportaInXMLRifNota02Doc, RiportaInXMLRifNota03Doc, RiportaInXMLRifNotaDoc, "
sSQL = sSQL & "RiportaInXMLRifIstrMitt, RiportaInXMLRifVettSucc, RiportaInXMLRifAgenziaTrasp, "
sSQL = sSQL & "RiportaInXMLRifTargaAutoMezzo"
sSQL = sSQL & " FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Rip_InXMLRifLetteraIntento = fnNotNullN(rs!RiportaInXMLRifLetteraIntento)
    Rip_InXMLRifNoteIva = fnNotNullN(rs!RiportaInXMLRifNoteIva)
    Rip_InXMLRifNota01Doc = fnNotNullN(rs!RiportaInXMLRifNota01Doc)
    Rip_InXMLRifNota02Doc = fnNotNullN(rs!RiportaInXMLRifNota02Doc)
    Rip_InXMLRifNota03Doc = fnNotNullN(rs!RiportaInXMLRifNota03Doc)
    Rip_InXMLRifNotaDoc = fnNotNullN(rs!RiportaInXMLRifNotaDoc)
    Rip_InXMLRifIstrMitt = fnNotNullN(rs!RiportaInXMLRifIstrMitt)
    Rip_InXMLRifVettSucc = fnNotNullN(rs!RiportaInXMLRifVettSucc)
    Rip_InXMLRifAgenziaTrasp = fnNotNullN(rs!RiportaInXMLRifAgenziaTrasp)
    Rip_InXMLRifTargaAutoMezzo = fnNotNullN(rs!RiportaInXMLRifTargaAutoMezzo)
End If
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_RECUPERA_CONFIG_CAUS_XML:
    MsgBox Err.Description, vbCritical, "RECUPERA_CONFIG_CAUS_XML"
End Sub

