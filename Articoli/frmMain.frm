VERSION 5.00
Object = "{CF6397E3-591D-11D2-8B11-00C02680407E}#1.0#0"; "DMTFormExtender.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "&Conferma"
      Height          =   375
      Left            =   9000
      TabIndex        =   38
      Top             =   9960
      Width           =   1245
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   10440
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   9960
      Width           =   1245
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9855
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   17383
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "COOPERATIVA"
      TabPicture(0)   =   "frmMain.frx":4781A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraGestioneVivaio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraSermac2Vie"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "AZIENDA"
      TabPicture(1)   =   "frmMain.frx":47836
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraAzienda"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "FARMACIA"
      TabPicture(2)   =   "frmMain.frx":47852
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraSermac2Vie 
         Caption         =   "Parametri sermac 2 vie"
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
         Height          =   855
         Left            =   7680
         TabIndex        =   133
         Top             =   4320
         Width           =   3855
         Begin VB.TextBox txtCalSermac2Vie 
            Height          =   315
            Left            =   1680
            TabIndex        =   37
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calibro macchina"
            Height          =   315
            Index           =   17
            Left            =   120
            TabIndex        =   134
            Top             =   360
            Width           =   1485
         End
      End
      Begin VB.Frame fraGestioneVivaio 
         Caption         =   "Gestione vivaio"
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
         Height          =   4095
         Left            =   7680
         TabIndex        =   122
         Top             =   360
         Width           =   3855
         Begin DMTEDITNUMLib.dmtNumber txtNumeroPianali 
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTEDITNUMLib.dmtNumber txtNPiantePerPianale 
            Height          =   315
            Left            =   2040
            TabIndex        =   28
            Top             =   600
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTEDITNUMLib.dmtNumber txtDaSettimana 
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTEDITNUMLib.dmtNumber txtASettimana 
            Height          =   315
            Left            =   1440
            TabIndex        =   30
            Top             =   1200
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTEDITNUMLib.dmtNumber txtDiametro 
            Height          =   315
            Left            =   2760
            TabIndex        =   31
            Top             =   1200
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTDataCmb.DMTCombo cboTipoPianta 
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   1800
            Width           =   3615
            _ExtentX        =   6376
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
         Begin DMTDataCmb.DMTCombo cboArticoloCarrello 
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   2400
            Width           =   3615
            _ExtentX        =   6376
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
         Begin DMTDataCmb.DMTCombo cboArticoloPianale 
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Top             =   3000
            Width           =   3615
            _ExtentX        =   6376
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
         Begin DMTDataCmb.DMTCombo cboArticoloProlunga 
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   3600
            Width           =   2655
            _ExtentX        =   4683
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
         Begin DMTEDITNUMLib.dmtNumber txtQuantitaProlunga 
            Height          =   315
            Left            =   2880
            TabIndex        =   36
            Top             =   3600
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "Q.tà"
            Height          =   255
            Index           =   14
            Left            =   2880
            TabIndex        =   132
            ToolTipText     =   "Quantità per carrello"
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Articolo prolunga predefinito"
            Height          =   210
            Index           =   6
            Left            =   120
            TabIndex        =   131
            Top             =   3360
            Width           =   2775
         End
         Begin VB.Label Label9 
            Caption         =   "Articolo pianale predefinito"
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   130
            Top             =   2760
            Width           =   3615
         End
         Begin VB.Label Label9 
            Caption         =   "Carrello predefinito"
            Height          =   210
            Index           =   4
            Left            =   120
            TabIndex        =   129
            Top             =   2160
            Width           =   3495
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo pianta"
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   128
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Diametro"
            Height          =   210
            Index           =   2
            Left            =   2760
            TabIndex        =   127
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Settimana a"
            Height          =   210
            Index           =   1
            Left            =   1440
            TabIndex        =   126
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Settimana da"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   125
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Piante per pianale"
            Height          =   210
            Left            =   2040
            TabIndex        =   124
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Pianali per carrello"
            Height          =   210
            Left            =   120
            TabIndex        =   123
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   112
         Top             =   360
         Width           =   5250
         Begin DMTEDITNUMLib.dmtNumber txtNumeroRegistrazione 
            Height          =   315
            Left            =   3120
            TabIndex        =   64
            Top             =   2400
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            AllowEmpty      =   0   'False
         End
         Begin VB.TextBox txtNomePresidio 
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   5055
         End
         Begin VB.TextBox txtNumeroPresidio 
            Height          =   285
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox chkPresidio 
            Caption         =   "Presidio sanitario"
            Height          =   255
            Left            =   3120
            TabIndex        =   60
            Top             =   1200
            Width           =   1935
         End
         Begin DMTDataCmb.DMTCombo cboPrincipioAttivoFarm 
            Height          =   315
            Left            =   120
            TabIndex        =   63
            Top             =   2400
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin DMTDataCmb.DMTCombo cboClasse 
            Height          =   315
            Left            =   3120
            TabIndex        =   62
            Top             =   1800
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin DMTDataCmb.DMTCombo cboProduttore 
            Height          =   315
            Left            =   120
            TabIndex        =   61
            Top             =   1800
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin DMTEDITNUMLib.dmtNumber txtConversione 
            Height          =   315
            Left            =   1680
            TabIndex        =   59
            Top             =   1200
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            DecimalPlaces   =   5
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "Conversione"
            Height          =   255
            Index           =   6
            Left            =   1680
            TabIndex        =   121
            ToolTipText     =   "Conversione in Kg o in Lt dell'unità di misura della vendita"
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label6 
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   1800
            Width           =   4695
         End
         Begin VB.Label Label1 
            Caption         =   "Nome presidio"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label Label1 
            Caption         =   "Numero presidio"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   117
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Produttore"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   116
            Top             =   1560
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Classe"
            Height          =   255
            Index           =   12
            Left            =   3120
            TabIndex        =   115
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Principio attivo"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   114
            Top             =   2160
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Numero registrazione"
            Height          =   255
            Index           =   10
            Left            =   3120
            TabIndex        =   113
            Top             =   2160
            Width           =   2055
         End
      End
      Begin VB.Frame FraAzienda 
         Height          =   7455
         Left            =   -74880
         TabIndex        =   90
         Top             =   360
         Width           =   10890
         Begin VB.TextBox txt_P 
            Height          =   285
            Left            =   1800
            TabIndex        =   44
            Top             =   3720
            Width           =   615
         End
         Begin VB.TextBox txt_K 
            Height          =   285
            Left            =   2760
            TabIndex        =   45
            Top             =   3720
            Width           =   615
         End
         Begin VB.TextBox txt_N 
            Height          =   285
            Left            =   3720
            TabIndex        =   46
            Top             =   3720
            Width           =   615
         End
         Begin VB.CommandButton cmdNuovo 
            Caption         =   "Nuovo"
            Height          =   375
            Left            =   5160
            TabIndex        =   49
            Top             =   5400
            Width           =   1215
         End
         Begin VB.CommandButton cmdSalva 
            Caption         =   "Salva"
            Height          =   375
            Left            =   5160
            TabIndex        =   48
            Top             =   6120
            Width           =   1215
         End
         Begin VB.CommandButton cmdElimina 
            Caption         =   "Elimina"
            Height          =   375
            Left            =   5160
            TabIndex        =   50
            Top             =   6840
            Width           =   1215
         End
         Begin VB.TextBox txtDoseVolAcqua 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6600
            TabIndex        =   51
            Top             =   1560
            Width           =   4215
         End
         Begin VB.CheckBox chkPassaporto 
            Caption         =   "Passaporto"
            Height          =   255
            Left            =   6600
            TabIndex        =   54
            Top             =   3120
            Width           =   3495
         End
         Begin DMTDataCmb.DMTCombo cboTipoPassaporto 
            Height          =   315
            Left            =   6600
            TabIndex        =   55
            Top             =   3720
            Width           =   4215
            _ExtentX        =   7435
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
         Begin DMTDataCmb.DMTCombo cboUMVend 
            Height          =   315
            Left            =   8760
            TabIndex        =   53
            Top             =   2280
            Width           =   2055
            _ExtentX        =   3625
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
         Begin DMTDataCmb.DMTCombo cboUMAcq 
            Height          =   315
            Left            =   6600
            TabIndex        =   52
            Top             =   2280
            Width           =   2055
            _ExtentX        =   3625
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
         Begin DmtCodDescCtl.DmtCodDesc CDProduttore 
            Height          =   615
            Left            =   120
            TabIndex        =   47
            Top             =   4680
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   1085
            PropCodice      =   $"frmMain.frx":4786E
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"frmMain.frx":478BD
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"frmMain.frx":47914
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin DmtGridCtl.DmtGrid Griglia 
            Height          =   1815
            Left            =   120
            TabIndex        =   91
            Top             =   5400
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3201
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
         Begin DMTDataCmb.DMTCombo cboGruppoProdotti 
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   480
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
         Begin DMTDataCmb.DMTCombo cboFamigliaProdotti 
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   1560
            Width           =   3015
            _ExtentX        =   5318
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
         Begin DMTEDITNUMLib.dmtNumber txtGiorniDiCarenza 
            Height          =   315
            Left            =   3480
            TabIndex        =   43
            Top             =   2640
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
            DecimalPlaces   =   0
         End
         Begin DMTDataCmb.DMTCombo cboPrincipioAttivo 
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   2640
            Width           =   3255
            _ExtentX        =   5741
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
         Begin DMTDataCmb.DMTCombo cboTipoVarieta 
            Height          =   315
            Left            =   3360
            TabIndex        =   41
            Top             =   1560
            Width           =   2895
            _ExtentX        =   5106
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
         Begin DMTDataCmb.DMTCombo cboNomeScientifico 
            Height          =   315
            Left            =   6600
            TabIndex        =   56
            Top             =   4800
            Width           =   4215
            _ExtentX        =   7435
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
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Produttori articolo"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   103
            Top             =   4200
            Width           =   5655
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Articoli fertilizzanti"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   102
            Top             =   3120
            Width           =   5655
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Articoli agrofarmaci"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   101
            Top             =   2040
            Width           =   5655
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Articoli per vendite"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   100
            Top             =   960
            Width           =   5655
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Generalità"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   4
            Left            =   6840
            TabIndex        =   99
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Gestione passaport"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   5
            Left            =   6840
            TabIndex        =   95
            Top             =   2760
            Width           =   3735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Altri dati"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   6
            Left            =   6840
            TabIndex        =   93
            Top             =   4200
            Width           =   3735
         End
         Begin VB.Line Line6 
            Index           =   0
            X1              =   6600
            X2              =   10800
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo varietà"
            Height          =   255
            Index           =   13
            Left            =   3360
            TabIndex        =   111
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Principio attivo"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   110
            Top             =   2400
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "Giorni di carenza"
            Height          =   255
            Index           =   11
            Left            =   3480
            TabIndex        =   109
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Famiglia prodotti"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   108
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Gruppo equivalenza"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   3615
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   120
            X2              =   6240
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   120
            X2              =   6240
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line3 
            Index           =   1
            X1              =   120
            X2              =   6240
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "P"
            Height          =   255
            Index           =   8
            Left            =   1800
            TabIndex        =   106
            Top             =   3480
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "K"
            Height          =   255
            Index           =   7
            Left            =   2760
            TabIndex        =   105
            Top             =   3480
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   104
            Top             =   3480
            Width           =   615
         End
         Begin VB.Line Line4 
            Index           =   1
            X1              =   120
            X2              =   6240
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line Line5 
            X1              =   6480
            X2              =   6480
            Y1              =   7320
            Y2              =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Dose per volume acqua"
            Height          =   255
            Index           =   9
            Left            =   6600
            TabIndex        =   98
            Top             =   1320
            Width           =   3735
         End
         Begin VB.Label Label2 
            Caption         =   "Unita di misura di Acq."
            Height          =   255
            Index           =   8
            Left            =   6600
            TabIndex        =   97
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Unita di misura di Vend."
            Height          =   255
            Index           =   7
            Left            =   8760
            TabIndex        =   96
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Line Line6 
            Index           =   1
            X1              =   6600
            X2              =   10800
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo passaporto"
            Height          =   255
            Index           =   1
            Left            =   6600
            TabIndex        =   94
            Top             =   3480
            Width           =   3615
         End
         Begin VB.Line Line6 
            Index           =   2
            X1              =   6600
            X2              =   10800
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label Label5 
            Caption         =   "Nome scientifico"
            Height          =   255
            Index           =   4
            Left            =   6600
            TabIndex        =   92
            Top             =   4560
            Width           =   3615
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9375
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   7530
         Begin VB.CheckBox chkNoPrezzoMedio 
            Caption         =   "Non partecipa al prezzo medio"
            Height          =   255
            Left            =   3840
            TabIndex        =   141
            Top             =   6120
            Width           =   3495
         End
         Begin VB.CheckBox chkSituazImbConf 
            Alignment       =   1  'Right Justify
            Caption         =   "Situazione imballi conf."
            Height          =   255
            Left            =   4920
            TabIndex        =   25
            Top             =   8160
            Width           =   2415
         End
         Begin VB.CheckBox chkTraccImballi 
            Alignment       =   1  'Right Justify
            Caption         =   "Tracciabilità imballi"
            Height          =   255
            Left            =   4920
            TabIndex        =   24
            Top             =   7800
            Width           =   2415
         End
         Begin VB.CheckBox chkInLiquidazione 
            Caption         =   "Non inserire l'articolo in liquidazione"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   6120
            Width           =   3495
         End
         Begin DMTEDITNUMLib.dmtNumber txtPesoOP 
            Height          =   315
            Left            =   4800
            TabIndex        =   20
            Top             =   7080
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
         Begin DMTDataCmb.DMTCombo cboCategoriaOP 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   7080
            Width           =   4575
            _ExtentX        =   8070
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
         Begin DMTDataCmb.DMTCombo cboCalibro 
            Height          =   315
            Left            =   5280
            TabIndex        =   7
            Top             =   2760
            Width           =   2055
            _ExtentX        =   3625
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
         Begin DMTEDITNUMLib.dmtNumber txtTaraImballo 
            Height          =   315
            Left            =   2640
            TabIndex        =   22
            Top             =   8040
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTDataCmb.DMTCombo cboTipoProdotto 
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   600
            Width           =   3255
            _ExtentX        =   5741
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
         Begin DMTDataCmb.DMTCombo cboUM_Acq 
            Height          =   315
            Left            =   5520
            TabIndex        =   2
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
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
         Begin DMTDataCmb.DMTCombo cboUM_Vend 
            Height          =   315
            Left            =   3600
            TabIndex        =   1
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
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
         Begin DMTDataCmb.DMTCombo cboCategoria 
            Height          =   315
            Left            =   3360
            TabIndex        =   6
            Top             =   2760
            Width           =   1815
            _ExtentX        =   3201
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
         Begin DMTDataCmb.DMTCombo cboTipoLavorazione 
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   2760
            Width           =   3015
            _ExtentX        =   5318
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
         Begin DMTDataCmb.DMTCombo cboTipoImballo 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   8040
            Width           =   2415
            _ExtentX        =   4260
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
         Begin DMTDataCmb.DMTCombo cboTipoPrezzoMedio 
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   4560
            Width           =   7095
            _ExtentX        =   12515
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
         Begin DMTDataCmb.DMTCombo cboUM_Liq 
            Height          =   315
            Left            =   240
            TabIndex        =   11
            Top             =   5160
            Width           =   2055
            _ExtentX        =   3625
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
         Begin DMTEDITNUMLib.dmtNumber txtQuantitaPerCollo 
            Height          =   315
            Left            =   4560
            TabIndex        =   13
            Top             =   5160
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTEDITNUMLib.dmtNumber txtMoltiplicatore 
            Height          =   315
            Left            =   6000
            TabIndex        =   14
            Top             =   5160
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTDataCmb.DMTCombo cboTipoPesoArticolo 
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   3360
            Width           =   3015
            _ExtentX        =   5318
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
         Begin DmtCodDescCtl.DmtCodDesc CDNaturaTransazione 
            Height          =   615
            Left            =   120
            TabIndex        =   26
            Top             =   8760
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   1085
            PropCodice      =   $"frmMain.frx":4796E
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"frmMain.frx":479BD
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"frmMain.frx":47A1D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin DMTEDITNUMLib.dmtNumber txtNumeroFori 
            Height          =   315
            Left            =   3720
            TabIndex        =   23
            Top             =   8040
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin DMTDataCmb.DMTCombo cboAddebitoImballo 
            Height          =   315
            Left            =   5520
            TabIndex        =   68
            Top             =   7440
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
         Begin DMTDataCmb.DMTCombo cboImballoVend 
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Top             =   1560
            Width           =   7095
            _ExtentX        =   12515
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
         Begin DMTDataCmb.DMTCombo cboImballoConf 
            Height          =   315
            Left            =   240
            TabIndex        =   4
            Top             =   2160
            Width           =   7095
            _ExtentX        =   12515
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
         Begin DMTDataCmb.DMTCombo cboCatLiq 
            Height          =   315
            Left            =   2400
            TabIndex        =   12
            Top             =   5160
            Width           =   2055
            _ExtentX        =   3625
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
         Begin VB.CheckBox Check1 
            Caption         =   "Protocollo ICE"
            Height          =   255
            Left            =   4560
            TabIndex        =   18
            Top             =   5760
            Width           =   2895
         End
         Begin DMTDataCmb.DMTCombo cboGruppoEvasioneMix 
            Height          =   315
            Left            =   3360
            TabIndex        =   9
            Top             =   3360
            Width           =   3975
            _ExtentX        =   7011
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
         Begin DMTEDITNUMLib.dmtNumber txtPercAbbQtaConf 
            Height          =   315
            Left            =   240
            TabIndex        =   15
            Top             =   5760
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
         Begin DMTEDITNUMLib.dmtNumber txtPercAbbQtaScarto 
            Height          =   315
            Left            =   2160
            TabIndex        =   16
            Top             =   5760
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            DecFinalZeros   =   -1  'True
            AllowEmpty      =   0   'False
         End
         Begin VB.Label Label2 
            Caption         =   "% abb. Q.tà in  liq. scarto"
            Height          =   255
            Index           =   18
            Left            =   2160
            TabIndex        =   140
            ToolTipText     =   "Percentuale di abbatimento quantità conferita"
            Top             =   5520
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "% abb. Q.tà in  liq."
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   139
            ToolTipText     =   "Percentuale di abbatimento quantità conferita"
            Top             =   5520
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Gruppo per evasione Mix"
            Height          =   255
            Index           =   16
            Left            =   3360
            TabIndex        =   138
            Top             =   3120
            Width           =   3975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PARAMETRI PER O.P."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   73
            Top             =   6480
            Width           =   6495
         End
         Begin VB.Label Label2 
            Caption         =   "Cat. liquidazione"
            Height          =   255
            Index           =   15
            Left            =   2400
            TabIndex        =   137
            ToolTipText     =   "Categoria di liquidazione"
            Top             =   4920
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Imballo predefinito per conferimento/acquisto merce"
            Height          =   195
            Index           =   19
            Left            =   240
            TabIndex        =   136
            Top             =   1920
            Width           =   4530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Imballo predefinito per vendita/lavorazione"
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   135
            Top             =   1365
            Width           =   3720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero fori"
            Height          =   195
            Index           =   16
            Left            =   3720
            TabIndex        =   120
            Top             =   7800
            Width           =   1020
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PARAMETRI GENERALI PREDEFINITI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   74
            Top             =   1080
            Width           =   6495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PARAMETRI GENERALI PER IMBALLO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   71
            Top             =   7440
            Width           =   6495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PARAMETRI DI LIQUIDAZIONE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   70
            Top             =   3960
            Width           =   6495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PARAMETRI INTRASTAT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   69
            Top             =   8520
            Width           =   6495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo imballo"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   89
            Top             =   7800
            Width           =   2355
         End
         Begin VB.Label lblCodiceImballo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   1320
            Width           =   6855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo lavorazione standard"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   87
            Top             =   2520
            Width           =   2940
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria"
            Height          =   195
            Index           =   5
            Left            =   3360
            TabIndex        =   85
            Top             =   2520
            Width           =   1560
         End
         Begin VB.Label Label2 
            Caption         =   "U.M. di vendita"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   84
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "U.M. di acquisto"
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   83
            Top             =   360
            Width           =   1575
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   120
            X2              =   7320
            Y1              =   7560
            Y2              =   7560
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   120
            X2              =   7320
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo prodotto"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   82
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "Calibro"
            Height          =   255
            Index           =   2
            Left            =   5400
            TabIndex        =   81
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Prezzo medio di liquidazione"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   80
            Top             =   4320
            Width           =   7095
         End
         Begin VB.Label Label2 
            Caption         =   "U.M. di liquidazione"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   79
            Top             =   4920
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Q.tà per collo"
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   78
            Top             =   4920
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Moltiplicatore"
            Height          =   255
            Index           =   6
            Left            =   6000
            TabIndex        =   77
            Top             =   4920
            Width           =   1335
         End
         Begin VB.Label lblCategoriaOP 
            Caption         =   "Categoria merceologica O.P."
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   6840
            Width           =   3975
         End
         Begin VB.Label lblPesoStaticoOP 
            Caption         =   "Peso statico O.P."
            Height          =   255
            Left            =   4800
            TabIndex        =   75
            Top             =   6840
            Width           =   1575
         End
         Begin VB.Line Line3 
            Index           =   0
            X1              =   120
            X2              =   7320
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Line Line4 
            Index           =   0
            X1              =   120
            X2              =   7320
            Y1              =   6600
            Y2              =   6600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo peso predefinito"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   72
            Top             =   3120
            Width           =   3000
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   120
            X2              =   7320
            Y1              =   8640
            Y2              =   8640
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tara"
            Height          =   195
            Index           =   4
            Left            =   2640
            TabIndex        =   86
            Top             =   7800
            Width           =   390
         End
      End
   End
   Begin DMTFormExtenderLib.FormExtender FormExtender1 
      Left            =   120
      Top             =   5640
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SW_RESTORE = 9
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub SetForegroundWindow Lib "user32" (ByVal hWnd As Long)
Private Nuovo As Boolean
Private Link_TipoProdotto As Long
Private bLoading As Integer
Private ESISTENZA_OP As Boolean
Private Link_TipoImballo As Long
Private Link_TipoGrezzo As Long
Private Link_TipoScarto As Long
Private Link_TipoCaloPeso As Long
Private Link_TipoAumentoPeso As Long


''''VARIABILI AZIENDA''''''''''''''''''''''''''''''''

Private Link_TipoFitofarmaco As Long
Private Link_TipoFertilizzante As Long
Private Link_GruppoEquivalenzaArticolo As Long
Private rsGriglia As ADODB.Recordset
Private IDRigaProduttore As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Private MODULO_ATTIVATO As Long
Private MODULO_DESCRIZIONE As String
Private Const MODULO_CODICE As String = "GT009"

Private Sub cboFamigliaProdotti_Click()
    
    With Me.cboTipoVarieta
        Set .Database = cnDMT
        .DisplayField = "Varieta"
        .AddFieldKey "IDRV_PO01_Varieta"
        .SQL = "SELECT * FROM RV_PO01_Varieta WHERE IDRV_PO01_FamigliaProdotti=" & Me.cboFamigliaProdotti.CurrentID & "  ORDER BY Varieta"
        .Refresh
    End With
End Sub

Private Sub cboTipoProdotto_Click()

If bLoading = 1 Then
    If Link_TipoProdotto = Link_TipoImballo Then
        Me.cboImballoVend.Enabled = False
        Me.cboImballoConf.Enabled = False
        Me.cboTipoImballo.Enabled = True
        Me.cboAddebitoImballo.Enabled = True
        Me.cboTipoLavorazione.Enabled = False
        Me.Check1.Enabled = False
        Me.cboCategoria.Enabled = False
        Me.cboTipoPesoArticolo.Enabled = False
        Me.txtTaraImballo.Enabled = True
        Me.txtNumeroFori.Enabled = True
    Else
        Me.cboImballoVend.Enabled = True
        Me.cboImballoConf.Enabled = True
        Me.cboTipoImballo.Enabled = False
        Me.Check1.Enabled = True
        Me.cboTipoLavorazione.Enabled = True
        Me.cboAddebitoImballo.Enabled = False
        Me.cboCategoria.Enabled = True
        Me.txtTaraImballo.Enabled = False
        Me.cboTipoPesoArticolo.Enabled = True
        Me.txtNumeroFori.Enabled = False
    End If
Else
    If Me.cboTipoProdotto.CurrentID = Link_TipoImballo Then
        Me.cboImballoVend.Enabled = False
        Me.cboImballoConf.Enabled = False
        Me.cboTipoImballo.Enabled = True
        Me.cboAddebitoImballo.Enabled = True
        Me.cboTipoLavorazione.Enabled = False
        Me.Check1.Enabled = False
        Me.cboCategoria.Enabled = False
        Me.txtTaraImballo.Enabled = True
        Me.cboTipoPesoArticolo.Enabled = False
        Me.txtNumeroFori.Enabled = True
    Else
        Me.cboImballoVend.Enabled = True
        Me.cboImballoConf.Enabled = True
        Me.cboTipoImballo.Enabled = False
        Me.Check1.Enabled = True
        Me.cboTipoLavorazione.Enabled = True
        Me.cboAddebitoImballo.Enabled = False
        Me.cboCategoria.Enabled = True
        Me.txtTaraImballo.Enabled = False
        Me.cboTipoPesoArticolo.Enabled = True
        Me.txtNumeroFori.Enabled = False
    End If
End If

Select Case Me.cboTipoProdotto.CurrentID
    Case Link_TipoGrezzo
        Me.txtPercAbbQtaConf.Enabled = True
        Me.txtPercAbbQtaScarto.Enabled = True
    Case Link_TipoScarto
        Me.txtPercAbbQtaConf.Enabled = True
        Me.txtPercAbbQtaScarto.Enabled = False
    Case Link_TipoCaloPeso
        Me.txtPercAbbQtaConf.Enabled = True
        Me.txtPercAbbQtaScarto.Enabled = False
    Case Link_TipoAumentoPeso
        Me.txtPercAbbQtaConf.Enabled = True
        Me.txtPercAbbQtaScarto.Enabled = False
    Case Else
        Me.txtPercAbbQtaConf.Enabled = False
        Me.txtPercAbbQtaScarto.Enabled = False
End Select


End Sub

Private Sub cmdAnnulla_Click()
    Unload Me
End Sub

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click

    If Me.cboUM_Vend.CurrentID = 0 Then
        MsgBox "Inserire l'unità di misura di vendita", vbInformation, "Inserimento dati"
        Exit Sub
    End If
    
    If Me.cboUM_Acq.CurrentID = 0 Then
        MsgBox "Inserire l'unità di misura di acquisto", vbInformation, "Inserimento dati"
        Exit Sub
    End If


    CONFERMA_COOPERATIVA
    
    CONFERMA_AZIENDA
    
    CONFERMA_FARMACIA
    
    Unload Me

Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
    
End Sub



Private Sub cmdElimina_Click()
On Error Resume Next
Dim sSQL As String
Dim Testo As String

    Testo = "Sei sicuro di voler eliminare la riga?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Elinazione produttore") = vbNo Then Exit Sub


    sSQL = "DELETE FROM RV_PO01_ProduttorePerArticolo WHERE IDRV_PO01_ProduttorePerArticolo=" & IDRigaProduttore
    cnDMT.Execute sSQL
    fncGriglia
    If Not (rsGriglia.BOF And rsGriglia.EOF) Then
        Me.Griglia.MoveLast
    Else
        Me.CDProduttore.Load 0
        Nuovo = True
    End If
End Sub

Private Sub cmdNuovo_Click()
On Error Resume Next
Nuovo = True
Me.CDProduttore.Load 0
Me.CDProduttore.SetFocus
End Sub

Private Sub cmdSalva_Click()
On Error Resume Next
Dim sSQL As String
If Me.CDProduttore.KeyFieldID > 0 Then
    If Nuovo = True Then
        sSQL = "INSERT INTO RV_PO01_ProduttorePerArticolo ("
        sSQL = sSQL & "IDRV_PO01_ProduttorePerArticolo, IDArticolo, IDRV_PO01_Produttore) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_PO01_ProduttorePerArticolo", "IDRV_PO01_ProduttorePerArticolo") & ", "
        sSQL = sSQL & IDArticolo & ", "
        sSQL = sSQL & Me.CDProduttore.KeyFieldID & ")"
    Else
        sSQL = "UPDATE RV_PO01_ProduttorePerArticolo SET "
        sSQL = sSQL & "IDRV_PO01_Produttore=" & Me.CDProduttore.KeyFieldID & " "
        sSQL = sSQL & "WHERE IDRV_PO01_ProduttorePerArticolo=" & IDRigaProduttore
    End If
    
    cnDMT.Execute sSQL
Else
    MsgBox "Inserire il produttore", vbInformation, "Salvataggio dati"
End If

    fncGriglia
    If Not (rsGriglia.BOF And rsGriglia.EOF) Then
        Me.Griglia.MoveLast
    Else
        Me.CDProduttore.Load 0
        Nuovo = True
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load
bLoading = 1

    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    ConnessioneADODBLib
    OpenConnection
    PrelevaAzienda
    Screen.MousePointer = 11
    DoEvents
    
    INIT_COOPERATIVA
    INIT_AZIENDA
    INIT_FARMACIA
    'ESISTENZA_OP = GET_ESISTENZA_PROGRAMMA_OP(13)
    
    
    'ParametroImballo
    
    'fncTipoProdotto
    'fncUMVendita
    'fncUMAcquisto
    'fncUMLiquidazione
    'fncArticoloImballoVendita
    'fncArticoloImballoAcquisto
    'fncTipoImballo
    'fncTipolavorazione
    'fncAddebitoImballo
    'fncTipoCategoria
    'fncCalibro
    'fncTipoPrezzoMedio
    'fncTipoPesoArticolo
    'fncNaturaTransazione
    
    'If ESISTENZA_OP = True Then
    '    fncCategoriamMerceologicaOP
    'End If
    ' Prelevamento dei dati aggiuntivi dell'articolo
    
    PrelevaDati
    PrelevaDatiAzienda
    PrelevaDatiFarmacia
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    Screen.MousePointer = 0
    
    GET_MODULO_ATTIVATO MODULO_CODICE, 80
    
    If MODULO_ATTIVATO = 0 Then
        Me.chkTraccImballi.Enabled = False
    End If
    
bLoading = 0
Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseConnection
End Sub

Private Sub FormExtender1_FormActivate(ByVal Status As Boolean)
    If Not Status Then
        ShowWindow Me.hWnd, SW_RESTORE
        SetForegroundWindow Me.hWnd
        Me.Show
        Me.SetFocus
        cmdConferma.SetFocus
    End If
End Sub

Private Sub PrelevaDati()
On Error GoTo ERR_PrelevaDati
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT CodiceArticolo, Articolo, IDTipoProdotto, RV_POIDImballoVendita,RV_POIDImballoConferimento, RV_POProtocolloICE, RV_POIDTipoImballo, RV_POIDTipoLavorazione, "
    sSQL = sSQL & "IDUnitaDiMisuraVendita, IDUnitaDiMisuraAcquisto, RV_PONonLiquidare, RV_POIDNaturaTransazione,"
    sSQL = sSQL & "RV_POImballoPerAddebito, RV_POIDTipoCategoria, Tara, RV_POIDCalibro, RV_POIDTipoPrezzoMedio, RV_POIDTipoPesoArticolo, "
    sSQL = sSQL & "RV_POIDUnitaDiMisuraLiq, RV_POQuantitaPerCollo, RV_POMoltiplicatore, "
    sSQL = sSQL & "RV_PO13_PesoOP, RV_PO13_IDCategoriaMerceologicaOP, "
    sSQL = sSQL & "RV_POIDUnitaDiMisuraLiq, RV_POQuantitaPerCollo, RV_POMoltiplicatore, RV_PO01_NumeroFori, RV_POTracciabilitaImballo, "
    sSQL = sSQL & "RV_PO01_PianaliPerCarrello, RV_PO01_PiantePerPianale, RV_PO01_SettimanaDa, RV_PO01_SettimanaA, "
    sSQL = sSQL & "RV_PO01_IDTipoPianta, RV_PO01_DiametroVaso, RV_PO01_IDArticoloPianale, RV_PO01_IDArticoloProlunga, "
    sSQL = sSQL & "RV_PO01_IDTipoPedana, QuantitaProlunga, RV_POCalibroSermac2Vie, RV_POIDCategoriaLiquidazione, RV_POImballoSituazStampaConf, RV_POIDGruppoArticoloPerEvasioneMix, "
    sSQL = sSQL & "RV_POPercentualeAbbattimentoLiquidazione, RV_POPercentualeAbbattimentoLiquidazioneScarto, RV_PONonPartecipaPrezzoMedio "
    sSQL = sSQL & "FROM Articolo "
    sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
    
    Set rs = cnDMT.OpenResultset(sSQL)
    
    Me.Caption = rs!CodiceArticolo & " - " & rs!Articolo
    Me.cboTipoImballo.WriteOn IIf(IsNull(rs!RV_POIDTipoImballo), 0, rs!RV_POIDTipoImballo)
    'Me.CDImballoVendita.Load fnNotNullN(rs!RV_POIDImballoVendita)
    
    Me.cboImballoVend.WriteOn fnNotNullN(rs!RV_POIDImballoVendita)
    Me.cboImballoConf.WriteOn fnNotNullN(rs!RV_POIDImballoConferimento)
    Me.cboTipoLavorazione.WriteOn IIf(IsNull(rs!RV_POIDTipoLavorazione), 0, rs!RV_POIDTipoLavorazione)
    Me.Check1.Value = IIf(IsNull(rs!RV_POProtocolloICE), 0, fnNormBoolean(rs!RV_POProtocolloICE))
    Link_TipoProdotto = fnNotNullN(rs!IDTipoProdotto)
    Me.cboAddebitoImballo.WriteOn fnNotNullN(rs!RV_POImballoPerAddebito)
    Me.cboCategoria.WriteOn fnNotNullN(rs!RV_POIDTipoCategoria)
    Me.cboTipoProdotto.WriteOn Link_TipoProdotto
    Me.cboUM_Vend.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
    Me.cboUM_Acq.WriteOn fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
    Me.txtTaraImballo.Value = fnNotNullN(rs!Tara)
    Me.cboCalibro.WriteOn fnNotNullN(rs!RV_POIDCalibro)
    Me.cboTipoPrezzoMedio.WriteOn fnNotNullN(rs!RV_POIDTipoPrezzoMedio)
    Me.cboUM_Liq.WriteOn fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
    Me.txtMoltiplicatore.Value = fnNotNullN(rs!RV_POMoltiplicatore)
    Me.txtQuantitaPerCollo.Value = fnNotNullN(rs!RV_POQuantitaPerCollo)
    Me.cboTipoPesoArticolo.WriteOn fnNotNullN(rs!RV_POIDTipoPesoArticolo)
    Me.CDNaturaTransazione.Load fnNotNullN(rs!RV_POIDNaturaTransazione)
    Me.txtNumeroFori.Value = fnNotNullN(rs!RV_PO01_NumeroFori)
    Me.chkTraccImballi.Value = Abs(fnNotNullN(rs!RV_POTracciabilitaImballo))
    Me.cboCategoriaOP.WriteOn fnNotNullN(rs!RV_PO13_IDCategoriaMerceologicaOP)
    Me.txtPesoOP.Value = fnNotNullN(rs!RV_PO13_PesoOP)
    Me.cboCatLiq.WriteOn fnNotNullN(rs!RV_POIDCategoriaLiquidazione)
    Me.cboGruppoEvasioneMix.WriteOn fnNotNullN(rs!RV_POIDGruppoArticoloPerEvasioneMix)
    Me.chkNoPrezzoMedio.Value = Abs(fnNotNullN(rs!RV_PONonPartecipaPrezzoMedio))
    If fnNotNullN(rs!RV_PONonLiquidare) = 0 Then
        Me.chkInLiquidazione.Value = vbUnchecked
    Else
        Me.chkInLiquidazione.Value = vbChecked
    End If
    
    Me.txtNumeroPianali.Value = fnNotNullN(rs!RV_PO01_PianaliPerCarrello)
    Me.txtNPiantePerPianale.Value = fnNotNullN(rs!RV_PO01_PiantePerPianale)
    Me.txtDaSettimana.Value = fnNotNullN(rs!RV_PO01_SettimanaDa)
    Me.txtASettimana.Value = fnNotNullN(rs!RV_PO01_SettimanaA)
    Me.txtQuantitaProlunga.Value = fnNotNullN(rs!QuantitaProlunga)
    Me.txtDiametro.Value = fnNotNullN(rs!RV_PO01_DiametroVaso)
    Me.cboTipoPianta.WriteOn fnNotNullN(rs!RV_PO01_IDTipoPianta)
    Me.cboArticoloPianale.WriteOn fnNotNullN(rs!RV_PO01_IDArticoloPianale)
    Me.cboArticoloProlunga.WriteOn fnNotNullN(rs!RV_PO01_IDArticoloProlunga)
    Me.cboArticoloCarrello.WriteOn fnNotNullN(rs!RV_PO01_IDTipoPedana)
    Me.txtCalSermac2Vie.Text = fnNotNull(rs!RV_POCalibroSermac2Vie)
    Me.chkSituazImbConf.Value = fnNotNullN(rs!RV_POImballoSituazStampaConf)
    Me.txtPercAbbQtaConf.Value = fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazione)
    Me.txtPercAbbQtaScarto.Value = fnNotNullN(rs!RV_POPercentualeAbbattimentoLiquidazioneScarto)
    rs.CloseResultset
    Set rs = Nothing
    
    
    
Exit Sub
ERR_PrelevaDati:
    MsgBox Err.Description, vbCritical, "Preleva Dati"
    MsgBox "L'errore potrebbe essere causato dai seguenti motivi:" & vbCrLf & "1. Moduli aggiuntivi relativi al programma non aggiornati alla versione richiesta" & vbCrLf & "2. Installazione non andata a buon fine" & vbCrLf & "3. Altri problemi generici" & vbCrLf & "In ogni caso contattare il fornitore per risolvere il problema", vbInformation, "Problemi tecnici"
    
    
End Sub
Private Sub fncTipoProdotto()
On Error GoTo ERR_fncTipoImballo

Exit Sub
ERR_fncTipoImballo:
    MsgBox Err.Description, vbCritical, "fncTipoImballo"

End Sub
Private Sub fncTipoPesoArticolo()
On Error GoTo ERR_fncTipoImballo
    Dim sSQL As String
   

Exit Sub
ERR_fncTipoImballo:
    MsgBox Err.Description, vbCritical, "fncTipoImballo"

End Sub


Private Sub fncUMVendita()
On Error GoTo ERR_fncTipoImballo
    Dim sSQL As String
   
    With Me.cboUM_Vend
        Set .Database = cnDMT
        .DisplayField = "UnitaDiMisura"
        .AddFieldKey "IDUnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Refresh
    End With
Exit Sub
ERR_fncTipoImballo:
    MsgBox Err.Description, vbCritical, "fncTipoImballo"

End Sub
Private Sub fncUMAcquisto()
On Error GoTo ERR_fncTipoImballo
    Dim sSQL As String
   

Exit Sub
ERR_fncTipoImballo:
    MsgBox Err.Description, vbCritical, "fncTipoImballo"

End Sub
Private Sub fncUMLiquidazione()
On Error GoTo ERR_fncTipoImballo
    Dim sSQL As String
   

Exit Sub
ERR_fncTipoImballo:
    MsgBox Err.Description, vbCritical, "fncTipoImballo"

End Sub
Private Sub fncTipoImballo()
On Error GoTo ERR_fncTipoImballo
    Dim sSQL As String
   

Exit Sub
ERR_fncTipoImballo:
    MsgBox Err.Description, vbCritical, "fncTipoImballo"

End Sub
Private Sub fncTipoPrezzoMedio()
On Error GoTo ERR_fncTipoImballo
    Dim sSQL As String
   

Exit Sub
ERR_fncTipoImballo:
    MsgBox Err.Description, vbCritical, "fncTipoImballo"

End Sub
Private Sub fncAddebitoImballo()
On Error GoTo ERR_fncTipoImballo
    Dim sSQL As String
   

Exit Sub
ERR_fncTipoImballo:
    MsgBox Err.Description, vbCritical, "fncTipoImballo"

End Sub
Private Sub fncTipolavorazione()
On Error GoTo ERR_fncTipolavorazione
    Dim sSQL As String
   

Exit Sub
ERR_fncTipolavorazione:
    MsgBox Err.Description, vbCritical, "fncTipolavorazione"

End Sub
Private Sub fncTipoCategoria()
On Error GoTo ERR_fncTipoCategoria
    Dim sSQL As String
   

Exit Sub
ERR_fncTipoCategoria:
    MsgBox Err.Description, vbCritical, "fncTipoCategoria"

End Sub
Private Sub fncCalibro()
On Error GoTo ERR_fncCalibro
    Dim sSQL As String
   

Exit Sub
ERR_fncCalibro:
    MsgBox Err.Description, vbCritical, "fncCalibro"

End Sub
Private Sub fncArticoloImballoVendita()
On Error GoTo ERR_fncArticoloImballo
Dim sSQL As String
   
   
   
Exit Sub
ERR_fncArticoloImballo:
    MsgBox Err.Description, vbCritical, "fncArticoloImballo"

End Sub
Private Sub fncArticoloImballoAcquisto()
On Error GoTo ERR_fncArticoloImballo
Dim sSQL As String
   
   
   
Exit Sub
ERR_fncArticoloImballo:
    MsgBox Err.Description, vbCritical, "fncArticoloImballo"

End Sub

Private Sub ParametroImballo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoImballo, IDTipoGrezzo, IDTipoScarto, IDTipoCaloPeso, IDTipoAumentoPeso   "
sSQL = sSQL & " FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & VarIDFiliale
sSQL = sSQL & " AND IDUtente=0"

Set rs = cnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoImballo = fnNotNullN(rs!IDTipoImballo)
    Link_TipoGrezzo = fnNotNullN(rs!IDTipoGrezzo)
    Link_TipoScarto = fnNotNullN(rs!IDTipoScarto)
    Link_TipoCaloPeso = fnNotNullN(rs!IDTipoCaloPeso)
    Link_TipoAumentoPeso = fnNotNullN(rs!IDTipoAumentoPeso)
Else
    Link_TipoImballo = 0
    Link_TipoGrezzo = 0
    Link_TipoScarto = 0
    Link_TipoCaloPeso = 0
    Link_TipoAumentoPeso = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ReadRegistro()
    Utente = GetSetting("141391218", "Setting", "Valore1")
    Password = GetSetting("141391218", "Setting", "Valore2")
End Sub

Private Function GET_ESISTENZA_PROGRAMMA_OP(IDProgramma As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POProgramma FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=" & IDProgramma

Set rs = cnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_PROGRAMMA_OP = False
Else
    GET_ESISTENZA_PROGRAMMA_OP = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub fncCategoriamMerceologicaOP()
On Error GoTo ERR_fncCategoriamMerceologicaOP
    Dim sSQL As String
   

Exit Sub
ERR_fncCategoriamMerceologicaOP:
    MsgBox Err.Description, vbCritical, "fncCategoriamMerceologicaOP"

End Sub

Private Sub fncNaturaTransazione()
On Error GoTo ERR_fncArticoloImballo
Dim sSQL As String
   

   
   
Exit Sub
ERR_fncArticoloImballo:
    MsgBox Err.Description, vbCritical, "fncArticoloImballo"

End Sub

Private Sub INIT_COOPERATIVA()
Dim sSQL As String
   
   ParametroImballo
   
    With Me.cboTipoProdotto
        Set .Database = cnDMT
        .DisplayField = "TipoProdotto"
        .AddFieldKey "IDTipoProdotto"
        .SQL = "SELECT * FROM TipoProdotto WHERE IDAzienda=" & VarIDAzienda
        .Refresh
    End With


    With Me.cboUM_Vend
        Set .Database = cnDMT
        .DisplayField = "UnitaDiMisura"
        .AddFieldKey "IDUnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Refresh
    End With

    With Me.cboUM_Acq
        Set .Database = cnDMT
        .DisplayField = "UnitaDiMisura"
        .AddFieldKey "IDUnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Refresh
    End With

    With Me.cboUM_Liq
        Set .Database = cnDMT
        .DisplayField = "UnitaDiMisuraCoop"
        .AddFieldKey "IDRV_POUnitaDiMisuraCoop"
        .SQL = "SELECT * FROM RV_POUnitaDiMisuraCoop"
        .Refresh
    End With
    
'    'Imballo vendita
'    With Me.CDImballoVendita
'        Set .Database = oDb
'        '.HwndContainer = Me.hWnd
'        .CodeField = "CodiceArticolo"
'        .DescriptionField = "Articolo"
'        .KeyField = "IDArticolo"
'        .TableName = "Articolo"
'        .Filter = "VirtualDelete = 0 AND IDAzienda = " & VarIDAzienda & " AND IDTipoProdotto=" & Link_TipoImballo
'        .MenuFunctions("EseguiGestione").Enabled = False
'        .MenuFunctions("Ricerca").Enabled = False
'        .PropCodice.Caption = "Codice"
'        .PropDescrizione.Caption = "Descrizione"
'        .CodeCaption4Find = "Codice Articolo"
'        .DescriptionCaption4Find = "Descrizione Articolo"
'
'        .CodeIsNumeric = False
'    End With

    With Me.cboImballoVend
        Set .Database = cnDMT
        .DisplayField = "Descrizione"
        .AddFieldKey "IDArticolo"
        .SQL = "SELECT IDArticolo, CodiceArticolo + ' - ' + Articolo AS Descrizione "
        .SQL = .SQL & "FROM Articolo "
        .SQL = .SQL & "WHERE VirtualDelete = 0 AND IDAzienda = " & VarIDAzienda & " AND IDTipoProdotto=" & Link_TipoImballo
        .Refresh
    End With


'    'Imballo acquisto
'    With Me.CDImballoAcquisto
'        Set .Database = oDb
'        .CodeField = "CodiceArticolo"
'        .DescriptionField = "Articolo"
'        .KeyField = "IDArticolo"
'        .TableName = "Articolo"
'        .Filter = "VirtualDelete = 0 AND IDAzienda = " & VarIDAzienda & " AND IDTipoProdotto=" & Link_TipoImballo
'        .MenuFunctions("EseguiGestione").Enabled = False
'        .PropCodice.Caption = "Codice"
'        .PropDescrizione.Caption = "Descrizione"
'        .CodeCaption4Find = "Codice Articolo"
'        .DescriptionCaption4Find = "Descrizione Articolo"
'        .CodeIsNumeric = False
'    End With


    With Me.cboImballoConf
        Set .Database = cnDMT
        .DisplayField = "Descrizione"
        .AddFieldKey "IDArticolo"
        .SQL = "SELECT IDArticolo, CodiceArticolo + ' - ' + Articolo AS Descrizione "
        .SQL = .SQL & "FROM Articolo "
        .SQL = .SQL & "WHERE VirtualDelete = 0 AND IDAzienda = " & VarIDAzienda & " AND IDTipoProdotto=" & Link_TipoImballo
        .Refresh
    End With

    With Me.cboTipoImballo
        Set .Database = cnDMT
        .DisplayField = "TipoImballo"
        .AddFieldKey "IDRV_POTipoImballo"
        .SQL = "SELECT * FROM RV_POTipoImballo"
        .Refresh
    End With

    With Me.cboTipoLavorazione
        Set .Database = cnDMT
        .DisplayField = "TipoLavorazione"
        .AddFieldKey "IDRV_POTipoLavorazione"
        .SQL = "SELECT * FROM RV_POTipoLavorazione"
        .Refresh
    End With

    With Me.cboAddebitoImballo
        Set .Database = cnDMT
        .DisplayField = "SiNo"
        .AddFieldKey "IDRV_POSiNo"
        .SQL = "SELECT * FROM RV_POSiNo"
        .Refresh
    End With

    With Me.cboCategoria
        Set .Database = cnDMT
        .DisplayField = "TipoCategoria"
        .AddFieldKey "IDRV_POTipoCategoria"
        .SQL = "SELECT * FROM RV_POTipoCategoria"
        .Refresh
    End With

    With Me.cboCalibro
        Set .Database = cnDMT
        .DisplayField = "Calibro"
        .AddFieldKey "IDRV_POCalibro"
        .SQL = "SELECT * FROM RV_POCalibro"
        .Refresh
    End With

    With Me.cboTipoPrezzoMedio
        Set .Database = cnDMT
        .DisplayField = "TipoPrezzoMedio"
        .AddFieldKey "IDRV_POTipoPrezzoMedio"
        .SQL = "SELECT * FROM RV_POTipoPrezzoMedio ORDER BY IDRV_POTipoPrezzoMedio"
        .Refresh
    End With

    With Me.cboTipoPesoArticolo
        Set .Database = cnDMT
        .DisplayField = "TipoPesoArticolo"
        .AddFieldKey "IDRV_POTipoPesoArticolo"
        .SQL = "SELECT * FROM RV_POTipoPesoArticolo"
        .Refresh
    End With

    With Me.CDNaturaTransazione
        Set .Database = oDb
        .HwndContainer = Me.hWnd
        .CodeField = "Codice"
        .DescriptionField = "NaturaTransazione"
        .KeyField = "IDNaturaTransazione"
        .TableName = "NaturaTransazione"
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Natura di transazione"
        .CodeCaption4Find = "Codice "
        .DescriptionCaption4Find = "Natura di transazione"
        .CodeIsNumeric = False
    End With

    With Me.cboCategoriaOP
        Set .Database = cnDMT
        .DisplayField = "CategoriaMerceologicaOP"
        .AddFieldKey "IDRV_PO13CategoriaMerceologicaOP"
        .SQL = "SELECT * FROM RV_PO13CategoriaMerceologicaOP ORDER BY CategoriaMerceologicaOP"
        .Refresh
    End With
    
    With cboTipoPianta
        Set .Database = cnDMT
        .DisplayField = "TipoPianta"
        .AddFieldKey "IDRV_PO01_TipoPianta"
        .SQL = "SELECT * FROM RV_PO01_TipoPianta ORDER BY TipoPianta"
        .Refresh
    End With

    With cboArticoloCarrello
        Set .Database = cnDMT
        .DisplayField = "TipoPedana"
        .AddFieldKey "IDRV_POTipoPedana"
        .SQL = "SELECT * FROM RV_POIETipoPedana WHERE IDAzienda = " & VarIDAzienda & " ORDER BY TipoPedana"
        .Refresh
    End With
    
    With cboArticoloPianale
        Set .Database = cnDMT
        .DisplayField = "Articolo"
        .AddFieldKey "IDArticolo"
        .SQL = "SELECT * FROM Articolo "
        .SQL = .SQL & " WHERE IDAzienda = " & VarIDAzienda
        .SQL = .SQL & " AND IDTipoProdotto=" & Link_TipoImballo
        .SQL = .SQL & " ORDER BY Articolo"
        .Refresh
    End With

    With cboArticoloProlunga
        Set .Database = cnDMT
        .DisplayField = "Articolo"
        .AddFieldKey "IDArticolo"
        .SQL = "SELECT * FROM Articolo "
        .SQL = .SQL & " WHERE IDAzienda = " & VarIDAzienda
        .SQL = .SQL & " AND IDTipoProdotto=" & Link_TipoImballo
        .SQL = .SQL & " ORDER BY Articolo"
        .Refresh
    End With
    
    With Me.cboCatLiq
        Set .Database = cnDMT
        .DisplayField = "CategoriaLiquidazione"
        .AddFieldKey "IDRV_POCategoriaLiquidazione"
        .SQL = "SELECT * FROM RV_POCategoriaLiquidazione ORDER BY CategoriaLiquidazione"
        .Refresh
    End With
    
    With Me.cboGruppoEvasioneMix
        Set .Database = cnDMT
        .DisplayField = "GruppoArticoloPerEvasioneMix"
        .AddFieldKey "IDRV_POGruppoArticoloPerEvasioneMix"
        .SQL = "SELECT * FROM RV_POGruppoArticoloPerEvasioneMix ORDER BY GruppoArticoloPerEvasioneMix"
        .Refresh
    End With
    

End Sub
Private Sub INIT_AZIENDA()
Dim sSQL As String

    With Me.CDProduttore
        Set .Database = oDb
        .HwndContainer = Me.hWnd
        .CodeField = "CodiceProduttore"
        .DescriptionField = "Produttore"
        .KeyField = "IDRV_PO01_Produttore"
        .TableName = "RV_PO01_Produttore"
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice"
        .DescriptionCaption4Find = "Descrizione"
        .CodeIsNumeric = False
    End With

    With Me.cboGruppoProdotti
        Set .Database = cnDMT
        .DisplayField = "GruppoEquivalenzaArticolo"
        .AddFieldKey "IDGruppoEquivalenzaArticolo"
        .SQL = "SELECT * FROM GruppoEquivalenzaArticolo WHERE IDAzienda=" & VarIDAzienda & " ORDER BY GruppoEquivalenzaArticolo"
        .Refresh
    End With

    With Me.cboTipoVarieta
        Set .Database = cnDMT
        .DisplayField = "Varieta"
        .AddFieldKey "IDRV_PO01_Varieta"
        .SQL = "SELECT * FROM RV_PO01_Varieta ORDER BY Varieta"
        .Refresh
    End With

    With Me.cboPrincipioAttivo
        Set .Database = cnDMT
        .DisplayField = "PrincipioAttivo"
        .AddFieldKey "IDRV_PO01_PrincipioAttivo"
        .SQL = "SELECT * FROM RV_PO01_PrincipioAttivo ORDER BY PrincipioAttivo"
        .Refresh
    End With

    With Me.cboFamigliaProdotti
        Set .Database = cnDMT
        .DisplayField = "FamigliaProdotti"
        .AddFieldKey "IDRV_PO01_FamigliaProdotti"
        .SQL = "SELECT * FROM RV_PO01_FamigliaProdotti ORDER BY FamigliaProdotti"
        .Refresh
    End With

    With Me.cboUMAcq
        Set .Database = cnDMT
        .DisplayField = "UnitaDiMisura"
        .AddFieldKey "IDUnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura ORDER BY UnitaDiMisura"
        .Refresh
    End With

    With Me.cboUMVend
        Set .Database = cnDMT
        .DisplayField = "UnitaDiMisura"
        .AddFieldKey "IDUnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura ORDER BY UnitaDiMisura"
        .Refresh
    End With

    With Me.cboTipoPassaporto
        Set .Database = cnDMT
        .DisplayField = "TipoPassaporto"
        .AddFieldKey "IDRV_PO01_TipoPassaporto"
        .SQL = "SELECT * FROM RV_PO01_TipoPassaporto ORDER BY TipoPassaporto"
        .Refresh
    End With

    With Me.cboNomeScientifico
        Set .Database = cnDMT
        .DisplayField = "NomeScientifico"
        .AddFieldKey "IDRV_PO01_ArtNomeScientifico"
        .SQL = "SELECT * FROM RV_PO01_ArtNomeScientifico ORDER BY NomeScientifico"
        .Refresh
    End With
End Sub

Private Sub PrelevaDatiAzienda()
On Error GoTo ERR_PrelevaDati
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT CodiceArticolo, Articolo, RV_PO01_IDVarieta, IDGruppoEquivalenzaArticolo, "
    sSQL = sSQL & "RV_PO01_IDPrincipioAttivo, RV_PO01_GiorniCarenza, RV_PO01_IDFamigliaProdotti, "
    sSQL = sSQL & "RV_PO01_K, RV_PO01_P, RV_PO01_N, RV_PO01_DoseVolAcqua,  "
    sSQL = sSQL & "IDUnitaDiMisuraVendita, IDUnitaDiMisuraAcquisto, RV_PO01_Passaporto, "
    sSQL = sSQL & "RV_PO01_IDTipoPassaporto, RV_PO01_IDArtNomeScientifico "
    sSQL = sSQL & "FROM Articolo "
    sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
    
    Set rs = cnDMT.OpenResultset(sSQL)
    Me.Caption = rs!CodiceArticolo & " - " & rs!Articolo
    Link_GruppoEquivalenzaArticolo = fnNotNullN(rs!IDGruppoEquivalenzaArticolo)
    Me.cboGruppoProdotti.WriteOn Link_GruppoEquivalenzaArticolo
    
    Me.cboTipoVarieta.WriteOn fnNotNullN(rs!RV_PO01_IDVarieta)
    Me.cboFamigliaProdotti.WriteOn fnNotNullN(rs!RV_PO01_IDFamigliaProdotti)
    Me.cboPrincipioAttivo.WriteOn fnNotNullN(rs!RV_PO01_IDPrincipioAttivo)
    Me.txtGiorniDiCarenza.Value = fnNotNullN(rs!RV_PO01_GiorniCarenza)
    Me.txt_K.Text = fnNotNull(rs!RV_PO01_K)
    Me.txt_P.Text = fnNotNull(rs!RV_PO01_P)
    Me.txt_N.Text = fnNotNull(rs!RV_PO01_N)
    Me.txtDoseVolAcqua.Text = fnNotNull(rs!RV_PO01_DoseVolAcqua)
    Me.cboUMAcq.WriteOn fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
    Me.cboUMVend.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
    Me.chkPassaporto.Value = Abs(fnNotNullN(rs!RV_PO01_Passaporto))
    Me.cboTipoPassaporto.WriteOn fnNotNullN(rs!RV_PO01_IDTipoPassaporto)
    Me.cboNomeScientifico.WriteOn fnNotNullN(rs!RV_PO01_IDArtNomeScientifico)
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    fncGriglia
    
    Nuovo = True
Exit Sub
ERR_PrelevaDati:
    MsgBox Err.Description, vbCritical, "Preleva Dati azienda"
End Sub

Private Sub CONFERMA_AZIENDA()
On Error GoTo ERR_cmdConferma_Click
    Dim sSQL As String
    
    

    
    
    sSQL = "UPDATE Articolo SET "
    sSQL = sSQL & "RV_PO01_IDVarieta=" & Me.cboTipoVarieta.CurrentID & ", "
    sSQL = sSQL & "RV_PO01_IDPrincipioAttivo=" & Me.cboPrincipioAttivo.CurrentID & ", "
    sSQL = sSQL & "RV_PO01_IDFamigliaProdotti=" & Me.cboFamigliaProdotti.CurrentID & ", "
    sSQL = sSQL & "RV_PO01_GiorniCarenza=" & Me.txtGiorniDiCarenza.Value & ",  "
    sSQL = sSQL & "RV_PO01_K=" & fnNormString(Me.txt_K.Text) & ", "
    sSQL = sSQL & "RV_PO01_P=" & fnNormString(Me.txt_P.Text) & ", "
    sSQL = sSQL & "RV_PO01_N=" & fnNormString(Me.txt_N.Text) & ", "
    sSQL = sSQL & "RV_PO01_DoseVolAcqua=" & fnNormString(Me.txtDoseVolAcqua.Text) & ", "
    sSQL = sSQL & "RV_PO01_Passaporto=" & Me.chkPassaporto.Value & ", "
    sSQL = sSQL & "RV_PO01_IDTipoPassaporto=" & Me.cboTipoPassaporto.CurrentID & ", "
    sSQL = sSQL & "IDUnitaDiMisuraAcquisto=" & Me.cboUMAcq.CurrentID & ", "
    sSQL = sSQL & "IDUnitaDiMisuraVendita=" & Me.cboUMVend.CurrentID & ", "
    sSQL = sSQL & "IDGruppoEquivalenzaArticolo=" & Me.cboGruppoProdotti.CurrentID & ", "
    sSQL = sSQL & "RV_PO01_IDArtNomeScientifico=" & Me.cboNomeScientifico.CurrentID & " "
    sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
    
    cnDMT.Execute sSQL

Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub
Private Sub CONFERMA_COOPERATIVA()
On Error GoTo ERR_cmdConferma_Click
Dim sSQL As String
    
    sSQL = "UPDATE Articolo SET "
    sSQL = sSQL & "IDTipoProdotto=" & Me.cboTipoProdotto.CurrentID & ", "
    sSQL = sSQL & "IDUnitaDiMisuraVendita=" & Me.cboUM_Vend.CurrentID & ", "
    sSQL = sSQL & "IDUnitaDiMisuraAcquisto=" & Me.cboUM_Acq.CurrentID & ", "
'    sSQL = sSQL & "RV_POIDImballoVendita=" & Me.CDImballoVendita.KeyFieldID & ", "
    sSQL = sSQL & "RV_POIDImballoVendita=" & Me.cboImballoVend.CurrentID & ", "
    sSQL = sSQL & "RV_POIDImballoConferimento=" & Me.cboImballoConf.CurrentID & ", "
    sSQL = sSQL & "RV_POIDTipoImballo=" & Me.cboTipoImballo.CurrentID & ", "
    sSQL = sSQL & "RV_POImballoPerAddebito=" & Me.cboAddebitoImballo.CurrentID & ", "
    sSQL = sSQL & "RV_POIDTipoLavorazione=" & Me.cboTipoLavorazione.CurrentID & ", "
    sSQL = sSQL & "RV_POIDTipoCategoria=" & Me.cboCategoria.CurrentID & ", "
    sSQL = sSQL & "RV_POProtocolloICE=" & fnNormBoolean(Me.Check1.Value) & ", "
    sSQL = sSQL & "RV_POIDCalibro=" & Me.cboCalibro.CurrentID & ", "
    sSQL = sSQL & "RV_POIDTipoPrezzoMedio=" & Me.cboTipoPrezzoMedio.CurrentID & ", "
    sSQL = sSQL & "Tara=" & fnNormNumber(Me.txtTaraImballo.Text) & ", "
    sSQL = sSQL & "RV_POIDUnitaDiMisuraLiq=" & Me.cboUM_Liq.CurrentID & ", "
    sSQL = sSQL & "RV_POQuantitaPerCollo=" & Me.txtQuantitaPerCollo.Value & ", "
    sSQL = sSQL & "RV_PONonLiquidare=" & fnNormBoolean(Me.chkInLiquidazione.Value) & ", "
    sSQL = sSQL & "RV_POIDTipoPesoArticolo=" & Me.cboTipoPesoArticolo.CurrentID & ", "
    sSQL = sSQL & "RV_POIDNaturaTransazione=" & Me.CDNaturaTransazione.KeyFieldID & ", "
    sSQL = sSQL & "RV_PO01_NumeroFori=" & fnNormNumber(Me.txtNumeroFori.Text) & ", "
    sSQL = sSQL & "RV_POTracciabilitaImballo=" & Me.chkTraccImballi.Value & ", "
    sSQL = sSQL & "RV_PO13_IDCategoriaMerceologicaOP=" & Me.cboCategoriaOP.CurrentID & ", "
    sSQL = sSQL & "RV_PO13_PesoOP=" & fnNormNumber(Me.txtPesoOP.Value) & ", "
    sSQL = sSQL & "RV_POMoltiplicatore=" & fnNormNumber(Me.txtMoltiplicatore.Value) & ", "
    sSQL = sSQL & "RV_POIDGruppoArticoloPerEvasioneMix=" & Me.cboGruppoEvasioneMix.CurrentID & ", "
    sSQL = sSQL & "RV_PO01_PianaliPerCarrello=" & fnNormNumber(Me.txtNumeroPianali.Value) & ", "
    sSQL = sSQL & "RV_PO01_PiantePerPianale=" & fnNormNumber(Me.txtNPiantePerPianale.Value) & ", "
    sSQL = sSQL & "RV_PO01_SettimanaDa=" & fnNormNumber(Me.txtDaSettimana.Value) & ", "
    sSQL = sSQL & "RV_PO01_SettimanaA=" & fnNormNumber(Me.txtASettimana.Value) & ", "
    sSQL = sSQL & "RV_PO01_DiametroVaso=" & fnNormNumber(Me.txtDiametro.Value) & ", "
    sSQL = sSQL & "RV_PO01_IDTipoPianta=" & Me.cboTipoPianta.CurrentID & ", "
    sSQL = sSQL & "RV_PO01_IDArticoloPianale=" & Me.cboArticoloPianale.CurrentID & ", "
    sSQL = sSQL & "RV_PO01_IDArticoloProlunga=" & Me.cboArticoloProlunga.CurrentID & ", "
    sSQL = sSQL & "RV_PO01_IDTipoPedana=" & Me.cboArticoloCarrello.CurrentID & ", "
    sSQL = sSQL & "QuantitaProlunga=" & fnNormNumber(Me.txtQuantitaProlunga.Value) & ", "
    sSQL = sSQL & "RV_POCalibroSermac2Vie=" & fnNormString(Me.txtCalSermac2Vie.Text) & ", "
    sSQL = sSQL & "RV_POIDCategoriaLiquidazione=" & Me.cboCatLiq.CurrentID & ", "
    sSQL = sSQL & "RV_POImballoSituazStampaConf=" & Me.chkSituazImbConf.Value & ", "
    sSQL = sSQL & "RV_POPercentualeAbbattimentoLiquidazione=" & fnNormNumber(Me.txtPercAbbQtaConf.Value) & ", "
    sSQL = sSQL & "RV_POPercentualeAbbattimentoLiquidazioneScarto=" & fnNormNumber(Me.txtPercAbbQtaScarto.Value) & ", "
    sSQL = sSQL & "RV_PONonPartecipaPrezzoMedio=" & Abs(Me.chkNoPrezzoMedio.Value)
    sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
    
    cnDMT.Execute sSQL
    
Exit Sub

ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub

Public Sub fncGriglia()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    

    sSQL = "SELECT RV_PO01_ProduttorePerArticolo.IDRV_PO01_ProduttorePerArticolo, RV_PO01_ProduttorePerArticolo.IDRV_PO01_Produttore, "
    sSQL = sSQL & "RV_PO01_ProduttorePerArticolo.IDArticolo , RV_PO01_Produttore.Produttore, RV_PO01_Produttore.CodiceProduttore "
    sSQL = sSQL & "FROM RV_PO01_ProduttorePerArticolo LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PO01_Produttore ON RV_PO01_ProduttorePerArticolo.IDRV_PO01_Produttore = RV_PO01_Produttore.IDRV_PO01_Produttore "
    sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo
    OLDCursor = cnDMT.CursorLocation
    cnDMT.CursorLocation = 3
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, cnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
            'Set rsEvent = rsGriglia2.Data
    
        
    
        With Me.Griglia
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            
            .ColumnsHeader.Clear
            
                    .ColumnsHeader.Add "IDRV_PO01_ProduttorePerArticolo", "ID", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "IDRV_PO01_Produttore", "IDProduttore", dgInteger, False, 500, dgAlignleft
                    .ColumnsHeader.Add "CodiceProduttore", "Codice produttore", dgchar, True, 1500, dgAlignleft
                    .ColumnsHeader.Add "Produttore", "Produttore", dgchar, True, 3000, dgAlignleft
                    
                        
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    cnDMT.CursorLocation = OLDCursor
End Sub

Public Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    Dim VarData As String
    
    
    
    
    'Monta la query SQL per trovare il massimo valore della chiave primaria
    sSQL = "SELECT MAX (" & CampoKey & ") AS MaxID FROM " & Tabella ' & " WHERE " & >=" & VarData
    
    'Apertura del recordset
    Set rs = cnDMT.OpenResultset(sSQL)
    
    'Determina il primo progressivo disponibile
    fnGetNewKey = fnNotNullN(rs.adoColumns("MaxID")) + 1
    If fnGetNewKey <= 0 Then fnGetNewKey = 1

    'Chiude il recordset e distrugge l'oggetto.
    rs.CloseResultset
    Set rs = Nothing
    
End Function

Private Sub Griglia_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
On Error Resume Next
IDRigaProduttore = Me.Griglia("IDRV_PO01_produttorePerArticolo").Value
Me.CDProduttore.Load Me.Griglia("IDRV_PO01_Produttore").Value
Nuovo = False

End Sub
Private Sub INIT_FARMACIA()
Dim sSQL As String

    With Me.cboClasse
        Set .Database = cnDMT
        .DisplayField = "Classe"
        .AddFieldKey "IDRV_PO10_Classe"
        .SQL = "SELECT * FROM RV_PO10_Classe"
        .Refresh
    End With

    With Me.cboPrincipioAttivoFarm
        Set .Database = cnDMT
        .DisplayField = "PrincipioAttivo"
        .AddFieldKey "IDRV_PO10_PrincipioAttivo"
        .SQL = "SELECT * FROM RV_PO10_PrincipioAttivo"
        .Refresh
    End With

    With Me.cboProduttore
        Set .Database = cnDMT
        .DisplayField = "Produttore"
        .AddFieldKey "IDRV_PO10_Produttore"
        .SQL = "SELECT * FROM RV_PO10_Produttore"
        .Refresh
    End With
End Sub


Private Sub PrelevaDatiFarmacia()
On Error GoTo ERR_PrelevaDati
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_PO10_NomePresidio, RV_PO10_NumeroPresidio, RV_PO10_NumeroRegistrazione, "
    sSQL = sSQL & "RV_PO10_PresidioSanitario, RV_PO10_IDClasse, RV_PO10_IDPrincipioAttivo, "
    sSQL = sSQL & "RV_PO10_IDProduttore, RV_PO10_ConversioneKgLt "
    sSQL = sSQL & "FROM Articolo "
    sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
    Set rs = cnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Me.txtNomePresidio.Text = fnNotNull(rs!RV_PO10_NomePresidio)
        Me.txtNumeroPresidio.Text = fnNotNull(rs!RV_PO10_NumeroPresidio)
        Me.txtNumeroRegistrazione.Text = fnNotNull(rs!RV_PO10_NumeroRegistrazione)
        Me.chkPresidio.Value = fnNormBoolean(fnNotNullN(rs!RV_PO10_PresidioSanitario))
        Me.cboClasse.WriteOn fnNotNullN(rs!RV_PO10_IDClasse)
        Me.cboPrincipioAttivoFarm.WriteOn fnNotNullN(rs!RV_PO10_IDPrincipioAttivo)
        Me.cboProduttore.WriteOn fnNotNullN(rs!RV_PO10_IDProduttore)
        Me.txtConversione.Value = fnNotNullN(rs!RV_PO10_ConversioneKgLt)
    End If
Exit Sub
ERR_PrelevaDati:
    MsgBox Err.Description, vbCritical, "Preleva Dati farmacia"
End Sub
Private Sub CONFERMA_FARMACIA()
'On Error GoTo ERR_cmdConferma_Click
Dim sSQL As String
    
    
    sSQL = "UPDATE Articolo SET "
    sSQL = sSQL & "RV_PO10_NomePresidio=" & fnNormString(Me.txtNomePresidio.Text) & ", "
    sSQL = sSQL & "RV_PO10_NumeroPresidio=" & fnNormString(Me.txtNumeroPresidio.Text) & ", "
    sSQL = sSQL & "RV_PO10_NumeroRegistrazione=" & fnNormNumber(Me.txtNumeroRegistrazione.Text) & ", "
    sSQL = sSQL & "RV_PO10_PresidioSanitario=" & fnNormBoolean(Me.chkPresidio.Value) & ", "
    sSQL = sSQL & "RV_PO10_IDClasse=" & Me.cboClasse.CurrentID & ", "
    sSQL = sSQL & "RV_PO10_IDPrincipioAttivo=" & Me.cboPrincipioAttivoFarm.CurrentID & ", "
    sSQL = sSQL & "RV_PO10_IDProduttore=" & Me.cboProduttore.CurrentID & ", "
    sSQL = sSQL & "RV_PO10_ConversioneKgLt=" & fnNormNumber(Me.txtConversione.Text) & " "
    sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
    
    cnDMT.Execute sSQL
    
    
Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub
Private Sub GET_MODULO_ATTIVATO(Codice As String, IdentificativoProgramma As Long)
On Error GoTo ERR_GET_MODULO_ATTIVATO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Attivato, DescrizioneModulo FROM RV_POProgrammaModulo "
sSQL = sSQL & "WHERE CodiceModulo=" & fnNormString(Codice)
sSQL = sSQL & " AND IdentificazioneProgramma=" & IdentificativoProgramma

Set rs = cnDMT.OpenResultset(sSQL)

If rs.EOF Then
    MODULO_ATTIVATO = 0
    MODULO_DESCRIZIONE = ""
Else
    MODULO_ATTIVATO = Abs(fnNotNullN(rs!Attivato))
    MODULO_DESCRIZIONE = fnNotNull(rs!DescrizioneModulo)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_MODULO_ATTIVATO:
    MODULO_ATTIVATO = 0
    MODULO_DESCRIZIONE = ""
End Sub

