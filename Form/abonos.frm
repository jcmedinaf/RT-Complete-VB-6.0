VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmPagosPrestamos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos de Pretamos"
   ClientHeight    =   6585
   ClientLeft      =   4515
   ClientTop       =   915
   ClientWidth     =   8745
   Icon            =   "abonos.frx":0000
   LinkTopic       =   "Form47"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Modificar Cuotas"
      Height          =   6375
      Left            =   8760
      TabIndex        =   31
      Top             =   120
      Width           =   8535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Mensual"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   47
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Quincenal"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   46
         Top             =   1800
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   5280
         TabIndex        =   44
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   211419139
         UpDown          =   -1  'True
         CurrentDate     =   40282
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1320
         Width           =   6015
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "abonos.frx":1002
         Left            =   1080
         List            =   "abonos.frx":1026
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TxtMxC 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         TabIndex        =   32
         Top             =   600
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
         Height          =   375
         Left            =   3000
         TabIndex        =   36
         ToolTipText     =   "Guardar / Actualizar "
         Top             =   5880
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Guardar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":105E
         PICN            =   "abonos.frx":107A
         PICH            =   "abonos.frx":1309
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SystemOncoAmerica.DMGrid DMGrid2 
         Height          =   1455
         Left            =   1080
         TabIndex        =   37
         Top             =   2400
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2566
         Object.Width           =   5985
         Object.Height          =   1425
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin SystemOncoAmerica.DMGrid DMGrid3 
         Height          =   1455
         Left            =   1080
         TabIndex        =   39
         Top             =   4200
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2566
         Object.Width           =   5985
         Object.Height          =   1425
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin ChamaleonButton.ChameleonBtn BtnRecargar 
         Height          =   375
         Left            =   1080
         TabIndex        =   41
         ToolTipText     =   "Deshacer Operacion"
         Top             =   5880
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Verificar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":174A
         PICN            =   "abonos.frx":1766
         PICH            =   "abonos.frx":19CB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
         Height          =   375
         Left            =   5520
         TabIndex        =   48
         ToolTipText     =   "Deshacer Operacion"
         Top             =   5880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancelar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":1C54
         PICN            =   "abonos.frx":1C70
         PICH            =   "abonos.frx":1F52
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblSaldoRef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   195
         Left            =   4920
         TabIndex        =   50
         Top             =   1800
         Width           =   450
      End
      Begin VB.Line Line2 
         X1              =   7320
         X2              =   7320
         Y1              =   240
         Y2              =   6240
      End
      Begin VB.Line Line1 
         X1              =   840
         X2              =   840
         Y1              =   240
         Y2              =   6240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año de los Periodo:"
         Height          =   195
         Left            =   5280
         TabIndex        =   45
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empezar desde el Período:"
         Height          =   195
         Left            =   1080
         TabIndex        =   43
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuotas Modificadas:"
         Height          =   195
         Left            =   1080
         TabIndex        =   40
         Top             =   3960
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuotas Actuales:"
         Height          =   195
         Left            =   1080
         TabIndex        =   38
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cuotas:"
         Height          =   195
         Left            =   1080
         TabIndex        =   35
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto x Cuota:"
         Height          =   195
         Left            =   3000
         TabIndex        =   34
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   5760
      Width           =   8535
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7440
         TabIndex        =   18
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cerrar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":21A3
         PICN            =   "abonos.frx":21BF
         PICH            =   "abonos.frx":2388
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnDesHacer 
         Height          =   375
         Left            =   6240
         TabIndex        =   19
         ToolTipText     =   "Deshacer Operacion"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Deshacer"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":25BD
         PICN            =   "abonos.frx":25D9
         PICH            =   "abonos.frx":28BB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnImprimir 
         Height          =   375
         Left            =   3000
         TabIndex        =   20
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Imprimir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":2B0C
         PICN            =   "abonos.frx":2B28
         PICH            =   "abonos.frx":2C4D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregar 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Agregar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":2EDD
         PICN            =   "abonos.frx":2EF9
         PICH            =   "abonos.frx":3086
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnBorrar 
         Height          =   375
         Left            =   1320
         TabIndex        =   23
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Borrar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":32BB
         PICN            =   "abonos.frx":32D7
         PICH            =   "abonos.frx":347B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Pagos y abonos"
      Height          =   3975
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   8535
      Begin VB.TextBox TxtMporC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   6960
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   960
         Width           =   1455
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   3615
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   6376
         Object.Width           =   4905
         Object.Height          =   3585
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin VB.TextBox TxtSaldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox TxtTotAbonos 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox TxtMontoAbono 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1680
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   211419139
         CurrentDate     =   39990
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
         Height          =   375
         Left            =   6960
         TabIndex        =   24
         ToolTipText     =   "Guardar / Actualizar "
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Guardar"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":361A
         PICN            =   "abonos.frx":3636
         PICH            =   "abonos.frx":38C5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5280
         TabIndex        =   29
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   211419139
         CurrentDate     =   39990
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
         Height          =   375
         Left            =   5280
         TabIndex        =   49
         ToolTipText     =   "Recalcular Cuotas"
         Top             =   2760
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Recalcular Cuotas"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "abonos.frx":3D06
         PICN            =   "abonos.frx":3D22
         PICH            =   "abonos.frx":3F8C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto por Cuota:"
         Height          =   195
         Left            =   6960
         TabIndex        =   52
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelado el:"
         Height          =   195
         Left            =   5280
         TabIndex        =   30
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Abono:"
         Height          =   195
         Left            =   5280
         TabIndex        =   28
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Cobro:"
         Height          =   195
         Left            =   5280
         TabIndex        =   16
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto del Abono:"
         Height          =   195
         Left            =   6960
         TabIndex        =   15
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   195
         Left            =   6840
         TabIndex        =   14
         Top             =   3240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Abonos:"
         Height          =   195
         Left            =   5280
         TabIndex        =   13
         Top             =   3240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Datos de Empleado"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8535
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Height          =   1335
         Left            =   5280
         TabIndex        =   25
         Top             =   120
         Width           =   3135
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Prestamo:"
            Height          =   195
            Left            =   1080
            TabIndex        =   27
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   840
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtCedula 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TxtApellidos 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox TxtNombres 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cédula:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   600
      End
   End
   Begin MSComCtl2.DTPicker DTPicker11 
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   211419137
      CurrentDate     =   39990
   End
End
Attribute VB_Name = "FrmPagosPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPrestamo As New ADODB.Recordset
Dim RsTemp As New ADODB.Recordset
Dim Cambio
Dim NewReg
Dim i As Integer
Dim j As Integer
Dim TotAbonos As Double
Dim Msg
Dim MuestraDialogo As Boolean

Public IdRengPrestamo As Integer
Public IdEmpla As Integer
Public IdPrestamo As Integer
Public Mon As Double

Sub InitGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 5
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True

'DMGrid1.DColumnas(3).Visible = False
DMGrid1.DColumnas(5).Visible = False

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 35 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 25 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 25 / 100) - 300
DMGrid1.DColumnas(1).Caption = "Nro."
DMGrid1.DColumnas(2).Caption = "Fecha"
DMGrid1.DColumnas(3).Caption = "Monto xCuota"
DMGrid1.DColumnas(4).Caption = "Abono"
DMGrid1.DColumnas(5).Caption = "IdRengPrestamo"
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' carga las columnas y encabezados de columna
DMGrid2.Cols = DMGrid1.Cols
DMGrid2.Rows = DMGrid1.Rows
DMGrid2.DColumnas(1).Alignment = DMGrid1.DColumnas(1).Alignment
DMGrid2.DColumnas(2).Alignment = DMGrid1.DColumnas(2).Alignment
DMGrid2.DColumnas(3).Alignment = DMGrid1.DColumnas(3).Alignment
DMGrid2.DColumnas(4).Alignment = DMGrid1.DColumnas(4).Alignment
DMGrid2.DColumnas(1).Locked = DMGrid1.DColumnas(1).Locked
DMGrid2.DColumnas(2).Locked = DMGrid1.DColumnas(2).Locked
DMGrid2.DColumnas(3).Locked = DMGrid1.DColumnas(3).Locked

DMGrid2.DColumnas(5).Visible = DMGrid1.DColumnas(5).Visible

DMGrid2.DColumnas(1).Width = Val(DMGrid2.Width * 15 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid2.Width * 35 / 100)
DMGrid2.DColumnas(3).Width = Val(DMGrid2.Width * 25 / 100)
DMGrid2.DColumnas(4).Width = Val(DMGrid2.Width * 25 / 100) - 300
DMGrid2.DColumnas(1).Caption = DMGrid1.DColumnas(1).Caption
DMGrid2.DColumnas(2).Caption = DMGrid1.DColumnas(2).Caption
DMGrid2.DColumnas(3).Caption = DMGrid1.DColumnas(3).Caption
DMGrid2.DColumnas(4).Caption = DMGrid1.DColumnas(4).Caption
DMGrid2.DColumnas(5).Caption = DMGrid1.DColumnas(5).Caption

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' carga las columnas y encabezados de columna
DMGrid3.Cols = DMGrid1.Cols
DMGrid3.Rows = DMGrid1.Rows
DMGrid3.DColumnas(1).Alignment = DMGrid1.DColumnas(1).Alignment
DMGrid3.DColumnas(2).Alignment = DMGrid1.DColumnas(2).Alignment
DMGrid3.DColumnas(3).Alignment = DMGrid1.DColumnas(3).Alignment
DMGrid3.DColumnas(4).Alignment = DMGrid1.DColumnas(4).Alignment
DMGrid3.DColumnas(1).Locked = DMGrid1.DColumnas(1).Locked
DMGrid3.DColumnas(2).Locked = DMGrid1.DColumnas(2).Locked
DMGrid3.DColumnas(3).Locked = DMGrid1.DColumnas(3).Locked

DMGrid3.DColumnas(5).Visible = DMGrid1.DColumnas(5).Visible

DMGrid3.DColumnas(1).Width = Val(DMGrid3.Width * 15 / 100)
DMGrid3.DColumnas(2).Width = Val(DMGrid3.Width * 35 / 100)
DMGrid3.DColumnas(3).Width = Val(DMGrid3.Width * 25 / 100)
DMGrid3.DColumnas(4).Width = Val(DMGrid3.Width * 25 / 100) - 300
DMGrid3.DColumnas(1).Caption = DMGrid1.DColumnas(1).Caption
DMGrid3.DColumnas(2).Caption = DMGrid1.DColumnas(2).Caption
DMGrid3.DColumnas(3).Caption = DMGrid1.DColumnas(3).Caption
DMGrid3.DColumnas(4).Caption = DMGrid1.DColumnas(4).Caption
DMGrid3.DColumnas(5).Caption = DMGrid1.DColumnas(5).Caption

End Sub

Sub Llenar_Abonos()
Dim MxC As Double
CSql = "SELECT * FROM Prestamos WHERE IdEmpleado=" & IdEmpla & " AND IdPrestamos=" & IdPrestamo
Set RsPrestamo = CrearRS(CSql)

If RsPrestamo.RecordCount = 0 Then Exit Sub

If RsPrestamo.Fields("IdPrestamos").Value <> "" Then Label21.Caption = RsPrestamo.Fields("IdPrestamos").Value Else Label21.Caption = ""
If Not IsNull(RsPrestamo.Fields("Monto_Presta").Value) Then Mon = CDbl(RsPrestamo.Fields("Monto_Presta").Value) Else Mon = 0

TotAbonos = 0
TxtSaldo.Text = Format("0", "#,##0.00")
TxtTotAbonos.Text = Format("0", "#,##0.00")

If IdEmpla = 0 Then Exit Sub

CSql = "SELECT * FROM RenglonPrestamos WHERE IdEmpleado = " & IdEmpla & " AND IdPrestamo=" & IdPrestamo & " Order By FechaPago"
Set RsTemp = CrearRS(CSql)

DMGrid1.Rows = 0
If RsTemp.RecordCount > 0 Then
    Do While Not RsTemp.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = DMGrid1.Rows
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("FechaPago").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = Format(RsTemp.Fields("AbonoMax").Value, "#,##0.00")
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = Format(RsTemp.Fields("MontoAbono").Value, "#,##0.00")
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = Format(RsTemp.Fields("IdRengPrestamo").Value, "#,##0.00")
        
        TotAbonos = TotAbonos + CDbl(RsTemp.Fields("MontoAbono").Value)
        RsTemp.MoveNext
    Loop
End If

DMGrid1.PaintMGrid
TxtTotAbonos.Text = Format(TotAbonos, "#,##0.00")
TxtSaldo.Text = Format(CDbl(Mon) - TotAbonos, "#,##0.00")

End Sub

Sub Calcula_Periodos()
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
  Call Calcular_Periodos(Format(DTPicker3.Value, "yyyy"))
  Combo1.Clear
  For i = 0 To 23
      Combo1.AddItem "Período " & PyFs(i, 0) & ": " & PyFs(i, 1) & " => " & PyFs(i, 2)
      Combo1.ItemData(Combo1.NewIndex) = PyFs(i, 0)
  Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
End Sub

Private Sub BtnAgregar_Click()
Dim NumMayor As Integer
Dim TamDMGrid As Integer

NumMayor = 0
TamDMGrid = DMGrid1.Rows

If TamDMGrid > 0 Then
    For i = 1 To TamDMGrid
        If Val(DMGrid1.ValorCelda(i, 1)) > NumMayor Then
            NumMayor = Val(DMGrid1.ValorCelda(i, 1))
        End If
    Next i
    Label3.Caption = "No. de Abono: " & NumMayor + 1
Else
    Label3.Caption = "No. de Abono: 1"
End If

DTPicker1.Enabled = True
TxtMontoAbono.Enabled = True
DTPicker1.SetFocus
DTPicker2.Value = Now
BtnAgregar.Enabled = False
BtnImprimir.Enabled = False
BtnBorrar.Enabled = False
BtnGuardarActualizar.Enabled = True

NewReg = 1

End Sub

Private Sub BtnCerrar_Click()
FrmPrestamos.BtnSiguiente_Click
FrmPrestamos.BtnAnterior_Click
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
DTPicker2.Value = Now
BtnAgregar.Enabled = True
BtnImprimir.Enabled = True
BtnBorrar.Enabled = True
BtnGuardarActualizar.Enabled = False
TxtMontoAbono.Text = "0,00"
TxtMporC.Text = "0,00"
TxtMporC.Locked = True
TxtMporC.BackColor = &HC0C0C0
TxtMontoAbono.Enabled = False
DTPicker1.Value = Now
DTPicker1.Enabled = False
Llenar_Abonos
End Sub

Private Sub BtnGuardarActualizar_Click()

If Trim(TxtMontoAbono.Text) = "" Then MsgBox "Debe de ingresar un monto!", vbCritical + vbOKOnly, "Error": TxtMontoAbono.SetFocus: Exit Sub
If Trim(TxtMporC.Text) = "" Then MsgBox "Debe de ingresar un monto!", vbCritical + vbOKOnly, "Error": TxtMporC.SetFocus: Exit Sub
If Not IsNumeric(TxtMontoAbono.Text) Then MsgBox "Debe de ingresar solo números!", vbCritical + vbOKOnly, "Error": TxtMontoAbono.SetFocus: TxtMontoAbono.Text = "0,00": Exit Sub
If Not IsNumeric(TxtMporC.Text) Then MsgBox "Debe de ingresar solo números!", vbCritical + vbOKOnly, "Error": TxtMporC.SetFocus: TxtMporC.Text = "0,00": Exit Sub

'If Trim(TxtSaldo.Text) = "" Then MsgBox "Debe de ingresar un monto!", vbCritical + vbOKOnly, "Error": TxtSaldo.SetFocus: Exit Sub
'If Not IsNumeric(TxtSaldo.Text) Then MsgBox "Debe de ingresar solo números!", vbCritical + vbOKOnly, "Error": TxtSaldo.SetFocus: TxtSaldo.Text = "": Exit Sub

If CDbl(TxtMporC.Text) <= 0# Then
    MsgBox "Debe de ingresar un monto mayor a 0,00!", vbCritical + vbOKOnly, "Error"
    TxtMporC.SetFocus
    Exit Sub
End If

'If CDbl(TxtSaldo.Text) <= 0# Then
    'MsgBox "Ya el Prestamo esta cancelado por completo", vbCritical + vbOKOnly, "Error"
    'TxtMontoAbono.Text = "0,00"
    'TxtMontoAbono.Enabled = False
    'DTPicker1.Value = Now
    'DTPicker1.Enabled = False
    'BtnDesHacer_Click
    'Exit Sub
'End If
Dim C As String
Select Case NewReg

    Case Is = 1
    
        CSql = "Select MAX(IdRengPrestamo) + 1 as NuevoId From RenglonPrestamos"
        Set RsTemp = CrearRS(CSql)
        
        If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
            C = Val(RsTemp.Fields("NuevoId").Value)
        Else
            C = "1"
        End If
        
        CSql = "Select Count(*) From RenglonPrestamos WHERE IdPrestamo=" & IdPrestamo
        Set RsTemp = CrearRS(CSql)
        
        CSql = "Select * From RenglonPrestamos"
        Set RsTemp = CrearRS(CSql)
        
        RsTemp.AddNew
        RsTemp.Fields("IdRengPrestamo").Value = C
        RsTemp.Fields("IdPrestamo").Value = IdPrestamo
        RsTemp.Fields("IdEmpleado").Value = IdEmpla
        RsTemp.Fields("Cuota").Value = Label21.Caption
        RsTemp.Fields("FechaPago").Value = Format(DTPicker1.Value, "dd/MM/yyyy")
        'RsTemp.Fields("AbonoMax").Value = CDbl(TxtMporC.Text)
        RsTemp.Fields("MontoAbono").Value = CDbl(TxtMontoAbono.Text)
        RsTemp.Fields("IdUser").Value = IdUser
        RsTemp.Update
        
        MsgBox "Abono Guardado Satisfactoriamente!", vbInformation + vbOKOnly, "Operación exitosa!"
        
    Case Is = 0
        
        If IdRengPrestamo <= 0 Then
            MsgBox "Ocurrio un error inesperado al tratar de actualizar un abono del prestamo No. " & Label21.Caption, vbCritical + vbOKOnly, "Contacte al administrador."
            Exit Sub
        End If
        
        CSql = "Select * From RenglonPrestamos Where IdRengPrestamo =" & IdRengPrestamo
        Set RsTemp = CrearRS(CSql)

        'RsTemp.Fields("AbonoMax").Value = Replace(Replace(TxtMporC.Text, ".", ""), ",", ".")
        RsTemp.Fields("FechaAbono").Value = Format(DTPicker2.Value, "dd/MM/yyyy")
        RsTemp.Fields("MontoAbono").Value = Replace(Replace(TxtMontoAbono.Text, ".", ""), ",", ".")
        RsTemp.Fields("FechaPago").Value = Format(DTPicker1.Value, "dd/MM/yyyy")
        RsTemp.Fields("IdUser").Value = IdUser
        RsTemp.Update
        
        MsgBox "Abono Actualizado Satisfactoriamente", vbInformation + vbOKOnly, "Operación exitosa!"
        
End Select

BtnDesHacer_Click

CSql = "UPDATE Prestamos set Adeuda=" & Replace(Replace(CDbl(TxtSaldo.Text), ".", ""), ",", ".") & _
        ", Abonos=" & Replace(Replace(CDbl(TxtTotAbonos.Text), ".", ""), ",", ".") & " WHERE IdEmpleado=" & IdEmpla & " AND IdPrestamos=" & IdPrestamo
Set RsPrestamo = CrearRS(CSql)

End Sub

Private Sub BtnRecargar_Click()
Dim SaldoTot As Double
Dim MontoCout As Double
Dim ValorAcumulado As Double
Dim Anio As Integer
Dim Cuota As Integer
Dim Cuotas As Integer
Dim TamDMGrid As Integer
Dim TamDMGrid2 As Integer
Dim Band As Boolean
Dim UltimaFechaAbono As String

If Combo1.ListIndex = -1 Then
    MsgBox "Seleccione el período del comienzo del cobro!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo1.SetFocus
    Exit Sub
ElseIf Combo3.ListIndex = -1 Then
    MsgBox "Seleccione el número de cuotas!", vbExclamation + vbOKOnly, "Faltan datos!"
    Combo3.SetFocus
    Exit Sub
ElseIf Trim(TxtMxC.Text) = "" Then
    MsgBox "Ingrese el Monto para cada cuota!", vbExclamation + vbOKOnly, "Faltan datos!"
    TxtMxC.SetFocus
    Exit Sub
ElseIf Not IsNumeric(TxtMxC.Text) Then
    MsgBox "Ingrese sólo números para las cuotas!", vbExclamation + vbOKOnly, "Error!"
    TxtMxC.SetFocus
    TxtMxC.Text = "0,00"
    Exit Sub
ElseIf CDbl(TxtMxC.Text) = 0# Then
    MsgBox "Las cuotas deben ser mayor de CERO (0)!", vbExclamation + vbOKOnly, "Error!"
    TxtMxC.SetFocus
    Exit Sub
End If

SaldoTot = CDbl(TxtSaldo.Text)
MontoCout = CDbl(TxtMxC.Text)
Cuotas = Val(Combo3.ItemData(Combo3.ListIndex))
DMGrid3.Clear
DMGrid3.Rows = 0

CSql = "SELECT * FROM RenglonPrestamos WHERE MontoAbono<>0 AND IdPrestamo=" & Val(Label21.Caption) & " ORDER BY FechaPago"
Set RsTemp = CrearRS(CSql)
DMGrid3.Clear
DMGrid3.Rows = 0

UltimaFechaAbono = "01/01/1899"

If RsTemp.RecordCount <> 0 Then
    While Not RsTemp.EOF
        DMGrid3.Rows = DMGrid3.Rows + 1
        DMGrid3.ValorCelda(DMGrid3.Rows, 1) = RsTemp.Fields("Cuota").Value
        DMGrid3.ValorCelda(DMGrid3.Rows, 2) = RsTemp.Fields("FechaPago").Value
        DMGrid3.ValorCelda(DMGrid3.Rows, 3) = RsTemp.Fields("AbonoMax").Value
        DMGrid3.ValorCelda(DMGrid3.Rows, 4) = RsTemp.Fields("MontoAbono").Value
        DMGrid3.ValorCelda(DMGrid3.Rows, 5) = RsTemp.Fields("IdRengPrestamo").Value
        Cuota = Val(RsTemp.Fields("Cuota").Value) + 1
        UltimaFechaAbono = Format(CDate(RsTemp.Fields("FechaPago").Value), "dd/MM/yyyy")
        RsTemp.MoveNext
    Wend
Else
    Cuota = 1
End If

ValorAcumulado = 0

i = Combo1.ListIndex

If CDate(UltimaFechaAbono) >= (PyFs(i, 2)) Then
    MsgBox "La fecha de inicio de cobro debe ser MAYOR al del último abono efectuado!", vbCritical + vbOKOnly, "ERROR: Fecha caducada!"
    Combo1.SetFocus
    Exit Sub
End If

j = 0
Anio = Val(Format(DTPicker3.Value, "yyyy"))
Call Calcular_Periodos(Str(Anio))
Band = True
While Band
    
    DMGrid3.Rows = DMGrid3.Rows + 1
    DMGrid3.ValorCelda(DMGrid3.Rows, 1) = Cuota
    DMGrid3.ValorCelda(DMGrid3.Rows, 2) = PyFs(i, 2)
    
    ValorAcumulado = ValorAcumulado + MontoCout
    If CDbl(TxtSaldo.Text) < ValorAcumulado Then MontoCout = (CDbl(TxtSaldo.Text) - (ValorAcumulado - MontoCout))
    
    DMGrid3.ValorCelda(DMGrid3.Rows, 3) = MontoCout
    DMGrid3.ValorCelda(DMGrid3.Rows, 4) = "0,00"
    
    i = i + 1           ' Periodos
    j = j + 1           ' Contador para las cuotas
    Cuota = Cuota + 1   ' Seguidilla para las Cuotas
    
    If i >= 24 Then
        i = 0
        Anio = Anio + 1
        Call Calcular_Periodos(Str(Anio))
    End If
    
    ' si "j" es igual a la cantidad de las nuevas cuotas, entonces finaliza...
    If j = Cuotas Then Band = False
Wend

DMGrid3.PaintMGrid
' If Option1(1).Value = True Then
' Else
' End If

End Sub

Private Sub ChameleonBtn1_Click()
Dim NuevoId As Integer
Dim TamDMGrid As Integer
Dim resp As Integer
Dim Band As Boolean
Dim NCuot As Integer
Dim CantAbonoMax As Double

resp = MsgBox("Se procederá a guardar los cambios. " & Chr(13) & "Desea Continuar?", vbQuestion + vbYesNo, "Confirmar!")
If resp = vbNo Then Exit Sub
' Normaliza la tabla de los renglos de los prestamos...
    'CSql = ""
    'Set RsTemp = CrearRS(CSql)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
CSql = "DELETE FROM RenglonPrestamos WHERE MontoAbono=0 AND IdPrestamo=" & IdPrestamo
Set RsTemp = CrearRS(CSql)

CSql = "SELECT MAX(IdRengPrestamo)+1 As NuevoId FROM RenglonPrestamos"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    NuevoId = Val(RsTemp.Fields("NuevoId").Value)
Else
    NuevoId = 1
End If

TamDMGrid = DMGrid3.Rows
Band = False
NCuot = 0

For i = 1 To TamDMGrid
    If CDbl(DMGrid3.ValorCelda(i, 4)) = 0# Then
        
        CSql = "INSERT INTO RenglonPrestamos (IdRengPrestamo, IdPrestamo, IdEmpleado, Cuota, AbonoMax, " & _
            " MontoAbono, FechaPago, IdUser) VALUES(" & NuevoId & "," & IdPrestamo & "," & IdEmpla & "," & DMGrid3.ValorCelda(i, 1) & _
            "," & Replace(Replace(DMGrid3.ValorCelda(i, 3), ".", ""), ",", ".") & ",0,'" & DMGrid3.ValorCelda(i, 2) & "'," & IdUser & " )"
        
        Set RsTemp = CrearRS(CSql)
        NuevoId = NuevoId + 1
        
        If Val(DMGrid3.ValorCelda(i, 1)) > NCuot Then NCuot = Val(DMGrid3.ValorCelda(i, 1))
        
        If Band = False Then
            CantAbonoMax = CDbl(DMGrid3.ValorCelda(i, 3))
            Band = True
        End If
    End If
Next i

CSql = "UPDATE Prestamos SET Cuotas=" & NCuot & ", AbonoMax=" & Replace(Replace(CantAbonoMax, ".", ""), ",", ".") & " WHERE IdPrestamos=" & IdPrestamo
Set RsTemp = CrearRS(CSql)

MsgBox "Las Cuotas fueron modificadas correctamente!", vbInformation + vbOKOnly, "Operación Exitosa!"
Frame2.Enabled = False
Frame2.Visible = False
BtnDesHacer_Click

End Sub

Private Sub ChameleonBtn2_Click()
Frame2.Enabled = False
Frame2.Visible = False
End Sub

Private Sub ChameleonBtn3_Click()
BtnDesHacer_Click
Frame2.Top = 120
Frame2.Left = 120
Frame2.Enabled = True
Frame2.Visible = True

Dim TamDMGrid As Integer

TamDMGrid = DMGrid1.Rows
LblSaldoRef.Caption = Format(TxtSaldo.Text, "#,##0.00")

DMGrid2.Clear
DMGrid2.Rows = 0

DMGrid3.Clear
DMGrid3.Rows = 0

For i = 1 To TamDMGrid
    DMGrid2.Rows = DMGrid2.Rows + 1
    DMGrid2.ValorCelda(DMGrid2.Rows, 1) = DMGrid1.ValorCelda(i, 1)
    DMGrid2.ValorCelda(DMGrid2.Rows, 2) = DMGrid1.ValorCelda(i, 2)
    DMGrid2.ValorCelda(DMGrid2.Rows, 3) = DMGrid1.ValorCelda(i, 3)
    DMGrid2.ValorCelda(DMGrid2.Rows, 4) = DMGrid1.ValorCelda(i, 4)
    DMGrid2.ValorCelda(DMGrid2.Rows, 5) = DMGrid1.ValorCelda(i, 5)
Next i
DMGrid2.PaintMGrid
End Sub

Private Sub Combo3_Click()
Dim Valo As String

Valo = 0
If Frame1.Enabled = False Then Exit Sub
If MuestraDialogo = False Then Exit Sub

If Combo3.ListIndex = 9 Then
    Valo = InputBox("Ingrese la cantidad de Cuotas:", "Número de Cuotas.", "11")
    
    If IsNumeric(Valo) Then
        If Val(Valo) > 0 Then
            Combo3.ItemData(Combo3.ListIndex) = Valo
            Combo3.List(Combo3.ListIndex) = "Definida (" & Valo & ")"
        End If
    ElseIf Trim(Valo) <> "" Then
        MsgBox "Debe ingresar sólo números!", vbExclamation + vbOKOnly, "Error"
    End If
End If

If Trim(TxtSaldo.Text) <> "" Then
    If Valo = 0 Then
        TxtMxC.Text = Format(CDbl(TxtSaldo.Text) / Combo3.ItemData(Combo3.ListIndex), "#,##0.00")
    Else
        TxtMxC.Text = Format(CDbl(TxtSaldo.Text) / Valo, "#,##0.00")
    End If
End If
End Sub

Private Sub Combo3_GotFocus()
MuestraDialogo = True
End Sub

Private Sub Combo3_LostFocus()
MuestraDialogo = False
End Sub

Private Sub DMGrid1_DobleClick()

IdRengPrestamo = 0

If DMGrid1.Rows = 0 Then Exit Sub
If DMGrid1.Row = 0 Then Exit Sub

NewReg = 0
BtnAgregar.Enabled = False
IdRengPrestamo = Val(DMGrid1.ValorCelda(DMGrid1.Row, 5))
Label3.Caption = "Nº de Abono: " & Val(DMGrid1.ValorCelda(DMGrid1.Row, 1))
DTPicker1.Value = Format(DMGrid1.ValorCelda(DMGrid1.Row, 2), "dd/MM/yyyy")
TxtMontoAbono.Text = Format(DMGrid1.ValorCelda(DMGrid1.Row, 4), "#,##0.00")
TxtMporC.Text = Format(DMGrid1.ValorCelda(DMGrid1.Row, 3), "#,##0.00")
TxtMporC.Locked = True
TxtMporC.BackColor = &HC0C0C0
TxtMontoAbono.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Value = Now
BtnGuardarActualizar.Enabled = True


End Sub

Private Sub DTPicker3_Change()
Calcula_Periodos
End Sub

Private Sub DTPicker3_Click()
Calcula_Periodos
End Sub

Private Sub Form_Load()

Centrar Me
Mon = 0
InitGrid

Calcula_Periodos
Llenar_Abonos

If TxtSaldo.Text = "0,00" Then
    MsgBox "Ya el Prestamo esta cancelado por completo", vbInformation + vbOKOnly, "Información."
    BtnGuardarActualizar.Enabled = False
    BtnAgregar.Enabled = False
    BtnBorrar.Enabled = False
    TxtMontoAbono.Enabled = False
    DTPicker1.Enabled = False
    Exit Sub
End If

End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 48 To 57
Case Is = 13
Case 8
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmPrestamos.BtnSiguiente_Click
FrmPrestamos.BtnAnterior_Click
End Sub

Private Sub TxtMontoAbono_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 44 Then
    KeyAscii = 0
ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii = 44 And InStr(1, TxtMontoAbono.Text, ",") <> 0 Then
    KeyAscii = 0
End If
End Sub


Private Sub TxtMporC_DblClick()
If TxtMporC.BackColor = &HC0C0C0 Then
    TxtMporC.Locked = False
    TxtMporC.BackColor = vbWhite
Else
    TxtMporC.Locked = True
    TxtMporC.BackColor = &HC0C0C0
End If
End Sub

Private Sub TxtMporC_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 44 Then
    KeyAscii = 0
ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii = 44 And InStr(1, TxtMporC.Text, ",") <> 0 Then
    KeyAscii = 0
End If
End Sub
 
Private Sub TxtMxC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtMxC.Text) <> "" Then
        If IsNumeric(TxtMxC.Text) Then
            If Trim(TxtSaldo.Text) <> "" Then
                If IsNumeric(TxtSaldo.Text) Then
                
                    If CDbl(TxtSaldo.Text) < CDbl(TxtMxC.Text) Then TxtMxC.Text = "0,00": Exit Sub
                    Dim NCuotas As Integer
                    Dim Entero As Double
                    Entero = Round(CDbl(TxtSaldo.Text) / CDbl(TxtMxC.Text), 4)
                    
                    If Val(Entero) < Entero Then
                        NCuotas = Val(Entero) + 1
                    Else
                        NCuotas = Val(Entero)
                    End If
                    
                    If NCuotas > 10 Then
                        Combo3.List(9) = "Definida (" & NCuotas & ")"
                        Combo3.ItemData(9) = NCuotas
                        Combo3.ListIndex = 9
                    Else
                        For i = 0 To Combo3.ListCount
                            If Combo3.ItemData(i) = NCuotas Then
                                Combo3.ListIndex = i
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
        End If
    End If
ElseIf KeyAscii = 46 Then
    If InStr(1, TxtMxC.Text, ",") Then
        KeyAscii = 0
    Else
        KeyAscii = 44
    End If
End If
End Sub
