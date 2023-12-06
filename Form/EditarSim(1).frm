VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEditorParametrosSimulacion 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Parametros de Simulación"
   ClientHeight    =   9105
   ClientLeft      =   7170
   ClientTop       =   795
   ClientWidth     =   16590
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   16590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   8895
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   16335
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Height          =   8535
         Left            =   10800
         TabIndex        =   81
         Top             =   240
         Width           =   5415
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   960
            Top             =   5640
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   259
            ImageHeight     =   258
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   23
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":0000
                  Key             =   "Cabeza Pos Prono.jpg"
                  Object.Tag             =   "Cabeza Pos Prono.jpg"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":1EFC
                  Key             =   "Cabeza y Cara Supina.jpg"
                  Object.Tag             =   "Cabeza y Cara Supina.jpg"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":480E
                  Key             =   "Craneo y Cuello Angulo Izquierdo.jpg"
                  Object.Tag             =   "Craneo y Cuello Angulo Izquierdo.jpg"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":6E02
                  Key             =   "Craneo y cuello hiperextendido AP.JPG"
                  Object.Tag             =   "Craneo y cuello hiperextendido AP.JPG"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":8B43
                  Key             =   "Craneo y Cuello Hiperextendido lateral.jpg"
                  Object.Tag             =   "Craneo y Cuello Hiperextendido lateral.jpg"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":11AA4
                  Key             =   "Craneo y Cuello.jpg"
                  Object.Tag             =   "Craneo y Cuello.jpg"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":16392
                  Key             =   "Cuerpo de Hombre AP.JPG"
                  Object.Tag             =   "Cuerpo de Hombre AP.JPG"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":18E84
                  Key             =   "Cuerpo de Hombre PA.jpg"
                  Object.Tag             =   "Cuerpo de Hombre PA.jpg"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":1B1D4
                  Key             =   "Cuerpo de mujer1 Ap.jpg"
                  Object.Tag             =   "Cuerpo de mujer1 Ap.jpg"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":1D1EB
                  Key             =   "Cuerpo de Mujer AP.jpg"
                  Object.Tag             =   "Cuerpo de Mujer AP.jpg"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":1F9CE
                  Key             =   "Mama Derecha.jpg"
                  Object.Tag             =   "Mama Derecha.jpg"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":2215D
                  Key             =   "Mama Izquerda.jpg"
                  Object.Tag             =   "Mama Izquerda.jpg"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":24DD2
                  Key             =   "Miembros Inferiores.jpg"
                  Object.Tag             =   "Miembros Inferiores.jpg"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":27976
                  Key             =   "Miembros Superiores.jpg"
                  Object.Tag             =   "Miembros Superiores.jpg"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":2B713
                  Key             =   "Pared Costal Derecho.jpg"
                  Object.Tag             =   "Pared Costal Derecho.jpg"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":2DF60
                  Key             =   "Pared Costal Izquierda.jpg"
                  Object.Tag             =   "Pared Costal Izquierda.jpg"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":301B2
                  Key             =   "Perine.jpg"
                  Object.Tag             =   "Perine.jpg"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":31EFF
                  Key             =   "Pie.jpg"
                  Object.Tag             =   "Pie.jpg"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":35345
                  Key             =   "Torax Brazos Abajo1 PA.jpg"
                  Object.Tag             =   "Torax Brazos Abajo1 PA.jpg"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":36E72
                  Key             =   "Torax Brazos Abajo AP.jpg"
                  Object.Tag             =   "Torax Brazos Abajo AP.jpg"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":38DB6
                  Key             =   "Torax Brazos Abajo.jpg"
                  Object.Tag             =   "Torax Brazos Abajo.jpg"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":3C2C0
                  Key             =   "Torax Brazos Arriba AP1.jpg"
                  Object.Tag             =   "Torax Brazos Arriba AP1.jpg"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EditarSim.frx":40313
                  Key             =   "Torax Brazos Arriba AP.jpg"
                  Object.Tag             =   "Torax Brazos Arriba AP.jpg"
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Herramientas"
            Height          =   1695
            Left            =   120
            TabIndex        =   83
            Top             =   6720
            Width           =   5055
            Begin ChamaleonButton.ChameleonBtn BtnColor 
               Height          =   495
               Left            =   120
               TabIndex        =   84
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   873
               BTYPE           =   3
               TX              =   "Color"
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
               MICON           =   "EditarSim.frx":449AF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   1200
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin ChamaleonButton.ChameleonBtn BtnLimpiarDibujo 
               Height          =   495
               Left            =   120
               TabIndex        =   85
               Top             =   960
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   873
               BTYPE           =   3
               TX              =   "Limpiar Dibujo"
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
               MICON           =   "EditarSim.frx":449CB
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnOvalo 
               Height          =   495
               Left            =   2520
               TabIndex        =   86
               Top             =   960
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   873
               BTYPE           =   3
               TX              =   "Ovalo"
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
               BCOL            =   8421504
               BCOLO           =   8421504
               FCOL            =   0
               FCOLO           =   16711680
               MCOL            =   8454143
               MPTR            =   1
               MICON           =   "EditarSim.frx":449E7
               PICN            =   "EditarSim.frx":44A03
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnLinea 
               Height          =   495
               Left            =   1560
               TabIndex        =   87
               Top             =   960
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   873
               BTYPE           =   3
               TX              =   "Linea"
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
               MCOL            =   8454143
               MPTR            =   1
               MICON           =   "EditarSim.frx":4769F
               PICN            =   "EditarSim.frx":476BB
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnRectangulo 
               Height          =   495
               Left            =   3480
               TabIndex        =   88
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   873
               BTYPE           =   3
               TX              =   "Rect."
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
               MCOL            =   8454143
               MPTR            =   1
               MICON           =   "EditarSim.frx":4A336
               PICN            =   "EditarSim.frx":4A352
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Shape Shape3 
               BorderWidth     =   2
               FillStyle       =   0  'Solid
               Height          =   495
               Left            =   1560
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            DrawWidth       =   2
            ForeColor       =   &H80000008&
            Height          =   4800
            Left            =   555
            MousePointer    =   2  'Cross
            ScaleHeight     =   4770
            ScaleWidth      =   4275
            TabIndex        =   82
            Top             =   240
            Width           =   4305
            Begin VB.Line Line1 
               BorderWidth     =   2
               Visible         =   0   'False
               X1              =   1080
               X2              =   1080
               Y1              =   1920
               Y2              =   1080
            End
            Begin VB.Shape Shape1 
               BorderWidth     =   2
               Height          =   735
               Left            =   120
               Shape           =   2  'Oval
               Top             =   480
               Visible         =   0   'False
               Width           =   855
            End
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   855
         Left            =   120
         TabIndex        =   58
         Top             =   7920
         Width           =   10575
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   9360
            TabIndex        =   40
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
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
            MICON           =   "EditarSim.frx":4CFF4
            PICN            =   "EditarSim.frx":4D010
            PICH            =   "EditarSim.frx":4D1D9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
            Height          =   495
            Left            =   1320
            TabIndex        =   37
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "EditarSim.frx":4D40E
            PICN            =   "EditarSim.frx":4D42A
            PICH            =   "EditarSim.frx":4D6B9
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
            Height          =   495
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "EditarSim.frx":4DAFA
            PICN            =   "EditarSim.frx":4DB16
            PICH            =   "EditarSim.frx":4DCA3
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
            Height          =   495
            Left            =   8040
            TabIndex        =   39
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "EditarSim.frx":4DED8
            PICN            =   "EditarSim.frx":4DEF4
            PICH            =   "EditarSim.frx":4E1D6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnEliminar 
            Height          =   495
            Left            =   2640
            TabIndex        =   38
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "EditarSim.frx":4E427
            PICN            =   "EditarSim.frx":4E443
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
         Caption         =   "Parametros a Simular"
         Height          =   7575
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   10575
         Begin VB.ComboBox Combo25 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E7FB
            Left            =   9000
            List            =   "EditarSim.frx":4E80E
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   6960
            Width           =   1215
         End
         Begin VB.ComboBox Combo24 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E821
            Left            =   9000
            List            =   "EditarSim.frx":4E846
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   6600
            Width           =   1215
         End
         Begin VB.ComboBox Combo23 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E86D
            Left            =   9000
            List            =   "EditarSim.frx":4E886
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   6240
            Width           =   1215
         End
         Begin VB.ComboBox Combo22 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E89F
            Left            =   9000
            List            =   "EditarSim.frx":4E8AF
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   5880
            Width           =   1215
         End
         Begin VB.ComboBox Combo21 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E8BF
            Left            =   6120
            List            =   "EditarSim.frx":4E8E4
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   6600
            Width           =   1095
         End
         Begin VB.ComboBox Combo20 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E914
            Left            =   6120
            List            =   "EditarSim.frx":4E927
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   6240
            Width           =   1095
         End
         Begin VB.ComboBox Combo19 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E93A
            Left            =   6120
            List            =   "EditarSim.frx":4E962
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   5880
            Width           =   1095
         End
         Begin VB.ComboBox Combo18 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E98D
            Left            =   9000
            List            =   "EditarSim.frx":4E997
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   5400
            Width           =   1215
         End
         Begin VB.ComboBox Combo17 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E9A3
            Left            =   6000
            List            =   "EditarSim.frx":4E9AD
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   5400
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   6000
            TabIndex        =   25
            Top             =   4800
            Width           =   1575
         End
         Begin VB.ComboBox Combo16 
            Height          =   315
            ItemData        =   "EditarSim.frx":4E9B9
            Left            =   9000
            List            =   "EditarSim.frx":4E9F0
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   4800
            Width           =   615
         End
         Begin VB.ComboBox Combo15 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EA27
            Left            =   9000
            List            =   "EditarSim.frx":4EA3A
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   4320
            Width           =   615
         End
         Begin VB.ComboBox Combo14 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EA4D
            Left            =   6000
            List            =   "EditarSim.frx":4EA5A
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   4320
            Width           =   1215
         End
         Begin VB.ComboBox Combo13 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EA74
            Left            =   8400
            List            =   "EditarSim.frx":4EA7E
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   3840
            Width           =   1215
         End
         Begin VB.ComboBox Combo12 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EA8A
            Left            =   6000
            List            =   "EditarSim.frx":4EA94
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   3840
            Width           =   1215
         End
         Begin VB.ComboBox Combo11 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EAA0
            Left            =   8400
            List            =   "EditarSim.frx":4EAAA
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3360
            Width           =   1215
         End
         Begin VB.ComboBox Combo10 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EAB6
            Left            =   6000
            List            =   "EditarSim.frx":4EAC0
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3360
            Width           =   1215
         End
         Begin VB.ComboBox Combo9 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EACC
            Left            =   8400
            List            =   "EditarSim.frx":4EAE2
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2880
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   8400
            TabIndex        =   16
            ToolTipText     =   "Angulo en Grados"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox Combo8 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EAF8
            Left            =   6000
            List            =   "EditarSim.frx":4EB02
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2880
            Width           =   1215
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EB0E
            Left            =   6000
            List            =   "EditarSim.frx":4EB18
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2400
            Width           =   1215
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EB24
            Left            =   6000
            List            =   "EditarSim.frx":4EB2E
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EB3A
            Left            =   1680
            List            =   "EditarSim.frx":4EB44
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   6240
            Width           =   2055
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EB50
            Left            =   1320
            List            =   "EditarSim.frx":4EB5D
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   4440
            Width           =   2415
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EB75
            Left            =   1320
            List            =   "EditarSim.frx":4EB82
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   4080
            Width           =   2415
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EB97
            Left            =   1200
            List            =   "EditarSim.frx":4EBA4
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   3240
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "EditarSim.frx":4EBB9
            Left            =   1680
            List            =   "EditarSim.frx":4EC02
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1440
            Width           =   3735
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "RMN"
            Height          =   255
            Left            =   960
            TabIndex        =   5
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tomografia de RX"
            Height          =   255
            Left            =   960
            TabIndex        =   4
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1320
            TabIndex        =   10
            Top             =   4800
            Width           =   2415
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1695
            TabIndex        =   11
            Top             =   5355
            Width           =   2040
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1680
            TabIndex        =   12
            Top             =   5790
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   405
            Left            =   960
            TabIndex        =   6
            Top             =   2670
            Width           =   2775
         End
         Begin ChamaleonButton.ChameleonBtn BtnAyuda 
            Height          =   495
            Left            =   9720
            TabIndex        =   41
            ToolTipText     =   "Ayuda"
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   ""
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
            FCOLO           =   0
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "EditarSim.frx":4EDF6
            PICN            =   "EditarSim.frx":4EE12
            PICH            =   "EditarSim.frx":4F0B4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posición:"
            Height          =   195
            Left            =   8160
            TabIndex        =   80
            Top             =   7035
            Width           =   645
         End
         Begin VB.Shape Shape2 
            Height          =   2175
            Left            =   7320
            Top             =   5280
            Width           =   3015
         End
         Begin VB.Shape Shape11 
            Height          =   2175
            Left            =   4320
            Top             =   5280
            Width           =   3015
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rotación:"
            Height          =   195
            Left            =   8160
            TabIndex        =   79
            Top             =   6645
            Width           =   690
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Altura:"
            Height          =   195
            Left            =   8400
            TabIndex        =   78
            Top             =   6285
            Width           =   450
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "En la Mesa:"
            Height          =   195
            Left            =   8040
            TabIndex        =   77
            Top             =   5925
            Width           =   840
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ángulo:"
            Height          =   195
            Left            =   5400
            TabIndex        =   76
            Top             =   6645
            Width           =   540
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posición:"
            Height          =   195
            Left            =   5280
            TabIndex        =   75
            Top             =   6285
            Width           =   645
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Elevacion del Brazo:"
            Height          =   195
            Left            =   4440
            TabIndex        =   74
            Top             =   5925
            Width           =   1455
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Soporte de Muñeca:"
            Height          =   195
            Left            =   7440
            TabIndex        =   73
            Top             =   5430
            Width           =   1455
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Soporte Cabeza:"
            Height          =   195
            Left            =   7680
            TabIndex        =   72
            Top             =   4920
            Width           =   1185
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Otro:"
            Height          =   195
            Left            =   5520
            TabIndex        =   71
            Top             =   4920
            Width           =   345
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Soporte de Brazo:"
            Height          =   195
            Left            =   4680
            TabIndex        =   70
            Top             =   5430
            Width           =   1275
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mama:"
            Height          =   195
            Left            =   5400
            TabIndex        =   69
            Top             =   4370
            Width           =   480
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inclinación de la Mesa:"
            Height          =   195
            Left            =   7320
            TabIndex        =   68
            Top             =   4365
            Width           =   1635
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Baja Lengua:"
            Height          =   195
            Left            =   7320
            TabIndex        =   67
            Top             =   3420
            Width           =   945
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Soporte:"
            Height          =   195
            Left            =   7680
            TabIndex        =   66
            Top             =   2880
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Angulo:"
            Height          =   195
            Left            =   7680
            TabIndex        =   65
            Top             =   2415
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAC-LOK:"
            Height          =   195
            Left            =   7560
            TabIndex        =   64
            Top             =   3855
            Width           =   720
         End
         Begin VB.Line Line4 
            X1              =   240
            X2              =   3840
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Line Line3 
            X1              =   240
            X2              =   10320
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line2 
            X1              =   240
            X2              =   3840
            Y1              =   3810
            Y2              =   3810
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Otro:"
            Height          =   195
            Left            =   795
            TabIndex        =   63
            Top             =   4890
            Width           =   345
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Piernas:"
            Height          =   195
            Left            =   600
            TabIndex        =   62
            Top             =   4485
            Width           =   570
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Region Anatomica:"
            Height          =   195
            Left            =   240
            TabIndex        =   61
            Top             =   1500
            Width           =   1350
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Otro:"
            Height          =   195
            Left            =   600
            TabIndex        =   60
            Top             =   2760
            Width           =   345
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Por:"
            Height          =   195
            Left            =   3600
            TabIndex        =   59
            Top             =   420
            Width           =   285
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3960
            TabIndex        =   2
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Soporte de Craneo:"
            Height          =   195
            Left            =   4440
            TabIndex        =   57
            Top             =   2415
            Width           =   1380
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mascarilla:"
            Height          =   195
            Left            =   5040
            TabIndex        =   56
            Top             =   2010
            Width           =   750
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contraste:"
            Height          =   195
            Left            =   240
            TabIndex        =   55
            Top             =   6300
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distancia de cortes:"
            Height          =   195
            Left            =   240
            TabIndex        =   54
            Top             =   5880
            Width           =   1410
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Espesor de cortes:"
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   5400
            Width           =   1320
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brazos:"
            Height          =   195
            Left            =   600
            TabIndex        =   52
            Top             =   4125
            Width           =   525
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posición de:"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   3840
            Width           =   870
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Orientación:"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   3300
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Estudio:"
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   1860
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   420
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1320
            TabIndex        =   0
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1320
            TabIndex        =   1
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apoyo de Cuello:"
            Height          =   195
            Left            =   4680
            TabIndex        =   46
            Top             =   2925
            Width           =   1200
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Baja Hombro:"
            Height          =   195
            Left            =   4920
            TabIndex        =   45
            Top             =   3420
            Width           =   960
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Torax: Manubrio"
            Height          =   195
            Left            =   4680
            TabIndex        =   44
            Top             =   3855
            Width           =   1155
         End
      End
   End
End
Attribute VB_Name = "FrmEditorParametrosSimulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xpos As Integer, Ypos As Integer, Flag As Boolean
Dim Recuadrar As Boolean
Dim NombImagen As String
Dim RsSimulacion As New ADODB.Recordset
Dim CSql As String

Private Sub BtnAgregar_Click()
Dim NuevoId As String

ACCION = AGREGAR_REGISTRO

Limpiar_Campos

CSql = "SELECT MAX(Id)+1 as NuevoId FROM Tecnica3"
Set RsSimulacion = CrearRS(CSql)
    
If IsNull(RsSimulacion.Fields("NuevoId")) Then
    Label2 = "0"
    Else
    Label2 = RsSimulacion.Fields("NuevoId")
End If
Label3 = Format(Now, "DD/MM/YYYY")
Label17 = Usuario


BtnAgregar.Enabled = False
BtnEliminar.Enabled = False
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnColor_Click()
    CommonDialog1.ShowColor
    Label1.BackColor = CommonDialog1.Color
    Picture1.ForeColor = CommonDialog1.Color
    Shape1.BorderColor = CommonDialog1.Color
    Line1.BorderColor = CommonDialog1.Color
    Shape3.FillColor = CommonDialog1.Color
End Sub

Private Sub BtnDesHacer_Click()
BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
Cargar_Simulacion
End Sub



Private Sub BtnEliminar_Click()
On Error Resume Next
Dim RsBorrar As New ADODB.Recordset

p = MsgBox("Desea Eliminar el registro actual?", vbQuestion + vbYesNo, "Confirmar")
If p = 7 Then Exit Sub


CSql = "UPDATE Tecnica3 SET Estado=2 Where Id=" & Label2.Caption
'CSql = "DELETE FROM Tecnica3 WHERE Id=" & Label2.Caption
Set RsBorrar = CrearRS(CSql)

 MkDir (FotoSimul2)
 Shell ("attrib +r +s +h " & FotoSimul2)

Call FileCopy(FotoSimul & "\" & NombImagen, FotoSimul2 & "\" & NombImagen)
Call SetAttr(FotoSimul2 & "\" & NombImagen, vbSystem + vbHidden + vbReadOnly)
Kill (FotoSimul & "\" & NombImagen)
MsgBox "Los datos fueron borrados del registro!", vbInformation + vbOKOnly, "Operacion Exitosa"
Unload Me


End Sub

Private Sub BtnGuardarActualizar_Click()
On Error Resume Next
Dim NomImagen As String

'Agrega el registro
    '''''''''''''''''''''''''''''''
p = MsgBox("Desea guardar los cambios?", vbQuestion + vbYesNo, "Confirmar")
If p = 7 Then Exit Sub
    'IdUser
    'Check1.Value = 0
    'Check2.Value = 0
    'Text1.Text = ""
    'Combo1.ListIndex = -1
    'Combo2.ListIndex = -1
    'Combo3.ListIndex = -1
    'Combo4.ListIndex = -1
    'Combo5.ListIndex = -1
    'Combo6.ListIndex = -1
    'Combo7.ListIndex = -1
    'Combo8.ListIndex = -1
    'Combo9.ListIndex = -1
    'Combo10.ListIndex = -1
    'Combo11.ListIndex = -1
    'Combo12.ListIndex = -1
    'Combo13.ListIndex = -1
    'Combo14.ListIndex = -1
    'Combo15.ListIndex = -1
    'Combo16.ListIndex = -1
    'Combo17.ListIndex = -1
    'Combo18.ListIndex = -1
    'Combo19.ListIndex = -1
    'Combo20.ListIndex = -1
    'Combo21.ListIndex = -1
    'Combo22.ListIndex = -1
    'Combo23.ListIndex = -1
    'Combo24.ListIndex = -1
    'Combo25.ListIndex = -1
    'Text2.Text = ""
    'Text3.Text = ""
    'Text4.Text = ""
    'Text5.Text = ""
    'Text6.Text = ""

    If Reg_Actual(33) = "" And ACCION = EDITAR_REGISTRO Then
        MsgBox "Debe Seleccionar o Agregar un registro para poder Guardar los cambios!", vbExclamation + vbOKOnly, "Informacion"
        Exit Sub
    End If
    
    Select Case ACCION
        Case EDITAR_REGISTRO
        Case AGREGAR_REGISTRO
    End Select

    CSql = "SELECT MAX(Id)+1 as NuevoId FROM Tecnica3"
    Set RsSimulacion = CrearRS(CSql)

    If IsNull(RsSimulacion.Fields("NuevoId")) Then
        Label2 = "0"
        Else
        Label2 = RsSimulacion.Fields("NuevoId")
    End If

        CSql = "UPDATE Tecnica3 SET Estado=0 Where IdPaciente=" & FrmRadioTerapia.IdPaciente & " AND Estado=1"
        Set RsSimulacion = CrearRS(CSql)
        
        MkDir (FotoSimul2)
        Shell ("attrib +r +s +h " & FotoSimul2)
        
        Call FileCopy(FotoSimul & "\" & NombImagen, FotoSimul2 & "\" & NombImagen)
        Call SetAttr(FotoSimul2 & "\" & NombImagen, vbSystem + vbHidden + vbReadOnly)
        Kill (FotoSimul & "\" & NombImagen)


        CSql = "SELECT * FROM Tecnica3"
        Set RsSimulacion = CrearRS(CSql)

        'Label17 = " " & RsSimulacion.Fields("Nombre").Value & " " & RsSimulacion.Fields("Apellidos").Value
        RsSimulacion.AddNew
        RsSimulacion.Fields("Id") = Label2.Caption
        RsSimulacion.Fields("IdUser").Value = IdUser
        RsSimulacion.Fields("IdPaciente").Value = FrmRadioTerapia.IdPaciente
        RsSimulacion.Fields("Rx").Value = Check1.Value
        RsSimulacion.Fields("Rmn").Value = Check2.Value
        RsSimulacion.Fields("Otro").Value = Text1.Text
        RsSimulacion.Fields("RegAnatomica").Value = Combo1.List(Combo1.ListIndex)
        If Combo2.Text = "Supina" Then
            RsSimulacion.Fields("Orientacion").Value = 1
        ElseIf Combo2.Text = "Prono" Then RsSimulacion.Fields("Orientacion").Value = 2
        Else: RsSimulacion.Fields("Orientacion").Value = "0"
        End If
        If Combo3.Text = "Arriba" Then
            RsSimulacion.Fields("Brazos").Value = 1
        ElseIf Combo3.Text = "Abajo" Then RsSimulacion.Fields("Brazos").Value = 2
        Else: RsSimulacion.Fields("Brazos").Value = 0
        End If
        If Combo4.Text = "Abierta" Then
            RsSimulacion.Fields("Piernas").Value = 1
        ElseIf Combo4.Text = "Cerrada" Then RsSimulacion.Fields("Piernas").Value = 2
        Else: RsSimulacion.Fields("Piernas").Value = 0
        End If
        RsSimulacion.Fields("Otro2").Value = Text2.Text
        RsSimulacion.Fields("EspCortes").Value = Text3.Text
        RsSimulacion.Fields("DisCortes").Value = Text4.Text
        If Combo5.Text = "SI" Then RsSimulacion.Fields("Contraste").Value = 1 Else RsSimulacion.Fields("Contraste").Value = 0
        If Combo6.Text = "SI" Then RsSimulacion.Fields("Mascarilla").Value = 1 Else RsSimulacion.Fields("Mascarilla").Value = 0
        If Combo7.Text = "SI" Then RsSimulacion.Fields("SopCraneo").Value = 1 Else RsSimulacion.Fields("SopCraneo").Value = 0
        RsSimulacion.Fields("SopCraneoAng").Value = Text5.Text
        If Combo8.Text = "SI" Then RsSimulacion.Fields("ApoCuello").Value = 1 Else RsSimulacion.Fields("ApoCuello").Value = 0
        RsSimulacion.Fields("ApoCuelloAng").Value = Combo9.List(Combo9.ListIndex)
        If Combo10.Text = "SI" Then RsSimulacion.Fields("BajaHombro").Value = 1 Else RsSimulacion.Fields("BajaHombro").Value = 0
        If Combo11.Text = "SI" Then RsSimulacion.Fields("BajaLengua").Value = 1 Else RsSimulacion.Fields("BajaLengua").Value = 0
        If Combo12.Text = "SI" Then RsSimulacion.Fields("ToraxManubrio").Value = 1 Else RsSimulacion.Fields("ToraxManubrio").Value = 0
        If Combo13.Text = "SI" Then RsSimulacion.Fields("Vaclok").Value = 1 Else RsSimulacion.Fields("Vaclok").Value = 0
        If Combo14.Text = "Izquierda" Then
            RsSimulacion.Fields("Mama").Value = 1
        ElseIf Combo14.Text = "Derecha" Then RsSimulacion.Fields("Mama").Value = 2
        Else: RsSimulacion.Fields("Mama").Value = 0
        End If
        RsSimulacion.Fields("InclinaMesa").Value = Combo15.List(Combo15.ListIndex)
        RsSimulacion.Fields("Otro3").Value = Text6.Text
        RsSimulacion.Fields("SopCabeza").Value = Combo16.List(Combo16.ListIndex)
        If Combo17.Text = "SI" Then RsSimulacion.Fields("SopBrazo").Value = 1 Else RsSimulacion.Fields("SopBrazo").Value = 0
        If Combo18.Text = "SI" Then RsSimulacion.Fields("SopMuneca").Value = 1 Else RsSimulacion.Fields("SopMuneca").Value = 0
        RsSimulacion.Fields("SopBrazoElevacion").Value = Combo19.List(Combo19.ListIndex)
        RsSimulacion.Fields("SopBrazoPosicion").Value = Combo20.List(Combo20.ListIndex)
        RsSimulacion.Fields("SopBrazoAngulo").Value = Combo21.List(Combo21.ListIndex)
        RsSimulacion.Fields("SopMunecaMesa").Value = Combo22.List(Combo22.ListIndex)
        RsSimulacion.Fields("SopMunecaAltura").Value = Combo23.List(Combo23.ListIndex)
        RsSimulacion.Fields("SopMunecaRotacion").Value = Combo24.List(Combo24.ListIndex)
        RsSimulacion.Fields("SopMunecaPosicion").Value = Combo25.List(Combo25.ListIndex)
        RsSimulacion.Fields("Fecha").Value = Format(Now, "DD/MM/YYYY")
        RsSimulacion.Fields("Estado").Value = 1
        NomImagen = "" & Label2.Caption & Format(Now, "DDMMYYHHMMSS") & FrmRadioTerapia.IdPaciente & ".jpg"
        RsSimulacion.Fields("Imagen").Value = NomImagen
        RsSimulacion.Update

        Dim Imagen As IPictureDisp
        Set Imagen = Picture1.Image
        'MkDir (FotoSimul)
        SavePicture Imagen, FotoSimul & "\" & NomImagen
        Set Imagen = Nothing

        
If ACCION = EDITAR_REGISTRO Then
    If Check1.Value <> Reg_Actual(0) Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo RX de (" & Reg_Actual(0) & ") a (" & Check1.Value & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Check2.Value <> Reg_Actual(1) Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo RMN de (" & Reg_Actual(1) & ") a (" & Check2.Value & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(2) <> Text1.Text Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo OTRO de (" & Reg_Actual(2) & ") a (" & Text1.Text & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(3) <> Combo1.List(Combo1.ListIndex) Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo RegAnatomica de (" & Reg_Actual(3) & ") a (" & Combo1.List(Combo1.ListIndex) & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(4) <> Combo2.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo Orientacion de (" & Reg_Actual(4) & ") a (" & Combo2.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(5) <> Combo3.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo Brazos de (" & Reg_Actual(5) & ") a (" & Combo3.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(6) <> Combo4.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo Piernas de (" & Reg_Actual(6) & ") a (" & Combo4.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(7) <> Text2.Text Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo OTRO2 de (" & Reg_Actual(7) & ") a (" & Text2.Text & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(8) <> Text3.Text Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo EspCortes de (" & Reg_Actual(8) & ") a (" & Text3.Text & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(9) <> Text4.Text Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo DisCortes de (" & Reg_Actual(9) & ") a (" & Text4.Text & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(10) <> Combo5.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo Contraste de (" & Reg_Actual(10) & ") a (" & Combo5.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(11) <> Combo6.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo Mascarilla de (" & Reg_Actual(11) & ") a (" & Combo6.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(12) <> Combo7.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopCraneo de (" & Reg_Actual(12) & ") a (" & Combo7.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(13) <> Text5.Text Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopCraneoAng de (" & Reg_Actual(13) & ") a (" & Text5.Text & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(14) <> Combo8.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo ApoCuello de (" & Reg_Actual(14) & ") a (" & Combo8.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(15) <> Combo9.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo ApoCuelloAng de (" & Reg_Actual(15) & ") a (" & Combo9.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(16) <> Combo10.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo BajaHombro de (" & Reg_Actual(16) & ") a (" & Combo10.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(17) <> Combo11.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo BajaLengua de (" & Reg_Actual(17) & ") a (" & Combo11.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(18) <> Combo12.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo ToraxManubrio de (" & Reg_Actual(18) & ") a (" & Combo12.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(19) <> Combo13.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo Vaclok de (" & Reg_Actual(19) & ") a (" & Combo13.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(20) <> Combo14.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo Mama de (" & Reg_Actual(20) & ") a (" & Combo14.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(21) <> Combo15.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo InclinaMesa de (" & Reg_Actual(21) & ") a (" & Combo15.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(22) <> Text6.Text Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo Otro3 de (" & Reg_Actual(22) & ") a (" & Text6.Text & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(23) <> Combo16.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopCabeza de (" & Reg_Actual(23) & ") a (" & Combo16.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(24) <> Combo17.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopBrazo de (" & Reg_Actual(24) & ") a (" & Combo17.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(25) <> Combo18.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopMuneca de (" & Reg_Actual(25) & ") a (" & Combo18.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(26) <> Combo19.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopBrazoElevacion de (" & Reg_Actual(26) & ") a (" & Combo19.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(27) <> Combo20.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopBrazoPosicion de (" & Reg_Actual(27) & ") a (" & Combo20.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(28) <> Combo21.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopBrazoAngulo de (" & Reg_Actual(28) & ") a (" & Combo21.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(29) <> Combo22.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopMunecaMesa de (" & Reg_Actual(29) & ") a (" & Combo22.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(30) <> Combo23.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopMunecaAltura de (" & Reg_Actual(30) & ") a (" & Combo23.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(31) <> Combo24.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopMunecaRotacion de (" & Reg_Actual(31) & ") a (" & Combo24.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
    If Reg_Actual(32) <> Combo25.ListIndex Then Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "MODIFICAR", "Se modifico de la tabla TECNICA3 el Campo SopMunecaPosicion de (" & Reg_Actual(32) & ") a (" & Combo25.ListIndex & ") del paciente de Id=" & FrmRadioTerapia.IdPaciente & " y registro Id=" & Reg_Actual(33))
Else
    Call Enviar_Bitacora(IdUser, "RadioTerapia-PARAMETROS DE SIMULACION", "INGRESAR", "Se Ingreso en la tabla TECNICA3 un nuevo registro de Id=" & Label2.Caption)
End If
        MsgBox "El registro ha sido Modificado Correctamente!", vbInformation + vbOKOnly, "Operación Exitosa"
        Cargar_Simulacion
        BtnEliminar.Enabled = True

End Sub

Private Sub BtnLimpiarDibujo_Click()
Picture1.Cls
End Sub

Private Sub BtnLinea_Click()
Shape1.Shape = 1
BtnOvalo.BackColor = &HD8E9EC
BtnOvalo.BackOver = &HD8E9EC
BtnLinea.BackColor = &H808080
BtnLinea.BackOver = &H808080
BtnRectangulo.BackColor = &HD8E9EC
BtnRectangulo.BackOver = &HD8E9EC
End Sub

Private Sub BtnOvalo_Click()
Shape1.Shape = 2
BtnOvalo.BackColor = &H808080
BtnOvalo.BackOver = &H808080
BtnLinea.BackColor = &HD8E9EC
BtnLinea.BackOver = &HD8E9EC
BtnRectangulo.BackColor = &HD8E9EC
BtnRectangulo.BackOver = &HD8E9EC
End Sub

Private Sub BtnRectangulo_Click()
Shape1.Shape = 0
BtnOvalo.BackColor = &HD8E9EC
BtnOvalo.BackOver = &HD8E9EC
BtnLinea.BackColor = &HD8E9EC
BtnLinea.BackOver = &HD8E9EC
BtnRectangulo.BackColor = &H808080
BtnRectangulo.BackOver = &H808080
End Sub

Private Sub Combo1_Click()
Dim Calc As Integer
If Combo1.ListIndex > -1 Then
    'Picture1.Visible = True
    Picture1.Picture = ImageList1.ListImages(Combo1.ListIndex + 1).Picture
    Calc = Frame3.Width - Picture1.Width
    If Calc > 10 Then
        Calc = Calc / 2
        Picture1.Left = Calc
    End If
End If
End Sub

Private Sub Combo17_Click()
If Combo17.Text = "NO" Then
    Combo19.ListIndex = -1
    Combo20.ListIndex = -1
    Combo21.ListIndex = -1
    Combo19.Visible = False
    Combo20.Visible = False
    Combo21.Visible = False
    Label33.Visible = False
    Label34.Visible = False
    Label35.Visible = False
    Else
    Combo19.Visible = True
    Combo20.Visible = True
    Combo21.Visible = True
    Label33.Visible = True
    Label34.Visible = True
    Label35.Visible = True
End If
End Sub

Private Sub Combo18_Click()
If Combo18.Text = "NO" Then
    Combo22.ListIndex = -1
    Combo23.ListIndex = -1
    Combo24.ListIndex = -1
    Combo25.ListIndex = -1
    Combo22.Visible = False
    Combo23.Visible = False
    Combo24.Visible = False
    Combo25.Visible = False
    Label36.Visible = False
    Label37.Visible = False
    Label38.Visible = False
    Label39.Visible = False
    Else
    Combo22.Visible = True
    Combo23.Visible = True
    Combo24.Visible = True
    Combo25.Visible = True
    Label36.Visible = True
    Label37.Visible = True
    Label38.Visible = True
    Label39.Visible = True
End If
End Sub

Private Sub Combo7_Click()

If Combo7.Text = "NO" Then
    Text5.Text = ""
    Text5.Visible = False
    Label24.Visible = False
    Else
    Text5.Visible = True
    Label24.Visible = True
End If

End Sub

Private Sub Combo8_Click()
If Combo8.Text = "NO" Then
    Combo9.ListIndex = -1
    Combo9.Visible = False
    Label25.Visible = False
    Else
    Combo9.Visible = True
    Label25.Visible = True
End If
End Sub

Private Sub Figura_Click(Index As Integer)
    Shape1.Shape = Index
End Sub

Private Sub Form_Load()
    
    For i = 0 To 60
        Reg_Actual(i) = ""
    Next i

    If ACCION = AGREGAR_REGISTRO Then
       Me.Caption = "Agregar nuevo registro"
    ElseIf ACCION = EDITAR_REGISTRO Then
       Me.Caption = "Editar registro"
    End If

    Cargar_Simulacion

End Sub

Sub Cargar_Simulacion()
    
    Limpiar_Campos
    CSql = "SELECT Tecnica3.*, Usuarios.Nombre,Usuarios.Apellidos FROM Tecnica3 INNER JOIN Usuarios ON Tecnica3.IdUser = Usuarios.IdUsuario WHERE Tecnica3.IdPaciente=" & FrmRadioTerapia.IdPaciente & " AND Tecnica3.Estado=1 ORDER BY Id"
    Set RsSimulacion = CrearRS(CSql)
    If RsSimulacion.RecordCount = 0 Then
        ACCION = AGREGAR_REGISTRO
        MsgBox "El Paciente no tiene Parametros de Simulación!", vbExclamation + vbOKOnly, "Sin Parametros de Simulación"
        BtnEliminar.Enabled = False
        Picture1.Picture = Nothing
        ACCION = AGREGAR_REGISTRO
        Exit Sub
    End If

    While Not RsSimulacion.EOF
        RsSimulacion.MoveNext
    Wend

    RsSimulacion.MoveLast

    ACCION = EDITAR_REGISTRO
    Label2 = RsSimulacion.Fields("Id").Value   ' Id
    Label17 = " " & RsSimulacion.Fields("Nombre").Value & " " & RsSimulacion.Fields("Apellidos").Value
    ' iduser
    If RsSimulacion.Fields("Rx").Value = True Then Check1.Value = 1
    If RsSimulacion.Fields("Rmn").Value = True Then Check2.Value = 1
    Text1.Text = RsSimulacion.Fields("Otro").Value
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = RsSimulacion.Fields("RegAnatomica").Value Then Combo1.ListIndex = i: Exit For Else Combo1.ListIndex = -1
    Next i

    Combo2.ListIndex = CInt(RsSimulacion.Fields("Orientacion").Value)
'    For i = 0 To Combo2.ListCount - 1
'        If RsSimulacion.Fields("Orientacion").Value = "0" Then Combo2.ListIndex = 0: Exit For Else Combo2.ListIndex = 1
'    Next i

    Combo3.ListIndex = CInt(RsSimulacion.Fields("Brazos").Value)
'    For i = 0 To Combo3.ListCount - 1
'        If RsSimulacion.Fields("Brazos").Value = True Then Combo3.ListIndex = 0: Exit For Else Combo3.ListIndex = 1
'    Next i

    Combo4.ListIndex = CInt(RsSimulacion.Fields("Piernas").Value)
'    For i = 0 To Combo4.ListCount - 1
'        If RsSimulacion.Fields("Piernas").Value = True Then Combo4.ListIndex = 0: Exit For Else Combo4.ListIndex = 1
'    Next i

    Text2.Text = RsSimulacion.Fields("Otro2").Value
    Text3.Text = RsSimulacion.Fields("EspCortes").Value
    Text4.Text = RsSimulacion.Fields("DisCortes").Value
    For i = 0 To Combo5.ListCount - 1
        If RsSimulacion.Fields("Contraste").Value = True Then Combo5.ListIndex = 0: Exit For Else Combo5.ListIndex = 1
    Next i
    For i = 0 To Combo6.ListCount - 1
        If RsSimulacion.Fields("Mascarilla").Value = True Then Combo6.ListIndex = 0: Exit For Else Combo6.ListIndex = 1
    Next i
    For i = 0 To Combo7.ListCount - 1
        If RsSimulacion.Fields("SopCraneo").Value = True Then Combo7.ListIndex = 0: Exit For Else Combo7.ListIndex = 1
    Next i
    Text5.Text = RsSimulacion.Fields("SopCraneoAng").Value
    For i = 0 To Combo8.ListCount - 1
        If RsSimulacion.Fields("ApoCuello").Value = True Then Combo8.ListIndex = 0: Exit For Else Combo8.ListIndex = 1
    Next i
    For i = 0 To Combo9.ListCount - 1
        If RsSimulacion.Fields("ApoCuelloAng").Value = Combo9.List(i) Then Combo9.ListIndex = i: Exit For Else Combo9.ListIndex = 1
    Next i
    For i = 0 To Combo10.ListCount - 1
        If RsSimulacion.Fields("BajaHombro").Value = True Then Combo10.ListIndex = 0: Exit For Else Combo10.ListIndex = 1
    Next i
    For i = 0 To Combo11.ListCount - 1
        If RsSimulacion.Fields("BajaLengua").Value = True Then Combo11.ListIndex = 0: Exit For Else Combo11.ListIndex = 1
    Next i
    For i = 0 To Combo12.ListCount - 1
        If RsSimulacion.Fields("ToraxManubrio").Value = True Then Combo12.ListIndex = 0: Exit For Else Combo12.ListIndex = 1
    Next i
    For i = 0 To Combo13.ListCount - 1
        If RsSimulacion.Fields("Vaclok").Value = True Then Combo13.ListIndex = 0: Exit For Else Combo13.ListIndex = 1
    Next i

    Combo14.ListIndex = CInt(RsSimulacion.Fields("Mama").Value)
'    For i = 0 To Combo14.ListCount - 1
'        If RsSimulacion.Fields("Mama").Value = True Then Combo14.ListIndex = 0: Exit For Else Combo14.ListIndex = 1
'    Next i

    For i = 0 To Combo15.ListCount - 1
        If Combo15.List(i) = RsSimulacion.Fields("InclinaMesa").Value Then Combo15.ListIndex = i
    Next i
    Text6.Text = RsSimulacion.Fields("Otro3").Value
    For i = 0 To Combo16.ListCount - 1
        If Combo16.List(i) = RsSimulacion.Fields("SopCabeza").Value Then Combo16.ListIndex = i
    Next i
    For i = 0 To Combo17.ListCount - 1
        If RsSimulacion.Fields("SopBrazo").Value = True Then Combo17.ListIndex = 0: Exit For Else Combo17.ListIndex = 1
    Next i
    For i = 0 To Combo18.ListCount - 1
        If RsSimulacion.Fields("SopMuneca").Value = True Then Combo18.ListIndex = 0: Exit For Else Combo18.ListIndex = 1
    Next i
    For i = 0 To Combo19.ListCount - 1
        If Combo19.List(i) = RsSimulacion.Fields("SopBrazoElevacion").Value Then Combo19.ListIndex = i
    Next i
    For i = 0 To Combo20.ListCount - 1
        If Combo20.List(i) = RsSimulacion.Fields("SopBrazoPosicion").Value Then Combo20.ListIndex = i
    Next i
    For i = 0 To Combo21.ListCount - 1
        If Combo21.List(i) = RsSimulacion.Fields("SopBrazoAngulo").Value Then Combo21.ListIndex = i
    Next i
    For i = 0 To Combo22.ListCount - 1
        If Combo22.List(i) = RsSimulacion.Fields("SopMunecaMesa").Value Then Combo22.ListIndex = i
    Next i
    For i = 0 To Combo23.ListCount - 1
        If Combo23.List(i) = RsSimulacion.Fields("SopMunecaAltura").Value Then Combo23.ListIndex = i
    Next i
    For i = 0 To Combo24.ListCount - 1
        If Combo24.List(i) = RsSimulacion.Fields("SopMunecaRotacion").Value Then Combo24.ListIndex = i
    Next i
    For i = 0 To Combo25.ListCount - 1
        If Combo25.List(i) = RsSimulacion.Fields("SopMunecaPosicion").Value Then Combo25.ListIndex = i
    Next i
    
    Label3 = RsSimulacion.Fields("Fecha").Value

    If Not RsSimulacion.Fields("Imagen").Value = "" Then
        Picture1.Picture = LoadPicture(FotoSimul & "\" & RsSimulacion.Fields("Imagen").Value)
        NombImagen = RsSimulacion.Fields("Imagen").Value
        'Picture1.Visible = True
        Else
        Picture1.Picture = Nothing
    End If

    Reg_Actual(33) = RsSimulacion.Fields("Id").Value
    If Check1.Value = 1 Then Reg_Actual(0) = "1" Else Reg_Actual(0) = "0"   ' Rx
    If Check2.Value = 1 Then Reg_Actual(1) = "1" Else Reg_Actual(1) = "0"   ' Rmn
    Reg_Actual(2) = RsSimulacion.Fields("Otro").Value
    Reg_Actual(3) = RsSimulacion.Fields("RegAnatomica").Value
    Reg_Actual(4) = CInt(RsSimulacion.Fields("Orientacion").Value)
    Reg_Actual(5) = CInt(RsSimulacion.Fields("Brazos").Value)
    Reg_Actual(6) = CInt(RsSimulacion.Fields("Piernas").Value)
    Reg_Actual(7) = RsSimulacion.Fields("Otro2").Value
    Reg_Actual(8) = RsSimulacion.Fields("EspCortes").Value
    Reg_Actual(9) = RsSimulacion.Fields("DisCortes").Value
    Reg_Actual(10) = Combo5.ListIndex   ' Contraste
    Reg_Actual(11) = Combo6.ListIndex   ' Mascarilla
    Reg_Actual(12) = Combo7.ListIndex   ' SopCraneo
    Reg_Actual(13) = Text5.Text         ' SopCraneoAng
    Reg_Actual(14) = Combo8.ListIndex   ' ApoCuello
    Reg_Actual(15) = Combo9.ListIndex   ' ApoCuelloAng
    Reg_Actual(16) = Combo10.ListIndex  ' BajaHombro
    Reg_Actual(17) = Combo11.ListIndex  ' BajaLengua
    Reg_Actual(18) = Combo12.ListIndex  ' ToraxManubrio
    Reg_Actual(19) = Combo13.ListIndex  ' Vaclok
    Reg_Actual(20) = Combo14.ListIndex  ' Mama
    Reg_Actual(21) = Combo15.ListIndex  ' InclinaMesa
    Reg_Actual(22) = Text6.Text         ' Otro3
    Reg_Actual(23) = Combo16.ListIndex  ' SopCabeza
    Reg_Actual(24) = Combo17.ListIndex  ' SopBrazo
    Reg_Actual(25) = Combo18.ListIndex  ' SopMuneca
    Reg_Actual(26) = Combo19.ListIndex  ' SopBrazoElevacion
    Reg_Actual(27) = Combo10.ListIndex  ' SopBrazoPosicion
    Reg_Actual(28) = Combo21.ListIndex  ' SopBrazoAngulo
    Reg_Actual(29) = Combo22.ListIndex  ' SopMunecaMesa
    Reg_Actual(30) = Combo23.ListIndex  ' SopMunecaAltura
    Reg_Actual(31) = Combo24.ListIndex  ' SopMunecaRotacion
    Reg_Actual(32) = Combo25.ListIndex  ' SopMunecaPosicion

End Sub

Sub Limpiar_Campos()
    
    Label2 = ""
    Label17 = ""
    ' iduser
    Check1.Value = 0
    Check2.Value = 0
    Text1.Text = ""
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Combo4.ListIndex = -1
    Combo5.ListIndex = -1
    Combo6.ListIndex = -1
    Combo7.ListIndex = -1
    Combo8.ListIndex = -1
    Combo9.ListIndex = -1
    Combo10.ListIndex = -1
    Combo11.ListIndex = -1
    Combo12.ListIndex = -1
    Combo13.ListIndex = -1
    Combo14.ListIndex = -1
    Combo15.ListIndex = -1
    Combo16.ListIndex = -1
    Combo17.ListIndex = -1
    Combo18.ListIndex = -1
    Combo19.ListIndex = -1
    Combo20.ListIndex = -1
    Combo21.ListIndex = -1
    Combo22.ListIndex = -1
    Combo23.ListIndex = -1
    Combo24.ListIndex = -1
    Combo25.ListIndex = -1
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Label3 = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
    Case vbLeftButton
        Xpos = X
        Ypos = Y
        Shape1.Move X, Y
        Line1.x1 = X
        Line1.y1 = Y
        Recuadrar = True
End Select
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ancho As Integer, Alto As Integer, Sup As Integer, Izq As Integer
   If Recuadrar Then
        Select Case Shape1.Shape
            Case 0, 2
                If X - Xpos > 0 Then
                    Ancho = X - Xpos
                    Izq = Xpos
                Else
                    Ancho = Xpos - X
                    Izq = X
                End If
                If Y - Ypos > 0 Then
                    Alto = Y - Ypos
                    Sup = Ypos
                Else
                    Alto = Ypos - Y
                    Sup = Y
                End If
                Shape1.Move Izq, Sup, Ancho, Alto
                If Shape1.Visible = False Then Shape1.Visible = True
            Case 1
                Line1.x2 = X
                Line1.y2 = Y
                If Line1.Visible = False Then Line1.Visible = True
        End Select
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Radio As Currency
Shape1.Visible = False
Line1.Visible = False

    Recuadrar = False

    Select Case Shape1.Shape
        Case 2
            If Shape1.Height > Shape1.Width Then
                Radio = Shape1.Height / 2
            Else
                Radio = Shape1.Width / 2
            End If
            Picture1.Circle (Shape1.Left + Shape1.Width / 2, Shape1.Top + Shape1.Height / 2), _
            Radio, , , , Shape1.Height / Shape1.Width
        Case 0
            Picture1.Line (Shape1.Left, Shape1.Top)- _
            (Shape1.Left + Shape1.Width, Shape1.Top + Shape1.Height), , B
        Case 1
            Picture1.Line (Line1.x1, Line1.y1)-(Line1.x2, Line1.y2)
    End Select

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) Or Len(Text5.Text) >= 5 Then
    KeyAscii = 0
End If
End Sub
