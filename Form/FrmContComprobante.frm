VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContComprobante 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprabante Contable"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11550
   Icon            =   "FrmContComprobante.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   7200
      Width           =   11415
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   10080
         TabIndex        =   2
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
         MICON           =   "FrmContComprobante.frx":1002
         PICN            =   "FrmContComprobante.frx":101E
         PICH            =   "FrmContComprobante.frx":11E7
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
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Guardar / Actualizar "
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "FrmContComprobante.frx":141C
         PICN            =   "FrmContComprobante.frx":1438
         PICH            =   "FrmContComprobante.frx":16C7
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
         TabIndex        =   4
         ToolTipText     =   "Agregar"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "FrmContComprobante.frx":1B08
         PICN            =   "FrmContComprobante.frx":1B24
         PICH            =   "FrmContComprobante.frx":1CB1
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
         Left            =   8880
         TabIndex        =   5
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
         MICON           =   "FrmContComprobante.frx":1EE6
         PICN            =   "FrmContComprobante.frx":1F02
         PICH            =   "FrmContComprobante.frx":21E4
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
         Height          =   375
         Left            =   2400
         TabIndex        =   6
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
         MICON           =   "FrmContComprobante.frx":2435
         PICN            =   "FrmContComprobante.frx":2451
         PICH            =   "FrmContComprobante.frx":25F5
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
         Left            =   7200
         TabIndex        =   27
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
         MICON           =   "FrmContComprobante.frx":2794
         PICN            =   "FrmContComprobante.frx":27B0
         PICH            =   "FrmContComprobante.frx":28D5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BrnListaEmplesas 
         Height          =   375
         Left            =   3840
         TabIndex        =   28
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Lista de Empresas"
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
         MICON           =   "FrmContComprobante.frx":2B65
         PICN            =   "FrmContComprobante.frx":2B81
         PICH            =   "FrmContComprobante.frx":2E0A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnListaComprobantes 
         Height          =   375
         Left            =   5640
         TabIndex        =   29
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Comprobantes"
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
         MICON           =   "FrmContComprobante.frx":3225
         PICN            =   "FrmContComprobante.frx":3241
         PICH            =   "FrmContComprobante.frx":34CA
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
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.TextBox TxtNoMovimientos 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0"
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   1335
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   11175
         Begin VB.TextBox TxtNoItems 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "0,00"
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox TxtSaldo 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   8880
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "0,00"
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox TxtDetalle 
            Height          =   375
            Left            =   1200
            TabIndex        =   16
            Top             =   720
            Width           =   6255
         End
         Begin VB.TextBox TxtNoComprobante 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   5400
            TabIndex        =   18
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   60424195
            CurrentDate     =   40241
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Items:"
            Height          =   195
            Left            =   7800
            TabIndex        =   23
            Top             =   810
            Width           =   945
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo:"
            Height          =   195
            Left            =   8280
            TabIndex        =   21
            Top             =   330
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   4800
            TabIndex        =   17
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   810
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   330
            Width           =   990
         End
      End
      Begin VB.TextBox TxtDebe 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0,00"
         Top             =   6600
         Width           =   2055
      End
      Begin VB.TextBox TxtHaber 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   6600
         Width           =   2055
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   4815
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8493
         Object.Width           =   11145
         Object.Height          =   4785
         Cols            =   5
         Rows            =   1
         ScrollBar       =   1
         Editable        =   -1  'True
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregarCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Agregar Renglón"
         Top             =   6600
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "FrmContComprobante.frx":38E5
         PICN            =   "FrmContComprobante.frx":3901
         PICH            =   "FrmContComprobante.frx":3A8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnEliminarCuenta 
         Height          =   375
         Left            =   1200
         TabIndex        =   20
         ToolTipText     =   "Eliminar Renglón"
         Top             =   6600
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
         MICON           =   "FrmContComprobante.frx":3CC3
         PICN            =   "FrmContComprobante.frx":3CDF
         PICH            =   "FrmContComprobante.frx":3E83
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. de Movimientos:"
         Height          =   195
         Left            =   2640
         TabIndex        =   26
         Top             =   6690
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debe:"
         Height          =   195
         Left            =   5760
         TabIndex        =   11
         Top             =   6690
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Haber:"
         Height          =   195
         Left            =   8520
         TabIndex        =   9
         Top             =   6690
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmContComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IdEmpresa As Integer
Dim RegNew As Boolean
Dim RsTemp As Recordset
Public Cantidad As Double

Sub Calcular_Comprobante()
Dim TamDMGrid As Integer
Dim TotalDebe As Double
Dim TotalHaber As Double
Dim i As Integer

TamDMGrid = DMGrid1.Rows

TotalDebe = 0
TotalHaber = 0

' Ciclo para calcular los montos del DEBE y HABER de la fila 1 hasta la cantidad de fila del DMGrid1
For i = 1 To TamDMGrid

    'Si no es nulo el campo de fila "i" y columna 4 entonces
    If Not IsNull(DMGrid1.ValorCelda(i, 4)) Then
        ' Si el campo de fila "i" y columna 4 es diferente de "" entonces
        If Trim(DMGrid1.ValorCelda(i, 4)) <> "" Then
            TotalDebe = TotalDebe + CDbl(DMGrid1.ValorCelda(i, 4))
        Else
            'Si no es nulo el campo de fila "i" y columna 5 entonces
            If Not IsNull(DMGrid1.ValorCelda(i, 5)) Then
                ' Si el campo de fila "i" y columna 5 es diferente de "" entonces
                If Trim(DMGrid1.ValorCelda(i, 5)) <> "" Then TotalHaber = TotalHaber + CDbl(DMGrid1.ValorCelda(i, 5))
            End If
        End If
    Else
        'Si no es nulo el campo de fila "i" y columna 5 entonces
        If Not IsNull(DMGrid1.ValorCelda(i, 5)) Then
            ' Si el campo de fila "i" y columna 5 es diferente de "" entonces
            If Trim(DMGrid1.ValorCelda(i, 5)) <> "" Then TotalHaber = TotalHaber + CDbl(DMGrid1.ValorCelda(i, 5))
        End If
    End If
Next i
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

TxtDebe.Text = Format(TotalDebe, "#,##0.00")
TxtHaber.Text = Format(TotalHaber, "#,##0.00")
TxtSaldo.Text = Format(TotalDebe - TotalHaber, "#,##0.00")
Cantidad = TotalDebe - TotalHaber
End Sub

Public Sub Blanqueo()
    DMGrid1.Clear
    DMGrid1.Rows = 0
    TxtDetalle.Text = ""
    DTPicker1.Value = Now
    TxtSaldo.Text = Format(0, "#,##0.00")
    TxtDebe.Text = Format(0, "#,##0.00")
    TxtHaber.Text = Format(0, "#,##0.00")
    TxtNoMovimientos.Text = 0
    TxtNoItems.Text = 0
    DMGrid1.PaintMGrid
End Sub
Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 5
DMGrid1.Rows = 0

DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(4).Alignment = 1
DMGrid1.DColumnas(5).Alignment = 1

DMGrid1.DColumnas(4).IsNumber = True
DMGrid1.DColumnas(5).IsNumber = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 40 / 100) - 300
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Detalle de Movimiento"
DMGrid1.DColumnas(3).Caption = "Referencia"
DMGrid1.DColumnas(4).Caption = "Debe"
DMGrid1.DColumnas(5).Caption = "Haber"
End Sub
 
 
Sub Cargar_Comprobantes()
Dim IdComprobante As Integer
Blanqueo
If IdEmpresa = 0 Then Exit Sub

CSql = "SELECT * FROM ContComprobantes WHERE IdEmpresa=" & IdEmpresa & " AND Activo='1'" ' ORDER BY IdRengComprobante"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then: RegNew = True: Exit Sub

RegNew = False

IdComprobante = RsTemp.Fields("IdComprobante").Value

TxtNoComprobante.Text = RsTemp.Fields("NroComprobante").Value
TxtDetalle.Text = RsTemp.Fields("Detalle").Value
DTPicker1.Value = CDate(RsTemp.Fields("Fecha").Value)
TxtSaldo.Text = RsTemp.Fields("Saldo").Value

Call Cargar_Renglones(IdComprobante)

End Sub

Public Sub Cargar_Renglones(IdComprobante As Integer)
CSql = "SELECT * FROM ContComprobantesReng WHERE IdComprobante=" & IdComprobante & " ORDER BY IdRengComprobante"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

Dim CantDebe As Double
Dim CantHaber As Double
Dim i As Integer

i = 0
While Not RsTemp.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("Formato").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Detalle").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("Referencia").Value
    
    If Val(RsTemp.Fields("Tipo").Value) = 0 Then
        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsTemp.Fields("Cantidad").Value
        CantDebe = CantDebe + CDbl(RsTemp.Fields("Cantidad").Value)
    Else
        DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsTemp.Fields("Cantidad").Value
        CantHaber = CantHaber + CDbl(RsTemp.Fields("Cantidad").Value)
    End If
    i = i + 1
    RsTemp.MoveNext
Wend

TxtNoMovimientos.Text = i
TxtDebe.Text = Format(CantDebe, "#,##0.00")
TxtHaber.Text = Format(CantHaber, "#,##0.00")
DMGrid1.PaintMGrid
End Sub

Private Sub BrnListaEmplesas_Click()
BtnDesHacer_Click
Tipo = "Comprobante"
FrmContListaEmpresas.Show vbModal, FrmPrincipal
End Sub

Private Sub BtnAgregar_Click()
If IdEmpresa = 0 Then MsgBox "Debe Seleccionar una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub
RegNew = True
Blanqueo
BtnAgregar.Enabled = False
BtnEliminar.Enabled = False

CSql = "SELECT MAX(NroComprobante)+1 as NuevoId FROM ContComprobantes WHERE IdEmpresa=" & IdEmpresa
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    TxtNoComprobante.Text = Trim(RsTemp.Fields("NuevoId").Value)
Else
    TxtNoComprobante.Text = "1"
End If

End Sub

Private Sub BtnAgregarCuenta_Click()
If IdEmpresa = 0 Then MsgBox "Debe Seleccionar una empresa!", vbExclamation + vbOKOnly, "Error":  Exit Sub

'If DMGrid1.Rows = 0 Then DMGrid1.Rows = DMGrid1.Rows + 1:    DMGrid1.PaintMGrid

'If DMGrid1.ValorCelda(DMGrid1.Rows, 1) <> "" Then
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.PaintMGrid
'End If

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Form_Load
BtnAgregar.Enabled = True
BtnEliminar.Enabled = True
End Sub

Private Sub BtnEliminar_Click()
Dim resp As Byte
If IdEmpresa = 0 Then MsgBox "Debe Seleccionar una empresa!", vbExclamation + vbOKOnly, "Error":  Exit Sub
If Val(TxtNoComprobante.Text) = 0 Then MsgBox "Seleccione un comprobante de pago!", vbExclamation + vbOKOnly, "Error": Exit Sub

resp = MsgBox("Se procedera a eliminar el Comprobante Nro " & TxtNoComprobante.Text & ", Desea Continuar?", vbQuestion + vbYesNo, "Confirmar!")
If resp = vbNo Then Exit Sub


CSql = "UPDATE ContComprobantes SET Activo='0' WHERE IdEmpresa=" & IdEmpresa & " AND IdComprobante=" & Val(TxtNoComprobante.Text)
Set RsTemp = CrearRS(CSql)

BtnDesHacer_Click

MsgBox "El Comprobante ha sido eliminado.", vbInformation + vbOKOnly, "Operación Exitosa."
End Sub

Private Sub BtnEliminarCuenta_Click()
If DMGrid1.Rows <= 0 Or DMGrid1.Row = 0 Then Exit Sub
DMGrid1.RowDelete (DMGrid1.Row)
DMGrid1.PaintMGrid
Call Calcular_Comprobante
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim resp As Byte
Dim NuevoId As Integer
Dim NuevoIdReng As Integer
Dim i As Integer
Dim TotalNeto As Double

If IdEmpresa = 0 Then MsgBox "Debe Seleccionar una empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub
If Val(TxtNoComprobante.Text) = 0 Then MsgBox "Seleccione un comprobante de pago antes de guardar!", vbExclamation + vbOKOnly, "Error": Exit Sub

' MMMMMMMM VALIDACIONES MMMMMMMM

If Val(TxtNoComprobante.Text) = 0 Then
    MsgBox "Error en el número del comprobante!", vbExclamation + vbOKOnly, "Error"
    Exit Sub
ElseIf Trim(TxtDetalle.Text) = "" Then
    MsgBox "Debe ingresar el Detalle para el comprobante!", vbExclamation + vbOKOnly, "Error"
    TxtDetalle.SetFocus
    Exit Sub
ElseIf DMGrid1.Rows = 0 Then
    MsgBox "Debe al menos un item en el comprobante!", vbExclamation + vbOKOnly, "Error"
    TxtDetalle.SetFocus
    Exit Sub
ElseIf Trim(DMGrid1.ValorCelda(1, 1)) = "" Then
    MsgBox "Debe al menos un item en el comprobante!", vbExclamation + vbOKOnly, "Error"
    TxtDetalle.SetFocus
    Exit Sub
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
resp = MsgBox("Se procedera a guardar el Comprobante Nro " & TxtNoComprobante.Text & ", Desea Continuar?", vbQuestion + vbYesNo, "Confirmar!")
If resp = vbNo Then Exit Sub

CSql = "SELECT MAX(IdComprobante)+1 as NuevoId FROM ContComprobantes"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    NuevoId = Val(RsTemp.Fields("NuevoId").Value)
Else
    NuevoId = 1
End If

If CDbl(TxtDebe.Text) > CDbl(TxtHaber.Text) Then
    TotalNeto = CDbl(TxtDebe.Text)
Else
    TotalNeto = CDbl(TxtHaber.Text)
End If

Dim TamDMGrid As Integer
Dim Formato As String
Dim Detalle As String
Dim Referencia As String
Dim Cantid As Double
Dim Tipo As Byte

    ' Normaliza los Id de la tabla de los renglones de los comprobantes...
    CSql = "SELECT * FROM ContComprobantesReng ORDER BY IdRengComprobante"
    Set RsTemp = CrearRS(CSql)
    
    If RsTemp.RecordCount <> 0 Then
        i = 0
        While Not RsTemp.EOF
            i = i + 1
            RsTemp.Fields("IdRengComprobante") = i
            RsTemp.Update
            NuevoIdReng = i + 1
            RsTemp.MoveNext
        Wend
    Else
        NuevoIdReng = 1
    End If
If BtnAgregar.Enabled = False Then
    RegNew = True
Else
    RegNew = False
End If
If RegNew Then
    CSql = "INSERT INTO ContComprobantes (IdComprobante, IdEmpresa, NroComprobante, Fecha, Detalle, Total," & _
        "Saldo, IdUser, Activo) VALUES (" & NuevoId & ", " & IdEmpresa & ", " & Val(TxtNoComprobante.Text) & _
        ", '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "', '" & TxtDetalle.Text & "', " & Replace(TotalNeto, ",", ".") & _
        " , " & Replace(Replace(TxtSaldo.Text, ".", ""), ",", ".") & ", " & IdUser & ", '1')"
    Set RsTemp = CrearRS(CSql)
    
Else
    CSql = "UPDATE ContComprobantes SET  Fecha='" & Format(DTPicker1.Value, "dd/MM/yyyy") & _
        "', Detalle='" & TxtDetalle.Text & "', Total=" & Replace(TotalNeto, ",", ".") & "," & _
        "Saldo=" & Replace(Replace(TxtSaldo.Text, ".", ""), ",", ".") & ", IdUser=" & IdUser & ", Activo='1' " & _
        " WHERE IdEmpresa=" & IdEmpresa & " AND NroComprobante=" & Val(TxtNoComprobante.Text)
    Set RsTemp = CrearRS(CSql)
    
End If

    ' Sentencia que elimina los renglones del comprobante a actualizar, si es un
    ' nuevo comprobante entonces no ocurre nada
    CSql = "DELETE FROM ContComprobantesReng WHERE IdComprobante=" & Val(TxtNoComprobante.Text) & " AND IdEmpresa=" & IdEmpresa
    Set RsTemp = CrearRS(CSql)

    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMM Ingresa los Renglones del Comprobante MMMMMMMMMMMMMMMMMMM
    
    Dim NumLinea As Integer
    TamDMGrid = DMGrid1.Rows
    For i = 1 To TamDMGrid
        
        Formato = DMGrid1.ValorCelda(i, 1)
        Detalle = DMGrid1.ValorCelda(i, 2)
        Referencia = DMGrid1.ValorCelda(i, 3)
        NumLinea = NumLinea + 1
        
        If Trim(Formato) = "" Then Exit For
        
        If Not IsNull(DMGrid1.ValorCelda(i, 4)) Then
            If Val(DMGrid1.ValorCelda(i, 4)) <> 0 Then
                Tipo = 0
                Cantid = CDbl(DMGrid1.ValorCelda(i, 4))
            Else
                Tipo = 1
                Cantid = CDbl(DMGrid1.ValorCelda(i, 5))
            End If
        Else
            Tipo = 1
            Cantid = CDbl(DMGrid1.ValorCelda(i, 5))
        End If
        
        CSql = "INSERT INTO ContComprobantesReng (IdRengComprobante, IdEmpresa, IdComprobante, Formato, " & _
            " Detalle, Referencia, Tipo, Cantidad, Auxiliar,Linea) VALUES (" & NuevoIdReng & ", " & IdEmpresa & _
            ", " & Val(TxtNoComprobante.Text) & ", '" & Formato & "', '" & Detalle & _
            "', '" & Referencia & "', " & Tipo & ", " & Replace(Replace(Cantid, ".", ""), ",", ".") & ", 0," & NumLinea & ")"
        Set RsTemp = CrearRS(CSql)
        
        NuevoIdReng = NuevoIdReng + 1
    Next i
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


Form_Load
MsgBox "Los cambios han sido guardardos!", vbInformation + vbOKOnly, "Operación Exitosa."
End Sub

Private Sub BtnImprimir_Click()
If IdEmpresa = 0 Then MsgBox "Debe Seleccionar una empresa!", vbExclamation + vbOKOnly, "Error":  Exit Sub
End Sub

Private Sub BtnListaComprobantes_Click()
If IdEmpresa = 0 Then MsgBox "Debe Seleccionar una empresa!", vbExclamation + vbOKOnly, "Error":  Exit Sub
BtnDesHacer_Click
Tipo = "Comprobante"
FrmContListaComprobantes.IdEmpresa = IdEmpresa
FrmContListaComprobantes.Show vbModal, FrmPrincipal
End Sub

Private Sub DMGrid1_AfterColEdit(ByVal lRow As Integer, ByVal lCol As Integer, Cancelar As Boolean)
Calcular_Comprobante
End Sub

Private Sub DMGrid1_BeforeColEdit(ByVal lRow As Integer, ByVal lCol As Integer)
If lCol = 5 Then DMGrid1.ValorCelda(lRow, lCol) = Cantidad: DMGrid1.TextEdit = Cantidad: DMGrid1.PaintMGrid
End Sub


Private Sub DMGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If IdEmpresa = 0 Then MsgBox "Debe Seleccionar una empresa!", vbExclamation + vbOKOnly, "Error": KeyCode = 0: Exit Sub

If DMGrid1.Col <> 1 And KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight Then
    If Trim(DMGrid1.ValorCelda(DMGrid1.Row, 1)) = "" Then KeyCode = 0: Exit Sub
End If

If KeyCode = 13 And DMGrid1.Col = 5 Then
    If DMGrid1.ValorCelda(DMGrid1.Row, 1) <> "" Then
        If DMGrid1.Rows = DMGrid1.Row Then
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.PaintMGrid
            DMGrid1.Row = DMGrid1.Rows
            DMGrid1.Col = 1
        Else
            DMGrid1.Row = DMGrid1.Row + 1
            DMGrid1.Col = 1
        End If
    End If
    Calcular_Comprobante
    Exit Sub
End If

If DMGrid1.Col = 5 And KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight Then
    If DMGrid1.ValorCelda(DMGrid1.Row, 4) <> "" Then KeyCode = 0
ElseIf DMGrid1.Col = 4 And KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight Then
    If DMGrid1.ValorCelda(DMGrid1.Row, 5) <> "" Then KeyCode = 0
End If

If KeyCode = vbKeyF1 And DMGrid1.Col = 1 Then
    Tipo = "Comprobante"
    FrmContListaPDC.IdEmpresa = IdEmpresa
    FrmContListaPDC.Show vbModal, FrmPrincipal
ElseIf KeyCode = vbKeyF1 And DMGrid1.Col = 2 And DMGrid1.ValorCelda(DMGrid1.Row, 1) <> "" Then
    Tipo = "Comprobante"
    FrmContListaDetallesMov.IdEmpresa = IdEmpresa
    FrmContListaDetallesMov.Show vbModal, FrmPrincipal
End If
End Sub

Private Sub Form_Activate()
Tipo = "Comprobante"
End Sub

Private Sub Form_Load()
IniDMGrid
Blanqueo

If IdEmpresa <> 0 Then
    CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmpresa
    Set RsTemp = CrearRS(CSql)
    
    If RsTemp.RecordCount <> 0 Then
        FrmContComprobante.Caption = "Comprabante Contable para la empresa '" & RsTemp.Fields("Nombre").Value & "'"
    End If
End If

Cargar_Comprobantes
End Sub
