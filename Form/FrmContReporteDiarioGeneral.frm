VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReporteDiarioGeneral 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diario General"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   Icon            =   "FrmContReporteDiarioGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   4335
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         ToolTipText     =   "Cerrar "
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
         MICON           =   "FrmContReporteDiarioGeneral.frx":1002
         PICN            =   "FrmContReporteDiarioGeneral.frx":101E
         PICH            =   "FrmContReporteDiarioGeneral.frx":11E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnResultados 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Resultados"
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
         MICON           =   "FrmContReporteDiarioGeneral.frx":141C
         PICN            =   "FrmContReporteDiarioGeneral.frx":1438
         PICH            =   "FrmContReporteDiarioGeneral.frx":16CA
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmContReporteDiarioGeneral.frx":1857
         Left            =   1560
         List            =   "FrmContReporteDiarioGeneral.frx":18B8
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmContReporteDiarioGeneral.frx":192F
         Left            =   1560
         List            =   "FrmContReporteDiarioGeneral.frx":1993
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox TxtLineaDetalle 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   56164355
         CurrentDate     =   40254
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período a imprimir:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde el día:"
         Height          =   195
         Left            =   465
         TabIndex        =   7
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta el día:"
         Height          =   195
         Left            =   510
         TabIndex        =   6
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lineas de detalle:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1905
         Visible         =   0   'False
         Width           =   1245
      End
   End
End
Attribute VB_Name = "FrmContReporteDiarioGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim i As Integer

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnResultados_Click()

Tipo = "ReporteDiarioGeneral"

Dim DDesde As String
Dim DHasta As String
Dim Periodo As String
Dim Cond_SQL As String
Dim NCompr As Integer
Dim IniCompr As Integer
Dim FinCompr As Integer
Dim CantDebe As Double
Dim CantHaber As Double
Dim Band As Boolean
Dim AgregoTotal As Boolean

If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0: Combo2.ListIndex = Combo2.ListCount - 1

Periodo = Format(DTPicker1.Value, "/MM/yyyy")

DDesde = Combo1.List(Combo1.ListIndex) & Periodo
DHasta = Combo2.List(Combo2.ListIndex) & Periodo

If Val(Format(DHasta, "00")) > Val(Format(DateSerial(Year(CDate(DDesde)), Month(CDate(DDesde)) + 1, 0), "dd")) Then
    DHasta = Format(DateSerial(Year(CDate(DDesde)), Month(CDate(DDesde)) + 1, 0), "dd") & Periodo
End If

CSql = "SELECT * FROM ContReporteDiarioGeneral WHERE Fecha >= '" & DDesde & "' AND Fecha <= '" & DHasta & _
"' ORDER BY Fecha, NroComprobante"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

With FrmContReportes

.DMGrid1.Clear
.DMGrid1.Rows = 0

.DMGrid1.Cols = 8
.DMGrid1.Rows = 0

.DMGrid1.DColumnas(1).Caption = "Día"
.DMGrid1.DColumnas(2).Caption = "Nro"
.DMGrid1.DColumnas(3).Caption = "Comprobante / Detalle"
.DMGrid1.DColumnas(4).Caption = "Referencia"
.DMGrid1.DColumnas(5).Caption = "Código"
.DMGrid1.DColumnas(6).Caption = "Descripción de la Cuenta"
.DMGrid1.DColumnas(7).Caption = "Débitos"
.DMGrid1.DColumnas(8).Caption = "Créditos"

.DMGrid1.DColumnas(1).Alignment = 1
.DMGrid1.DColumnas(2).Alignment = 1
.DMGrid1.DColumnas(4).Alignment = 1
.DMGrid1.DColumnas(7).Alignment = 1
.DMGrid1.DColumnas(8).Alignment = 1

.DMGrid1.DColumnas(7).IsNumber = True
.DMGrid1.DColumnas(8).IsNumber = True

.DMGrid1.DColumnas(1).Width = Val(.DMGrid1.Width * 5 / 100)
.DMGrid1.DColumnas(2).Width = Val(.DMGrid1.Width * 5 / 100)
.DMGrid1.DColumnas(3).Width = Val(.DMGrid1.Width * 20 / 100)
.DMGrid1.DColumnas(4).Width = Val(.DMGrid1.Width * 10 / 100) - 300
.DMGrid1.DColumnas(5).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(6).Width = Val(.DMGrid1.Width * 20 / 100)
.DMGrid1.DColumnas(7).Width = Val(.DMGrid1.Width * 15 / 100)
.DMGrid1.DColumnas(8).Width = Val(.DMGrid1.Width * 15 / 100)


Band = False
While Not RsTemp.EOF

    .DMGrid1.Rows = .DMGrid1.Rows + 1
    
    .DMGrid1.RowBackColor 1, RGB(221, 221, 221)
    If RsTemp.Fields("NroComprobante").Value <> NCompr Then
        
        If .DMGrid1.Rows > 0 Then
            For i = IniCompr To FinCompr
                If Val(.DMGrid1.ValorCelda(i, 7)) = 0 Then
                    If Val(.DMGrid1.ValorCelda(i, 8)) <> 0 Then CantHaber = CantHaber + CDbl(.DMGrid1.ValorCelda(i, 8))
                Else
                    If Val(.DMGrid1.ValorCelda(i, 7)) <> 0 Then CantDebe = CantDebe + CDbl(.DMGrid1.ValorCelda(i, 7))
                End If
            Next i
            'If Band Then .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221) Else .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(255, 255, 255)
            .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
            
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = "Total Comprobante: "
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = CantDebe
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = CantHaber
        
            AgregoTotal = True
            
            If CantDebe <> 0 And CantHaber <> 0 Then
                .DMGrid1.Rows = .DMGrid1.Rows + 1
                .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
                .DMGrid1.Rows = .DMGrid1.Rows + 1
            End If
            CantHaber = 0
            CantDebe = 0
        End If
        
        If AgregoTotal Then .DMGrid1.Rows = .DMGrid1.Rows + 1
        
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(162, 162, 162)
        
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = Day(RsTemp.Fields("Fecha").Value)
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = RsTemp.Fields("NroComprobante").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp.Fields("Detalle").Value
        IniCompr = .DMGrid1.Rows
        
    Else
    
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
        
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp.Fields("Detalle2").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = RsTemp.Fields("Referencia").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = RsTemp.Fields("Formato").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = RsTemp.Fields("Nombre").Value
        FinCompr = .DMGrid1.Rows
    End If
    
    If AgregoTotal Then
        
        .DMGrid1.Rows = .DMGrid1.Rows + 1
        AgregoTotal = False
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp.Fields("Detalle2").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = RsTemp.Fields("Referencia").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = RsTemp.Fields("Formato").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = RsTemp.Fields("Nombre").Value
    End If
    
    If Val(RsTemp.Fields("Tipo").Value) = 0 Then
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = RsTemp.Fields("Cantidad").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = ""
    Else
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = RsTemp.Fields("Cantidad").Value
    End If
    
    NCompr = RsTemp.Fields("NroComprobante").Value
    RsTemp.MoveNext
    
    If RsTemp.EOF Then
        If .DMGrid1.Rows > 0 Then
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            For i = IniCompr To FinCompr
                If Val(.DMGrid1.ValorCelda(i, 7)) = 0 Then
                    If Val(.DMGrid1.ValorCelda(i, 8)) <> 0 Then CantHaber = CantHaber + CDbl(.DMGrid1.ValorCelda(i, 8))
                Else
                    If Val(.DMGrid1.ValorCelda(i, 7)) <> 0 Then CantDebe = CantDebe + CDbl(.DMGrid1.ValorCelda(i, 7))
                End If
            Next i
            
            .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(162, 162, 162)
            '.DMGrid1.RowForeColor .DMGrid1.Rows - 1, vbMagenta
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = "Total Comprobante: "
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = CantDebe
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = CantHaber
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            CantHaber = 0
            CantDebe = 0
        End If
    End If

Wend

.DMGrid1.RowDelete 1
.DMGrid1.PaintMGrid
.Show vbModal, FrmPrincipal

End With

End Sub

Private Sub Combo1_Change()
Combo2_Change
End Sub

Private Sub Combo1_Click()
Combo2_Change
End Sub

Private Sub Combo2_Change()
If Val(Combo2.List(Combo2.ListIndex)) <= Val(Combo1.List(Combo1.ListIndex)) Then
    Combo2.ListIndex = Combo1.ListIndex + 1
End If
End Sub

Private Sub Combo2_Click()
Combo2_Change
End Sub

Private Sub Form_Load()
Centrar Me
CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then Frame1.Caption = Trim(RsTemp.Fields("Nombre")) & " RIF " & Trim(RsTemp.Fields("Rif"))

End Sub

Private Sub TxtLineaDetalle_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
