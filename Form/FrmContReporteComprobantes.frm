VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReporteComprobantes 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Comprobantes"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "FrmContReporteComprobantes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Mostrar uno por página"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   2640
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Detallado"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   13
         Top             =   2280
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Resumido"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   12
         Top             =   2280
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmContReporteComprobantes.frx":1002
         Left            =   1560
         List            =   "FrmContReporteComprobantes.frx":1066
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmContReporteComprobantes.frx":10DF
         Left            =   1560
         List            =   "FrmContReporteComprobantes.frx":1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55246851
         CurrentDate     =   40254
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55246851
         CurrentDate     =   40254
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta la Fecha:"
         Height          =   195
         Left            =   285
         TabIndex        =   11
         Top             =   930
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta el Nro:"
         Height          =   195
         Left            =   510
         TabIndex        =   9
         Top             =   1860
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde el Nro:"
         Height          =   195
         Left            =   465
         TabIndex        =   8
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde la Fecha:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   450
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   4335
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
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
         MICON           =   "FrmContReporteComprobantes.frx":11B7
         PICN            =   "FrmContReporteComprobantes.frx":11D3
         PICH            =   "FrmContReporteComprobantes.frx":139C
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
         TabIndex        =   2
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
         MICON           =   "FrmContReporteComprobantes.frx":15D1
         PICN            =   "FrmContReporteComprobantes.frx":15ED
         PICH            =   "FrmContReporteComprobantes.frx":187F
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
End
Attribute VB_Name = "FrmContReporteComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim RsTemp2 As Recordset
Dim i As Integer
Dim j As Integer

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnResultados_Click()

If Option1(0).Value = True Then
    CSql = "SELECT NroComprobante, Fecha, Detalle, Tipo, SUM(Cantidad) AS Total FROM ContReporteComprobanteDetallado " & _
    "WHERE (IdEmpresa = " & IdEmprs & ") AND (Fecha >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "') AND (Fecha <= '" & Format(DTPicker2.Value, "dd/MM/yyyy") & _
    "') AND NroComprobante>=" & Val(Combo1.List(Combo1.ListIndex)) & " AND NroComprobante<=" & Val(Combo2.List(Combo2.ListIndex)) & " AND ACTIVO='1' GROUP BY NroComprobante, Fecha, Detalle, Tipo ORDER BY Fecha, NroComprobante"
Else
    CSql = "SELECT NroComprobante, Fecha, Detalle, Total, Saldo, Formato, Detalle2, Referencia, Tipo, Cantidad, IdEmpresa FROM ContReporteComprobanteDetallado " & _
    "WHERE (IdEmpresa = " & IdEmprs & ") AND (Fecha >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "') AND (Fecha <= '" & Format(DTPicker2.Value, "dd/MM/yyyy") & _
    "') AND NroComprobante>=" & Val(Combo1.List(Combo1.ListIndex)) & " AND NroComprobante<=" & Val(Combo1.List(Combo1.ListIndex)) & " AND ACTIVO='1' GROUP BY NroComprobante, Fecha, Detalle, Tipo, Detalle, Total, Saldo, Formato, Detalle2, Referencia, " & _
    " Cantidad, IdEmpresa ORDER BY Fecha, NroComprobante"
End If

Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub


With FrmContReportes

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMM Inicializa el DMGrid del formulario del reporte MMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
.DMGrid1.Clear
.DMGrid1.Rows = 0

.DMGrid1.Cols = 9
.DMGrid1.Rows = 0

.DMGrid1.DColumnas(1).Caption = "Fecha"
.DMGrid1.DColumnas(2).Caption = "Nro"
.DMGrid1.DColumnas(3).Caption = "Descripción"
.DMGrid1.DColumnas(4).Caption = "Débitos"
.DMGrid1.DColumnas(5).Caption = "Créditos"
.DMGrid1.DColumnas(9).Caption = "NroCompro"
.DMGrid1.DColumnas(9).Visible = False

.DMGrid1.DColumnas(4).Alignment = 1
.DMGrid1.DColumnas(5).Alignment = 1

.DMGrid1.DColumnas(4).IsNumber = True
.DMGrid1.DColumnas(5).IsNumber = True

.DMGrid1.DColumnas(1).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(2).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(3).Width = Val(.DMGrid1.Width * 40 / 100) - 300
.DMGrid1.DColumnas(4).Width = Val(.DMGrid1.Width * 20 / 100)
.DMGrid1.DColumnas(5).Width = Val(.DMGrid1.Width * 20 / 100)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Dim TamDMGrid As Integer
Dim Band As Boolean
Dim ContaReporte As Integer

ContaReporte = 0

While Not RsTemp.EOF
    If Option1(0).Value = True Then
    
        TamDMGrid = .DMGrid1.Rows
        .DMGrid1.RowBackColor 1, RGB(162, 162, 162)
        
        ' Se configura la bandera a FALSE para que de acuerdo al ciclo la active o no
        Band = False
        
        ' Ciclo que verifica si el comprobante ya esta en lista, de ser asi, activa la bandera a TRUE
        ' para que mas adelante saber que NO se agregara otra fila sino que se usara la misma fila
        ' en la cual se encuentra el comprobante...
        For i = 1 To TamDMGrid
            If Val(.DMGrid1.ValorCelda(i, 2)) = Val(RsTemp.Fields("NroComprobante").Value) And _
               Format(.DMGrid1.ValorCelda(i, 1), "dd/MM/yyyy") = Format(RsTemp.Fields("Fecha").Value, "dd/MM/yyyy") Then
                Band = True
                Exit For
            End If
        Next i
        
        ' Si Band esta activo entonces no agrega otra fila sino que trabaja en la misma fila
        If Band Then
            If Val(RsTemp.Fields("Tipo").Value) = 0 Then
                .DMGrid1.ValorCelda(i, 4) = RsTemp.Fields("Total").Value
            Else
                .DMGrid1.ValorCelda(i, 5) = RsTemp.Fields("Total").Value
            End If
        Else
        ' en el caso contrario al condicional entonces me agrega una fila nueva junto con el comprobante.
            
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = Format(Trim(RsTemp.Fields("Fecha").Value), "dd/MM/yyyy")
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = RsTemp.Fields("NroComprobante").Value
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = Trim(RsTemp.Fields("Detalle").Value)
            
            ' condicional para saber si es Débito (Tipo=0) ó Crédito (Tipo=1)
            If Val(RsTemp.Fields("Tipo").Value) = 0 Then
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = RsTemp.Fields("Total").Value
            Else
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = RsTemp.Fields("Total").Value
            End If
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
        End If
    Else
        ' Si se pide un reporte DETALLADO de comprobantes entonces hacer los procesos pertinentes
        Mostrar_Comprobantes_Detallados
        RsTemp.MoveLast
    End If
    ContaReporte = ContaReporte + 1
    RsTemp.MoveNext
Wend
.DMGrid1.PaintMGrid

If Option1(0).Value = True Then
    Tipo = "ReporteComprobantes"
Else
    Tipo = "ReporteComprobantes2"
End If
.Show vbModal, FrmPrincipal
End With

End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Sub Mostrar_Comprobantes_Detallados()
Tipo = "ReporteComprobantesDetallados"

Dim DDesde As String
Dim DHasta As String
Dim Periodo As String
Dim Cond_SQL As String
Dim NCompr As Integer
Dim IniCompr As Integer
Dim FinCompr As Integer
Dim ContaReporte As Integer
Dim CantDebe As Double
Dim CantHaber As Double
Dim Band As Boolean
Dim AgregoTotal As Boolean

If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0: Combo2.ListIndex = Combo2.ListCount - 1

DDesde = Format(DTPicker1.Value, "dd/MM/yyyy")
DHasta = Format(DTPicker2.Value, "dd/MM/yyyy")

CSql = "SELECT * FROM ContReporteDiarioGeneral WHERE Fecha >= '" & DDesde & "' AND Fecha <= '" & DHasta & "' AND NroComprobante>=" & Combo1.List(Combo1.ListIndex) & " AND NroComprobante<=" & Combo2.List(Combo2.ListIndex) & " ORDER BY Fecha, NroComprobante"
Set RsTemp2 = CrearRS(CSql)

If RsTemp2.RecordCount = 0 Then Exit Sub

ContaReporte = 0

With FrmContReportes

.DMGrid1.Clear
.DMGrid1.Rows = 0

.DMGrid1.Cols = 9
.DMGrid1.Rows = 0

.DMGrid1.DColumnas(1).Caption = "Fecha"
.DMGrid1.DColumnas(2).Caption = "Nro"
.DMGrid1.DColumnas(3).Caption = "Comprobante / Detalle"
.DMGrid1.DColumnas(4).Caption = "Referencia"
.DMGrid1.DColumnas(5).Caption = "Código"
.DMGrid1.DColumnas(6).Caption = "Descripción de la Cuenta"
.DMGrid1.DColumnas(7).Caption = "Débitos"
.DMGrid1.DColumnas(8).Caption = "Créditos"
.DMGrid1.DColumnas(9).Caption = "NrCompr"
.DMGrid1.DColumnas(9).Visible = False

.DMGrid1.DColumnas(1).Alignment = 1
.DMGrid1.DColumnas(2).Alignment = 1
.DMGrid1.DColumnas(4).Alignment = 1
.DMGrid1.DColumnas(7).Alignment = 1
.DMGrid1.DColumnas(8).Alignment = 1

.DMGrid1.DColumnas(7).IsNumber = True
.DMGrid1.DColumnas(8).IsNumber = True

.DMGrid1.DColumnas(1).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(2).Width = Val(.DMGrid1.Width * 5 / 100)
.DMGrid1.DColumnas(3).Width = Val(.DMGrid1.Width * 20 / 100)
.DMGrid1.DColumnas(4).Width = Val(.DMGrid1.Width * 9 / 100)
.DMGrid1.DColumnas(5).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(6).Width = Val(.DMGrid1.Width * 20 / 100) - 300
.DMGrid1.DColumnas(7).Width = Val(.DMGrid1.Width * 13 / 100)
.DMGrid1.DColumnas(8).Width = Val(.DMGrid1.Width * 13 / 100)


Band = False
IniCompr = 1

While Not RsTemp2.EOF

    .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
    .DMGrid1.Rows = .DMGrid1.Rows + 1
    .DMGrid1.RowBackColor 1, RGB(162, 162, 162)
    
    If RsTemp2.Fields("NroComprobante").Value <> NCompr Then
        
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
            
            If CantDebe > 0# And CantHaber <> 0# Then
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
                .DMGrid1.Rows = .DMGrid1.Rows + 1
                .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
                .DMGrid1.Rows = .DMGrid1.Rows + 1
            End If
            CantHaber = 0
            CantDebe = 0
            ContaReporte = ContaReporte + 1
        End If
        
        If AgregoTotal Then .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte: .DMGrid1.Rows = .DMGrid1.Rows + 1
        
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(162, 162, 162)
        
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = Format(Trim(RsTemp2.Fields("Fecha").Value), "dd/MM/yyyy")
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = RsTemp2.Fields("NroComprobante").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp2.Fields("Detalle").Value
        IniCompr = .DMGrid1.Rows
    Else
    
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
        
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp2.Fields("Detalle2").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = RsTemp2.Fields("Referencia").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = RsTemp2.Fields("Formato").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = RsTemp2.Fields("Nombre").Value
        FinCompr = .DMGrid1.Rows
    End If
    
    If AgregoTotal Then
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
        .DMGrid1.Rows = .DMGrid1.Rows + 1
        AgregoTotal = False
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp2.Fields("Detalle2").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = RsTemp2.Fields("Referencia").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = RsTemp2.Fields("Formato").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = RsTemp2.Fields("Nombre").Value
    End If
    
    If Val(RsTemp2.Fields("Tipo").Value) = 0 Then
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = RsTemp2.Fields("Cantidad").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = ""
    Else
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = RsTemp2.Fields("Cantidad").Value
    End If
    .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
    NCompr = RsTemp2.Fields("NroComprobante").Value
    RsTemp2.MoveNext

    If RsTemp2.EOF Then
        If .DMGrid1.Rows > 0 Then
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            For i = IniCompr To FinCompr
                If Val(.DMGrid1.ValorCelda(i, 7)) = 0 Then
                    If Val(.DMGrid1.ValorCelda(i, 8)) <> 0 Then CantHaber = CantHaber + CDbl(.DMGrid1.ValorCelda(i, 8))
                Else
                    If Val(.DMGrid1.ValorCelda(i, 7)) <> 0 Then CantDebe = CantDebe + CDbl(.DMGrid1.ValorCelda(i, 7))
                End If
            Next i
            'If Band Then .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221) Else .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(255, 255, 255)
            .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(162, 162, 162)
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = "Total Comprobante: "
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = CantDebe
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = CantHaber
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 9) = ContaReporte
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            CantHaber = 0
            CantDebe = 0
        End If
    End If

Wend
.DMGrid1.RowDelete 1
.DMGrid1.PaintMGrid

End With

End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Private Sub Combo1_Click()
Combo2_Change
End Sub
Private Sub Combo2_Change()
If Val(Combo2.List(Combo2.ListIndex)) <= Val(Combo1.List(Combo1.ListIndex)) Then
    Combo2.ListIndex = Combo1.ListIndex
End If
End Sub
Private Sub Combo2_Click()
Combo2_Change
End Sub
Private Sub Combo1_Change()
Combo2_Change
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Private Sub DTPicker1_Change()
DTPicker1_Click
End Sub
Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
DTPicker1_Click
End Sub
Private Sub DTPicker1_Click()
If CDate(DTPicker1.Value) > CDate(DTPicker2.Value) Then
    DTPicker2.Value = DTPicker1.Value
End If
End Sub
Private Sub DTPicker2_Change()
DTPicker1_Click
End Sub
Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
DTPicker1_Click
End Sub
Private Sub DTPicker2_Click()
DTPicker1_Click
End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Private Sub Form_Load()
Dim NCompbts As Integer
Dim i As Integer

Centrar Me

CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then Frame1.Caption = Trim(RsTemp.Fields("Nombre")) & " RIF " & Trim(RsTemp.Fields("Rif"))

CSql = "SELECT * FROM ContComprobantes WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

DTPicker1.Value = "01" & Format(Now, "/MM/yyyy")
DTPicker2.Value = Format(DateSerial(Year(CDate(Now)), Month(CDate(Now)) + 1, 0), "dd/MM/yyyy")

Combo1.Clear
Combo2.Clear
Combo1.AddItem ""
Combo2.AddItem ""
If RsTemp.RecordCount = 0 Then Exit Sub

Combo1.Clear
Combo2.Clear

NCompbts = Val(RsTemp.RecordCount)

For i = 1 To NCompbts
    Combo1.AddItem i
    Combo2.AddItem i
Next i

End Sub
