VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReporteMayorAnalitico 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mayor Analitico"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   Icon            =   "FrmContReporteMayorAnalitico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox Check5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Ajustado por inflación"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Resumir lineas de detalle"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   3360
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Todas las transacciones"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Incluir terceros"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Incluir la columna saldo del mes"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmContReporteMayorAnalitico.frx":1002
         Left            =   1560
         List            =   "FrmContReporteMayorAnalitico.frx":1066
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmContReporteMayorAnalitico.frx":10DF
         Left            =   1560
         List            =   "FrmContReporteMayorAnalitico.frx":1140
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
         Format          =   55508995
         CurrentDate     =   40254
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55508995
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
         Caption         =   "Hasta la cuenta:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1860
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde la cuenta:"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde la Fecha:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   450
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2400
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
         MICON           =   "FrmContReporteMayorAnalitico.frx":11B7
         PICN            =   "FrmContReporteMayorAnalitico.frx":11D3
         PICH            =   "FrmContReporteMayorAnalitico.frx":139C
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
         MICON           =   "FrmContReporteMayorAnalitico.frx":15D1
         PICN            =   "FrmContReporteMayorAnalitico.frx":15ED
         PICH            =   "FrmContReporteMayorAnalitico.frx":187F
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
Attribute VB_Name = "FrmContReporteMayorAnalitico"
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

Dim DDesde As String
Dim DHasta As String
Dim Cond_SQL As String
Dim NForma As String
Dim IniCompr As Integer
Dim FinCompr As Integer
Dim CantDebe As Double
Dim CantHaber As Double
Dim Band As Boolean
Dim AgregoTotal As Boolean

Dim ContaReporte As Integer

If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0: Combo2.ListIndex = Combo2.ListCount - 1

DDesde = Format(DTPicker1.Value, "dd/MM/yyyy")
DHasta = Format(DTPicker2.Value, "dd/MM/yyyy")

CSql = "SELECT * FROM ContReporteMayorAnalitico WHERE Fecha >= '" & DDesde & "' AND Fecha <= '" & DHasta & "' AND IdEmpresa=" & IdEmprs & " ORDER BY Formato,Fecha, NroComprobante"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

With FrmContReportes

.DMGrid1.Clear
.DMGrid1.Rows = 0

.DMGrid1.Cols = 8
.DMGrid1.Rows = 0

.DMGrid1.DColumnas(1).Caption = "Fecha"
.DMGrid1.DColumnas(2).Caption = "Nro Comprobt."
.DMGrid1.DColumnas(3).Caption = "Detalle del Movto."
.DMGrid1.DColumnas(4).Caption = "Referencia"
.DMGrid1.DColumnas(5).Caption = "Débitos"
.DMGrid1.DColumnas(6).Caption = "Créditos"
.DMGrid1.DColumnas(7).Caption = "IdTipo"
.DMGrid1.DColumnas(8).Caption = "Linea"
.DMGrid1.DColumnas(7).Visible = False
.DMGrid1.DColumnas(8).Visible = False

.DMGrid1.DColumnas(1).Alignment = 1
.DMGrid1.DColumnas(2).Alignment = 1
.DMGrid1.DColumnas(4).Alignment = 1
.DMGrid1.DColumnas(5).Alignment = 1
.DMGrid1.DColumnas(6).Alignment = 1

.DMGrid1.DColumnas(5).IsNumber = True
.DMGrid1.DColumnas(6).IsNumber = True

.DMGrid1.DColumnas(1).Width = Val(.DMGrid1.Width * 20 / 100)
.DMGrid1.DColumnas(2).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(3).Width = Val(.DMGrid1.Width * 30 / 100) - 300
.DMGrid1.DColumnas(4).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(5).Width = Val(.DMGrid1.Width * 15 / 100)
.DMGrid1.DColumnas(6).Width = Val(.DMGrid1.Width * 15 / 100)

Band = False
IniCompr = 1
ContaReporte = 0

While Not RsTemp.EOF

    .DMGrid1.Rows = .DMGrid1.Rows + 1
    .DMGrid1.RowBackColor 1, RGB(162, 162, 162)
    
    If RsTemp.Fields("Formato").Value <> NForma Then
        
        If .DMGrid1.Rows > 0 And NForma <> "" Then
            For i = IniCompr To FinCompr
                If EsNumerico(.DMGrid1.ValorCelda(i, 5)) Then
                    If CDbl(.DMGrid1.ValorCelda(i, 5)) > 0# Then CantDebe = CantDebe + CDbl(.DMGrid1.ValorCelda(i, 5))
                Else
                    If EsNumerico(.DMGrid1.ValorCelda(i, 6)) Then CantHaber = CantHaber + CDbl(.DMGrid1.ValorCelda(i, 6))
                End If
            Next i
            
            'If Band Then .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221) Else .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(255, 255, 255)
            .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
            
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = "Total Comprobante: "
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = CantDebe
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = CantHaber
            AgregoTotal = True
            
            If CantDebe <> 0 Or CantHaber <> 0 Then
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ContaReporte
                .DMGrid1.Rows = .DMGrid1.Rows + 1
                .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ContaReporte
                .DMGrid1.Rows = .DMGrid1.Rows + 1
            End If
            CantHaber = 0
            CantDebe = 0
            ContaReporte = ContaReporte + 1
        End If
        
        If AgregoTotal Then .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ContaReporte: .DMGrid1.Rows = .DMGrid1.Rows + 1
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(162, 162, 162)
        
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = RsTemp.Fields("Formato").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = Trim(RsTemp.Fields("DetallePDC").Value)
        IniCompr = .DMGrid1.Rows
        'If NForma = "" Then .DMGrid1.Rows = .DMGrid1.Rows + 1: .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(162, 162, 162)
    Else
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
        
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = Format(RsTemp.Fields("Fecha").Value, "dd/MM/yy")
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = RsTemp.Fields("NroComprobante").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp.Fields("Detalle2").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = RsTemp.Fields("Referencia").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = RsTemp.Fields("Linea").Value
        FinCompr = .DMGrid1.Rows
    End If
    
    If AgregoTotal Or NForma = "" Then
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ContaReporte
        .DMGrid1.Rows = .DMGrid1.Rows + 1
        AgregoTotal = False
        .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221)
        
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = Format(RsTemp.Fields("Fecha").Value, "dd/MM/yy")
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = RsTemp.Fields("NroComprobante").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp.Fields("Detalle2").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = RsTemp.Fields("Referencia").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 8) = RsTemp.Fields("Linea").Value
        
        FinCompr = .DMGrid1.Rows
    End If
    
    If Val(RsTemp.Fields("Tipo").Value) = 0 Then
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = RsTemp.Fields("Cantidad").Value
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = ""
    Else
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = ""
        .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = RsTemp.Fields("Cantidad").Value
    End If
    
    NForma = RsTemp.Fields("Formato").Value
    RsTemp.MoveNext

    If RsTemp.EOF Then
        If .DMGrid1.Rows > 0 Then
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ContaReporte
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            For i = IniCompr To FinCompr
                If Val(.DMGrid1.ValorCelda(i, 5)) = 0 Then
                    If Val(.DMGrid1.ValorCelda(i, 6)) <> 0 Then CantHaber = CantHaber + CDbl(.DMGrid1.ValorCelda(i, 6))
                Else
                    If Val(.DMGrid1.ValorCelda(i, 5)) <> 0 Then CantDebe = CantDebe + CDbl(.DMGrid1.ValorCelda(i, 5))
                End If
            Next i
            'If Band Then .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(221, 221, 221) Else .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(255, 255, 255)
            .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(162, 162, 162)
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = "Total Comprobante: "
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = CantDebe
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 6) = CantHaber
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ContaReporte
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            CantHaber = 0
            CantDebe = 0
        End If
    End If
    .DMGrid1.ValorCelda(.DMGrid1.Rows, 7) = ContaReporte
Wend

.DMGrid1.PaintMGrid
Tipo = "ReporteMayorAnalitico"
.Show vbModal, FrmPrincipal

End With
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Function EsNumerico(ByVal Cad As String) As Boolean
Dim cont As Byte
Dim ii As Integer

If IsNull(Cad) Then EsNumerico = False: Exit Function
If Trim(Cad) = "" Then EsNumerico = False: Exit Function

For ii = 1 To Len(Cad)
    If Not IsNumeric(Mid(Cad, ii, 1)) Then cont = cont + 1
Next ii

If cont > 1 Then EsNumerico = False Else EsNumerico = True

End Function
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Private Sub Combo1_Click()
Combo2_Change
End Sub
Private Sub Combo2_Change()
Dim Formt1 As String
Dim Formt2 As String
Dim Caracter As String
Dim Band As Boolean
Dim i As Integer
Dim Diferencia As Integer

Formt1 = Combo1.List(Combo1.ListIndex)
Formt2 = Combo2.List(Combo2.ListIndex)
Band = False

For i = 1 To Len(Formt1)
    Caracter = Mid(Formt1, i, 1)
    If Not IsNumeric(Caracter) Then
        Formt1 = Replace(Formt1, Caracter, "")
        Formt2 = Replace(Formt2, Caracter, "")
        Band = True
        Exit For
    End If
Next i

If Not Band Then
    For i = 1 To Len(Formt2)
        Caracter = Mid(Formt2, i, 1)
        If Not IsNumeric(Caracter) Then
            Formt1 = Replace(Formt1, Caracter, "")
            Formt2 = Replace(Formt2, Caracter, "")
            Exit For
        End If
    Next i
End If

Diferencia = Len(Formt1) - Len(Formt2)
Caracter = ""

For i = 1 To Abs(Diferencia)
    Caracter = Caracter & "0"
Next i

If Diferencia >= 0 Then
    Formt2 = Formt2 & Caracter
Else
    Formt1 = Formt1 & Caracter
End If

If Val(Formt1) > Val(Formt2) Then
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
Centrar Me

CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then Frame1.Caption = Trim(RsTemp.Fields("Nombre")) & " RIF " & Trim(RsTemp.Fields("Rif"))

CSql = "SELECT * FROM ContPDC WHERE IdEmpresa=" & IdEmprs & " AND Activo='1' ORDER BY Identificador"
Set RsTemp = CrearRS(CSql)

DTPicker1.Value = "01" & Format(Now, "/MM/yyyy")
DTPicker2.Value = Format(DateSerial(Year(CDate(Now)), Month(CDate(Now)) + 1, 0), "dd/MM/yyyy")

Combo1.Clear
Combo2.Clear

Combo1.AddItem " "
Combo2.AddItem " "

If RsTemp.RecordCount = 0 Then Exit Sub

Combo1.Clear
Combo2.Clear

While Not RsTemp.EOF
    Combo1.AddItem Trim(RsTemp.Fields("Identificador").Value)
    Combo1.ItemData(Combo1.NewIndex) = RsTemp.Fields("IdPDC").Value
    Combo2.AddItem Trim(RsTemp.Fields("Identificador").Value)
    Combo2.ItemData(Combo2.NewIndex) = RsTemp.Fields("IdPDC").Value
    RsTemp.MoveNext
Wend

End Sub
