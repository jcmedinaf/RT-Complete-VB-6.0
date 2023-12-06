VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReporteDiarioLegal 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diario Legal"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "FrmContReporteDiarioLegal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   4335
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   3240
         TabIndex        =   9
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
         MICON           =   "FrmContReporteDiarioLegal.frx":1002
         PICN            =   "FrmContReporteDiarioLegal.frx":101E
         PICH            =   "FrmContReporteDiarioLegal.frx":11E7
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
         TabIndex        =   10
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
         MICON           =   "FrmContReporteDiarioLegal.frx":141C
         PICN            =   "FrmContReporteDiarioLegal.frx":1438
         PICH            =   "FrmContReporteDiarioLegal.frx":16CA
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
         ItemData        =   "FrmContReporteDiarioLegal.frx":1857
         Left            =   1890
         List            =   "FrmContReporteDiarioLegal.frx":18BB
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Incluir cuentas con saldo CERO ""0"""
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1800
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   56164355
         CurrentDate     =   40254
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56164355
         CurrentDate     =   40254
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel de Jerarquía:"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   1380
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del reporte:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   930
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período del reporte:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   450
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FrmContReporteDiarioLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim RsTemp2 As Recordset
Dim i As Integer
Dim j As Integer
Dim Band As Boolean

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnResultados_Click()
Dim Formt As String
Dim Spdr As String
Dim TamDMGrid As Integer
Dim Band As Boolean
Dim NivelAnt As String

' Sentencia que devuelve una categoria de un P.D.C. por el debe y por el haber
CSql = "SELECT Identificador, Nombre, Tipo, SUM(Cantidad) AS Total From ContReporteDiarioLegal WHERE Fecha <='" & DTPicker2.Value & "' GROUP BY Identificador, Nombre, Tipo ORDER BY Identificador"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

With FrmContReportes

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Inicializa el DMGrid del formulario del reporte
.DMGrid1.Clear
.DMGrid1.Rows = 0

.DMGrid1.Cols = 4
.DMGrid1.Rows = 0

.DMGrid1.DColumnas(1).Caption = "Formato"
.DMGrid1.DColumnas(2).Caption = "Descripción del PDC"
.DMGrid1.DColumnas(3).Caption = "Débitos"
.DMGrid1.DColumnas(4).Caption = "Créditos"

.DMGrid1.DColumnas(3).Alignment = 1
.DMGrid1.DColumnas(4).Alignment = 1

.DMGrid1.DColumnas(3).IsNumber = True
.DMGrid1.DColumnas(4).IsNumber = True

.DMGrid1.DColumnas(1).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(2).Width = Val(.DMGrid1.Width * 50 / 100) - 350
.DMGrid1.DColumnas(3).Width = Val(.DMGrid1.Width * 20 / 100)
.DMGrid1.DColumnas(4).Width = Val(.DMGrid1.Width * 20 / 100)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

    While Not RsTemp.EOF
    
        Formt = Trim(RsTemp.Fields("Identificador").Value)  ' almacena en la variable "Formt" el formato del registro actual.
        
        ' Ciclo para mostrar los niveles anteriores del formato en la variable "Formt".
        For i = 1 To Len(Formt)
            Spdr = Mid(Formt, i, 1)
            
            If Not IsNumeric(Spdr) Or i = Len(Formt) Then
                ' sentencia que busca en el registro el nivel anterior al formato del registro actual.
                If i = Len(Formt) Then
                    CSql = "SELECT Identificador, Nombre FROM ContPDC WHERE Identificador='" & Mid(Formt, 1, i) & "'"
                Else
                    CSql = "SELECT Identificador, Nombre FROM ContPDC WHERE Identificador='" & Mid(Formt, 1, i - 1) & "'"
                End If
                Set RsTemp2 = CrearRS(CSql)
                
                ' condicional para salir del Ciclo FOR en el caso de que la sentencia no devuelva nada.
                  If RsTemp2.RecordCount = 0 Then Exit For
                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                
                TamDMGrid = .DMGrid1.Rows
                Band = True
                
                ' Ciclo para verificar si ya se ingreso el formato del nivel anterior, de ser asi, la variable
                ' BAND sera FALSE, indicando que ya se introdujo, de lo contrario se mantendra VERDADERA
                For j = 1 To TamDMGrid
                    If .DMGrid1.ValorCelda(j, 1) = Trim(RsTemp2.Fields("Identificador").Value) Then Band = False
                Next j
                
                ' Si la variable BAND es VERDADERA, entonces agrega el nivel anterior.
                If RsTemp.RecordCount <> 0 And Band Then
                    .DMGrid1.RowBackColor 1, RGB(255, 255, 255)
                    .DMGrid1.Rows = .DMGrid1.Rows + 1
                    .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = Trim(RsTemp2.Fields("Identificador").Value)
                    .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = DuplicaChr(" ", i) & RsTemp2.Fields("Nombre").Value
                End If
            End If
        Next i
        
        If Val(RsTemp.Fields("Tipo").Value) = 0 Then
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = RsTemp.Fields("Total").Value
        Else
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = RsTemp.Fields("Total").Value
        End If
        RsTemp.MoveNext
    Wend
    .DMGrid1.PaintMGrid
    Calcular_Categorias
    Tipo = "ReporteDiarioLegal"
    .Show vbModal, FrmPrincipal
End With

End Sub

Sub Calcular_Categorias()
Dim TamDMGrid As Integer
Dim NivelBase As String
Dim CantDebe  As Double
Dim CantHaber As Double

Dim Valor1 As Double
Dim Valor2 As Double

Dim AntValor1 As Double
Dim AntValor2 As Double

With FrmContReportes.DMGrid1

    TamDMGrid = .Rows
    
    For i = 1 To TamDMGrid

Volver:
        AntValor1 = Valor1
        AntValor2 = Valor2
        
        If Not IsNull(.ValorCelda(i, 3)) Then
            If Val(.ValorCelda(i, 3)) <> 0 Then Valor1 = CDbl(.ValorCelda(i, 3)) Else Valor1 = 0
        Else
            Valor1 = 0
        End If
        If Not IsNull(.ValorCelda(i, 4)) Then
            If Val(.ValorCelda(i, 4)) <> 0 Then Valor2 = CDbl(.ValorCelda(i, 4)) Else Valor2 = 0
        Else
            Valor2 = 0
        End If
        CantDebe = CantDebe + Valor1
        CantHaber = CantHaber + Valor2
         
        ' Condicional para saber si la fila actual es un grupo (es un grupo cuando el debe y el haber es CERO "0")
        If Valor1 = 0 And Valor2 = 0 Then
            ' Condicional para saber si anteriormente hubieron valores, si los hubieron, entonces
            ' colocar la sumatoria de los mismos...
            If AntValor1 <> 0 Or AntValor2 <> 0 Then
                .RowInsert (i)
                .ValorCelda(i, 2) = DuplicaChr(" ", InStr(1, NivelBase, Trim(NivelBase), vbTextCompare)) & "Total para " & Trim(NivelBase)
                .ValorCelda(i, 3) = CantDebe
                .ValorCelda(i, 4) = CantHaber
                .RowInsert (i + 1)
                i = i + 1
                i = i + 1
                CantDebe = 0
                CantHaber = 0
                TamDMGrid = .Rows
                If Band Then Band = False Else Band = True
            End If
            NivelBase = .ValorCelda(i, 2)   ' Almacena el nombre del nivel anterior inmediato...
        ElseIf i = TamDMGrid Then
            .Rows = .Rows + 1
            .ValorCelda(.Rows, 2) = DuplicaChr(" ", InStr(1, NivelBase, Trim(NivelBase), vbTextCompare)) & "Total para " & Trim(NivelBase)
            .ValorCelda(.Rows, 3) = CantDebe
            .ValorCelda(.Rows, 4) = CantHaber
            .Rows = .Rows + 1
            CantDebe = 0
            CantHaber = 0
        End If
        .PaintMGrid
        
        If i < TamDMGrid Then i = i + 1: GoTo Volver
    Next i
End With

End Sub

Function DuplicaChr(Caracter As String, i As Integer) As String
Dim Devolver As String
Dim ii As Integer
For ii = 1 To i
    Devolver = Devolver & Caracter
Next ii
DuplicaChr = Devolver
End Function

Private Sub Form_Load()
Centrar Me
CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)
If RsTemp.RecordCount <> 0 Then Frame1.Caption = Trim(RsTemp.Fields("Nombre")) & " RIF " & Trim(RsTemp.Fields("Rif"))
End Sub
