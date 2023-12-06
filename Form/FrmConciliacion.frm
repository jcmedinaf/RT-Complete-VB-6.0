VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmConciliacion 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliación"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   Icon            =   "FrmConciliacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   4575
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   3480
         TabIndex        =   20
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
         MICON           =   "FrmConciliacion.frx":1002
         PICN            =   "FrmConciliacion.frx":101E
         PICH            =   "FrmConciliacion.frx":11E7
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
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Guardar / Actualizar"
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
         MICON           =   "FrmConciliacion.frx":141C
         PICN            =   "FrmConciliacion.frx":1438
         PICH            =   "FrmConciliacion.frx":16C7
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00EAEFEF&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox TxtFechaUltimaConciliacion 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxtDiferencia 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox TxtSaldoEstadoCuenta 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox TxtDepositosCreditosTransitos 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0,00"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox TxtChequesDebitosTransitos 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0,00"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox TxtDepositosCreditos 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox TxtChequesDebitos 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox TxtSaldoAnteriorConciliado 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   720
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPickerFecha 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52625409
         CurrentDate     =   40236
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ultima Conciliación:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   330
         Width           =   1875
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   4530
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Estado Cuenta:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   4050
         Width           =   1545
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depósitos o Créditos Tránsitos:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   3450
         Width           =   2190
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques o Débitos tránsitos:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2970
         Width           =   2025
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depósitos o Créditos:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2370
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques o Débitos:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1890
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Anterior Conciliado:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   810
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Conciliación:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1290
         Width           =   1395
      End
   End
End
Attribute VB_Name = "FrmConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsEstadoCuenta As New ADODB.Recordset
Dim RsConciliar As New ADODB.Recordset
Dim RsActualiza As New ADODB.Recordset

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
Dim X, j As Integer

'CSql = "SELECT EstadoDeCuenta.* FROM EstadoDeCuenta INNER JOIN Movi_BanCaja ON EstadoDeCuenta.IdCajaBanco = Movi_BanCaja.IdCajaBanco AND " & _
       "EstadoDeCuenta.Monto_Mov = Movi_BanCaja.Monto_Mov AND EstadoDeCuenta.n_comprobante = Movi_BanCaja.n_comprobante AND " & _
       "EstadoDeCuenta.Fecha_Transa = Movi_BanCaja.Fecha_Transa " & _
       "WHERE EstadoDeCuenta.Conciliado = 0 And EstadoDeCuenta.IdCajaBanco='" & FrmConcilicacionBancaria.TxtCodigo.Text & "'"

CSql = "SELECT * From Movi_BanCaja WHERE Conciliado = 0 And IdCajaBanco='" & FrmConcilicacionBancaria.TxtCodigo.Text & "'"


Set RsConciliar = CrearRS(CSql)

If RsConciliar.RecordCount > 0 Then
    X = RsConciliar.RecordCount
    
    
    ReDim ArreConciliado(Val(X), 3) As String
    
    For X = 1 To X
    
            ArreConciliado(X, 1) = RsConciliar.Fields("IdCajaBanco").Value
            ArreConciliado(X, 2) = RsConciliar.Fields("N_Comprobante").Value
            ArreConciliado(X, 3) = RsConciliar.Fields("Conciliado").Value
    Next X
    
    
    For X = 1 To X
    
'        CSql = "Update Movi_BanCaja Set Conciliado=1, FechaConciliacion='" & DateTime.Date & "' Where Idcajabanco='" & ArreConciliado(X, 1) & "' And N_Comprobante='" & ArreConciliado(X, 2) & "'"
'        Set RsActualiza = CrearRS(CSql)
    
        CSql = "Update Movi_BanCaja Set Conciliado=1, FechaConciliacion='" & DateTime.Date & "' Where Idcajabanco='" & ArreConciliado(X, 1) & "' And N_Comprobante='" & ArreConciliado(X, 2) & "'"
        Set RsActualiza = CrearRS(CSql)
    Next X
    
    MsgBox "Conciliación realizada con Exito!", vbInformation + vbOKOnly, "Conciliación"
    FrmConcilicacionBancaria.Movi
    FrmConcilicacionBancaria.Conciliacion
    FrmConcilicacionBancaria.TotalConciliado
    FrmConcilicacionBancaria.TotalNoConciliado
    FrmConcilicacionBancaria.ChequesTransitos
    
 
    Unload Me
Else
    MsgBox "No Hay Movimientos Pendientes por realizarle la Conciliación!", vbInformation + vbOKOnly, "Error Conciliación"
    Unload Me
End If
End Sub

Private Sub Form_Load()
'**********************
' Busca la fecha ultima conciliacion
CSql = "Select distinct (FechaConciliacion) as FechaUltimaConciliacion From Movi_BanCaja Where Conciliado=1 And IdCajaBanco='" & FrmConcilicacionBancaria.TxtCodigo.Text & "' order by FechaConciliacion desc"
Set RsEstadoCuenta = CrearRS(CSql)
If RsEstadoCuenta.RecordCount > 0 Then
    RsEstadoCuenta.MoveFirst
    TxtFechaUltimaConciliacion.Text = Format(RsEstadoCuenta.Fields("FechaUltimaConciliacion").Value, "dd/mm/yyyy")
Else
    TxtFechaUltimaConciliacion.Text = ""
End If
'**********************
' Suma el total de todos los montos conciliados
CSql = "Select Sum (Monto_Mov) as Monto_MovConciliacion From Movi_BanCaja Where Conciliado=1 And IdCajaBanco='" & FrmConcilicacionBancaria.TxtCodigo.Text & "'"
Set RsEstadoCuenta = CrearRS(CSql)

If RsEstadoCuenta.RecordCount > 0 Then
    If Not IsNull(RsEstadoCuenta.Fields("Monto_MovConciliacion").Value) Then
        TxtSaldoAnteriorConciliado.Text = Format(RsEstadoCuenta.Fields("Monto_MovConciliacion").Value, "#,##0.00")
    Else
        TxtSaldoAnteriorConciliado.Text = Format(0, "#,##0.00")
    End If
End If
'**********************
'Suma el total de todos los cheques y debitos
CSql = "Select Sum (Monto_Mov) as TotalChequesDebitos From Movi_BanCaja Where Tipo_Mov=2 or Tipo_Mov=6 And IdCajaBanco='" & FrmConcilicacionBancaria.TxtCodigo.Text & "'"
Set RsEstadoCuenta = CrearRS(CSql)

If RsEstadoCuenta.RecordCount > 0 Then
    If Not IsNull(RsEstadoCuenta.Fields("TotalChequesDebitos").Value) Then
        TxtChequesDebitos.Text = Format(RsEstadoCuenta.Fields("TotalChequesDebitos").Value, "#,##0.00")
    Else
        TxtChequesDebitos.Text = Format(0, "#,##0.00")
    End If
End If
'**********************
'Suma el total de todos los depositos y creditos
CSql = "Select Sum (Monto_Mov) as TotalDepositosCreditos From Movi_BanCaja Where Tipo_Mov=3 or Tipo_Mov=5 And IdCajaBanco='" & FrmConcilicacionBancaria.TxtCodigo.Text & "'"
Set RsEstadoCuenta = CrearRS(CSql)

If RsEstadoCuenta.RecordCount > 0 Then
    If Not IsNull(RsEstadoCuenta.Fields("TotalDepositosCreditos").Value) Then
        TxtDepositosCreditos.Text = Format(RsEstadoCuenta.Fields("TotalDepositosCreditos").Value, "#,##0.00")
    Else
        TxtDepositosCreditos.Text = Format(0, "#,##0.00")
    End If
End If
'**********************
'resta el total de los cheques y debitos menos el total de los depositos y creditos
TxtDiferencia.Text = CDbl(TxtChequesDebitos.Text) - CDbl(TxtDepositosCreditos.Text)
End Sub
