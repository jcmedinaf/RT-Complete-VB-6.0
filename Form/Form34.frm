VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form34 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asiento de  Pago"
   ClientHeight    =   8880
   ClientLeft      =   5715
   ClientTop       =   675
   ClientWidth     =   9525
   LinkTopic       =   "Form34"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9525
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   8160
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Retenciones hechas al cobro"
      Height          =   2055
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Agregar"
      Top             =   4440
      Width           =   9255
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   1680
         Picture         =   "Form34.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Eliminar"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   360
         Picture         =   "Form34.frx":03A8
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Agregar"
         Top             =   1440
         Width           =   1095
      End
      Begin Oncoamerica.DMGrid DMGrid1 
         Height          =   1095
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1931
         Object.Width           =   8610
         Object.Height          =   1065
         Rows            =   5
         AllowAddNew     =   -1  'True
         FindMode        =   3
         Editable        =   -1  'True
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7320
      Picture         =   "Form34.frx":07BB
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Salir"
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar Cobro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5520
      TabIndex        =   27
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos del Movimiento de Caja o Banco"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   6600
      Width           =   5055
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Text            =   "INGRESO"
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form34.frx":0D3D
         Left            =   1800
         List            =   "Form34.frx":0D3F
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto del Movimiento"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Movimiento"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Caja / Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos del Cobro"
      Height          =   3015
      Left            =   4320
      TabIndex        =   1
      Top             =   1320
      Width           =   5055
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   1920
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form34.frx":0D41
         Left            =   2040
         List            =   "Form34.frx":0D51
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   66912257
         CurrentDate     =   39940
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cheque / Deposito"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Operacion"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma del Pago"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto a Cancelar"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos de la Factura"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   25
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Restante"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Abonado"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Docuento"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Documento"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   120
      Picture         =   "Form34.frx":0D80
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bdata As New ADODB.Recordset
Dim a5
Dim camb1

Sub carga_retencion()
DMGrid1.Rows = 2
DMGrid1.Clear
DMGrid1.PaintMGrid

CSql = "select * from cobros where n_factura = " & N_FACTUR & " and tipo = 99"
bdata.Open CSql, CADENA

If Not (bdata.EOF) Then
bdata.MoveFirst
i = 1
Do While Not bdata.EOF
DMGrid1.Rows = DMGrid1.Rows + 1
DMGrid1.PaintMGrid
DMGrid1.ValorCelda(i, 1) = bdata.Fields("idcobro")
DMGrid1.ValorCelda(i, 2) = bdata.Fields("tipo_ret")
DMGrid1.ValorCelda(i, 3) = bdata.Fields("n_retencion")
DMGrid1.ValorCelda(i, 4) = bdata.Fields("Fecha_cob")
DMGrid1.ValorCelda(i, 5) = bdata.Fields("monto")
bdata.MoveNext

i = i + 1
Loop

End If

bdata.Close
DMGrid1.PaintMGrid


End Sub
Sub grid2()
    'carga las columnas y encabezados de columna
    DMGrid1.Cols = 5
    'DMGrid1.Rows = 1
    DMGrid1.DColumnas(5).Alignment = 1

    'DMGrid1.DColumnas(2).Locked = True
    DMGrid1.DColumnas(5).IsNumber = True

    DMGrid1.DColumnas(1).Width = 300
    DMGrid1.DColumnas(2).Width = 1600
    DMGrid1.DColumnas(3).Width = 1600
    DMGrid1.DColumnas(4).Width = 1000
    DMGrid1.DColumnas(5).Width = 1600
    DMGrid1.DColumnas(1).Caption = "ID"
    DMGrid1.DColumnas(2).Caption = "Tipo retencion"
    DMGrid1.DColumnas(3).Caption = "Nº Retencion"
    DMGrid1.DColumnas(4).Caption = "Fecha"
    DMGrid1.DColumnas(5).Caption = "Monto"
    

End Sub
Private Sub Combo1_Click()
Text7.Text = Text6.Text
End Sub

Sub Calcular()
CSql = "SELECT SUM(monto) as monto2 FROM cobros WHERE N_factura = " & N_FACTUR

bdata.Open CSql, CADENA
If IsNull(bdata.Fields("monto2")) Then Text8.Text = "0.00": a5 = 0 Else Text8.Text = Format(bdata.Fields("monto2"), "##,##0.00"): a5 = bdata.Fields("monto2")
bdata.Close
Call QuitarCaracter(Text3.Text)
a3 = CArac
Text9.Text = Format(a3 - a5, "##,##0.00")
Text6.Text = Text9.Text
Text7.Text = Text9.Text
End Sub
Private Sub Command1_Click()
Call Text6_LostFocus

If Combo1.ListIndex = -1 Then
    Msg = "No ha seleccionado la forma de pago"
    MsgBox Msg
    Exit Sub
End If

If Combo2.ListIndex = -1 Then
    Msg = "No ha seleccionado el Banco o Caja donde va a asentar el ingreso"
    MsgBox Msg
    Exit Sub
End If

If Text6.Text = "" Or CDec(Text6.Text) = 0 Then Exit Sub
Call QuitarCaracter(Text6.Text)
T6 = CArac
Call Quitar(T6)
T6 = CArac

CSql = "insert into movi_bancaja(idcajabanco, ingr_egr, monto_mov, tipo_mov, fecha_transa) values(" & Combo2.ItemData(Combo2.ListIndex) & ", 1," & T6 & "," & Combo1.ListIndex & ",#" & CDate(DTPicker1.Value) & "#)"
bdata.Open CSql, CADENA

CSql = "SELECT MAX(idmovcajabanco) as IDMOV FROM movi_bancaja"
bdata.Open CSql, CADENA
idmov = bdata.Fields(0)
bdata.Close

CSql = "insert into cobros(n_factura, fecha_cob, monto, form_pag, n_comprobante, idmovcajabanco, tipo) values(" & N_FACTUR & ",#" & CDate(DTPicker1.Value) & "#," & T6 & "," & Combo1.ListIndex & ",'" & Text4.Text & "'," & idmov & ", 1)"
bdata.Open CSql, CADENA
Msg = "Datos registrados satisfactoriamente"
MsgBox Msg
Unload Me


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
With Form36
.Text1.Text = Text1.Text
.Show 1
End With

Call Calcular
Call carga_retencion
camb1 = 1

End Sub


Private Sub Command4_Click()
d = DMGrid1.Row
CSql = "delete from cobros where idcobro = " & DMGrid1.ValorCelda(d, 1)
bdata.Open CSql, CADENA
DMGrid1.RowDelete (d)
DMGrid1.PaintMGrid
Msg = "Retención Eliminada Satisfactoriamente"
MsgBox Msg, vbOKOnly, "Eliminado"

Call Calcular
End Sub

Private Sub DMGrid1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()

camb1 = 0
Call Calcular
CSql = "Select * from cajas_bancos"
bdata.Open CSql, CADENA
If bdata.EOF Then bdata.Close: Exit Sub

bdata.MoveFirst
Do While Not bdata.EOF
Combo2.AddItem bdata.Fields(1)
Combo2.ItemData(Combo2.NewIndex) = bdata.Fields(0)
bdata.MoveNext
Loop
bdata.Close
Call grid2
Call carga_retencion
End Sub

Private Sub Text3_Change()
Call Calcular
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then KeyAscii = 0

End Sub

Private Sub Text6_Change()
Text7.Text = Text6.Text

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 48 To 57
Exit Sub
Case Is = 13
Exit Sub
Case 45 To 46
Exit Sub
Case Is = 8
Exit Sub
Case Else
KeyAscii = 0
End Select

End Sub

Private Sub Text6_LostFocus()
Text6.Text = Format(Text6.Text, "##,##0.00")
Text7.Text = Text6.Text
End Sub


Private Sub Timer1_Timer()
If camb1 = 1 Then
Call Calcular
Call carga_retencion
camb1 = 0
End If
End Sub


