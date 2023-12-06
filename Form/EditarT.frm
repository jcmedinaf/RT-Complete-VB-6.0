VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form21 
   Caption         =   "Form21"
   ClientHeight    =   5775
   ClientLeft      =   6450
   ClientTop       =   2700
   ClientWidth     =   13440
   LinkTopic       =   "Form21"
   ScaleHeight     =   5775
   ScaleWidth      =   13440
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   13335
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "EditarT.frx":0000
         Height          =   855
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1508
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Campo"
            Caption         =   "Nº Campo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Intrucciones"
            Caption         =   "Intrucciones para Cuadrar Campos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   6660.284
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1440
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "EditarT.frx":0015
         Height          =   975
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   1720
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "Campo"
            Caption         =   "Nº Campo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "SAD"
            Caption         =   "SADo SSD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Direccion"
            Caption         =   "Direccion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Upper"
            Caption         =   "Upper(mm)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Lower"
            Caption         =   "Lower(mm)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Gantry"
            Caption         =   "Gantry"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Colimador"
            Caption         =   "Colimador"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Camilla"
            Caption         =   "Camilla"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Bandeja"
            Caption         =   "Bandeja"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Bloque"
            Caption         =   "Bloque"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Compensa"
            Caption         =   "Compensador"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "Cuña"
            Caption         =   "Cuña"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "Bolus"
            Caption         =   "Bolus"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "Inicial"
            Caption         =   "Inicial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column13 
               Alignment       =   2
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column14 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar"
         Height          =   375
         Index           =   2
         Left            =   12000
         TabIndex        =   3
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Editar"
         Height          =   375
         Index           =   1
         Left            =   10920
         TabIndex        =   2
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nuevo"
         Height          =   375
         Index           =   0
         Left            =   9840
         TabIndex        =   1
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   0
      Picture         =   "EditarT.frx":002A
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD58 As New ADODB.Connection
Dim cnN As New ADODB.Recordset


Sub CargarDataGrid(dg As DataGrid)
    
    dg.MarqueeStyle = dbgHighlightRow
    Set dg.DataSource = cnN
    dg.Refresh
End Sub

Public Sub IniciarConexion()
If BD58.State = adStateOpen Then BD58.Close
     BD58.CursorLocation = adUseClient
        BD58.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              "z:\oa.mdb" & ";Persist Security Info=False"
End Sub
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0: Call Agregar
        Case 1: Call Editar
        Case 2: Call Eliminar
        Case 3: Unload Me
    End Select
End Sub


Private Sub Form_Load()
   
    SQL = ""
    IdPac1 = ""
      Call IniciarConexion
If cnN.State = adStateOpen Then cnN.Close
   cnN.Open "select * from Tecnica2", BD58, adOpenStatic, adLockOptimistic
    
  With DataGrid1
   .AllowUpdate = False
 End With
    
    Call CargarDataGrid(DataGrid1)
    Call CargarDataGrid(DataGrid2)
  
        
End Sub

Private Sub Eliminar()

    If DataGrid1.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
     
    With DataGrid1
        If MsgBox("Se va a eliminar el registro : está seguro ", _
                    vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            
                        BD58.Delete
                        
            ' Actualiza el recordset
            BD58.Update
            DataGrid1.Refresh
        End If
    End With
End Sub

' agrega uno nuevo
'''''''''''''''''''''''
Sub Agregar()
    
    With Form22
        Form22.ACCION = AGREGAR_REGISTRO1
        Form22.Label3 = Format(Date, "mm/dd/yyyy")
        Form22.Show vbModal
        DataGrid1.Refresh
        DataGrid2.Refresh
    End With
    
End Sub


'Abre el formulario para Editar el registro seleccionado
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editar()

    Dim i As Integer
    
    If DataGrid1.Row = -1 Then Exit Sub
    If DataGrid2.Row = -1 Then Exit Sub
    
    With Form22
        ' obtiene el elemento seleccionado, el id
        .Label2 = cnN("Idpaciente")
        
        ' llena los campos
        For i = 3 To 14
            .Text1(i).Text = cnN(i)
        Next
    If cnN(15) = True Then
    .Check1.Value = 1
    Else
    .Check1.Value = 0
    End If
    
        If cnN(16) = True Then
         .Check2.Value = 1
        Else
            .Check2.Value = 0
            End If
        If cnN(17) = True Then
            .Check3.Value = 1
                Else
            .Check3.Value = 0
        End If
        If cnN(18) = True Then
        .Check4.Value = 1
        Else
        .Check4.Value = 0
        End If
        .Label3 = cnN(19)
        .ACCION = EDITAR_REGISTRO
        
        .Show vbModal
        DataGrid1.Refresh
        
    End With

End Sub








