VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DColumnas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolDColumnas As Collection
'El objeto de colección se crea en el método Initialize de la clase:
Private Sub Class_Initialize()
  Set mcolDColumnas = New Collection
End Sub
'Los métodos del módulo de clase delegan en los métodos predeterminados del objeto Collection de Visual Basic.
Public Property Get Count() As Integer
  Count = mcolDColumnas.Count
End Property
' Método Add para crear nuevos objetos.
Public Function Add(mCaption As String) As DColumna
Dim mDColumna As New DColumna
mDColumna.Caption = mCaption
mcolDColumnas.Add mDColumna
Set Add = mDColumna
End Function

Public Sub Delete(ByVal Index As Variant)
    mcolDColumnas.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As DColumna
Attribute Item.VB_UserMemId = 0
    Set Item = mcolDColumnas.Item(Index)
End Function

'Atributos del procedimiento marcarlo como oculto y asignarle un Id. de procedimiento de –4, para poder utilizar For Each … Next.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolDColumnas.[_NewEnum]
End Function

