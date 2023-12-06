Attribute VB_Name = "Funciones"
Option Explicit
Dim ICont As Integer
Dim RsTemp As Recordset

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMM   FUNCIONES DE NOMINA MMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

' Calcula el Numero de Lunes que hay dentro de dos fechas
Public Function LunesDelPeriodo(ByRef FechaInicio As String, ByRef FechaFin As String) As Integer
Dim FechaI As Date
Dim FechaF As Date
Dim BuffTemp As Integer

FechaI = CDate(FechaInicio)
FechaF = CDate(FechaFin)

While Not DateDiff("d", FechaI, FechaF) = -1
    If Weekday(FechaI) = 2 Then
        BuffTemp = BuffTemp + 1
    End If
    FechaI = FechaI + 1
Wend
LunesDelPeriodo = BuffTemp
End Function

' Calcula el Numero de Lunes que hay dentro del MES de la fecha contenida en la vareiable "Fecha"
Public Function LunesDelMes(ByRef Fecha As String) As Integer
Dim FechaI As Date
Dim FechaF As Date
Dim BuffTemp As Integer

FechaI = "01/" & Format(CDate(Fecha), "mm/yyyy")
FechaF = Format(DateSerial(Year(CDate(FechaI)), Month(CDate(FechaI)) + 1, 0), "dd/MM/yyyy")

While Not DateDiff("d", FechaI, FechaF) = -1
    If Weekday(FechaI) = 2 Then
        BuffTemp = BuffTemp + 1
    End If
    FechaI = FechaI + 1
Wend
LunesDelMes = BuffTemp
End Function
