VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Facilities make strings"
Option Explicit

Private Str As String
Private newLine As Boolean

Private Sub Class_Initialize()
    Str = Empty
    newLine = False
End Sub

Public Function Add(ByVal sText As String)
    Str = Str & sText
    If newLine Then
        Str = Str & vbNewLine
    End If
End Function

Public Function Clear()
    Str = Empty
End Function

Public Property Get Text() As String
Attribute Text.VB_Description = "Return your String"
Attribute Text.VB_UserMemId = 0
    Text = Str
End Property

Public Property Get autoNewLine() As Boolean
    autoNewLine = newLine
End Property

Public Property Let autoNewLine(ByVal vValue As Boolean)
    newLine = vValue
End Property


