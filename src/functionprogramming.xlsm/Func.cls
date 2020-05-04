VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public m_FuncName As String
Public m_Params As Variant

Private Sub Class_Initialize()
  m_Params = Array()
End Sub


Public Sub Init(FunctionName As String, Params As Variant)
  m_FuncName = FunctionName
  
  If (LBound(Params) > UBound(Params)) Then
    Exit Sub
  End If
  
  ReDim m_Params(LBound(Params) To UBound(Params))
  
  Dim C As Long
  For C = LBound(Params) To UBound(Params)
    Call setOrLet(m_Params(C), Params(C))
  Next
End Sub

Public Function setParam(ParamArray Params()) As Func
Attribute setParam.VB_UserMemId = 0
'=Default Property
  Dim newFunc As Func
  Set newFunc = New Func
  
  Dim prevUbound As Long
  prevUbound = UBound(m_Params)
  
  Dim addParamCount As Long
  addParamCount = UBound(Params) - LBound(Params) + 1
    
  Dim newParams As Variant
  ReDim newParams(LBound(m_Params) To prevUbound + addParamCount)
  
  Dim C As Long
  For C = LBound(m_Params) To prevUbound
    Call setOrLet(newParams(C), m_Params(C))
  Next
  
  For C = prevUbound + 1 To UBound(newParams)
    Call setOrLet(newParams(C), Params(C - prevUbound + LBound(Params) - 1))
  Next
  
  newFunc.Init m_FuncName, newParams
  
  Set setParam = newFunc
End Function


Public Function Run() As Variant
  Dim LB As Long
  Dim UB As Long
  
  LB = LBound(m_Params)
  UB = UBound(m_Params)
  
  Select Case UB - LB
    Case -1
      Call setOrLet(Run, Application.Run(m_FuncName))
    Case 0
      Call setOrLet(Run, Application.Run(m_FuncName, m_Params(LB)))
    Case 1
      Call setOrLet(Run, Application.Run(m_FuncName, m_Params(LB), m_Params(LB + 1)))
    Case 2
      Call setOrLet(Run, Application.Run(m_FuncName, m_Params(LB), m_Params(LB + 1), m_Params(LB + 2)))
    Case 3
      Call setOrLet(Run, Application.Run(m_FuncName, m_Params(LB), m_Params(LB + 1), m_Params(LB + 2), m_Params(LB + 3)))
  End Select
  
End Function


