Attribute VB_Name = "FuncInit"
Option Explicit

Public Function F(FunctionName As String, ParamArray Params()) As Func
  Dim ret As Func
  Set ret = New Func
  
  If (LBound(Params) > UBound(Params)) Then
    Call ret.Init(FunctionName, Array())
  Else
  
    Dim newParams As Variant
    ReDim newParams(LBound(Params) To UBound(Params))
    
    Dim C As Long
    For C = LBound(Params) To UBound(Params)
      Call setOrLet(newParams(C), Params(C))
    Next
  
    Call ret.Init(FunctionName, newParams)
  End If
  
  Set F = ret
End Function

Public Sub setOrLet(ByRef Var, ByRef Value)
  If IsObject(Value) Then
    Set Var = Value
  Else
    Var = Value
  End If
End Sub


