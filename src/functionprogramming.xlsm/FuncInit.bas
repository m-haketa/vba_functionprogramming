Attribute VB_Name = "FuncInit"
Option Explicit

Public Function F(FunctionName As String, ParamArray Params()) As Func
  Dim arrParams
  arrParams = Params
  
  Dim ret As Func
  Set ret = New Func
    
  Call ret.Init(FunctionName, arrParams)
  
  Set F = ret
End Function

Public Sub setOrLet(ByRef Var, ByRef Value)
  If IsObject(Value) Then
    Set Var = Value
  Else
    Var = Value
  End If
End Sub


