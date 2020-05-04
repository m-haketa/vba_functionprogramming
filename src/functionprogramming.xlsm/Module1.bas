Attribute VB_Name = "Module1"
Option Explicit

Sub test()
  Dim addFunc
  Set addFunc = F("add")
  
  Dim add2
  Set add2 = addFunc(2)
  
  Debug.Print add2(3).Run

End Sub


Sub test2()
  Dim largerThan10
  Set largerThan10 = F("largerThan")(10)
  
  Dim arr
  arr = Array(1, 15, 8, 13)
  
  Dim filterLargerThan10
  Set filterLargerThan10 = F("filter")(largerThan10)
  
  Dim arr2
  arr2 = filterLargerThan10(arr).Run
End Sub

Sub test3()
  Dim add2
  Set add2 = F("add", 2)
  
  Debug.Print add2(3).Run

End Sub

Sub maptest()
  Dim addFunc
  Set addFunc = F("add")
  
  Dim add2
  Set add2 = addFunc(2)
  
  Dim arr
  arr = Array(1, 15, 8, 13)
  
  Dim arr2
  arr2 = F("map")(add2, arr).Run

End Sub

Sub composetest()
  Dim add10
  Set add10 = F("add")(10)

  Dim multi3
  Set multi3 = F("multi")(3)
  
  Dim add10multi2Func
  Set add10multi2Func = F("compose")(multi3, add10)
  
  Dim ret
  ret = add10multi2Func(2).Run
  
  Debug.Print ret
End Sub

Sub composetest2()
  Dim add2
  Set add2 = F("add")(2)

  Dim multi3
  Set multi3 = F("multi")(3)
  
  Dim add2multi3Func
  Set add2multi3Func = F("compose")(multi3, add2)
  
  Dim ret
  ret = add2multi3Func(4).Run
  
  Debug.Print ret
End Sub

Sub sortTest()
  
  Dim arr
  arr = Array(10, 1, 6, 8, 4)
  
  Dim ret
  ret = F("sort")(F("numComp"), arr).Run

End Sub


Sub sortTest2()
  
  Dim arr
  arr = Array(Array(10, "a"), Array(1, "b"), Array(6, "c"), Array(8, "d"), Array(4, "e"))
  
  Dim ret
  ret = F("sort")(F("firstarrComp"), arr).Run

End Sub


Function multi(a, b)
  multi = a * b
End Function

Function add(a, b)
  add = a + b
End Function

Function largerThan(condNum, data) As Boolean
  largerThan = condNum < data
End Function

Function numComp(a, b) As Long
  If a > b Then numComp = 1
  If a = b Then numComp = 0
  If a < b Then numComp = -1
End Function

Function firstarrComp(a, b) As Long
  If a(0) > b(0) Then firstarrComp = 1
  If a(0) = b(0) Then firstarrComp = 0
  If a(0) < b(0) Then firstarrComp = -1
End Function


'小さい順に並び替え
Function sort(Comp As Func, iArr As Variant) As Variant
  Dim idxArr
  ReDim idxArr(LBound(iArr) To UBound(iArr))
  
  Dim C As Long
  For C = LBound(iArr) To UBound(iArr)
    idxArr(C) = C
  Next
  
  Dim C1 As Long
  Dim C2 As Long
  
  For C1 = LBound(iArr) To UBound(iArr) - 1
    For C2 = C1 + 1 To UBound(iArr)
      Dim sortCond
      sortCond = Comp(CVar(iArr(idxArr(C1))), CVar(iArr(idxArr(C2)))).Run
      
      'C1のほうが大きい場合、入れ替え
      If sortCond > 0 Then
        Dim Temp
        Temp = idxArr(C1)
        idxArr(C1) = idxArr(C2)
        idxArr(C2) = Temp
      End If
      
    Next
  Next

  Dim oArr As Variant
  ReDim oArr(LBound(iArr) To UBound(iArr))
  
  For C = LBound(iArr) To UBound(iArr)
    oArr(C) = iArr(idxArr(C))
  Next

  sort = oArr
End Function


Function filter(Cond As Func, iArr As Variant) As Variant
  Dim oC As Long
  oC = LBound(iArr) - 1
  
  Dim oArr As Variant
  ReDim oArr(LBound(iArr) To UBound(iArr))
  
  Dim iC As Long
  For iC = LBound(iArr) To UBound(iArr)
    If Cond(iArr(CVar(iC))).Run Then
      oC = oC + 1
      Call setOrLet(oArr(oC), iArr(CVar(iC)))
    End If
  Next
  
  If oC >= LBound(iArr) Then
    ReDim Preserve oArr(LBound(iArr) To oC)
  Else
    oArr = Array()
  End If
  
  filter = oArr
End Function

Function map(mapFunc As Func, iArr As Variant) As Variant
  Dim oArr As Variant
  ReDim oArr(LBound(iArr) To UBound(iArr))
  
  Dim C As Long
  For C = LBound(iArr) To UBound(iArr)
    Call setOrLet(oArr(C), mapFunc(CVar(iArr(C))).Run)
  Next
  
  map = oArr
End Function

Function compose(ParamArray FuncsAndParams() As Variant) As Variant
  Dim Params As Variant
  
  Dim minC As Long
  minC = LBound(FuncsAndParams)
   
  Dim funcMinC As Long
  funcMinC = minC
  
  Dim paramMinC As Long
  paramMinC = minC

  Dim maxC As Long
  maxC = UBound(FuncsAndParams)


  '最初のいくつかの引数は関数。
  '関数ではない引数が出てくるまでparamMinCを加算
  Do While TypeName(FuncsAndParams(paramMinC)) = "Func"
    paramMinC = paramMinC + 1
  Loop
  
  '関数が1つもない場合はエラー
  If funcMinC = paramMinC Then
    Err.Raise Number:=65000, Description:="関数が入力されていません"
  End If

  '最後の関数の添え字を確定
  Dim funcMaxC As Long
  funcMaxC = paramMinC - 1

  '最後の関数を格納
  Dim ret As Variant
  Set ret = FuncsAndParams(funcMaxC)

  Dim C As Long
  C = paramMinC
  '残りはすべて最後の関数に対するパラメータなので、最後の関数に適用していく
  For C = paramMinC To maxC
    If TypeName(FuncsAndParams(C)) = "func" Then
      Err.Raise Number:=65000, Description:="最後の関数に対する引数を入力してください"
    Else
      Set ret = ret(FuncsAndParams(C))
    End If
  Next

  '最後の関数を起動して、結果をRetに格納
  Call setOrLet(ret, ret.Run)
  
  '以下、順番に前の関数を起動
  For C = funcMaxC - 1 To funcMinC Step -1
    Call setOrLet(ret, FuncsAndParams(C)(ret).Run)
  Next
  
  Call setOrLet(compose, ret)
End Function

