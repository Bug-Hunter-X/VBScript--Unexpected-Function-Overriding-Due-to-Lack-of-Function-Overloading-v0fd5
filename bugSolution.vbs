Function overloadedFunction(param1, optional param2)
  If IsMissing(param2) Then
    MsgBox "One parameter version: " & param1
  Else
    MsgBox "Two parameter version: " & param1 & ", " & param2
  End If
End Function

'OR

Function overloadedFunctionOneParam(param1)
  MsgBox "One parameter version: " & param1
End Function

Function overloadedFunctionTwoParams(param1, param2)
  MsgBox "Two parameter version: " & param1 & ", " & param2
End Function

'Calling the functions
overloadedFunction "single parameter"
overloadedFunction "param1", "param2"

overloadedFunctionOneParam "single parameter"
overloadedFunctionTwoParams "param1", "param2"