Option Explicit
' Mathematical expressions rules

Public Function Parse(ByVal sExpr As String, Optional ByVal iRecur As Integer = 0)
    Dim cTree As Collection
    
    If iRecur = 0 Then
        ' Used as breakpoint
        iRecur = 0
    End If
    
    Set cTree = ParseBinary(sExpr, ",")
    If cTree Is Nothing Then
        Set cTree = ParseBinary(sExpr, "+")
        If cTree Is Nothing Then
            Set cTree = ParseBinary(sExpr, "-")
            If cTree Is Nothing Then
                Set cTree = ParseBinary(sExpr, "*")
                If cTree Is Nothing Then
                    Set cTree = ParseBinary(sExpr, "/")
                    If cTree Is Nothing Then
                        Set cTree = ParseBinary(sExpr, "\")
                        If cTree Is Nothing Then
                            Set cTree = ParseUnary(sExpr, "+")
                            If cTree Is Nothing Then
                                Set cTree = ParseUnary(sExpr, "-")
                                If cTree Is Nothing Then
                                    Set cTree = ParseBinary(sExpr, "^")
                                    If cTree Is Nothing Then
                                        Set cTree = ParseBrackets(sExpr)
                                        If cTree Is Nothing Then
                                            Set cTree = ParseFunction(sExpr)
                                            If cTree Is Nothing Then
                                                Set cTree = ParseToken(sExpr)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    Set Parse = cTree
End Function


Public Function ParseBinary(ByVal sExpr As String, ByVal sOp As String) As Collection
    Dim i As Integer, j As Integer, p As Integer
    Dim c As String
    Dim sExpr1 As String, sExpr2 As String
    Dim cExpr1 As Collection, cExpr2 As Collection, cTree As Collection
    
    p = 0
    For i = 1 To Len(sExpr)
        c = Mid(sExpr, i, 1)
        If c = "(" Then p = p + 1
        If c = ")" Then p = p - 1
        
        If p < 0 Then Exit For
        If p = 0 And c = sOp Then
            sExpr1 = Left(sExpr, i - 1)
            sExpr2 = Right(sExpr, Len(sExpr) - i)
            Set cExpr1 = Parse(sExpr1, 1)
            Set cExpr2 = Parse(sExpr2, 1)
            Exit For
        End If
    Next
    
    If cExpr1 Is Nothing Or cExpr2 Is Nothing Then
        Exit Function
    End If
    
    Set cTree = New Collection
    cTree.Add sOp
    cTree.Add cExpr1
    If cExpr2(1) = sOp Then
        For i = 2 To cExpr2.Count
            cTree.Add cExpr2(i)
        Next
    Else
        cTree.Add cExpr2
    End If
    
    Set ParseBinary = cTree
End Function


Public Function ParseUnary(ByVal sExpr As String, ByVal sOp As String) As Collection
    Dim i As Integer, j As Integer
    Dim c As String
    Dim cExpr As Collection, cTree As Collection
    
    sExpr = Trim(sExpr)
    c = Left(sExpr, 1)
    If c <> sOp Then
        Exit Function
    End If
    
    sExpr = Right(sExpr, Len(sExpr) - 1)
    Set cExpr = Parse(sExpr, 1)
    If cExpr Is Nothing Then
        Exit Function
    End If
    
    Select Case sOp
    Case "+"
        Set cTree = cExpr
    Case "-"
        Set cTree = New Collection
        cTree.Add "(-)"
        cTree.Add cExpr
    End Select
    
    Set ParseUnary = cTree
End Function


Public Function ParseBrackets(ByVal sExpr As String) As Collection
    sExpr = Trim(sExpr)
    If Left(sExpr, 1) = "(" And Right(sExpr, 1) = ")" Then
        sExpr = Mid(sExpr, 2, Len(sExpr) - 2)
        Set ParseBrackets = Parse(sExpr, 1)
    End If
End Function


Public Function ParseToken(ByVal sExpr As String) As Collection
    Dim cTree As Collection
    Dim i As Integer
    Dim c As String
    Dim sExprOld As String
    
    If False Then
    sExprOld = sExpr
    Do
        sExpr = Replace(sExpr, "- ", " -")
        If sExprOld = sExpr Then Exit Do
        sExprOld = sExpr
    Loop
    End If
    
    sExpr = Trim(sExpr)
    If Len(sExpr) = 0 Then Exit Function
    For i = 1 To Len(sExpr)
        c = UCase(Mid(sExpr, i, 1))
        If ("A" <= c And c <= "Z") Or ("0" <= c And c <= "9") Or InStr(1, ".%_", c) Then
        Else
            Exit Function
        End If
    Next
    
    Set cTree = New Collection
    cTree.Add "TOKEN"
    cTree.Add sExpr
    
    Set ParseToken = cTree
End Function


Public Function ParseFunction(ByVal sExpr As String) As Collection
    Dim iStart As Integer
    Dim sFunc As String
    Dim sArgs As String
    Dim cArgs As Collection
    Dim cExpr1 As Collection, cExpr2 As Collection, cTree As Collection
    
    sExpr = Trim(sExpr)
    iStart = InStr(1, sExpr, "(")
    If iStart = 0 Or Right(sExpr, 1) <> ")" Then
        Exit Function
    End If
    sFunc = Trim(Left(sExpr, iStart - 1))
    sArgs = Mid(sExpr, iStart + 1, Len(sExpr) - iStart - 1)
    
    Set cArgs = Parse(sArgs)
    If cArgs Is Nothing Then
        Exit Function
    End If
        
    Set cTree = New Collection
    cTree.Add "FUNC"
    cTree.Add sFunc
    cTree.Add cArgs
    
    Set ParseFunction = cTree
End Function


Public Function TreeEval(ByVal cTree As Collection) As Double
    Dim sOp As String, i As Integer
    Dim dParam As Double
    Dim dResult As Double, sErrorMsg As String
    
    sErrorMsg = "Syntax error"
    On Error GoTo EvalError
    
    sOp = cTree(1)
    
    Select Case sOp
    Case "TOKEN":
        Select Case UCase(cTree(2))
        Case "PI"
            dResult = 3.141592654
        ' Put here you variables definitions
        Case Else
            dResult = Val(cTree(2))
        End Select
    Case "(-)":
        dResult = -TreeEval(cTree(2))
    Case "+"
        dResult = TreeEval(cTree(2))
        For i = 3 To cTree.Count
            dResult = dResult + TreeEval(cTree(i))
        Next
    Case "-":
        dResult = TreeEval(cTree(2))
        For i = 3 To cTree.Count
            dResult = dResult - TreeEval(cTree(i))
        Next
    Case "\"
        dResult = TreeEval(cTree(2))
        For i = 3 To cTree.Count
            dResult = dResult Mod TreeEval(cTree(i))
        Next
    Case "*"
        dResult = TreeEval(cTree(2))
        For i = 3 To cTree.Count
            dResult = dResult * TreeEval(cTree(i))
        Next
    Case "/"
        dResult = TreeEval(cTree(2))
        For i = 3 To cTree.Count
            dParam = TreeEval(cTree(i))
            If dParam = 0 Then
                sErrorMsg = "Division by 0"
                GoTo EvalError
            Else
                dResult = dResult / dParam
            End If
        Next
    Case "^"
        dResult = TreeEval(cTree(2))
        For i = 3 To cTree.Count
            dParam = TreeEval(cTree(i))
            If dResult < 0 And dParam < 1 Then
                sErrorMsg = "Power negative"
                GoTo EvalError
            Else
                dResult = dResult ^ dParam
            End If
        Next
    Case "FUNC":
        dParam = TreeEval(cTree(3))
        Select Case UCase(cTree(2))
        Case "COS"
            dResult = Cos(dParam)
        Case "SIN"
            dResult = Sin(dParam)
        ' Put here your functions
        Case Else
            sErrorMsg = "Function not implemented"
            GoTo EvalError
        End Select
    End Select
    
    TreeEval = dResult
    Exit Function

EvalError:
    MsgBox sErrorMsg
End Function
