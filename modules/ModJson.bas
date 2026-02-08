Attribute VB_Name = "ModJson"
'---------------------------------------------------
' VBA JSON Parser
'---------------------------------------------------
Option Explicit
Private p&, token, dic

Function ParseJson(json$, Optional Key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj Key Else ParseArr Key
    Set ParseJson = dic
End Function

Function ParseObj(Key$)
    Do: p = p + 1
        Select Case token(p)
        Case "]"
        Case "[":  ParseArr Key
        Case "{"
            If token(p + 1) = "}" Then
                p = p + 1
                dic.Add Key, "null"
            Else
                ParseObj Key
            End If
                
        Case "}":  Key = ReducePath(Key): Exit Do
        Case ":":  Key = Key & "." & token(p - 1)
        Case ",":  Key = ReducePath(Key)
        Case Else: If token(p + 1) <> ":" Then dic.Add Key, token(p)
        End Select
    Loop
End Function

Function ParseArr(Key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
        Case "}"
        Case "{":  ParseObj Key & ArrayID(e)
        Case "[":  ParseArr Key
        Case "]":  Exit Do
        Case ":":  Key = Key & ArrayID(e)
        Case ",":  e = e + 1
        Case Else: 'dic.Add key & ArrayID(e), Token(p)
        End Select
    Loop
End Function


'---------------------------------------------------
' Support Functions
'---------------------------------------------------
Function Tokenize(s$)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function
Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
    Dim c&, m, n, v
    With CreateObject("vbscript.regexp")
        .Global = bGlobal
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = Pattern
        If .Test(s) Then
            Set m = .Execute(s)
            ReDim v(1 To m.Count)
            For Each n In m
                c = c + 1
                v(c) = n.value
                If bGroup1Bias Then If Len(n.submatches(0)) Or n.value = """""" Then v(c) = n.submatches(0)
            Next
        End If
    End With
    RExtract = v
End Function

Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function

Function ReducePath$(Key$)
    If InStr(Key, ".") Then ReducePath = Left(Key, InStrRev(Key, ".") - 1) Else ReducePath = Key
End Function

Function ListPaths(dic)
    Dim s$, v
    For Each v In dic
        s = s & v & " --> " & dic(v) & vbLf
    Next
    'Debug.Print s
End Function

Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.Keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function

Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.Keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
        z = GetFilteredValues(dic, cols(j - 1))
        For i = 1 To UBound(z)
            w(i, j) = z(i)
        Next
    Next
    GetFilteredTable = w
End Function

Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
    End With
End Function

