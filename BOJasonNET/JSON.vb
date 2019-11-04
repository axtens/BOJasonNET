Option Explicit On

<Serializable>
Public Class RecursiveDict
    Inherits Dictionary(Of Object, RecursiveDict)

    Public Sub New()
    End Sub
End Class

Public Class Json
    ' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
    ' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
    ' BSD Licensed

    Const InvalidJson As Integer = 1
    Const InvalidObject As Integer = 2
    Const InvalidArray As Integer = 3
    Const InvalidBoolean As Integer = 4
    Const InvalidNull As Integer = 5
    Const InvalidKey As Integer = 6
    Const InvalidRpcCall As Integer = 7

    Private _psError As Integer

    '
    '   parse string and create JSON object
    '
    Public Function Parse(ByRef str As String) As RecursiveDict
        Dim index As Integer = 1
        Dim result As New RecursiveDict()
        _psError = 0
        'On Error Resume Next
        Call SkipChar(str, index)
        Select Case Mid(str, index, 1)
            Case "{"
                result = ParseObject(str, index)
            Case "["
                result = ParseArray(str, index)
            Case Else
                _psError = InvalidJson
        End Select
        Parse = result
    End Function


    '
    '   parse collection of key/value
    '

    Private Function ParseObject(ByRef str As String, ByRef index As Integer) As RecursiveDict

        Dim result = New RecursiveDict()
        Dim sKey As String

        ' "{"
        Call SkipChar(str, index)
        If Mid(str, index, 1) <> "{" Then
            _psError = InvalidObject
            ParseObject = result
            Exit Function
        End If

        index += 1

        Do
            Call SkipChar(str, index)
            If "}" = Mid(str, index, 1) Then
                index += 1
                Exit Do
            ElseIf "," = Mid(str, index, 1) Then
                index += 1
                Call SkipChar(str, index)
            ElseIf index > Len(str) Then
                _psError = InvalidObject
                Exit Do
            End If

            ' add key/value pair
            sKey = ParseKey(str, index)
            'On Error Resume Next

            result.Add(sKey, ParseValue(str, index))
            'If Err.Number <> 0 Then
            'psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
            'Exit Do
            'End If
        Loop
        ParseObject = result
    End Function

    '
    '   parse list
    '
    Private Function ParseArray(ByRef str As String, ByRef index As Integer) As RecursiveDict

        Dim result = New RecursiveDict()
        Dim counter As Integer = 0

        ' "["
        Call SkipChar(str, index)
        If Mid(str, index, 1) <> "[" Then
            _psError = InvalidArray
            ParseArray = result
            Exit Function
        End If

        index += 1

        Do

            Call SkipChar(str, index)
            If "]" = Mid(str, index, 1) Then
                index += 1
                Exit Do
            ElseIf "," = Mid(str, index, 1) Then
                index += 1
                Call SkipChar(str, index)
            ElseIf index > Len(str) Then
                _psError = InvalidArray
                Exit Do
            End If

            ' add value
            'On Error Resume Next
            result.Add(counter, ParseValue(str, index))
            counter += 1
            'If Err.Number <> 0 Then
            'psErrors = psErrors & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
            'Exit Do
            'End If
        Loop
        ParseArray = result
    End Function

    '
    '   parse string / number / object / array / true / false / null
    '
    Private Function ParseValue(ByRef str As String, ByRef index As Integer) As RecursiveDict

        Call SkipChar(str, index)
        Dim c As Char = Mid(str, index, 1)

        Dim res As RecursiveDict
        Select Case c
            Case "{"
                res = ParseObject(str, index)
            Case "["
                res = ParseArray(str, index)
            Case """", "'"
                res = ParseString(str, index)
            Case "t", "f"
                res = ParseBoolean(str, index)
            Case "n"
                res = ParseNull(str, index)
            Case Else
                res = ParseNumber(str, index)
        End Select
        ParseValue = res
    End Function

    '
    '   parse string
    '
    Private Function ParseString(ByRef str As String, ByRef index As Integer) As RecursiveDict

        Dim quote As String
        Dim sChar As Char
        Dim code As String

        Dim sb As New List(Of String)
        Call SkipChar(str, index)
        quote = Mid(str, index, 1)
        index += 1

        Do While index > 0 And index <= Len(str)
            sChar = Mid(str, index, 1)
            Select Case (sChar)
                Case "\"
                    index += 1
                    sChar = Mid(str, index, 1)
                    Select Case (sChar)
                        Case """", "\", "/", "'"
                            sb.Add(sChar)
                            index += 1
                        Case "b"
                            sb.Add(vbBack)
                            index += 1
                        Case "f"
                            sb.Add(vbFormFeed)
                            index += 1
                        Case "n"
                            sb.Add(vbLf)
                            index += 1
                        Case "r"
                            sb.Add(vbCr)
                            index += 1
                        Case "t"
                            sb.Add(vbTab)
                            index += 1
                        Case "u"
                            index += 1
                            code = Mid(str, index, 4)
                            sb.Add(ChrW(Val("&h" + code)))
                            index += 4
                    End Select
                Case quote
                    index += 1
                    Exit Do
                Case Else
                    sb.Add(sChar)
                    index += 1
            End Select
        Loop
        Dim r = New RecursiveDict
        If sb.Count = 0 Then
            r.Add("", Nothing)
        Else
            r.Add(Join(sb.ToArray(), String.Empty), Nothing)
        End If
        ParseString = r
    End Function

    '
    '   parse number
    '
    Private Function ParseNumber(ByRef str As String, ByRef index As Integer) As RecursiveDict
        Dim result As Decimal = 0
        Dim sValue As String = String.Empty

        Call SkipChar(str, index)
        Do While index > 0 And index <= Len(str)
            Dim sChar As String = Mid(str, index, 1)

            If InStr("+-0123456789.eE", sChar) Then
                sValue &= sChar
                index += 1
            Else
                result = CDec(sValue)
                Exit Do
            End If
        Loop

        Dim r = New RecursiveDict From {
            {result, Nothing}
        }
        ParseNumber = r
    End Function

    '
    '   parse true / false
    '
    Private Function ParseBoolean(ByRef str As String, ByRef index As Integer) As RecursiveDict
        Call SkipChar(str, index)
        Dim result As Boolean
        If Mid(str, index, 4) = "true" Then
            result = True
            index += 4
        ElseIf Mid(str, index, 5) = "false" Then
            result = False
            index += 5
        Else
            _psError = InvalidBoolean
            result = False
        End If

        Dim r = New RecursiveDict From {
            {result, Nothing}
        }
        ParseBoolean = r
    End Function

    '
    '   parse null
    '
    Private Function ParseNull(ByRef str As String, ByRef index As Integer) As RecursiveDict
        Dim result As VariantType = vbNull
        Call SkipChar(str, index)
        If Mid(str, index, 4) = "null" Then
            result = vbNull
            index += 4
        Else
            _psError = InvalidNull
        End If

        Dim r As New RecursiveDict From {
            {result, Nothing}
        }
        ParseNull = r
    End Function

    Private Function ParseKey(ByRef str As String, ByRef index As Integer) As String
        Dim result As String = String.Empty
        Dim dquote As Boolean
        Dim squote As Boolean
        Call SkipChar(str, index)
        Do While index > 0 And index <= Len(str)
            Dim sChar As String = Mid(str, index, 1)
            Dim sKey As String
            Select Case (sChar)
                Case """"
                    dquote = Not dquote
                    index += 1
                    If Not dquote Then
                        Call SkipChar(str, index)
                        sKey = Mid(str, index, 1)
                        If sKey <> ":" Then
                            _psError = InvalidKey
                            Exit Do
                        End If
                    End If
                Case "'"
                    squote = Not squote
                    index += 1
                    If Not squote Then
                        Call SkipChar(str, index)
                        sKey = Mid(str, index, 1)
                        If sKey <> ":" Then
                            _psError = InvalidKey
                            Exit Do
                        End If
                    End If
                Case ":"
                    index += 1
                    If Not dquote And Not squote Then
                        Exit Do
                    Else
                        result &= sChar
                    End If
                Case Else
                    If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", sChar) Then
                    Else
                        result &= sChar
                    End If
                    index += 1
            End Select
        Loop
        ParseKey = result
    End Function

    '
    '   skip special sCharacter
    '
    Private Sub SkipChar(ByRef str As String, ByRef index As Integer)
        Dim bComment As Boolean
        Dim bStartComment As Boolean
        Dim bLongComment As Boolean
        Do While index > 0 And index <= Len(str)
            Select Case Mid(str, index, 1)
                Case vbCr, vbLf
                    If Not bLongComment Then
                        bStartComment = False
                        bComment = False
                    End If

                Case vbTab, " ", "(", ")"

                Case "/"
                    If Not bLongComment Then
                        If bStartComment Then
                            bStartComment = False
                            bComment = True
                        Else
                            bStartComment = True
                            bComment = False
                            bLongComment = False
                        End If
                    Else
                        If bStartComment Then
                            bLongComment = False
                            bStartComment = False
                            bComment = False
                        End If
                    End If

                Case "*"
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                        bLongComment = True
                    Else
                        bStartComment = True
                    End If

                Case Else
                    If Not bComment Then
                        Exit Do
                    End If
            End Select

            index += 1
        Loop
    End Sub

    Public Function AsUnicode(str As String) As String

        Dim x As Integer
        Dim uStr As New List(Of String)
        Dim uChrCode As Integer
        Dim strLen As Integer
        strLen = Len(str)
        For x = 1 To strLen
            uChrCode = Asc(Mid(str, x, 1))
            Select Case uChrCode
                Case 8 ' backspace
                    uStr.Add("\b")
                Case 9 ' tab
                    uStr.Add("\t")
                Case 10 ' line feed
                    uStr.Add("\n")
                Case 12 ' formfeed
                    uStr.Add("\f")
                Case 13 ' carriage return
                    uStr.Add("\r")
                Case 34 ' quote
                    uStr.Add("\""")
                Case 39 ' apostrophe
                    uStr.Add("\'")
                Case 92 ' backslash
                    uStr.Add("\\")
                Case 123, 125 ' "{" and "}"
                    uStr.Add("\u" & Right$("0000" & Hex$(uChrCode), 4))
                Case Is < 32, Is > 127 ' non-ascii sCharacters
                    uStr.Add("\u" & Right$("0000" & Hex$(uChrCode), 4))
                Case Else
                    uStr.Add(Chr(uChrCode))
            End Select
        Next
        AsUnicode = Join(uStr.ToArray(), String.Empty)
    End Function

    Public Function GetKey(entry As RecursiveDict, Optional limit As Integer = 0) As Object
        Dim i As Integer = 0
        Dim o As Object = Nothing
        For Each pair As KeyValuePair(Of Object, RecursiveDict) In entry
            o = pair.Key
            If i = limit Then
                Return o
            Else
                i += 1
            End If
        Next
        Return o
    End Function
End Class

