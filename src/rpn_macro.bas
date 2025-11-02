Option Explicit

' ================================================================
'  RPN CALCULATOR (Polish locale, numbers with comma)
'
'  Purpose
'  -------
'  This module provides a single worksheet function `RPN(...)` that
'  evaluates expressions written in Reverse Polish Notation (postfix
'  notation). It also provides two helper macros for interactive use:
'  - RPN_OknoMsg      → ask for an RPN string, show the result in a dialog
'  - RPN_DoKomorki    → ask for an RPN string, write the result to the
'                       currently selected cell
'
'  Design
'  ------
'  1. The public entry point for Calc is the function:
'         =RPN("3,5 4,5 +")
'     Calc passes ONE string. We do NOT pass multiple arguments.
'
'  2. The string is first split into tokens by `RPN_Tokenize(...)`.
'     We support three separators, in this order:
'        - semicolon ;     (preferred, works best with Polish decimal comma)
'        - space     " "
'        - comma     ,
'     Example:
'        "3,5;4,5;+;2;*"  →  ["3,5","4,5","+","2","*"]
'
'  3. The list of tokens is then evaluated by `RPN_Eval(...)`.
'     This is the stack engine. It pushes numbers, recognizes a small
'     fixed set of operators, and pops values from the stack.
'
'  4. If, at the end, the stack does NOT contain exactly one value,
'     an error string is returned.
'
'  Locale assumptions
'  ------------------
'  - We assume Polish/continental style numbers with comma as decimal
'    separator, e.g. "3,14", "10,5".
'  - We do NOT convert decimal dots to commas.
'  - If the user wants to separate tokens with commas, they must NOT
'    use commas inside numbers at the same time. The recommended format
'    is to use semicolons for tokens:
'        =RPN("10,5;2;/")
'
'  Mixed mode: numbers + text
'  --------------------------
'  To make RPN more useful, we allow tokens that are neither numbers
'  nor operators to be treated as plain text. Then operators work in
'  two paths:
'  - numeric path if both operands are numeric
'  - text path otherwise
'
'  Implemented semantics:
'  - "+" : numbers → addition
'          text    → concatenation
'          "foo" "bar" +  → "foobar"
'
'  - "*" : numbers → multiplication
'          text × number → repeat text N times
'              "ab" 3 *  → "ababab"
'          text × text   → concatenate
'
'  - "-" : numbers → subtraction
'          text    → NOT supported here (we return error for text)
'
'  - "/" : numbers → division
'          text    → NOT supported here (we return error for text)
'
'  - "^" : numbers → power
'
'  Notes
'  -----
'  - The engine is intentionally simple. It does not know operator
'    precedence, because RPN does not need it. Order is defined by the
'    sequence of tokens.
'  - Errors are returned as strings starting with "ERR:". Calc will
'    display them as text. This is deliberate: we want the user to see
'    what went wrong (unknown token, too few arguments, division by zero).
'  - The code keeps the stack in a Variant array and grows it in blocks
'    of 16 elements when needed.
'
'  Examples
'  --------
'  =RPN("3,5 4,5 +")           → 8
'  =RPN("3,5;4,5;+;2;*")       → 16
'  =RPN("ale kot +")           → "alekot"
'  =RPN("ale kot + 4 *")       → "alekotalekotalekotalekot"
'  =RPN("10,5;2;/")            → 5,25
'
'  Integration
'  -----------
'  - You can assign the macro `RPN_OknoMsg` or `RPN_DoKomorki` to a
'    keyboard shortcut via Tools → Customize → Keyboard.
'  - Only Sub procedures can be bound to shortcuts. Functions like
'    `RPN(...)` cannot, because they expect an argument.
'
'  Extending
'  ---------
'  - To add more text operators (e.g. REMOVE, LEFT, RIGHT), add new
'    cases in `RPN_Eval(...)` just like we did for "+" and "*".
'  - To add unary math functions (SIN, COS, SQRT), add a detection
'    in the operator block and pop only one value instead of two.
'
' ================================================================

Function RPN(Optional expr As Variant) As Variant
    Dim txt As String
    Dim tokens As Variant

    If IsMissing(expr) Then
        RPN = "ERR: brak wyrażenia"
        Exit Function
    End If

    txt = CStr(expr)
    tokens = RPN_Tokenize(txt)
    RPN = RPN_Eval(tokens)
End Function

Private Function RPN_Tokenize(s As String) As Variant
    Dim txt As String
    Dim parts() As String
    Dim i As Long

    txt = Trim(s)
    If txt = "" Then
        RPN_Tokenize = Array()
        Exit Function
    End If

    If InStr(txt, ";") > 0 Then
        parts = Split(txt, ";")
    ElseIf InStr(txt, " ") > 0 Then
        parts = Split(txt, " ")
    Else
        parts = Split(txt, ",")
    End If

    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim(parts(i))
    Next i

    RPN_Tokenize = parts
End Function

Private Function RPN_Eval(tokens As Variant) As Variant
    Dim st() As Variant
    Dim sp As Long
    Dim i As Long
    Dim tok As String
    Dim n As Long
    Dim j As Long
    Dim src As String
    Dim out As String

    ReDim st(0 To 0)
    sp = -1

    For i = LBound(tokens) To UBound(tokens)
        tok = tokens(i)

        If tok <> "" Then

            ' numeric literal
            If IsNumeric(tok) Then
                sp = sp + 1
                If sp > UBound(st) Then
                    ReDim Preserve st(0 To UBound(st) + 16)
                End If
                st(sp) = CDbl(tok)

            ' known operator
            ElseIf tok = "+" Or tok = "-" Or tok = "*" Or tok = "/" Or tok = "^" Then

                If sp < 1 Then
                    RPN_Eval = "ERR: too few args"
                    Exit Function
                End If

                Select Case tok
                    Case "+"
                        If IsNumeric(st(sp - 1)) And IsNumeric(st(sp)) Then
                            st(sp - 1) = CDbl(st(sp - 1)) + CDbl(st(sp))
                        Else
                            st(sp - 1) = CStr(st(sp - 1)) & CStr(st(sp))
                        End If
                        sp = sp - 1

                    Case "-"
                        If Not IsNumeric(st(sp - 1)) Or Not IsNumeric(st(sp)) Then
                            RPN_Eval = "ERR: - only for numbers"
                            Exit Function
                        End If
                        st(sp - 1) = CDbl(st(sp - 1)) - CDbl(st(sp))
                        sp = sp - 1

                    Case "*"
                        If IsNumeric(st(sp - 1)) And IsNumeric(st(sp)) Then
                            st(sp - 1) = CDbl(st(sp - 1)) * CDbl(st(sp))
                            sp = sp - 1
                        ElseIf (Not IsNumeric(st(sp - 1))) And IsNumeric(st(sp)) Then
                            n = CLng(st(sp))
                            If n < 0 Then
                                RPN_Eval = "ERR: * < 0"
                                Exit Function
                            End If
                            src = CStr(st(sp - 1))
                            out = ""
                            For j = 1 To n
                                out = out & src
                            Next j
                            st(sp - 1) = out
                            sp = sp - 1
                        Else
                            st(sp - 1) = CStr(st(sp - 1)) & CStr(st(sp))
                            sp = sp - 1
                        End If

                    Case "/"
                        If Not IsNumeric(st(sp - 1)) Or Not IsNumeric(st(sp)) Then
                            RPN_Eval = "ERR: / only for numbers"
                            Exit Function
                        End If
                        If CDbl(st(sp)) = 0 Then
                            RPN_Eval = "ERR: division by zero"
                            Exit Function
                        End If
                        st(sp - 1) = CDbl(st(sp - 1)) / CDbl(st(sp))
                        sp = sp - 1

                    Case "^"
                        If Not IsNumeric(st(sp - 1)) Or Not IsNumeric(st(sp)) Then
                            RPN_Eval = "ERR: ^ only for numbers"
                            Exit Function
                        End If
                        st(sp - 1) = CDbl(st(sp - 1)) ^ CDbl(st(sp))
                        sp = sp - 1
                End Select

            ' unknown token → treat as text
            Else
                sp = sp + 1
                If sp > UBound(st) Then
                    ReDim Preserve st(0 To UBound(st) + 16)
                End If
                st(sp) = tok
            End If

        End If
    Next i

    If sp <> 0 Then
        RPN_Eval = "ERR: stack has " & (sp + 1) & " items"
    Else
        RPN_Eval = st(0)
    End If
End Function

Sub RPN_DoKomorki()
    Dim s As String
    Dim v As Variant

    s = InputBox("Enter RPN expression (e.g. 3,5 4,5 +)", "RPN → cell")
    If s = "" Then Exit Sub

    v = RPN(s)

    On Error Resume Next
    ThisComponent.CurrentSelection.String = v
End Sub

Sub RPN_OknoMsg()
    Dim s As String
    Dim v As Variant

    s = InputBox("Enter RPN expression", "RPN")
    If s = "" Then Exit Sub

    v = RPN(s)
    MsgBox "Result: " & v, 64, "RPN"
End Sub

