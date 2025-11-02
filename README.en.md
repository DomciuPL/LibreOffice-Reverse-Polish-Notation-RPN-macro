# libreoffice-rpn-pl (English)

This project provides a simple Reverse Polish Notation (RPN) calculator implemented as a **LibreOffice Basic** macro for **LibreOffice Calc**. It is tailored to the **Polish numeric format** (decimal comma), but the logic is explicit and easy to adapt. The macro can evaluate numeric RPN expressions and, in addition, can perform very simple text operations (concatenation and text repetition). It can be used both as a worksheet function and via dialog-based macros.

## Features

- language: **LibreOffice Basic**
- target: **LibreOffice Calc**
- numbers use **comma** as decimal separator, e.g. `3,5`, `10,5`
- expression is passed as **one string**: `=RPN("3,5 4,5 +")`
- supported token delimiters (in this order):
  1. semicolon `;` – recommended
  2. single space `" "`
  3. comma `,` – only if it is not used inside a number
- classic RPN stack evaluation: no operator precedence, order is defined by token sequence
- two helper macros: one writes the result to the current cell, the other shows it in a message box

## Supported operators

Current version recognizes these operators **only**:

- `+`
  - numbers → addition
  - text → concatenation  
    example: `ala kot +` → `alakot`
- `-`
  - numbers → subtraction
  - text → **not supported** (returns an error)
- `*`
  - numbers → multiplication
  - text × number → repeat text  
    example: `ala 3 *` → `alaalaala`
  - text × text → concatenate
- `/`
  - numbers → division, with division-by-zero check
  - text → **not supported**
- `^`
  - numbers → exponentiation

**Note:** the character `÷` is **not** recognized in this version. If you want to type `÷` instead of `/`, you must add it to the operator-check section.

## Examples

```text
=RPN("3,5 4,5 +")
→ 8

=RPN("3,5;4,5;+;2;*")
→ 16

=RPN("10,5;2;/")
→ 5,25

=RPN("ala kot +")
→ alakot

=RPN("ala kot + 4 *")
→ alakotalekotalekotalekot
```

Infix expression with parentheses:

```text
(3,5 + 4,5) * 2
```

corresponds to this RPN:

```text
3,5 4,5 + 2 *
```

## Source code (LibreOffice Basic)

```basic
Option Explicit

' RPN worksheet function
Function RPN(Optional expr As Variant) As Variant
    Dim txt As String
    Dim tokens As Variant

    If IsMissing(expr) Then
        RPN = "ERR: no expression"
        Exit Function
    End If

    txt = CStr(expr)
    tokens = RPN_Tokenize(txt)
    RPN = RPN_Eval(tokens)
End Function

' Split input string into tokens
Private Function RPN_Tokenize(s As String) As Variant
    Dim txt As String
    Dim parts() As String
    Dim i As Long

    txt = Trim(s)
    If txt = "" Then
        RPN_Tokenize = Array()
        Exit Function
    End If

    ' 1) prefer semicolon
    If InStr(txt, ";") > 0 Then
        parts = Split(txt, ";")
    ' 2) then space
    ElseIf InStr(txt, " ") > 0 Then
        parts = Split(txt, " ")
    ' 3) finally comma
    Else
        parts = Split(txt, ",")
    End If

    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim(parts(i))
    Next i

    RPN_Tokenize = parts
End Function

' Core RPN stack evaluator
Private Function RPN_Eval(tokens As Variant) As Variant
    Dim st() As Variant
    Dim sp As Long
    Dim i As Long
    Dim tok As String
    Dim n As Long, j As Long
    Dim src As String, out As String

    ReDim st(0 To 0)
    sp = -1

    For i = LBound(tokens) To UBound(tokens)
        tok = tokens(i)

        If tok <> "" Then

            ' number (with comma in Polish locale)
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
                        ' numeric addition OR text concatenation
                        If IsNumeric(st(sp - 1)) And IsNumeric(st(sp)) Then
                            st(sp - 1) = CDbl(st(sp - 1)) + CDbl(st(sp))
                        Else
                            st(sp - 1) = CStr(st(sp - 1)) & CStr(st(sp))
                        End If
                        sp = sp - 1

                    Case "-"
                        If IsNumeric(st(sp - 1)) And IsNumeric(st(sp)) Then
                            st(sp - 1) = CDbl(st(sp - 1)) - CDbl(st(sp))
                            sp = sp - 1
                        Else
                            RPN_Eval = "ERR: - numbers only"
                            Exit Function
                        End If

                    Case "*"
                        ' number * number
                        If IsNumeric(st(sp - 1)) And IsNumeric(st(sp)) Then
                            st(sp - 1) = CDbl(st(sp - 1)) * CDbl(st(sp))
                            sp = sp - 1
                        ' text * number → repeat
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
                        ' text * text → concatenate
                        Else
                            st(sp - 1) = CStr(st(sp - 1)) & CStr(st(sp))
                            sp = sp - 1
                        End If

                    Case "/"
                        If Not IsNumeric(st(sp - 1)) Or Not IsNumeric(st(sp)) Then
                            RPN_Eval = "ERR: / numbers only"
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
                            RPN_Eval = "ERR: ^ numbers only"
                            Exit Function
                        End If
                        st(sp - 1) = CDbl(st(sp - 1)) ^ CDbl(st(sp))
                        sp = sp - 1

                End Select

            ' unknown token → treat as text literal, push
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

' Dialog → current cell
Sub RPN_DoKomorki()
    Dim s As String, v As Variant
    s = InputBox("Enter RPN expression (e.g. 3,5 4,5 +)", "RPN → cell")
    If s = "" Then Exit Sub
    v = RPN(s)
    On Error Resume Next
    ThisComponent.CurrentSelection.String = v
End Sub

' Dialog → message box
Sub RPN_OknoMsg()
    Dim s As String, v As Variant
    s = InputBox("Enter RPN expression", "RPN")
    If s = "" Then Exit Sub
    v = RPN(s)
    MsgBox "Result: " & v, 64, "RPN"
End Sub
```

## Installation

1. Open **LibreOffice Calc**.
2. Go to **Tools → Macros → Edit Macros…**.
3. Choose either **My Macros → Standard** or **This Document → Standard**.
4. Create a new module and paste the code above.
5. Save the document as `.ods` so that the macro is stored with the file.
6. Optionally: **Tools → Customize → Keyboard** and bind `RPN_OknoMsg` or `RPN_DoKomorki` to a shortcut.

## Suggested repository layout

```text
.
├── README.pl.md      # Polish version
├── README.en.md      # this file
├── src
│   └── rpn_macro.bas
```

## License

MIT (recommended).
