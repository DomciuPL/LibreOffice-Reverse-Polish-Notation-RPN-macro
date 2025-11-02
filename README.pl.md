# libreoffice-rpn-pl

Makro w języku LibreOffice Basic realizujące kalkulator w notacji odwrotnej polskiej (RPN) dla arkusza **LibreOffice Calc**, dostosowane do polskiego formatu liczb (przecinek jako separator dziesiętny). Projekt łączy w jednym module obsługę liczb oraz prostą obsługę tekstu (konkatenacja i powielanie łańcucha). Może być używany zarówno jako funkcja arkuszowa (`=RPN("...")`), jak i jako makro wywoływane z poziomu interfejsu (okno dialogowe → wynik do komórki).

## Cechy

- język: **LibreOffice Basic**
- środowisko: **LibreOffice Calc**
- zapis liczb: **z przecinkiem**, np. `3,5`, `10,5`
- zapis RPN: **jeden łańcuch znaków**
- separatory tokenów (w kolejności sprawdzania):
  1. `;` – zalecany,
  2. pojedyncza spacja,
  3. `,` – tylko gdy nie jest częścią liczby
- stosowana mechanika: klasyczny stos RPN – tokeny przetwarzane od lewej do prawej, każdy operator zużywa dane ze stosu i odkłada wynik
- możliwość wprowadzenia wyrażenia z okna dialogowego i wpisania wyniku do bieżącej komórki

## Obsługiwane operatory

Aktualna wersja rozpoznaje **wyłącznie** poniższe symbole operatorów:

- `+`
  - liczby → dodawanie
  - tekst → konkatenacja (np. `ala kot +` → `alakot`)
- `-`
  - liczby → odejmowanie
  - tekst → **nieobsługiwane** (zwracany błąd)
- `*`
  - liczby → mnożenie
  - tekst × liczba → powtórzenie tekstu (np. `ala 3 *` → `alaalaala`)
  - tekst × tekst → konkatenacja
- `/`
  - liczby → dzielenie, z kontrolą dzielenia przez zero
  - tekst → **nieobsługiwane**
- `^`
  - liczby → potęgowanie

**Uwaga:** znak `÷` **nie jest** w tej wersji rozpoznawany. Jeśli użytkownik chce wprowadzać `÷`, należy go dopisać w sekcji sprawdzania operatorów.

## Przykłady

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

Wyrażenie z nawiasami w zapisie infiksowym:

```text
(3,5 + 4,5) * 2
```

w RPN ma postać:

```text
3,5 4,5 + 2 *
```

## Kod źródłowy

```basic
Option Explicit

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

            If IsNumeric(tok) Then
                sp = sp + 1
                If sp > UBound(st) Then
                    ReDim Preserve st(0 To UBound(st) + 16)
                End If
                st(sp) = CDbl(tok)

            ElseIf tok = "+" Or tok = "-" Or tok = "*" Or tok = "/" Or tok = "^" Then

                If sp < 1 Then
                    RPN_Eval = "ERR: za mało argumentów"
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
                        If IsNumeric(st(sp - 1)) And IsNumeric(st(sp)) Then
                            st(sp - 1) = CDbl(st(sp - 1)) - CDbl(st(sp))
                            sp = sp - 1
                        Else
                            RPN_Eval = "ERR: - tylko liczby"
                            Exit Function
                        End If

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
                            RPN_Eval = "ERR: / tylko liczby"
                            Exit Function
                        End If
                        If CDbl(st(sp)) = 0 Then
                            RPN_Eval = "ERR: dzielenie przez zero"
                            Exit Function
                        End If
                        st(sp - 1) = CDbl(st(sp - 1)) / CDbl(st(sp))
                        sp = sp - 1

                    Case "^"
                        If Not IsNumeric(st(sp - 1)) Or Not IsNumeric(st(sp)) Then
                            RPN_Eval = "ERR: ^ tylko liczby"
                            Exit Function
                        End If
                        st(sp - 1) = CDbl(st(sp - 1)) ^ CDbl(st(sp))
                        sp = sp - 1
                End Select

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
        RPN_Eval = "ERR: na stosie " & (sp + 1) & " wartości"
    Else
        RPN_Eval = st(0)
    End If
End Function

Sub RPN_DoKomorki()
    Dim s As String, v As Variant
    s = InputBox("Enter RPN expression (e.g. 3,5 4,5 +)", "RPN → cell")
    If s = "" Then Exit Sub
    v = RPN(s)
    On Error Resume Next
    ThisComponent.CurrentSelection.String = v
End Sub

Sub RPN_OknoMsg()
    Dim s As String, v As Variant
    s = InputBox("Enter RPN expression", "RPN")
    If s = "" Then Exit Sub
    v = RPN(s)
    MsgBox "Result: " & v, 64, "RPN"
End Sub
```

## Instalacja

1. Otwórz LibreOffice Calc.
2. Narzędzia → Makra → Edytuj makra…
3. Wybierz „Moje makra” albo „[ten_arkusz].ods” → Standard → nowy moduł.
4. Wklej powyższy kod i zapisz.
5. Zapisz arkusz jako `.ods`, aby zachować makro.
6. (opcjonalnie) przypisz `RPN_OknoMsg` / `RPN_DoKomorki` do skrótu.

## Licencja

MIT (proponowana).
