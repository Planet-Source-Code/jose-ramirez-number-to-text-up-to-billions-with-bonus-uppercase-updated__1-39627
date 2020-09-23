Attribute VB_Name = "NumToText"
'=====================================================================
'Module for converting long numbers to text.
'If you use this in one of your applications,
'the About box of the program must show the following:

'Number to Text algorithm kindly provided by:
'José Pablo Ramírez Vargas
'San José, Costa Rica
'webJose@developerfusion.com

'Or in Spanish:
'Algoritmo de conversión de números a letras provisto amablemente por:
'José Pablo Ramírez Vargas
'San José, Costa Rica
'webJose@developerfusion.com
'=====================================================================
Option Explicit

'=======================================
'Section for upper case alternate method
'=======================================
Private Const c_iLowerCaseBit As Integer = &H20 'Unicode version (2 bytes)

Private Type SafeArray1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'==================================
'End of upper case alternate method
'==================================

'Languages enum
Public Enum LanguagesEnum
    LangSpanish
    LangEnglish
End Enum

'Format enum
Public Enum FormatsEnum
    FormatNone
    FormatUpperTitle
    FormatUpperFirstLetter
    FormatUpperAll
End Enum

Public Function NumberToText(ByVal lNumber As Long, Optional lLang As LanguagesEnum = LangEnglish, Optional lFormat As FormatsEnum = FormatNone) As String

Dim bCount As Byte
Dim bOrder As Byte
Dim strFinal As String
Dim strTemp As String
Dim strPass As String
Dim lNumberOrder As Long
Dim lTemp As Long

    If (lNumber = 0) Then
        strFinal = TextStringNumber(0, lLang, True)
        FormatOutput strFinal, lFormat
        NumberToText = strFinal
        Exit Function
    End If
    strFinal = ""
    'Get the number order (base 10)
    'This number is being rounded up by VB when log > x.5
    'so special consideration was needed while appending
    'the order qualifier.
    'Why not use int(log10(lnumber))??  Well, I tried, but VB
    'keeps saying that int(log10(1000))=2!!!!!
    lNumberOrder = Log10(lNumber)
    'Get the number of 10^3 orders
    bOrder = CByte(lNumberOrder \ 3)
    For bCount = 0 To bOrder
        strPass = ""
        'Get the 10^3 part we are on
        lTemp = (lNumber \ 10 ^ (3 * bCount)) * 10 ^ (3 * bCount)
        'The next line may overflow, this is why the error trapping.
        'As far as my testing goes, it can safely be ignored.
        On Error Resume Next
        lTemp = lTemp - (lNumber \ 10 ^ (3 * (bCount + 1))) * 10 ^ (3 * (bCount + 1))
        On Error GoTo 0
        lTemp = lTemp / 10 ^ (3 * bCount)
        'This part is english-only}
        If (lLang = LangEnglish) Then
            'Convert the hundreds
            strTemp = TextStringNumber(lTemp \ 100, lLang)
            'Append...
            strPass = strPass & IIf((strPass <> "") And (strTemp <> ""), Space(1), "") & strTemp
            'Append the hundreds qualifier
            If (strTemp <> "") Then
                strPass = strPass & IIf(strPass <> "", Space(1), "") & TextHundredQualifier(lTemp \ 100, lLang, ((lTemp Mod 100) = 0))
            End If
            'Now get the number string for everything below 100.
            strTemp = TextStringNumber(lTemp Mod 100, lLang, False, ((lTemp \ 100) > 0))
        ElseIf (lLang = LangSpanish) Then
            'spanish version
            'Append the hundreds qualifier
            strPass = strPass & IIf(strPass <> "", Space(1), "") & TextHundredQualifier(lTemp \ 100, lLang, ((lTemp Mod 100) = 0))
            'Now get the number string for everything below 100.
            strTemp = TextStringNumber(lTemp Mod 100, lLang, False, (bCount > 0))
        End If
        'Append...
        strPass = strPass & IIf((strPass <> "") And (strTemp <> ""), Space(1), "") & strTemp
        'Now get the order qualifier, if applicable
        If (lTemp > 0) Then
            strTemp = TextOrderQualifier(bCount, lLang, (lTemp > 1))
        End If
        'Append...
        strPass = strPass & IIf((strPass <> "") And (strTemp <> ""), Space(1), "") & strTemp
        'Now add to the main string
        strFinal = strPass & IIf(strPass <> "", Space(1), "") & strFinal
    Next bCount
    FormatOutput strFinal, lFormat
    NumberToText = strFinal
End Function

Private Function Log10(ByVal lNumber As Long) As Double
    Log10 = Log(CDbl(lNumber)) / Log(10#)
End Function

Private Function TextStringNumber(ByVal lNumber As Long, ByVal lLang As LanguagesEnum, Optional ByVal bShowZero As Boolean = False, Optional ByVal bSpecial As Boolean = False) As String
    If (lLang = LangEnglish) Then
        TextStringNumber = TextStringNumberEn(lNumber, bShowZero, bSpecial)
    ElseIf (lLang = LangSpanish) Then
        TextStringNumber = TextStringNumberEs(lNumber, bShowZero, bSpecial)
    End If
End Function

Private Function TextStringNumberEn(ByVal lNumber As Long, Optional bShowZero As Boolean = False, Optional bUseAnd As Boolean = False) As String

Dim arr20AndAbove(2 To 9) As String

    arr20AndAbove(2) = "twenty"
    arr20AndAbove(3) = "thirty"
    arr20AndAbove(4) = "forty"
    arr20AndAbove(5) = "fifty"
    arr20AndAbove(6) = "sixty"
    arr20AndAbove(7) = "seventy"
    arr20AndAbove(8) = "eighty"
    arr20AndAbove(9) = "ninety"
    Select Case lNumber
        Case 0
            TextStringNumberEn = IIf(bShowZero, "zero", "")
        Case 1
            TextStringNumberEn = "one"
        Case 2
            TextStringNumberEn = "two"
        Case 3
            TextStringNumberEn = "three"
        Case 4
            TextStringNumberEn = "four"
        Case 5
            TextStringNumberEn = "five"
        Case 6
            TextStringNumberEn = "six"
        Case 7
            TextStringNumberEn = "seven"
        Case 8
            TextStringNumberEn = "eight"
        Case 9
            TextStringNumberEn = "nine"
        Case 10
            TextStringNumberEn = "ten"
        Case 11
            TextStringNumberEn = "eleven"
        Case 12
            TextStringNumberEn = "twelve"
        Case 13
            TextStringNumberEn = "thirteen"
        Case 14, 16, 17, 19
            TextStringNumberEn = TextStringNumberEn(lNumber - 10) & "teen"
        Case 15
            TextStringNumberEn = "fifteen"
        Case 18
            TextStringNumberEn = "eighteen"
        Case Is > 19
            TextStringNumberEn = arr20AndAbove((lNumber \ 10)) & IIf(lNumber > (lNumber \ 10) * 10, "-", "") & TextStringNumberEn(lNumber - (lNumber \ 10) * 10)
    End Select
    If (TextStringNumberEn <> "") And bUseAnd Then
        TextStringNumberEn = "and " & TextStringNumberEn
    End If
End Function

Private Function TextOrderQualifier(ByVal bOrder As Byte, ByVal lLang As LanguagesEnum, Optional bPlural As Boolean = False)
    If (lLang = LangEnglish) Then
        Select Case bOrder
            Case 0
                TextOrderQualifier = ""
            Case 1
                TextOrderQualifier = "thousand"
            Case 2
                TextOrderQualifier = "million"
            Case 3
                'This is the US billion:  1000 millions
                TextOrderQualifier = "billion" & IIf(bPlural, "s", "")
        End Select
    ElseIf (lLang = LangSpanish) Then
        Select Case bOrder
            Case 0
                TextOrderQualifier = ""
            Case 1
                TextOrderQualifier = "mil"
            Case 2
                TextOrderQualifier = "mill" & IIf(bPlural, "ones", "ón")
            Case 3
                'In my country (Costa Rica), a billion is one million millions
                'and the US billion is, literally translated, a thousand millions.
                'That is the form I am using, therefore this case returns the same
                'as case 1.
                TextOrderQualifier = "mil"
        End Select
    End If
End Function

Public Function TextHundredQualifier(ByVal lNumber As Long, ByVal lLang As LanguagesEnum, Optional bUseSingleOneHundred As Boolean = False) As String
    If (lLang = LangEnglish) Then
        TextHundredQualifier = "hundred"
    ElseIf (lLang = LangSpanish) Then
        Select Case lNumber
            Case 0
                TextHundredQualifier = ""
            Case 1
                TextHundredQualifier = IIf(bUseSingleOneHundred, "cien", "ciento")
            Case 2 To 4, 6, 8
                TextHundredQualifier = TextStringNumberEs(lNumber) & "cientos"
            Case 5
                TextHundredQualifier = "quinientos"
            Case 7
                TextHundredQualifier = "setecientos"
            Case 9
                TextHundredQualifier = "novecientos"
            Case Else
                TextHundredQualifier = "###This case should never happen!###"
        End Select
    End If
End Function

Private Function TextStringNumberEs(ByVal lNumber As Long, Optional ByVal bShowZero As Boolean = False, Optional bUseAlternateOne As Boolean = False) As String

Dim arr30AndAbove(3 To 9) As String

    arr30AndAbove(3) = "treinta"
    arr30AndAbove(4) = "cuarenta"
    arr30AndAbove(5) = "cincuenta"
    arr30AndAbove(6) = "sesenta"
    arr30AndAbove(7) = "setenta"
    arr30AndAbove(8) = "ochenta"
    arr30AndAbove(9) = "noventa"
    Select Case lNumber
        Case 0
            TextStringNumberEs = IIf(bShowZero, "cero", "")
        Case 1
            TextStringNumberEs = IIf(bUseAlternateOne, "un", "uno")
        Case 2
            TextStringNumberEs = "dos"
        Case 3
            TextStringNumberEs = "tres"
        Case 4
            TextStringNumberEs = "cuatro"
        Case 5
            TextStringNumberEs = "cinco"
        Case 6
            TextStringNumberEs = "seis"
        Case 7
            TextStringNumberEs = "siete"
        Case 8
            TextStringNumberEs = "ocho"
        Case 9
            TextStringNumberEs = "nueve"
        Case 10
            TextStringNumberEs = "diez"
        Case 11
            TextStringNumberEs = "once"
        Case 12
            TextStringNumberEs = "doce"
        Case 13
            TextStringNumberEs = "trece"
        Case 14
            TextStringNumberEs = "catorce"
        Case 15
            TextStringNumberEs = "quince"
        Case 16 To 19
            TextStringNumberEs = "dieci" & TextStringNumberEs(lNumber - 10)
        Case 20
            TextStringNumberEs = "veinte"
        Case 21, 24 To 29
            TextStringNumberEs = "veinti" & TextStringNumberEs(lNumber - (lNumber \ 10) * 10)
        Case 22
            TextStringNumberEs = "veintidós"
        Case 23
            TextStringNumberEs = "veintitrés"
        Case Is > 29
            TextStringNumberEs = arr30AndAbove((lNumber \ 10)) & IIf(lNumber > (lNumber \ 10) * 10, " y ", "") & TextStringNumberEs(lNumber - (lNumber \ 10) * 10)
    End Select
End Function

Private Sub FormatOutput(strText As String, ByVal lFormat As FormatsEnum)

Dim sa As SafeArray1D
Dim arrString() As Integer
Dim lCount As Long
Dim bFirstLetter As Boolean

    If (strText = "") Or (lFormat = FormatNone) Then
        Exit Sub
    End If
    sa.cbElements = 2
    sa.cDims = 1
    sa.cElements = Len(strText)
    sa.lLbound = 1
    sa.pvData = StrPtr(strText)
    CopyMemory ByVal VarPtrArray(arrString), VarPtr(sa), 4
    If (lFormat = FormatUpperFirstLetter) Then
        arrString(sa.lLbound) = arrString(sa.lLbound) And Not (c_iLowerCaseBit)
    Else
        For lCount = 1 To sa.cElements
            bFirstLetter = (lCount = 1)
            If (lCount > 1) Then
                bFirstLetter = bFirstLetter Or (arrString(lCount - 1) = &H20)
            End If
            If (lFormat = FormatUpperAll) Or (bFirstLetter And (lFormat = FormatUpperTitle)) Then
                Select Case arrString(lCount)
                    Case &H61 To &H7A, &HE0 To &HFD
                        arrString(lCount) = arrString(lCount) And Not (c_iLowerCaseBit)
                End Select
            End If
        Next lCount
    End If
    CopyMemory ByVal VarPtrArray(arrString), 0&, 4
End Sub
