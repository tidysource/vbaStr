Attribute VB_Name = "str_"
Option Explicit

'=============================
'Utility functions for strings
'=============================

'Returns a substring
'--------------------
'startI
'endI is not included in substring
'
'Should have same behaviour as https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/slice
'Note: Does not work for arrays (yet), only strings
Function slice( _
                str As String, startI As Integer, _
                Optional endI As Variant _
                ) As String
    'Default endI is length of str
    If IsMissing(endI) Then
        endI = Len(str)
    End If

    'For negative values start counting from end of string
    If startI < 0 Then
        startI = Len(str) + startI
    End If
    If endI < 0 Then
        endI = Len(str) + endI
    End If

    'Adjust to make base 0
    startI = startI + 1
    endI = endI + 1
    If endI - startI > 0 And startI > -1 Then
        slice = Mid(str, startI, endI - startI)
    Else
        slice = ""
    End If
End Function

'Normalize new line
'------------------
Function normalizeNewLines(str As String, Optional newLineChar As String = vbNewLine) As String
    'Replace all new line characters with vbCr
    str = Replace(str, vbLf, vbCr)
    str = Replace(str, vbCrLf, vbCr)
    str = Replace(str, vbNewLine, vbCr)
    str = Replace(str, "\n", vbCr)
    str = Replace(str, "\r", vbCr)
    str = Replace(str, "\n", vbCr)
    str = Replace(str, "\r\n", vbCr)

    'Replace vbCr to desired new line characted
    str = Replace(str, vbCr, newLineChar)

    normalizeNewLines = str
End Function

'Trims the string on the left for a given substring
'--------------------------------------------------
Function trimStrLeft(str As String, rmStr As String) As String
    Do While InStr(str, rmStr) = 1 And Len(str) > 0
        str = Right(str, Len(str) - Len(rmStr))
    Loop
    trimStrLeft = str
End Function

'Trims the string on the right for a given substring
'---------------------------------------------------
Function trimStrRight(str As String, rmStr As String) As String
    Dim length As Integer
    length = Len(str) - Len(rmStr)

    Do While InStrRev(str, rmStr) = length + 1 And length > -1 And Len(str) > 0
        str = Left(str, Len(str) - Len(rmStr))
        length = Len(str) - Len(rmStr)
    Loop

    trimStrRight = str
End Function

'Trims the string on both sides
'------------------------------
Function trimStr(str As String, rmStr As String) As String
    str = trimStrLeft(str, rmStr)
    str = trimStrRight(str, rmStr)

    trimStr = str
End Function

'Removes double or multiple occurances of a substring
'----------------------------------------------------
'Note: It keeps a single occurance
Function singleStr(str As String, rmStr As String) As String
    Dim doubleStr As String
    doubleStr = rmStr & rmStr

    Do While InStr(str, doubleStr) > 0
        str = Replace(str, doubleStr, rmStr)
    Loop

    singleStr = str
End Function
