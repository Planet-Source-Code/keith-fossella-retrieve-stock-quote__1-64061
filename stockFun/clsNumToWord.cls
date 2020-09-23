VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumToWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim digits() As String
Dim digitsDecimal() As String
Dim wordString As String

Public Function numToWord(someNumber As String, isCurrency As Boolean) As String
    
    Dim i As Long
    Dim decimalLocation As Long
    Dim leftOfDecimal As String
    Dim rightOfDecimal As String
    Dim temp As String
    Dim temp2 As String
    Dim temp3 As String
    
    Dim decimalPoint As Boolean
    Dim blnErr As Boolean
    Dim blnNegative As Boolean
    Dim blnPercent As Boolean
    
    someNumber = Replace(someNumber, ",", "")
    someNumber = Replace(someNumber, "$", "")
    
    If InStr(1, someNumber, "-") Then
        blnNegative = True
        someNumber = Replace(someNumber, "-", "")
    End If
    
    If InStr(1, someNumber, "%") Then
        blnPercent = True
        someNumber = Replace(someNumber, "%", "")
    End If
    
    If InStr(1, someNumber, ".") Then
        decimalLocation = InStr(1, someNumber, ".")
        If InStr(decimalLocation + 1, someNumber, ".") Then
            numToWord = "Multi Decimaled"
            Exit Function
        Else
            decimalPoint = True
        End If
    End If
    
    If Not (IsNumeric(someNumber)) Then
            numToWord = "Not Numeric"
            Exit Function
    End If
    

   
    If decimalPoint Then
    
        rightOfDecimal = Mid(someNumber, decimalLocation + 1)
        leftOfDecimal = Left(someNumber, decimalLocation - 1)
        
        ReDim digits(Len(leftOfDecimal))
        ReDim digitsDecimal(Len(rightOfDecimal))
        temp = StrReverse(rightOfDecimal)
        For i = 0 To UBound(digitsDecimal)
            digitsDecimal(i) = Mid(temp, i + 1, 1)
        Next i
        
        temp = StrReverse(leftOfDecimal)
        
        For i = 0 To UBound(digits)
            digits(i) = Mid(temp, i + 1, 1)
        Next i
        
    Else
        ReDim digits(Len(someNumber))
        temp = StrReverse(someNumber)
        For i = 0 To UBound(digits)
            digits(i) = Mid(temp, i + 1, 1)
        Next i
    End If
      
    
    Select Case UBound(digits)
        Case 1
            wordString = getTrueName(digits(0))
        Case 2
            wordString = combo(digits(1) + digits(0))
        Case 3
            wordString = hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 4
            wordString = thousand(getTrueName(digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 5
            wordString = thousand(combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 6
            wordString = thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 7
            wordString = million(getTrueName(digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 8
            wordString = million(combo(digits(7) + digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 9
            wordString = million(hundred(getTrueName(digits(8))) + combo(digits(7) + digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 10
            wordString = billion(getTrueName(digits(9))) + million(hundred(getTrueName(digits(8))) + combo(digits(7) + digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 11
            wordString = billion(combo(digits(10) + digits(9))) + million(hundred(getTrueName(digits(8))) + combo(digits(7) + digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 12
            wordString = billion(hundred(getTrueName(digits(11))) + combo(digits(10) + digits(9))) + million(hundred(getTrueName(digits(8))) + combo(digits(7) + digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 13
            wordString = trillion(getTrueName(digits(12))) + billion(hundred(getTrueName(digits(11))) + combo(digits(10) + digits(9))) + million(hundred(getTrueName(digits(8))) + combo(digits(7) + digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 14
            wordString = trillion(combo(digits(13) + digits(12))) + billion(hundred(getTrueName(digits(11))) + combo(digits(10) + digits(9))) + million(hundred(getTrueName(digits(8))) + combo(digits(7) + digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case 15
            wordString = trillion(hundred(getTrueName(digits(14))) + combo(digits(13) + digits(12))) + billion(hundred(getTrueName(digits(11))) + combo(digits(10) + digits(9))) + million(hundred(getTrueName(digits(8))) + combo(digits(7) + digits(6))) + thousand(hundred(getTrueName(digits(5))) + combo(digits(4) + digits(3))) + hundred(getTrueName(digits(2))) + combo(digits(1) + digits(0))
        Case Else
            If UBound(digits) > 15 Then
                blnErr = True
            Else
                wordString = ""
            End If
            
    End Select
        
    If decimalPoint Then
        If isCurrency Then
            Select Case UBound(digitsDecimal)
                Case 1
                    wordString = wordString + cents(getTrueName(digitsDecimal(0)))
                Case Else
                    wordString = wordString + cents(combo(digitsDecimal(UBound(digitsDecimal) - 1) + digitsDecimal(UBound(digitsDecimal) - 2)))
            End Select
        Else
            Select Case UBound(digitsDecimal)
                Case 1
                    wordString = wordString + tenths(getTrueName(digitsDecimal(0)))
                Case 2
                    wordString = wordString + hundreths(combo(digitsDecimal(1) + digitsDecimal(0)))
                Case Else
                    wordString = wordString + thousanths(hundred(getTrueName(digitsDecimal(UBound(digitsDecimal) - 1))) + combo(digitsDecimal(UBound(digitsDecimal) - 2) + digitsDecimal(UBound(digitsDecimal) - 3)))
            End Select
        End If
    Else
        If isCurrency Then
            If (wordString <> "") And (wordString <> "One") Then
                wordString = wordString + " Dollars"
            ElseIf wordString <> "" Then
                wordString = wordString + " Dollar"
            End If
        End If
    End If
 

    If blnErr = True Then
        numToWord = "To many digits left of decimal"
    Else
        If blnNegative And wordString <> "" Then
            numToWord = "Negative " + wordString
        ElseIf blnPercent And wordString <> "" Then
            numToWord = wordString + " Percent"
        Else
            numToWord = wordString
        End If
    End If
                
End Function

Private Function getTrueName(inData As String) As String
    Select Case inData
        Case 0
            getTrueName = ""
        Case 1
            getTrueName = "One"
        Case 2
            getTrueName = "Two"
        Case 3
            getTrueName = "Three"
        Case 4
            getTrueName = "Four"
        Case 5
            getTrueName = "Five"
        Case 6
            getTrueName = "Six"
        Case 7
            getTrueName = "Seven"
        Case 8
            getTrueName = "Eight"
        Case 9
            getTrueName = "Nine"
        Case Else
            getTrueName = ""
    End Select
    Exit Function
End Function


Private Function combo(inData As String) As String
    
    Dim halfString As String
    
    Select Case Mid(inData, 1, 1)  'digits(1)
        Case 0
            combo = getTrueName(Mid(inData, 2, 1))  'digits(0))
        Case 1
           combo = teens(Mid(inData, 2, 1))
        Case 2
            halfString = "Twenty"
            combo = halfString + " " + getTrueName(Mid(inData, 2, 1))
        Case 3
            halfString = "Thirty"
            combo = halfString + " " + getTrueName(Mid(inData, 2, 1))
        Case 4
            halfString = "Fourty"
            combo = halfString + " " + getTrueName(Mid(inData, 2, 1))
        Case 5
            halfString = "Fifty"
            combo = halfString + " " + getTrueName(Mid(inData, 2, 1))
        Case 6
            halfString = "Sixty"
            combo = halfString + " " + getTrueName(Mid(inData, 2, 1))
        Case 7
            halfString = "Seventy"
            combo = halfString + " " + getTrueName(Mid(inData, 2, 1))
        Case 8
            halfString = "Eighty"
            combo = halfString + " " + getTrueName(Mid(inData, 2, 1))
        Case 9
            halfString = "Ninety"
            combo = halfString + " " + getTrueName(Mid(inData, 2, 1))
        Case Else
            combo = ""
    End Select
End Function

Private Function cents(inData As String) As String
If inData <> "" Then
    If (wordString <> "") And (wordString <> "One") Then
        cents = " Dollars and " + inData + " Cent" + IIf((inData <> "") And (inData <> "One"), "s", "")
    ElseIf wordString <> "" Then
        cents = " Dollar and " + inData + " Cent" + IIf((inData <> "") And (inData <> "One"), "s", "")
    Else
        cents = inData + " Cent" + IIf((inData <> "") And (inData <> "One"), "s", "")
    End If
Else
    If (wordString <> "") And (wordString <> "One") Then
        cents = inData + " Dollars"
    ElseIf wordString <> "" Then
        cents = inData + " Dollar"
    Else
        cents = ""
    End If
    
End If
End Function

Private Function hundred(inData As String) As String
If inData <> "" Then
    hundred = inData + " Hundred "
Else
    hundred = inData + ""
End If
End Function


Private Function thousand(inData As String) As String
    If inData <> "" Then
        thousand = inData + " Thousand "
    Else
        thousand = ""
    End If
End Function

Private Function million(inData As String) As String
    If inData <> "" Then
        million = inData + " Million "
    Else
        million = ""
    End If
End Function

Private Function billion(inData As String) As String
    If inData <> "" Then
        billion = inData + " Billion "
    Else
        billion = ""
    End If
End Function

Private Function trillion(inData As String) As String
    If inData <> "" Then
        trillion = inData + " Trillion "
    Else
        trillion = ""
    End If
End Function

Private Function hundreths(inData As String) As String
    If inData <> "" Then
        hundreths = IIf(wordString <> "", " and ", "") + inData + " Hundreths"
    Else
        hundreths = ""
    End If
End Function

Private Function thousanths(inData As String) As String
    If inData <> "" Then
        thousanths = IIf(wordString <> "", " and ", "") + inData + " Thousanths"
    Else
        thousanths = ""
    End If
End Function

Private Function tenths(inData As String) As String
    If inData <> "" Then
        tenths = IIf(wordString <> "", " and ", "") + inData + " Tenth" + IIf((inData <> "") And (inData <> "One"), "s", "")
    Else
        tenths = ""
    End If
End Function

Private Function teens(inData As String) As String
Select Case inData
    Case 0
        teens = "Ten"
    Case 1
        teens = "Eleven"
    Case 2
        teens = "Twelve"
    Case 3
        teens = "Thirteen"
    Case 4
        teens = "Fourteen"
    Case 5
        teens = "Fifteen"
    Case 6
        teens = "Sixteen"
    Case 7
        teens = "Seventeen"
    Case 8
        teens = "Eighteen"
    Case 9
        teens = "Nineteen"
    Case Else
        teens = ""
    End Select
End Function
