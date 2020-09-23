VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStockQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tool As MSXML2.XMLHTTP

Public Function getQuote(symbol As String) As Double
On Error GoTo errorHandler

    Dim num As Long
    Dim strNew As String
    Dim coName As String
    Dim inData As String

    Set tool = New MSXML2.XMLHTTP
    
    With tool
        .open "GET", "http://quotes.nasdaq.com/quote.dll?mode=stock&page=quick&symbol=" & symbol, False
        .send
        inData = .responseText
    End With
    

    num = InStr(1, inData, "$&nbsp;") + 7
    
    strNew = Mid(inData, num, InStr(num, inData, "<") - num)
   
    Set tool = Nothing
    
    getQuote = Trim(strNew)

    Exit Function
    
errorHandler:
    
    Set tool = Nothing
    getQuote = -1

End Function

Public Function getCompanyName(symbol As String) As String
On Error GoTo errorHandler

    Dim num As Long
    Dim coName As String
    Dim inData As String
    Dim findd As String
    Dim a As Long

    Set tool = New MSXML2.XMLHTTP
    
    With tool
        .open "GET", "http://quotes.nasdaq.com/quote.dll?mode=stock&page=multi&symbol=" & symbol, False
        .send
        inData = .responseText
    End With
    
    findd = "symbolname =" & Chr(34)
    num = InStr(1, inData, findd)
    a = InStr(num, inData, ";")
    
    coName = Mid(inData, num + 13, a - (num + 14))
    
    Set tool = Nothing
    
    getCompanyName = Trim(coName)

    Exit Function
    
errorHandler:
    
    Set tool = Nothing
    getCompanyName = ""

End Function
