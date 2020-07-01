Attribute VB_Name = "Module1"
Option Explicit

Dim Ticker As String
Dim DateValue As String
Dim OpenValue As Double
Dim HighValue As Double
Dim LowValue As Double
Dim CloseValue As Double
Dim Volume As Double

Dim InputRowNumber As Long
Dim OutputRowNumber As Integer

Dim CumulativeVolume As Double
Dim YearOpenValue As Double
Dim YearCloseValue As Double
Dim YearPercentChange As Double
Dim PriceChange As Double

Dim NextTicker As String
Dim PreviousTicker As String

Sub Stock()
    MsgBox ("Started")
    OutputRowNumber = 2
    InputRowNumber = 2
    NextTicker = Range("A2").Value
    Do While NextTicker <> ""
        Ticker = Range("A" & CStr(InputRowNumber)).Value
        DateValue = Range("B" & CStr(InputRowNumber)).Value
        OpenValue = Range("C" & CStr(InputRowNumber)).Value
        HighValue = Range("D" & CStr(InputRowNumber)).Value
        LowValue = Range("E" & CStr(InputRowNumber)).Value
        CloseValue = Range("F" & CStr(InputRowNumber)).Value
        Volume = Range("G" & CStr(InputRowNumber)).Value
        
        NextTicker = Range("A" & CStr(InputRowNumber + 1)).Value
        PreviousTicker = Range("A" & CStr(InputRowNumber - 1)).Value
        
        If InputRowNumber = 2 Then
            '''' First row with Ticker Symbol E, when InputRowNumber = 2:
            ' Start accumulating Volume for the year.
            CumulativeVolume = Volume
            ' Save this row's OpenValue for when we want to calculate Yearly Change.
            YearOpenValue = OpenValue
        End If
        
        If Ticker <> PreviousTicker And InputRowNumber > 2 Then
            '''' First row for a specific Ticker Symbol:
            ' Start accumulating Volume for the year.
            CumulativeVolume = Volume
            ' Save this row's OpenValue for when we want to calculate Yearly Change.
            YearOpenValue = OpenValue
        End If
        
        If Ticker = NextTicker Then
            '''' Middle rows with a specific Ticker Symbol, when InputRowNumber not first and not last:
            ' Continue adding up volume.
            CumulativeVolume = CumulativeVolume + Volume
        End If
        
        If Ticker <> NextTicker Then
            '''' Last row for a specific Ticker Symbol:
            ' Finish adding up volume.
            CumulativeVolume = CumulativeVolume + Volume
            ' Get this row's CloseValue for calculating Yearly Change.
            YearCloseValue = CloseValue
            ' Calculate % Yearly Change.
            PriceChange = YearCloseValue - YearOpenValue
            YearPercentChange = (YearCloseValue - YearOpenValue) - 1
            ' Print Ticker, CumulativeVolume, YearPercentChange on Row 2.
            Range("J" & CStr(OutputRowNumber)).Value = Ticker
            Range("K" & CStr(OutputRowNumber)).Value = PriceChange
            Range("L" & CStr(OutputRowNumber)).Value = YearPercentChange
            Range("M" & CStr(OutputRowNumber)).Value = CumulativeVolume
            OutputRowNumber = OutputRowNumber + 1
        End If

        InputRowNumber = InputRowNumber + 1
    Loop

    MsgBox ("Stopped")

End Sub


