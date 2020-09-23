Attribute VB_Name = "Module1"
Public Function ConvertDate(dateToConvert As Date)
Dim NewDay As String
Dim NewMonth As String

Select Case Day(dateToConvert)
    Case 1
        NewDay = "1st"
    Case 2
        NewDay = "2nd"
    Case 3
        NewDay = "3rd"
    Case 4
        NewDay = "4th"
    Case 5
        NewDay = "5th"
    Case 6
        NewDay = "6th"
    Case 7
        NewDay = "7th"
    Case 8
        NewDay = "8th"
    Case 9
        NewDay = "9th"
    Case 10
        NewDay = "10th"
    Case 11
        NewDay = "11th"
    Case 12
        NewDay = "12th"
    Case 13
        NewDay = "13th"
    Case 14
        NewDay = "14th"
    Case 15
        NewDay = "15th"
    Case 16
        NewDay = "16th"
    Case 17
        NewDay = "17th"
    Case 18
        NewDay = "18th"
    Case 19
        NewDay = "19th"
    Case 20
        NewDay = "20th"
    Case 21
        NewDay = "21st"
    Case 22
        NewDay = "22nd"
    Case 23
        NewDay = "23rd"
    Case 24
        NewDay = "24th"
    Case 25
        NewDay = "25th"
    Case 26
        NewDay = "26th"
    Case 27
        NewDay = "27th"
    Case 28
        NewDay = "28th"
    Case 29
        NewDay = "29th"
    Case 30
        NewDay = "30th"
    Case 31
        NewDay = "31st"
End Select

Select Case Month(dateToConvert)
    Case 1
        NewMonth = "January"
    Case 2
        NewMonth = "Febuary"
    Case 3
        NewMonth = "March"
    Case 4
        NewMonth = "April"
    Case 5
        NewMonth = "May"
    Case 6
        NewMonth = "June"
    Case 7
        NewMonth = "July"
    Case 8
        NewMonth = "August"
    Case 9
        NewMonth = "September"
    Case 10
        NewMonth = "October"
    Case 11
        NewMonth = "November"
    Case 12
        NewMonth = "December"
End Select

If Len(Year(dateToConvert)) = 2 Then
    newyear = "20" & Str(Year(dateToConvert))
Else
    newyear = Str(Year(dateToConvert))
End If

ConvertDate = NewMonth & " " & NewDay & ", " & newyear

End Function

