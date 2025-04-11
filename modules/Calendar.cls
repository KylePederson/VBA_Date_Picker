VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 1
Option Explicit

Private monthDays() As Date
Private viewDate As Date
Public viewLevel As Integer
Private caledarDisplay As String

Private uiWidth As Double
Private uiHeight As Double
Private listOfMonths() As String

Dim Events() As New ButtonEvents


Private Sub Class_Initialize()
    viewLevel = 1
    
    '# - create the month list
    Dim monthNames As String
        monthNames = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"
    listOfMonths = Split(monthNames, ",")
    
    '# - get button frame dimensions
    uiWidth = form_calendar.fm_days.Width
    uiHeight = form_calendar.fm_days.Height
    
    '# - initialize
    viewDate = DateAdd("d", -day(Date) + 1, Date)
    Call UpdateCalendar
    
End Sub


Function ChangeViewDate(ByVal direction As Integer)
    '# - based on view level, increase or decrease viewDate
    Select Case viewLevel
        Case 1
            viewDate = DateAdd("m", 1 * direction, viewDate)
            Call UpdateCalendar
        Case 2
            viewDate = DateAdd("yyyy", 1 * direction, viewDate)
            Call DisplayMonths
        Case 3
            viewDate = DateAdd("yyyy", 9 * direction, viewDate)
            Call DisplayYears
    End Select
End Function


Private Function UpdateCalendar()
    Call ClearCalendar      '# - remove all buttons from calendar
    Call GetCalendar        '# - get all days in month
    Call DisplayDays        '# - generate new buttons
    Call SetDisplay         '# - set calendar display
End Function


'# - gets all days in the month and pushes to monthDays array
Private Function GetCalendar()
    
    '# - clear and resize array to 1
    ReDim monthDays(1)
    
    '# - set date increment to first of month
    Dim newDate As Date
        newDate = viewDate
    
    '# - counter used to redim array
    Dim counter As Integer
        counter = 1
    
    '# - push all days in month to the monthDays array
    Do While Month(newDate) = Month(viewDate)
        ReDim Preserve monthDays(counter)   '# - extend the array size
        monthDays(counter) = newDate        '# - add the new date
        newDate = DateAdd("d", 1, newDate)  '# - increment date
        counter = counter + 1               '# - increment counter
    Loop
    
End Function


Private Function DisplayDays()

    Dim btn As MSForms.CommandButton
    Dim dayOfWeek As Integer
    Dim weekOfMonth As Integer
        weekOfMonth = 1
    form_calendar.fm_day_names.Visible = True
    
    '# - clear and resize events array
    ReDim Events(1 To 1)
    Dim counter As Integer
        counter = 1
    Dim i As Integer
    Dim j As Integer
    '# - create buttons for days of month
    For i = LBound(monthDays) To UBound(monthDays)

        dayOfWeek = weekDay(monthDays(i))
        
        '# - create button and position
        Set btn = CreateButton("btn-" & monthDays(i))
        Call PositionButton(btn, Str(i), dayOfWeek, weekOfMonth, 7, 6)
        
        '# - add button events
        ReDim Preserve Events(1 To counter)
        Set Events(i).SelectDate = btn
        counter = counter + 1
        If dayOfWeek = 7 Then weekOfMonth = weekOfMonth + 1
    Next i
End Function


Private Function ClearCalendar()

    '# - clear events array
    ReDim Events(1 To 1)
    
    '# - remove all controls from fm_days
    Dim element As Control
    For Each element In form_calendar.fm_days.Controls
        form_calendar.fm_days.Controls.Remove element.Name
    Next element
    
End Function


'# - sets the display based on view level
Private Function SetDisplay()

    Dim displayValue As String
    Select Case viewLevel
        Case 1
            displayValue = monthName(Month(viewDate)) & " " & Year(viewDate)
        Case 2
            displayValue = Year(viewDate)
        Case 3
            displayValue = Year(viewDate) - 4 & " - " & Year(viewDate) + 4
    End Select
    
    form_calendar.btn_month_display.Caption = displayValue
    
End Function


Function DisplayMonths()
    
    Call ClearCalendar
    form_calendar.fm_day_names.Visible = False
    viewLevel = 2
    
    Dim btn As MSForms.CommandButton
    Dim counter As Integer
    
    ReDim Events(1)
    counter = 1

    Dim i As Integer
    Dim j As Integer
    For i = 1 To 3
        For j = 1 To 4
            Set btn = CreateButton("month-" & counter)
            Call PositionButton(btn, listOfMonths(counter - 1), j, i, 4, 3)
            counter = counter + 1
            ReDim Preserve Events(1 To counter)
            Set Events(counter).SelectMonth = btn
        Next j
    Next i

    Call SetDisplay

End Function

'# - display
Function DisplayYears()
    viewLevel = 3
    Call ClearCalendar
    Dim yearsDisplay As String
    Dim btn As MSForms.CommandButton
    Dim yearOffset As Integer
        yearOffset = -4

    Dim yearValue As Integer
    Dim counter As Integer
        counter = 1
    
    Dim i As Integer
    Dim j As Integer
    For i = 1 To 3
        For j = 1 To 3
        
            yearValue = Year(viewDate) + yearOffset
            yearOffset = yearOffset + 1
        
            Set btn = CreateButton("year-" & counter)
            Call PositionButton(btn, Str(yearValue), j, i, 3, 3)
            
            ReDim Preserve Events(1 To counter)
            Set Events(counter).SelectYear = btn
            
            counter = counter + 1
            
        Next j
    Next i
    
    Call SetDisplay
    
End Function


'# - sets the viewDate based on selected month
Function ChangeMonthFromSelection(ByVal selectedMonth As String)

    viewLevel = 1
    Dim monthNumber As Integer
    Dim counter As Integer
        counter = 1
        
    Dim monthName As Variant
    For Each monthName In listOfMonths
        If monthName = selectedMonth Then
            monthNumber = counter
            Exit For
        End If
        counter = counter + 1
    Next monthName
            
    viewDate = dateValue(monthNumber & "-" & "1" & "-" & Year(viewDate))
    Call UpdateCalendar
    
End Function


'# - sets the viewDate based on selected year
Function ChangeYearFromSelection(ByVal selectedYear As Integer)

    viewLevel = 2
    viewDate = dateValue("1-1-" & selectedYear)
    Call DisplayMonths
    
End Function


'# - creates and returns a commandbutton
Private Function CreateButton(ByVal btnName As String) As MSForms.CommandButton

    Set CreateButton = form_calendar.fm_days.Controls _
    .Add("Forms.CommandButton.1", btnName, True)
    
End Function


'# - sets the button position
Private Function PositionButton( _
    ByVal btn As MSForms.CommandButton, _
    btnCaption As String, x As Integer, y As Integer, _
    row As Integer, col As Integer)

    Dim btnWidth As Integer
        btnWidth = uiWidth / row
        
    Dim btnHeight As Integer
        btnHeight = uiHeight / col

    With btn
        .Caption = btnCaption
        .Width = btnWidth
        .Height = btnHeight
        .left = x * btnWidth - btnWidth
        .top = y * btnHeight - btnHeight
    End With

End Function


'# - prints the selected date to the home sheet
Function PrintSelectedDate(ByVal dayValue As Integer)
    Dim selectedDay As Date
        selectedDay = DateAdd("d", dayValue - 1, viewDate)
    ThisWorkbook.Worksheets("Home").Range("B6").Value = "Date Picked: " & selectedDay
End Function







