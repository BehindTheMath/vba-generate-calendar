Attribute VB_Name = "GenerateCalendar"
'********************************************
' (c) 2016 Behind The Math
' Licensed under the MIT License
'********************************************

Option Explicit

Public MaxEvents As Integer, EventsSheet As Worksheet, CalendarSheet As Worksheet, RecurringSheet As Worksheet, NumWeeks As Integer, LastDay As Integer

' Change these column numbers to reflect the Events data table
Const EventNameColumn As Integer = 1, EventStartDateColumn As Integer = 3, EventStartTimeColumn As Integer = 5
Const EventEndDateColumn As Integer = 7, EventEndTimeColumn As Integer = 9, EventDurationColumn As Integer = 11, RecurringColumn As Integer = 12

Sub GenerateCalendar()
    Dim FirstDayOfWeek As Integer, ThisMonth As Integer, ThisYear As Integer, DayOfWeekCounter As Integer, DateCounter As Integer, EventListRowCounter As Integer
    Dim x As Integer, TopRow As Integer
    Dim EventData As Variant
    Dim StartDay As Date
    Dim DaysEvents As Collection, Events As New Collection
    
    Set EventsSheet = Worksheets("Events")
    Set CalendarSheet = Worksheets("Calendar")
    Set RecurringSheet = Worksheets("Recurring")
    ThisYear = Year(EventsSheet.Cells(2, EventStartDateColumn))
    ThisMonth = Month(EventsSheet.Cells(2, EventStartDateColumn))
    StartDay = DateSerial(ThisYear, ThisMonth, 1)
    NumWeeks = 0
    
    ' Unprotect sheet if it had a previous calendar to prevent errors.
    CalendarSheet.Protect Contents:=False
    ' Prevent screen from flashing while drawing the calendar.
    Application.ScreenUpdating = False
    
    ' Clear any previous data.
    CalendarSheet.Cells.Clear
    
    ' Setup the headers
    SetupHeaders StartDay
    
    ' Get on which day of the week the month starts.
    FirstDayOfWeek = Weekday(StartDay)
    ' Get the last date of the month
    LastDay = Day(DateSerial(ThisYear, ThisMonth + 1, 1) - 1)
    DateCounter = 1
    TopRow = 3
    
    ' If there are recurring events
    If EventsSheet.Cells(2, RecurringColumn).End(xlDown) <> vbNullString Then
        ParseRecurring
        Set Events = LoadEvents(Worksheets("Recurring"))
        Worksheets("Recurring").Cells.Clear
    Else
        Set Events = LoadEvents(EventsSheet)
    End If
    
    Do
        For DayOfWeekCounter = FirstDayOfWeek To 7
            ' Write the dates
            With CalendarSheet.Cells(TopRow, DayOfWeekCounter)
                .Value = DateCounter
                .Font.Size = 12
                .Font.Bold = True
                .RowHeight = 20
                .HorizontalAlignment = xlRight
                .IndentLevel = 1
            End With
            
            ' Write events
            Set DaysEvents = Nothing
            ' Get this day's events (if there are any)
            On Error Resume Next
            Set DaysEvents = Events(Str(DateCounter))
            On Error GoTo 0
            ' If there are events on this day
            If Not DaysEvents Is Nothing Then
                EventListRowCounter = 0
                ' Go through this day's events and write them
                For Each EventData In DaysEvents
                    EventListRowCounter = EventListRowCounter + 1
                    CalendarSheet.Cells(TopRow + EventListRowCounter, DayOfWeekCounter) = EventData
                Next EventData
            End If
            
            DateCounter = DateCounter + 1
            ' If we reached the end of the month, stop.
            If DateCounter > LastDay Then
                NumWeeks = NumWeeks + 1
                Exit Do
            End If
        Next DayOfWeekCounter
        
        NumWeeks = NumWeeks + 1
        FirstDayOfWeek = 1
        TopRow = TopRow + MaxEvents + 1
    Loop
    
    ' Set row height
    For x = 1 To NumWeeks
        CalendarSheet.Range(CalendarSheet.Cells(3 + x + MaxEvents * (x - 1), 1), CalendarSheet.Cells(3 + MaxEvents * x + (x - 1), 1)).RowHeight = 15
    Next x
    
    DrawBorders
    
    ' Set the print area
    SetupPage
    
    ' Turn off gridlines.
    ActiveWindow.DisplayGridlines = False
    ' Protect sheet to prevent overwriting the dates.
    CalendarSheet.Protect Contents:=True, UserInterfaceOnly:=True

    ' Resize window to show all of calendar (may have to be adjusted
    ' for video configuration).
    ActiveWindow.WindowState = xlMaximized
    ActiveWindow.ScrollRow = 1

    ' Allow screen to redraw with calendar showing.
    Application.ScreenUpdating = True
    
    Set Events = Nothing: Set EventsSheet = Nothing: Set CalendarSheet = Nothing: Set DaysEvents = Nothing: Set RecurringSheet = Nothing
End Sub

Sub SetupHeaders(ByVal StartDay As Date)
    Dim x As Integer

    ' Create the month and year title.
    With CalendarSheet.Range("A1:G1")
        .Merge
        .Value = Format(StartDay, "mmmm yyyy")
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .RowHeight = 35
        .NumberFormat = "mmmm yyyy"
    End With
    
    ' Format A2:G2 for the days of week labels.
    With CalendarSheet.Range("A2:G2")
        .HorizontalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = True
        .RowHeight = 20
        .ColumnWidth = 35
    End With
    ' Write days of week in A2:G2.
    For x = 1 To 7
        CalendarSheet.Cells(2, x) = WeekdayName(x)
    Next x
End Sub

Sub ParseRecurring()
    Dim CurDate As Integer, LastRow As Integer, OriginalLastRow As Integer, CurOriginalRow As Integer, DateCounter As Integer
    
    LastRow = GetLastRow(EventsSheet)
    OriginalLastRow = LastRow

    With RecurringSheet
        ' Clear any old data from the Recurring sheet
        .Cells.Clear
        ' Copy the data from the Events sheet to the Recurring sheet so we can manipulate it without affecting the original data
        EventsSheet.Range("A1", .Cells(OriginalLastRow, RecurringColumn).Address).Copy .Range("A1")
        
        ' For each row of the original data
        For CurOriginalRow = 1 To OriginalLastRow
            ' If this event is recurring
            If .Cells(CurOriginalRow, RecurringColumn) <> vbNullString Then
                ' Get the date of the original event
                CurDate = Day(.Cells(CurOriginalRow, EventStartDateColumn))
                ' What is the frequency that it recurs
                Select Case LCase(.Cells(CurOriginalRow, RecurringColumn))
                    Case "daily"
                        ' For each subsequent day
                        For DateCounter = CurDate To LastDay
                            ' Copy the data
                            .Range(.Cells(CurOriginalRow, 1), .Cells(CurOriginalRow, RecurringColumn - 1)).Copy .Cells(LastRow + DateCounter - CurDate + 1, 1)
                            'Update the day to the new day
                            .Cells(LastRow + (DateCounter - CurDate) + 1, EventStartDateColumn) = .Cells(CurOriginalRow, EventStartDateColumn) + (DateCounter - CurDate) + 1
                            .Cells(LastRow + (DateCounter - CurDate) + 1, EventEndDateColumn) = .Cells(CurOriginalRow, EventEndDateColumn) + (DateCounter - CurDate) + 1
                        Next DateCounter
                        LastRow = LastRow + DateCounter - CurDate - 1
                    Case "weekly"
                        ' If there are more dates to recur on
                        If LastDay - CurDate >= 7 Then
                            ' For each week
                            For DateCounter = 7 To LastDay - CurDate Step 7
                                ' Copy the data
                                .Range(.Cells(CurOriginalRow, 1), .Cells(CurOriginalRow, RecurringColumn - 1)).Copy .Cells(LastRow + (DateCounter / 7), 1)
                                'Update the day to the new day
                                .Cells(LastRow + (DateCounter / 7), EventStartDateColumn) = .Cells(CurOriginalRow, EventStartDateColumn) + DateCounter
                                .Cells(LastRow + (DateCounter / 7), EventEndDateColumn) = .Cells(CurOriginalRow, EventEndDateColumn) + DateCounter
                            Next DateCounter
                            LastRow = LastRow + ((DateCounter - 7) / 7)
                        End If
                End Select
            End If
        Next CurOriginalRow
    End With
End Sub

Function LoadEvents(ByRef sheet As Worksheet) As Collection
    Dim RowCounter As Integer, CurDate As Integer, CurMonth As Integer
    Dim EventData As String, LastDate As String, EventDuration As String
    Dim MonthsEvents As New Collection

    SortEvents sheet:=sheet
    RowCounter = 2
    CurDate = Day(sheet.Cells(RowCounter, EventStartDateColumn))
    CurMonth = Month(sheet.Cells(RowCounter, EventStartDateColumn))
    LastDate = "0"

    Do While sheet.Cells(RowCounter, EventStartDateColumn) <> vbNullString
        ' If the next event is from a different month, stop
        If Month(sheet.Cells(RowCounter, EventStartDateColumn)) <> CurMonth Then Exit Do
        
        ' Get the next event
        EventDuration = Format(sheet.Cells(RowCounter, EventDurationColumn), "h:mm")
        ' Formula for calculating duration:
        'EventDuration = Int(DateDiff("n", Sheet.Cells(RowCounter, EventStartTimeColumn), Sheet.Cells(RowCounter, EventEndTimeColumn)) / 60)
        'EventDuration = Event Duration & ":" & DateDiff("n", Sheet.Cells(RowCounter, EventStartTimeColumn), Sheet.Cells(RowCounter, EventEndTimeColumn)) Mod 60
        EventData = sheet.Cells(RowCounter, EventNameColumn) & ": " & Format(sheet.Cells(RowCounter, EventStartTimeColumn), "h:mm AMPM") & " - "
        EventData = EventData & Format(sheet.Cells(RowCounter, EventEndTimeColumn), "h:mm AMPM") & " (" & EventDuration & ")"

        If LastDate <> Str(CurDate) Then
            LastDate = Str(CurDate)
            MonthsEvents.Add New Collection, LastDate
        End If
        
        MonthsEvents(LastDate).Add EventData
        If MonthsEvents(LastDate).Count > MaxEvents Then MaxEvents = MonthsEvents(LastDate).Count

        ' Advance to the next row
        RowCounter = RowCounter + 1
        CurDate = Day(sheet.Cells(RowCounter, EventStartDateColumn))
    Loop
    
    Set LoadEvents = MonthsEvents
    
    Set MonthsEvents = Nothing
End Function

Sub SortEvents(ByRef sheet As Worksheet)
    With sheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sheet.Columns(EventStartDateColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=sheet.Columns(EventStartTimeColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=sheet.Columns(EventEndDateColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sheet.Range(sheet.Cells(1, 1), sheet.Cells(GetLastRow(sheet), RecurringColumn))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub DrawBorders()
    Dim x As Integer
    
    ' Draw outside and vertical borders
    With CalendarSheet.Range(CalendarSheet.Cells(1, 1), CalendarSheet.Cells(2 + NumWeeks * (MaxEvents + 1), 7))
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlRight).Weight = xlThick
        .Borders(xlRight).ColorIndex = xlAutomatic
    End With
    With CalendarSheet.Range(CalendarSheet.Cells(1, 1), CalendarSheet.Cells(2 + NumWeeks * (MaxEvents + 1), 1))
        .Borders(xlLeft).Weight = xlThick
        .Borders(xlLeft).ColorIndex = xlAutomatic
    End With
    
    ' Draw border above the weekday names
    With CalendarSheet.Range("A2:G2")
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
    End With
    
    ' Draw borders above and below the dates
    For x = 1 To NumWeeks
        With CalendarSheet.Range(CalendarSheet.Cells(3 + ((MaxEvents + 1) * (x - 1)), 1), CalendarSheet.Cells(3 + ((MaxEvents + 1) * (x - 1)), 7))
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
            .Borders(xlEdgeBottom).Weight = xlThick
            .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        End With
    Next x
End Sub

Sub SetupPage()
    Worksheets("Calendar").Select
    ' Switch to Page Break Preview mode
    ActiveWindow.View = xlPageBreakPreview
    With CalendarSheet
        ' Remove old page breaks
        .ResetAllPageBreaks
        ' Set landscape
        .PageSetup.Orientation = xlLandscape
        ' Set page area
        .PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(2 + NumWeeks * (MaxEvents + 1), 7)).Address
        ' Move page breaks if necessary
        If .VPageBreaks.Count Then .VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
        If .HPageBreaks.Count Then Set .HPageBreaks(1).Location = .Range("$A$53")
    End With
    ' Switch back to Normal View
    ActiveWindow.View = xlNormalView
End Sub

Function GetLastRow(ByRef sheet As Worksheet) As Integer
    ' Refresh UsedRange
    'sheet.UsedRange
    'GetLastRow = sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row
    GetLastRow = sheet.Cells(1, EventNameColumn).CurrentRegion.Rows.Count
End Function