Attribute VB_Name = "GenerateCalendar"
'********************************************
' (c) 2016 Behind The Math
' Licensed under the MIT License
'********************************************

Option Explicit

Public MaxEvents As Integer, EventsSheet As Worksheet, CalendarSheet As Worksheet, NumWeeks As Integer

' Change these column numbers to reflect the Events data table
Const EventNameColumn As Integer = 1, EventDateColumn As Integer = 3, EventStartTimeColumn As Integer = 5, EventEndTimeColumn As Integer = 9, EventDurationColumn As Integer = 11

Sub GenerateCalendar()
    Dim FirstDayOfWeek As Integer, ThisMonth As Integer, ThisYear As Integer, DayOfWeekCounter As Integer, DateCounter As Integer, EventListRowCounter As Integer
    Dim x As Integer, TopRow As Integer, LastDay As Integer
    Dim EventData As Variant
    Dim StartDay As Date
    Dim DaysEvents As Collection, Events As New Collection
    
    Set EventsSheet = Worksheets("Events")
    Set CalendarSheet = Worksheets("Calendar")
    ThisYear = Year(EventsSheet.Cells(2, EventDateColumn))
    ThisMonth = Month(EventsSheet.Cells(2, EventDateColumn))
    StartDay = DateSerial(ThisYear, ThisMonth, 1)
    NumWeeks = 0
    
    ' Unprotect sheet if it had a previous calendar to prevent errors.
    CalendarSheet.Protect Contents:=False
    ' Prevent screen from flashing while drawing the calendar.
    Application.ScreenUpdating = False
    
    ' Clear any previous data.
    CalendarSheet.Cells.Clear
    
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

    ' Get on which day of the week the month starts.
    FirstDayOfWeek = Weekday(StartDay)
    ' Get the last date of the month
    LastDay = Day(DateSerial(ThisYear, ThisMonth + 1, 1) - 1)
    DateCounter = 1
    TopRow = 3
    
    Set Events = LoadEvents
    
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
        Range(Cells(3 + x + MaxEvents * (x - 1), 1), Cells(3 + MaxEvents * x + (x - 1), 1)).RowHeight = 15
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
    
    Set Events = Nothing: Set EventsSheet = Nothing: Set CalendarSheet = Nothing: Set DaysEvents = Nothing
End Sub

Function LoadEvents() As Collection
    Dim RowCounter As Integer, CurDate As Integer, CurMonth As Integer
    Dim EventData As String, LastDate As String, EventDuration As String
    Dim MonthsEvents As New Collection, DaysEvents As Collection

    SortEvents
    RowCounter = 2
    CurDate = Day(EventsSheet.Cells(RowCounter, EventDateColumn))
    CurMonth = Month(EventsSheet.Cells(RowCounter, EventDateColumn))
    Set DaysEvents = New Collection
    LastDate = "0"

    Do While EventsSheet.Cells(RowCounter, EventDateColumn) <> ""
        ' If the next event is from a different month, stop
        If Month(EventsSheet.Cells(RowCounter, EventDateColumn)) <> CurMonth Then Exit Do
        
        ' Get the next event
        EventDuration = Format(EventsSheet.Cells(RowCounter, EventDurationColumn), "h:mm")
        ' Formula for calculating duration:
        'EventDuration = Int(DateDiff("n", EventsSheet.Cells(RowCounter, EventStartTimeColumn), EventsSheet.Cells(RowCounter, EventEndTimeColumn)) / 60)
        'EventDuration = Event Duration & ":" & DateDiff("n", EventsSheet.Cells(RowCounter, EventStartTimeColumn), EventsSheet.Cells(RowCounter, EventEndTimeColumn)) Mod 60
        EventData = EventsSheet.Cells(RowCounter, EventNameColumn) & ": " & Format(EventsSheet.Cells(RowCounter, EventStartTimeColumn), "h:mm AMPM") & " - "
        EventData = EventData & Format(EventsSheet.Cells(RowCounter, EventEndTimeColumn), "h:mm AMPM") & " (" & EventDuration & ")"

        If LastDate <> Str(CurDate) Then
            LastDate = Str(CurDate)
            MonthsEvents.Add New Collection, LastDate
        End If
        
        MonthsEvents(LastDate).Add EventData
        If MonthsEvents(LastDate).Count > MaxEvents Then MaxEvents = MonthsEvents(LastDate).Count

        ' Advance to the next row
        RowCounter = RowCounter + 1
        CurDate = Day(EventsSheet.Cells(RowCounter, EventDateColumn))
    Loop
    
    Set LoadEvents = MonthsEvents
    
    Set MonthsEvents = Nothing: Set DaysEvents = Nothing
End Function

Sub SortEvents()
    EventsSheet.Sort.SortFields.Clear
    EventsSheet.Sort.SortFields.Add Key:=Columns(EventDateColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    EventsSheet.Sort.SortFields.Add Key:=Columns(EventStartTimeColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    EventsSheet.Sort.SortFields.Add Key:=Columns(EventEndTimeColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With EventsSheet.Sort
        .SetRange Range(Cells(1, 1), Cells(getLastRow(ActiveSheet), 11))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Function getLastRow(sheet As Worksheet) As Integer
    ' Refresh UsedRange
    'sheet.UsedRange
    getLastRow = sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).Row
End Function

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
        If .HPageBreaks.Count Then Set .HPageBreaks(1).Location = Range("$A$53")
    End With
    ActiveWindow.View = xlNormalView
End Sub
