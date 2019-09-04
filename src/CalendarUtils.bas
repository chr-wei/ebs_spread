Attribute VB_Name = "CalendarUtils"
'  This macro collection lets you organize your tasks and schedules
'  for you with the evidence based design (EBS) approach by Joel Spolsky.
'
'  Copyright (C) 2019  Christian Weihsbach
'  This program is free software; you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation; either version 3 of the License, or
'  (at your option) any later version.
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'  You should have received a copy of the GNU General Public License
'  along with this program; if not, write to the Free Software Foundation,
'  Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301  USA
'
'  Christian Weihsbach, weihsbach.c@gmail.com

Option Explicit

Enum DateExtremum
    ceLatest = -1
    ceEarliest = 1
End Enum



Sub Test_GetCalItems()
    Call CalendarUtils.GetCalItems("Me", False)
End Sub



Sub Test_MultiMapHoursToDate()
    Debug.Print ""
    Dim contributor As String: contributor = "Me"

    Dim oItems As Outlook.Items
    Set oItems = CalendarUtils.GetCalItems(contributor, Constants.BUSY_AT_OPTIONAL_APPOINTMENTS)

    Dim multiHours() As Double
    multiHours = Utils.CopyVarArrToDoubleArr(Array(6, 7, 60, 120, 2, 6, 10, 0, 5, 40))

    Dim startingTime As Date
    startingTime = CDate("06/20/2019 07:00")

    Dim hourSetpoints() As Double
    'hourSetpoints = Utils.CopyVarArrToDoubleArr(Array(0, 0, 0))

    Dim dateSetpoints() As Date
    'dateSetpoints = Utils.CopyVarArrToDateArr(Array(startingTime, startingTime, startingTime))

    Dim allRetrievedDates() As Date
    Debug.Print "#### Test with speed improvement ####"

    allRetrievedDates = CalendarUtils.MultiMapHoursToDate(contributor, oItems, _
        multiHours, _
        startingTime, _
        hourSetpoints, dateSetpoints, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))

    Dim idx As Integer
    For idx = 0 To UBound(allRetrievedDates)
        Debug.Print "Mapped hour " & multiHours(idx) & " to " & allRetrievedDates(idx)
    Next idx

    'Now test map without speed improvement
    Dim mapHour As Double
    Dim retrievedDate As Date
    Debug.Print "#### Test without speed improvement ####"

    mapHour = 6
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
        mapHour, _
        startingTime, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 7
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 60
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 120
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 2
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 6
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 10
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 0
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 5
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate

    mapHour = 40
    retrievedDate = CalendarUtils.MapHoursToDate(contributor, oItems, _
    mapHour, _
    startingTime, _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
    SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    Debug.Print "Mapped hour " & mapHour & " to " & retrievedDate
End Sub



Sub Test_MapDateToHours()
    'A test method which maps dates to hours
    'https://docs.microsoft.com/de-de/office/vba/outlook/how-to/search-and-filter/search-the-calendar-for-appointments-that-occur-partially-or-entirely-in-a-given
         
    Dim contributor As String
    contributor = "Me"
    
    Dim oItems As Outlook.Items
    Set oItems = CalendarUtils.GetCalItems(contributor, Constants.BUSY_AT_OPTIONAL_APPOINTMENTS)
    
    Debug.Print "##### Test_MapDateToHours on: " + CStr(Now) + " #####"

    Dim startingDate As Date
    Dim endingDate As Date
    
    startingDate = CDate("06/17/2019 08:00")
    endingDate = CDate("06/18/2019 08:00")
    Debug.Print "Testing 1: " & CStr(startingDate) & " to " & endingDate & " -> " & MapDateToHours(contributor, oItems, _
        endingDate, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset)) 'Result nok.
        
    startingDate = CDate("06/17/2019 08:00")
    endingDate = CDate("06/18/2019 10:00")
    Debug.Print "Testing 2: " & CStr(startingDate) & " to " & endingDate & " -> " & MapDateToHours(contributor, oItems, _
        endingDate, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset)) 'Result ok.
        
    startingDate = CDate("06/17/2019 06:00")
    endingDate = CDate("06/18/2019 10:00")
    Debug.Print "Testing 3: " & CStr(startingDate) & " to " & endingDate & " -> " & MapDateToHours(contributor, oItems, _
        endingDate, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset)) 'Result ok.
        
    startingDate = CDate("06/17/2019 06:00")
    endingDate = CDate("06/19/2019 06:00")
    Debug.Print "Testing 4: " & CStr(startingDate) & " to " & endingDate & " -> " & MapDateToHours(contributor, oItems, _
        endingDate, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset)) 'Result ok.
        
    startingDate = CDate("06/24/2019 06:00")
    endingDate = CDate("06/28/2019 19:00")
    Debug.Print "Testing 5: " & CStr(startingDate) & " to " & endingDate & " -> " & MapDateToHours(contributor, oItems, _
        endingDate, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset)) 'Result ok.
        
    startingDate = CDate("07/15/2019 22:00")
    endingDate = CDate("07/25/2019 00:00")
    Debug.Print "Testing 6: " & CStr(startingDate) & " to " & endingDate & " -> " & MapDateToHours(contributor, oItems, _
        endingDate, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
    
    startingDate = CDate("11/08/2019 17:20")
    endingDate = CDate("14/08/2019 17:20")
    Debug.Print "Testing 7: " & CStr(startingDate) & " to " & endingDate & " -> " & MapDateToHours(contributor, oItems, _
        endingDate, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
End Sub




Function MultiMapHoursToDate( _
    contributor As String, _
    appointmentList As Outlook.Items, _
    inHours() As Double, _
    inStart As Date, _
    Optional inHourSetpoints As Variant, Optional inDateSetpoints As Variant, _
    Optional ByVal appointmentOnsetHours As Double = 0, _
    Optional ByVal appointmentOffsetHours As Double = 0, _
    Optional ByVal maxIterations As Integer = 32767) As Date()
    
    'Maps multi hours to a date regarding the standard outlook calendar. Events, working hours, preparation and post-action times are
    'taken into account. Data fetching approach:
    '  (1) Data from outlook and from 'Settings'sheet
    '  (2) No outlook events but working hours and on-offset hours from the 'Settings' sheet
    '  (3) No outlook events but working hours, and on-offset hours from defined constants in code
    '
    'Input args:
    '   contributor:            The name of the contributor. Needed to get leisure times for the switched 'current'day in the algorithm
    '   appointmentList:        List of appointments coming from outlook
    '   inHours:                Pass an array with remaining hours to map. This is the main input
    '   inHourSetpoints:        Remaining hours for which dates have been mapped already (see inDateSetpoints)
    '   inDateSetpoints:        Mapped dates as result of a previous map call. These setpoints help speeding up the calculation for the passed 'inHours' array
    '   inStart:                The time the mapping is started from
    '   appointmentOnsetHours:  The hours one needs prior to the event to prepare
    '   apptoinmentOffsetHours: The hours one needs after the event to analyze the event
    '
    'Output args:
    '  MultiMapHoursToDate:  Array of dates corresponding to the hours from the input array
    
    'Init output
    Dim allMappedDates() As Date
    MultiMapHoursToDate = allMappedDates
    
    'Check args
    If Not Base.IsArrayAllocated(inHours) Then Exit Function
    
    Dim useSetpoints As Boolean: useSetpoints = False
    Dim bndSetpoints As Integer: bndSetpoints = 0
    If Base.IsArrayAllocated(inHourSetpoints) And StrComp(TypeName(inHourSetpoints), "Double()") = 0 And _
        Base.IsArrayAllocated(inDateSetpoints) And StrComp(TypeName(inDateSetpoints), "Date()") = 0 Then
        'Check if double and dates arrays were passed and initialized
        If (UBound(inHourSetpoints)) = UBound(inDateSetpoints) And _
            (LBound(inHourSetpoints) = LBound(inDateSetpoints)) Then
            'Check if bounds of setpoints are the same
            bndSetpoints = UBound(inHourSetpoints)
            useSetpoints = True
        End If
    End If
    
    Dim mappedDate As Date
    Dim hourIdx As Integer
    
    Dim chosenStart As Date
    Dim associatedRemainingHours As Double
    Dim chosenHours As Double
    
    Dim thisProcessedHours As Variant
    Dim thisMappedDates As Variant
    
    Dim allHourSetpoints() As Double
    Dim allDateSetpoints() As Date
    
    For hourIdx = 0 To UBound(inHours)
        'Map hours to date for every given input
        Dim improveSpeed As Boolean: improveSpeed = False
        
        thisProcessedHours = Base.ExtractSubArray(inHours, LBound(inHours), LBound(inHours) + hourIdx - 1)
        thisMappedDates = allMappedDates 'Use allMappedDates directly without sub array extraction as the allMappedDates-array only contains processed results
            
        If useSetpoints Then
            'If external setpoints are available: Concatenate the internal (the already processed 'inHours' / 'mappedDates') setpoints and
            'external setpoints
            allHourSetpoints = Utils.CopyVarArrToDoubleArr(Base.ConcatToArray(inHourSetpoints, thisProcessedHours))
            allDateSetpoints = Utils.CopyVarArrToDateArr(Base.ConcatToArray(inDateSetpoints, thisMappedDates))
        Else
            'Only use internal setpoints here
            allHourSetpoints = Utils.CopyVarArrToDoubleArr(thisProcessedHours)
            allDateSetpoints = Utils.CopyVarArrToDateArr(thisMappedDates)
        End If
        
        If Base.IsArrayAllocated(allHourSetpoints) And Base.IsArrayAllocated(allDateSetpoints) Then
            'If any setpoint is given (interal or external)
            Call Base.QuickSort(allDateSetpoints, ceDescending, , , allHourSetpoints)
            
            'Now search for a setpoint that has lower 'inHour' value. From this starting point the remaining hours can be mapped to a date
            Dim searchIdx As Integer: searchIdx = 0
            Dim lowerValueFound As Boolean: lowerValueFound = True
            While allHourSetpoints(searchIdx) > inHours(hourIdx) And lowerValueFound
                searchIdx = searchIdx + 1
                If searchIdx < UBound(allHourSetpoints) Then
                    'Even at the end of the array no value could be found which is lower than the hours to map. Search unsuccessful
                    lowerValueFound = False
                End If
            Wend
            
            'Check if the date belonging to the found hour setpoint is not zero. If it is, then the date has not been mapped successfully.
            '(inHour = 0 <=> mappedDate = inStart) therefore 'normal / non improved mode' can be used
            If lowerValueFound And (allDateSetpoints(searchIdx) <> CDate(0)) Then improveSpeed = True
        Else
            improveSpeed = False
        End If
        
        If improveSpeed Then
            'The closest starting time will be taken for the next map: The algorithm either chooses the common starting time given
            '   inStart(n) -> remainingHours(n) => mappedDate(n)
            '
            'OR a setpoint of a later starting time in conjuction with already mapped hours. With the 'later' starting time the amount of remaining
            'hours is reduced -> calculation speed will increase
            '   mappedDate(x) + map([remainingHours(n) - remainingHours(x)]) => mappedDate(n)
            '   Condition: The remaining hours of n are greater than the remaining hours of x. Both n and x have to share a common starting time
            
            chosenStart = allDateSetpoints(searchIdx)
            chosenHours = inHours(hourIdx) - allHourSetpoints(searchIdx)
        Else
            'Just use the normal starting time with the hours you want to map, no speed improvement
            chosenStart = inStart
            chosenHours = inHours(hourIdx)
        End If
        
        'Finally map the (chosen)hours to a date
        mappedDate = CalendarUtils.MapHoursToDate( _
            contributor, _
            appointmentList, _
            chosenHours, _
            chosenStart, _
            appointmentOnsetHours, _
            appointmentOffsetHours, _
            maxIterations)
        
        'Increase the resulting array with every run
        ReDim Preserve allMappedDates(0 To hourIdx)
        allMappedDates(hourIdx) = mappedDate
    Next hourIdx
    
    MultiMapHoursToDate = allMappedDates
End Function



Function MapHoursToDate( _
    contributor As String, _
    appointmentList As Outlook.Items, _
    ByVal remainingHours As Double, _
    ByVal startingTime As Date, _
    Optional ByVal appointmentOnsetHours As Double = 0, _
    Optional ByVal appointmentOffsetHours As Double = 0, _
    Optional ByVal maxIterations As Integer = 32767) As Date
    
    'Maps hours to a date regarding the standard outlook calendar. Events, working hours, preparation and post-action times are
    'taken into account. E.g.:
    'Map(01/01/2019 + 8h) = 02/01/2019
    'Given events in the calendar are assumed to be 'non-working-time'. Together with working hours an end date after a task is finished
    'can be calculated. Steps of the following algorithm can be described as:
    '  (0) Set initial values and sort data structures
    '  (1) While the final value has not been calculated (remainingHours > 0), iterate
    '      (1.1)   Fetch event from calendar, fetch leisure times from calendar ('events' which block you from working after you left work)
    '      (1.2)   Sort the events and blockify them (overlapping events and events too close to each other become a block to make further calculation easier
    '      (1.3)   From the given start time calculate the time difference up to the next coming event block. The time difference is substracted
    '              from remainingHours and eventually leads to exiting the loop
    '  (2) Do post processing of value
    
    'Data fetching approach:
    '  (1) Data from outlook and from 'Settings' sheet
    '  (2) No outlook events but working hours and on-offset hours from the 'Settings'sheet
    '  (3) No outlook events but working hours, and on-offset hours from defined constants in code
    '
    'Input args:
    '  contributor:            The name of the contributor. Needed to get leisure times for the switched 'current'day in the algorithm
    '  appointmentList:        List of appointments coming from outlook
    '  remainingHours:         The hours for which a date shall be calculated
    '  startingTime:           The time from when to calculate
    '  appointmentOnsetHours:  The hours one needs prior to the event to prepare
    '  apptoinmentOffsetHours: The hours one needs after the event to analyze the event
    '  maxIterations:          The iteration count after which calculation stops. For the default 32768 events are examined at max.
    '                          Prohibits while loop to get out of control
    '
    'Output args:
    '  MapHoursToDate:  Date mapped from input time
    
    'Init output
    MapHoursToDate = CDate(0)
    
    'Check args
    If remainingHours < 0 Or _
        appointmentOnsetHours < 0 Or _
        appointmentOffsetHours < 0 Or _
        maxIterations < 1 Then
        Exit Function
    End If
    
    Dim iterIdx As Integer
        
    Dim deltaHoursToBlock As Double
    Dim spanRestriction As String
    Dim sortedList As Outlook.Items
    
    Dim blockEnd As Date
    Dim blockStart As Date
    
    Dim preWorkLeisureStart As Date
    Dim preWorkLeisureEnd As Date
    
    Dim postWorkLeisureStart As Date
    Dim postWorkLeisureEnd As Date
    
    'Use a copy of the passed list to prohibit modification of the passed list
    Set sortedList = appointmentList
    
    'This flag marks if no outlook calendar items are available (anymore). If true expensive and errornous calls to outlook can be omitted.
    Dim noCalItemsFlag As Boolean
    
    If sortedList Is Nothing Then
        noCalItemsFlag = True
    Else
        noCalItemsFlag = False
    End If
    
    If Not noCalItemsFlag Then
        'A sorted list is needed for the algorithm to work
        sortedList.Sort "[Start]"
    End If
    
    'Iterate as long as 'remainingHours' are haven't been decreased to zero.
    iterIdx = 0
    While remainingHours >= 0 And iterIdx < maxIterations
        'Debug info
        'Debug.Print "run: " & iterIdx & "," & Format(startingTime, "mm/dd/yyyy hh:mm AMPM")
        
        'Reset next event block limits
        blockStart = CDate(0)
        blockEnd = CDate(0)
        
        If Not noCalItemsFlag Then
            'Get the next event block from outlook cal.
            Call CalendarUtils.GetNextAppointmentBlock(sortedList, startingTime, _
                blockStart, blockEnd, _
                appointmentOnsetHours, appointmentOffsetHours)
            If blockStart = CDate(0) Or blockEnd = CDate(0) Then
                noCalItemsFlag = True
            End If

        End If
        
        'Take working hours into account. As one cannot add code leisure events to the restricted list the above actions (sorting and restricting)
        'have to be done here again manually.
        Call CalendarUtils.GetLeisureAppointments(startingTime, contributor, _
            preWorkLeisureStart, preWorkLeisureEnd, _
            postWorkLeisureStart, postWorkLeisureEnd)
            
        Call CalendarUtils.FilterAppointmentForMinEnd(startingTime, preWorkLeisureStart, preWorkLeisureEnd) 'see also restriction from above
        Call CalendarUtils.FilterAppointmentForMinEnd(startingTime, postWorkLeisureStart, postWorkLeisureEnd) 'see also restriction from above
        
        'Blockify the outlook and leisure time events or return the earliest of them if they are not overlapping. Question here: What is the next event?
        Call CalendarUtils.MergeOrReturnEarliest(blockStart, blockEnd, preWorkLeisureStart, preWorkLeisureEnd, _
            blockStart, blockEnd) 'combine leisure times with 'normal'events to a block
        Call CalendarUtils.MergeOrReturnEarliest(blockStart, blockEnd, postWorkLeisureStart, postWorkLeisureEnd, _
            blockStart, blockEnd)

        'Start calculating the remaining time left from the captured data
        If blockStart = CDate(0) Then
            'Stop if no next block limit is found. Calculate straight forward from starting time
            MapHoursToDate = startingTime + remainingHours / 24
            Exit Function
        ElseIf Not (blockStart <= startingTime) Then
            'Calculate the remaining time.
            deltaHoursToBlock = (blockStart - startingTime) * 24
            remainingHours = remainingHours - deltaHoursToBlock
        End If
        'If non of the above applied: Skipped remaining time calculation because starting time is in the middle of the appointment block
        'and continue with updated starting time. This happens in first iteration if the starting time was chosen to be during an appointment.
        
        'Update the starting time for the next run (do this in every case)
        startingTime = blockEnd
        iterIdx = iterIdx + 1
    Wend
    
    If iterIdx = maxIterations Then
        'Error reached max iterations
        Exit Function
    End If
    'Remaining hours are only zero or negative here. 'Negative' remaining hours determine the point of time
    'in between two appointment blocks and are 'added'to the next block's start time to give a value prior to the next block's start time.
    MapHoursToDate = blockStart + remainingHours / 24
    
    'Debug info
    'Debug.Print MapHoursToDate
End Function



Function MapDateToHours(contributor As String, appointmentList As Outlook.Items, _
    ByVal endingDate As Date, _
    ByVal startingDate As Date, _
    Optional ByVal appointmentOnsetHours As Double = 0, _
    Optional ByVal appointmentOffsetHours As Double = 0, _
    Optional ByVal maxIterations As Integer = 32767) As Double
    
    'Maps a date to hours regarding the standard outlook calendar. Events, working hours, preparation and post-action times are
    'taken into account. E.g.:
    'Map(01/01/2019 -> 02/01/2019) = 8h (working)
    'Given events in the calendar are assumed to be 'non-working-time'. Starting date together with working hours and end date gives the free working time
    'in between. This algorithm is more or less the inverse of the 'MapHoursToDate'function.
    'Steps of the following algorithm can be described as:
    '  (0) Set initial values and sort data structures
    '  (1) Calculate the time difference of start and and date. It is the absolut maximum time value (if you worked 24/7 and had no events distracting you)
    '  (2) 'Eat up'the calculated max time with events and leisure time (they diminish the max value) as long as there are events distracting you. Iterate:
    '      (1.1)   Fetch event from calendar, fetch leisure times from calendar ('events'which block you from working after you left work)
    '      (1.2)   Sort the events and blockify them (overlapping events and events too close to each other become a block to make further calculation easier
    '      (1.3)   From the given start time calculate the time difference up to the next coming event block. The time difference is substracted
    '              from remainingHours and eventually leads to exiting the loop
    '  (2) Do post processing of value
    
    'Data fetching approach:
    '  (1) Data from outlook and from 'Settings'sheet
    '  (2) No outlook events but working hours and on-offset hours from the 'Settings'sheet
    '  (3) No outlook events but working hours, and on-offset hours from defined constants in code
    '
    'Input args:
    '  contributor:            The name of the contributor. Needed to get leisure times for the switched 'current'day in the algorithm
    '  appointmentList:        List of appointments coming from outlook
    '  remainingHours:         The hours for which a date shall be calculated
    '  startingTime:           The time from when to calculate
    '  appointmentOnsetHours:  The hours one needs prior to the event to prepare
    '  apptoinmentOffsetHours: The hours one needs after the event to analyze the event
    '  maxIterations:          The iteration count after which calculation stops. For the default 32768 events are examined at max.
    '                          Prohibits while loop to get out of control
    '  (arg list is the same as in 'MapHoursToDate'
    '
    'Output args:
    '  LeanMapDateToHours:         Time between to points of time which you can use to work
    
    'Init output
    MapDateToHours = -1
    
    'Check args
    If endingDate < startingDate Or _
        appointmentOnsetHours < 0 Or _
        appointmentOffsetHours < 0 Or _
        maxIterations < 1 Then
        Exit Function
    End If
    
    Dim timeDifference As Double
    
    'Init the timeDifference
    timeDifference = endingDate - startingDate
        
    Dim iterIdx As Integer
    Dim deltaHoursToBlock As Double
    Dim spanRestriction As String
    Dim sortedList As Outlook.Items

    Dim nextBlock As Outlook.AppointmentItem
    Dim preWorkLeisure As Outlook.AppointmentItem
    Dim postWorkLeisure As Outlook.AppointmentItem
    
    Dim blockStart As Date
    Dim blockEnd As Date
    
    Dim preWorkLeisureStart As Date
    Dim preWorkLeisureEnd As Date
    
    Dim postWorkLeisureStart As Date
    Dim postWorkLeisureEnd As Date

    'Use a copy of the passed list to prohibit modification of the passed list
    Set sortedList = appointmentList

    Dim noCalItemsFlag As Boolean
    If sortedList Is Nothing Then
        noCalItemsFlag = True
    Else
        noCalItemsFlag = False
    End If

    'A sorted list is needed for the algorithm to work
    If Not noCalItemsFlag Then
        sortedList.Sort "[Start]"
    End If

    iterIdx = 0
    
    Dim stopFlag As Boolean
    stopFlag = False
    
    Do
        'Loop through all events ending in between the starting and end date - substract the hours from the total difference
        'Debug info
        'Debug.Print "Run of LeanMapDateToHours: " & iterIdx & "," & Format(startingDate, "mm/dd/yyyy hh:mm AMPM")

        'Reset next block
        blockStart = CDate(0)
        blockEnd = CDate(0)

        If noCalItemsFlag Then
            'There are no further outlook events that have to be taken into consideration.
            'Still leisure time (pseudo) events are generated below
        Else
            'Search for the next outlook event block start and end limits
            Call CalendarUtils.GetNextAppointmentBlock(sortedList, startingDate, _
                blockStart, blockEnd, _
                appointmentOnsetHours, appointmentOffsetHours)
        End If

        'Take working hours into account
        Call CalendarUtils.GetLeisureAppointments(startingDate, contributor, preWorkLeisureStart, preWorkLeisureEnd, postWorkLeisureStart, postWorkLeisureEnd)
        Call CalendarUtils.FilterAppointmentForMinEnd(startingDate, preWorkLeisureStart, preWorkLeisureEnd)
        Call CalendarUtils.FilterAppointmentForMinEnd(startingDate, postWorkLeisureStart, postWorkLeisureEnd)
        
        'Blockify the outlook and leisure time events or return the limits of the earliest event if they are not overlapping. Question here: What is the next event?
        Call CalendarUtils.MergeOrReturnEarliest(blockStart, blockEnd, preWorkLeisureStart, preWorkLeisureEnd, blockStart, blockEnd) 'Error here
        Call CalendarUtils.MergeOrReturnEarliest(blockStart, blockEnd, postWorkLeisureStart, postWorkLeisureEnd, blockStart, blockEnd)
        Dim blockTime As Double
        
        'Now calculate how much time the next block 'removes' from the remaining max hours / time difference (the remaining hours decrease every iteration)
        If blockStart <> CDate(0) Then
            If blockEnd > endingDate Then
                If blockStart > endingDate Then
                    'The next found event starts and ends after the ending date so it does not reduce the time difference (it is not taken into account)
                    blockTime = 0
                Else
                    'The given initial endingDate lies within the found event. Use the found event-end as time limit otherwise you would substract a value too high
                    blockTime = endingDate - blockStart
                End If
                
                'Break the loop as we are finished.
                stopFlag = True
            ElseIf blockStart < startingDate Then
                'The starting time intersects with the found block. Only use the time span between starting time and block end as substracted time span.
                '(This should only be the case in first iteration if the passed starting time intersects an event)
                blockTime = blockEnd - startingDate
            Else
                'Substract time of block from total time difference.
                blockTime = blockEnd - blockStart
            End If
            
            'Decrease the searched timeDifference (will be output when no events are left to iterate over)
            timeDifference = timeDifference - blockTime
            
            'Update the starting time for the next run to search for the next events
            startingDate = blockEnd
        End If

        iterIdx = iterIdx + 1
    Loop Until (stopFlag Or timeDifference < 0) And iterIdx <= maxIterations

    If iterIdx = maxIterations Then
        'Error reached max iterations
        Exit Function
    End If
    
    'Limit time difference to zero
    timeDifference = Base.Max(0, timeDifference)
    
    'Rescale date difference to hours
    MapDateToHours = timeDifference * 24
End Function



Function GetNextAppointmentBlock( _
    appointmentList As Outlook.Items, ByVal startTime As Date, _
    ByRef blockStart As Date, ByRef blockEnd As Date, _
    Optional appointmentOnsetHours As Double = 0, _
    Optional appointmentOffsetHours As Double = 0)
    
    'This function returns the limits of the next 'block' of appointments. A block is a bunch of appointments overlapping or nearly overlapping
    'within a threshhold. The threshold is a sum of the time which you need to analyze a previous event + the time you need to prepare for the next event.
    '
    'Input args:
    '   appointmentList:            List of all outlook events. The list has to be sorted according to appointments' start dates
    '   startTime:                  The point of time the next block is searched from
    '   appointmentOnsetHours:      The hours one needs prior to the event to prepare
    '   apptoinmentOffsetHours:     The hours one needs after the event to analyze the event
    '
    'Output args:
    '   blockStart:                 The start date of the block
    '   blockEnd:                   The end date of the block
    
    'Init output
    blockStart = CDate(0)
    blockEnd = CDate(0)
        
    'Check args
    If appointmentList Is Nothing Then
        Exit Function
    End If
        
    'Start block detection: Find event starting after or at the given time
    Dim appointment As Outlook.AppointmentItem
    Dim dFormatString As String
    Dim dSep As String: dSep = Application.International(xlDateSeparator)
    Select Case Application.International(xlDateOrder)
        Case 0 ' = month-day-year
            dFormatString = "mm" & dSep & "dd" & dSep & "yyyy hh:mm AMPM"
        Case 1 ' = day-month-year
            dFormatString = "dd" & dSep & "mm" & dSep & "yyyy hh:mm AMPM"
        Case 2 ' = year-month-day
            dFormatString = "yyyy" & dSep & "mm" & dSep & "dd hh:mm AMPM"
    End Select
    Set appointment = appointmentList.Find("[Start] >= '" & Format(startTime, dFormatString) & "'")
    
    If appointment Is Nothing Then Exit Function
    startTime = appointment.Start - appointmentOnsetHours / 24
    blockStart = startTime
    
    Dim oneStart As Date
    Dim oneEnd As Date
    Dim twoStart As Date
    Dim twoEnd As Date
    
    Dim apptsAreOverlapping As Boolean: apptsAreOverlapping = True
    Do Until Not apptsAreOverlapping
        'Loop while events are overlapping. Call 'FindNext' to get the next event in the row
        oneStart = appointment.Start
        oneEnd = appointment.End
        Set appointment = appointmentList.FindNext ''GetNext' won't work here
        
        If Not appointment Is Nothing Then
            twoStart = appointment.Start
            twoEnd = appointment.End
            apptsAreOverlapping = _
                CalendarUtils.AppointmentsAreOverlapping(oneStart, oneEnd, twoStart, twoEnd, appointmentOnsetHours + appointmentOffsetHours)
        Else
            'Appointments are not overlapping anymore - stop search
            apptsAreOverlapping = False
        End If
    Loop
    
    blockEnd = oneEnd + appointmentOffsetHours / 24
End Function



Function AppointmentsAreOverlapping( _
    ByVal startOne As Date, ByVal endOne As Date, _
    ByVal startTwo As Date, ByVal endTwo As Date, _
    Optional minDeltaHours As Double = 0#) As Boolean
    
    'This function checks if two appointments are overlapping. A threshold can be defined to mark close events as overlapping.
    '
    'Input args:
    '   startOne:       Start time of first appointment
    '   endOne:         End time of first appointment
    '   startTwo:       Start time of second appointment
    '   endTwo:         End time of second appointment
    '   minDeltaHours:  Threshold. Events have at least to have a time span of 'minDeltaHours'in between them to not to be set as overlapping
    '
    'Output args:
    '  AppointmentsAreOverlapping:    True/False
    
    Dim hStart As Date
    Dim hEnd As Date
    
    'Debug info
    'Debug.Print "Checking against: '" + apptOne.Subject + "'and '"; apptTwo.Subject + "'"
    
    If startTwo < startOne Then
        'Swap appointments to make the earlier event the first event. This is the standard for the comparison
        hStart = startTwo
        startTwo = startOne
        startOne = hStart
        
        hEnd = endTwo
        endTwo = endOne
        endOne = hEnd
    End If
    
    If startTwo < endOne + minDeltaHours / 24 Then
        'Check whether the events are overlapping or close together with a certain time threshold. Convert to days
        AppointmentsAreOverlapping = True
    Else
        AppointmentsAreOverlapping = False
    End If
End Function



Function GetWorkingHours(contributor As String, ByVal refDay As Date, ByRef workStart As Date, ByRef workEnd As Date)

    'This function returns the working hours of a day in passed workStart and workEnd variables
    'Init output
    '
    'Input args:
    '  contributor:    The name of the contributor for which the data should be collected
    '  refDay:         The day for which the working hours should be calculated
    '  workStart:      The returned start date
    '  workEnd:        The returned end date
    
    'Init output. Set values to be out of scope of the given ref day
    workStart = refDay - 1
    workEnd = workStart
    
    'Check args
    If StrComp(contributor, "") = 0 Then
        Exit Function
    End If
    
    refDay = CalendarUtils.GetStartOfDay(refDay)
    
    'Outlook data readout for working hours would be very good. Apparently it seems there is no such method to retrieve the values you can set
    'in the outlook settings in VBA. ToDo: Add method, if VBA-api gets updated
    Dim workingWeekdays() As Double
    workingWeekdays = SettingUtils.GetContributorWorkingDays(contributor)
    
    Dim workingDay As Variant
    For Each workingDay In workingWeekdays
        'Cycle through the array of working days e.g. Monday, Tuesday, Wednesday, Thursday, Friday
        If workingDay = Weekday(refDay) Then
            'If the ref day matches the weekday calculate the exact date of the start time regarding the ref day. e.g. 21/01/2019 08:00
            workStart = refDay + SettingUtils.GetContribWorkTimeSetting(contributor, ceWorkingStart)
            workEnd = refDay + SettingUtils.GetContribWorkTimeSetting(contributor, ceWorkingEnd)
            Exit Function
        End If
    Next workingDay
End Function



Function GetLeisureAppointments(ByVal refDay As Date, contributor As String, _
    ByRef preWorkLeisureStart As Date, ByRef preWorkLeisureEnd As Date, _
    ByRef postWorkLeisureStart As Date, ByRef postWorkLeisureEnd As Date)
    
    'This function returns two events' limits to block a day with leisure time to mark on which time spans a contributor is not working:
    '  (1) First event e.g. from 21/01/2019 0:00 to 21/01/2019 08:00 (time prior to start working)
    '  (2) Second event e.g. from 21/01/2019 6:00pm to 22/01/2019 00:00 (time after finishing work)
    '
    'Input args:
    '   refDay:                 The day for which the appointment limits should be calculated. You can pass any clock time, day date counts
    '   contributor:            The contributor whose leisure times shall be read
    '
    'Output args:
    '   preWorkLeisureStart:    The start date of leisure time before work
    '   preWorkLeisureEnd:      The end date of leisure time before work
    '   postWorkLeisureStart:   The start date of leisure time after work
    '   postWorkLeisureEnd:     The end date of leisure time after work
    
    'Init output
    preWorkLeisureStart = CDate(0)
    preWorkLeisureEnd = CDate(0)
    postWorkLeisureStart = CDate(0)
    postWorkLeisureEnd = CDate(0)
        
    'Check args
    'Do not check contributor name here - the standard values can also be returned if contributor name is empty if one whishes to do so.
    Dim workStart As Date
    Dim workEnd As Date
            
    refDay = CalendarUtils.GetStartOfDay(refDay)
    Call CalendarUtils.GetWorkingHours(contributor, refDay, workStart, workEnd)
    
    If workStart = workEnd Then
        'No working hours for the given day received - block the whole day
        preWorkLeisureStart = CalendarUtils.GetStartOfDay(refDay)
        preWorkLeisureEnd = CalendarUtils.GetEndOfDay(refDay)
    Else
        preWorkLeisureStart = refDay
        preWorkLeisureEnd = workStart
        
        postWorkLeisureStart = workEnd
        postWorkLeisureEnd = refDay + 1
    End If
End Function



Function MergeOrReturnEarliest( _
    ByVal apptOneStart As Date, ByVal apptOneEnd As Date, ByVal apptTwoStart As Date, ByVal apptTwoEnd As Date, _
    ByRef apptReturnStart As Date, ByRef apptReturnEnd As Date)

    'This function checks two passed events limits overlap. If the events are not overlapping the limits of the earliest event are returned, otherwise the
    'limits of the whole merged block are returned
    '
    'Input args:
    '   apptOneStart:       The start time of the first event
    '   apptOneEnd:         The end time of the first event
    '   apptTwoStart:       The start time of the second event
    '   apptTwoEnd:         The end time of the second event
    'Output args:
    '   apptReturnStart:    In case of overlapping events: The start date of the block. Otherwise: The start date of the earlier event
    '   apptReturnEnd:      In case of overlapping events: The end date of the block. Otherwise: The end date of the earlier event
    
    'Init output
    apptReturnStart = CDate(0)
    apptReturnEnd = CDate(0)
    
    'Check args. Any date not initialized?
    Dim apptOneInvalid As Boolean: apptOneInvalid = (apptOneStart = CDate(0)) Or (apptOneEnd = CDate(0))
    Dim apptTwoInvalid As Boolean: apptTwoInvalid = (apptTwoStart = CDate(0)) Or (apptTwoEnd = CDate(0))
    
    If apptOneInvalid And apptTwoInvalid Then
        'Input contains two invalid appointments
        Exit Function
    End If
      
    'If one event is invalid return the dates of the other one
    If apptOneInvalid Then
        apptReturnStart = apptTwoStart
        apptReturnEnd = apptTwoEnd
        Exit Function
    ElseIf apptTwoInvalid Then
        apptReturnStart = apptOneStart
        apptReturnEnd = apptOneEnd
        Exit Function
    End If
    
    'Two valid events were passed. Check overlapping and merge if necessary
    Dim earliestStart As Date
    Dim earliestEnd As Date
    Dim latestEnd As Date
    
    earliestStart = Base.Min(apptOneStart, apptTwoStart)
    
    If CalendarUtils.AppointmentsAreOverlapping(apptOneStart, apptOneEnd, apptTwoStart, apptTwoEnd) Then
        latestEnd = Base.Max(apptOneEnd, apptTwoEnd)
        apptReturnStart = earliestStart
        apptReturnEnd = latestEnd
    Else
        earliestEnd = Base.Min(apptOneEnd, apptTwoEnd)
        apptReturnStart = earliestStart
        apptReturnEnd = earliestEnd
    End If
End Function



Function FilterAppointmentForMinEnd(ByVal dateMinLimit As Date, ByRef apptStart As Date, ByRef apptEnd As Date)
    
    'This function returns invalid appointment limits (CDate(0)) if a passed appointment ends prior to a given date
    'Only limits of appointments occuring after the given date or overlapping with it will be returned
    '
    'Input args:
    '   dateMinLimit:   The given date used to filter the appointment
    '   apptStart:      The event start time
    '   apptEnd:        The event end time

    'Check args
    If (apptStart = CDate(0)) Or (apptEnd = CDate(0)) Or (apptEnd <= dateMinLimit) Then
        'Return invalid dates if any date is zero (wrong input)
        'Check functionality: Return invalid dates if the appointment's end is below date limit
        apptStart = CDate(0)
        apptEnd = CDate(0)
    End If
End Function



Function GetStartOfDay(refDay As Date) As Date
    'Set the reference day date to date DAY 0:00
    GetStartOfDay = DateSerial(Year(refDay), Month(refDay), day(refDay))
End Function



Function GetEndOfDay(refDay As Date) As Date
    'Set the reference day date to date DAY+1 0:00
    GetEndOfDay = CalendarUtils.GetStartOfDay(refDay) + 1
End Function



Function GetCalItems(contributor As String, useOptionalAppts As Boolean) As Outlook.Items
    'This function returns calendar items of a contributor if a folder id was specified. These appointments are filtered by busy status and category.
    'Recurrent appointments are included
    '
    'Input args:
    '  contributor:            The contributor whose events / calendar items shall be fetched
    '  useOptionalAppts:       Select if optional appointments are relevant as well
    '
    'Output args:
    '  GetCalItems:            The list of filtered cal items
    
    'Init output
    Set GetCalItems = Nothing
    
    'Check args
    If StrComp(contributor, "") = 0 Then Exit Function

    Dim filteredItems As Outlook.Items
    Dim cal As Outlook.Folder
    Dim restriction As String
    restriction = ""
    
    'Read the calendar id of the current contributor. This seems to be more robust than reading the mail address, receiving the outlook recipient and
    'reading the shared calendar folder afterwards.
    
    Dim calId As String
    Dim storId As String
    calId = SettingUtils.GetContributorCalIdSetting(contributor, storId)
    
    'Catch error reading calendar id
    On Error Resume Next
    If (StrComp(calId, "") <> 0) And StrComp(storId, "") <> 0 Then
        'Use cal id and store id to read calendar
        Set cal = Outlook.Session.GetFolderFromID(calId, storId)
    ElseIf (StrComp(calId, "") <> 0) Then
        'Use only cal id to read calendar
        Set cal = Outlook.Session.GetFolderFromID(calId)
    Else
        Exit Function
    End If
    If Err.Number <> 0 Then Debug.Print "Error reading calendar of cal id: " & calId
    Err.Clear
    On Error GoTo 0


    If Not cal Is Nothing Then
        'The calendar folder could be resolved. now read items
        Set filteredItems = cal.Items
        
        'Set restriction to get all appointments that make the user busy
        restriction = "[BusyStatus] = '" + CStr(olBusy) + "'" 'user is busy"
        restriction = restriction + " OR " + "[BusyStatus] = '" + CStr(olOutOfOffice) + "'" 'user is out of office
            
        If useOptionalAppts Then
            restriction = restriction + " OR " + "[BusyStatus] = '" + CStr(olTentative) + "'" 'user tentative
        Else
            'Do not add any further restrictions here.
        End If
        
        'Set restriction to not use appointments with a specificied 'blocker category'(marks that the appointment is used to get work done)
        Dim blockerCat As String
        blockerCat = SettingUtils.GetContributorGetWorkDoneCat(contributor)
        
        If StrComp(blockerCat, "") <> 0 Then
            restriction = "(" + restriction + ")" + " AND NOT [Categories] = '" + blockerCat + "'"
        End If
        
        filteredItems.IncludeRecurrences = True
        
        'Filter the items according to their restrictions
        Set filteredItems = filteredItems.Restrict(restriction)
        Set GetCalItems = filteredItems
    Else
        Set GetCalItems = Nothing
    End If
    
    'Debug info
    'Debug.Print "Count of filtered items is: " & filteredItems.Count
End Function



Function GetSelectedCalendarId(Optional ByRef storId As String, Optional ByRef folderPath As String) As String
    'Source: https://stackoverflow.com/questions/48789601/how-to-find-calendar-id
    
    'This function returns the id of a selected calendar folder as well as its path. Please do select the folder inside the explorer view
    '(calendar sheet with appointments) and not the list of calendars
    
    'Init output
    Dim id As String: id = ""
    folderPath = ""
    
    Dim fold As Outlook.Folder
    Set fold = Outlook.Application.ActiveExplorer.CurrentFolder
    
    If Not fold Is Nothing Then
        If fold.DefaultItemType = OlItemType.olAppointmentItem Then
            'Check if folder holds appointment items (one can also select mail folders etc.)
            id = fold.EntryID
            GetSelectedCalendarId = id
            storId = fold.storeId
            folderPath = fold.folderPath
        End If
    End If
    
    'Debug info
    'Debug.Print "Folder ID is: " & id
    'Debug.Print "Folder path is: " & folderPath
End Function

