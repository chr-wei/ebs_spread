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



Sub Test_MapHoursToDate()
    'A test method which maps hours to a dates.
    'https://docs.microsoft.com/de-de/office/vba/outlook/how-to/search-and-filter/search-the-calendar-for-appointments-that-occur-partially-or-entirely-in-a-given

    Dim contributor As String
    contributor = "Me"
    
    Dim oItems As Outlook.Items
    Set oItems = CalendarUtils.GetCalItems(contributor, Constants.BUSY_AT_OPTIONAL_APPOINTMENTS)
    
    Debug.Print "##### Test_MapHoursToDate on: " + CStr(Now) + " #####"
    
    Dim startingDate As Date
    Dim hours As Double
        
    hours = 1
    startingDate = CDate("06/20/2019 07:00")
    Debug.Print "Testing 5: " & startingDate & " +" & hours & "h -> " & MapHoursToDate(contributor, oItems, _
        hours, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
        
    startingDate = CDate("06/20/2019 08:00")
    
    Debug.Print "Testing 1: " & startingDate; " +" & hours & "h -> " & MapHoursToDate(contributor, oItems, _
        hours, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))

    startingDate = CDate("06/20/2019 09:00")
    Debug.Print "Testing 2: " & startingDate & " +" & hours & "h -> " & MapHoursToDate(contributor, oItems, _
        hours, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))

    startingDate = CDate("06/20/2019 10:00")
    Debug.Print "Testing 3: " & startingDate & " +" & hours & "h -> " & MapHoursToDate(contributor, oItems, _
        hours, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))

    startingDate = CDate("06/20/2019 11:00")
    Debug.Print "Testing 4: " & startingDate & " +" & hours & "h -> " & MapHoursToDate(contributor, oItems, _
        hours, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))

    startingDate = CDate("06/20/2019 19:00")
    Debug.Print "Testing 6: " & startingDate & " +" & hours & "h -> " & MapHoursToDate(contributor, oItems, _
        hours, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))

    startingDate = CDate("06/21/2019 19:00")
    Debug.Print "Testing 7: " & startingDate & " +" & hours & "h -> " & MapHoursToDate(contributor, oItems, _
        hours, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
        
    startingDate = CDate("06/17/2019 08:00")
    hours = 40
    Debug.Print "Testing 8: " & startingDate & " +" & hours & "h -> " & MapHoursToDate(contributor, oItems, _
        hours, startingDate, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
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
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset)) 'Result ok.
        
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
    multiRemainingHours() As Double, _
    startingTime As Date, _
    Optional ByVal appointmentOnsetHours As Double = 0, _
    Optional ByVal appointmentOffsetHours As Double = 0, _
    Optional ByVal maxIterations As Integer = 32767) As Date()
    
    'Maps multi hours to a date regarding the standard outlook calendar. Events, working hours, preparation and post-action times are
    'taken into account. Data fetching approach:
    '  (1) Data from outlook and from 'Settings'sheet
    '  (2) No outlook events but working hours and on- offset hours from the 'Settings'sheet
    '  (3) No outlook events but working hours, and on-offset hours from defined constants in code
    '
    'Input args:
    '  contributor:            The name of the contributor. Needed to get leisure times for the switched 'current'day in the algorithm
    '  appointmentList:        List of appointments coming from outlook
    '  multieRemainingHours:   Has to be sorted ascending as the algorithm starts with the last (lower) value und only does calculation
    '                          on the hours left
    '  appointmentOnsetHours:  The hours one needs prior to the event to prepare
    '  apptoinmentOffsetHours: The hours one needs after the event to analyze the event
    '  (arg list is the same as in 'MapHoursToDate'
    '
    'Output args:
    '  MultiMapHoursToDate:  Array of dates corresponding to the hours from the input array
    
    'Init output
    Dim allMappedDates() As Date
    MultiMapHoursToDate = allMappedDates
    
    'Check args
    If Not Base.IsArrayAllocated(multiRemainingHours) Then Exit Function
    
    ReDim allMappedDates(UBound(multiRemainingHours))
    
    Dim previousRemainingHours As Double
    Dim remainingHours As Double
    Dim mappedDate As Date
    
    'Init remainingHours and starting time
    previousRemainingHours = 0
    startingTime = Now
    
    Dim hourIdx As Integer
    
    For hourIdx = 0 To UBound(multiRemainingHours)
        'Cycle through all passed hours and calc a date for each of them
        
        'Read the value out of the array
        remainingHours = multiRemainingHours(hourIdx)
        
        mappedDate = CalendarUtils.MapHoursToDate( _
            contributor, _
            appointmentList, _
            remainingHours - previousRemainingHours, _
            startingTime, _
            appointmentOnsetHours, _
            appointmentOffsetHours, _
            maxIterations)

        allMappedDates(hourIdx) = mappedDate
        
        'Prepare variables for next run
        previousRemainingHours = remainingHours
        startingTime = mappedDate
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
    '  MapHoursToDate:  Date mapped from input time
    
    'Init output
    Dim endDate As Date
    MapHoursToDate = startingTime - 1
    
    'Check args
    If remainingHours < 0 Or _
        appointmentOnsetHours < 0 Or _
        appointmentOffsetHours < 0 Or _
        maxIterations < 1 Then
        Exit Function
    End If
    
    Dim iterIdx As Integer
    Dim nextBlock As Outlook.AppointmentItem
    Dim deltaHoursToBlock As Double
    Dim spanRestriction As String
    Dim restrictedList As Outlook.Items
    
    Dim preWorkLeisure As Outlook.AppointmentItem
    Dim postWorkLeisure As Outlook.AppointmentItem
    
    'Use a copy of the passed list to prohibit modification of the passed list
    Set restrictedList = appointmentList
    
    'This flag marks if no outlook calendar items are available (anymore). If true expensive and errornous calls to outlook can be omitted.
    Dim noCalItemsFlag As Boolean
    
    If restrictedList Is Nothing Then
        noCalItemsFlag = True
    Else
        noCalItemsFlag = False
    End If
    
    If Not noCalItemsFlag Then
        'A sorted list is needed for the algorithm to work
        restrictedList.Sort "[Start]"
    End If
    
    'Iterate as long as 'remainingHours'have got shrinked to zero.
    iterIdx = 0
    While remainingHours >= 0 And iterIdx < maxIterations
        'Debug info
        'Debug.Print "run: " & iterIdx & "," & Format(startingTime, "mm/dd/yyyy hh:mm AMPM")
        
        'Reset next event block
        Set nextBlock = Nothing
        
        If Not noCalItemsFlag Then
            'Only use appointments that have an ending date after the (iteration) starting date. The date changes every loop cycle.
            'If the list is already empty further restriction is causing an error.
            spanRestriction = "[End] > '" & Format(startingTime, "mm/dd/yyyy hh:mm AMPM") & "'"
            Set restrictedList = restrictedList.Restrict(spanRestriction)
            
            If restrictedList.Count = 0 Then
                'Problems can occur if the resticted list item count is zero and restriction is applied again in the next iteration.
                'Prohibit this with setting a flag.
                noCalItemsFlag = True
            Else
                'Get the next event block from outlook cal.
                Set nextBlock = CalendarUtils.GetNextAppointmentBlock(restrictedList, appointmentOnsetHours, appointmentOffsetHours)
                        
                'Debug info
                'If Not blockAppointment Is Nothing Then
                '   blockAppointment.Subject = "created block"
                '   blockAppointment.Save
                'End If
            End If
        End If
        
        'Take working hours into account. As one cannot add code leisure events to the restricted list the above actions (sorting and restricting)
        'have to be done here again manually.
        Call CalendarUtils.GetLeisureAppointments(startingTime, contributor, preWorkLeisure, postWorkLeisure)
        Set preWorkLeisure = CalendarUtils.FilterAppointmentForMinEnd(preWorkLeisure, startingTime) 'see also restriction from above
        Set postWorkLeisure = CalendarUtils.FilterAppointmentForMinEnd(postWorkLeisure, startingTime) 'see also restriction from above
        
        'Blockify the outlook and leisure time events or return the earliest of them if they are not overlapping. Question here: What is the next event?
        Set nextBlock = CalendarUtils.MergeOrReturnEarliest(nextBlock, preWorkLeisure) 'combine leisure times with 'normal'events to a block
        Set nextBlock = CalendarUtils.MergeOrReturnEarliest(nextBlock, postWorkLeisure) 'combine leisure times with 'normal'events to a block
        
        'Start calculating the remaining time left from the captured data
        If nextBlock Is Nothing Then
            'Stop if no next block is found. Calculate straight forward from starting time
            MapHoursToDate = startingTime + remainingHours / 24
            Exit Function
        ElseIf Not (nextBlock.Start <= startingTime) Then
            'Calculate the remaining time.
            deltaHoursToBlock = (nextBlock.Start - startingTime) * 24
            remainingHours = remainingHours - deltaHoursToBlock
        End If
        'If non of the above applied: Skipped remaining time calculation because starting time is in the middle of the appointment block
        'and continue with updated starting time. This happens in first iteration if the starting time was chosen to be during an appointment.
        
        'Update the starting time for the next run (do this in every case)
        startingTime = nextBlock.End
        iterIdx = iterIdx + 1
    Wend
    
    If iterIdx = maxIterations Then
        'Error reached max iterations
        Exit Function
    End If
    'Remaining hours are only zero or negative here. 'Negative'remaining hours determine the point of time
    'in between two appointment blocks and are 'added'to the next block's start time to give a value prior to the next block's start time.
    MapHoursToDate = nextBlock.Start + remainingHours / 24
    
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
    '  MapDateToHours:         Time between to points of time which you can use to work
    
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
    Dim nextBlock As Outlook.AppointmentItem
    Dim deltaHoursToBlock As Double
    Dim spanRestriction As String
    Dim restrictedList As Outlook.Items

    Dim preWorkLeisure As Outlook.AppointmentItem
    Dim postWorkLeisure As Outlook.AppointmentItem

    'Use a copy of the passed list to prohibit modification of the passed list
    Set restrictedList = appointmentList

    Dim noCalItemsFlag As Boolean
    If restrictedList Is Nothing Then
        noCalItemsFlag = True
    Else
        noCalItemsFlag = False
    End If

    'A sorted list is needed for the algorithm to work
    If Not noCalItemsFlag Then
        restrictedList.Sort "[Start]"
    End If

    iterIdx = 0
    
    Dim stopFlag As Boolean
    stopFlag = False
    
    Do
        'Loop through all events ending in between the starting and end date - substract the hours from the total difference
        'Debug info
        'Debug.Print "Run of MapDateToHours: " & iterIdx & "," & Format(startingDate, "mm/dd/yyyy hh:mm AMPM")

        'Reset next block
        Set nextBlock = Nothing

        If noCalItemsFlag Then
            'There are no further outlook events that have to be taken into consideration.
            'Still leisure time (pseudo) events are generated below
        Else
            'Only use appointments that have an ending date after the (iteration) starting date. The date changes every loop cycle
            spanRestriction = "[End] > '" & Format(startingDate, "mm/dd/yyyy hh:mm AMPM") & "'"
            Set restrictedList = restrictedList.Restrict(spanRestriction)
            If restrictedList.Count = 0 Then
                'Problems can occur if the resticted list item count is zero and restriction is applied again. Prohibit this with setting the flag
                noCalItemsFlag = True
            Else
                Set nextBlock = CalendarUtils.GetNextAppointmentBlock(restrictedList, appointmentOnsetHours, appointmentOffsetHours)
                'Debug info
                'If Not blockAppointment Is Nothing Then
                '   blockAppointment.Subject = "created block"
                '   blockAppointment.Save
                'End If
            End If
        End If

        'Take working hours into account
        Call CalendarUtils.GetLeisureAppointments(startingDate, contributor, preWorkLeisure, postWorkLeisure)
        Set preWorkLeisure = CalendarUtils.FilterAppointmentForMinEnd(preWorkLeisure, startingDate)
        Set postWorkLeisure = CalendarUtils.FilterAppointmentForMinEnd(postWorkLeisure, startingDate)
        
        'Blockify the outlook and leisure time events or return the earliest of them if they are not overlapping. Question here: What is the next event?
        Set nextBlock = CalendarUtils.MergeOrReturnEarliest(nextBlock, preWorkLeisure)
        Set nextBlock = CalendarUtils.MergeOrReturnEarliest(nextBlock, postWorkLeisure)

        Dim blockTime As Double
        
        'Now calculate how much time the next block 'removes'from the remaining max hours / time difference (the remaining hours decrease every iteration)
        If Not nextBlock Is Nothing Then
            If nextBlock.End > endingDate Then
                If nextBlock.Start > endingDate Then
                    'The next found event starts and ends after the ending date so it does not reduce the time difference (it is not taken into account)
                    blockTime = 0
                Else
                    'The given initial endingDate lies within the found event. Use the found event-end as time limit otherwise you would substract a value too high
                    blockTime = endingDate - nextBlock.Start
                End If
                
                'Break the loop as we are finished.
                stopFlag = True
            ElseIf nextBlock.Start < startingDate Then
                'The starting time intersects with the found block. Only use the time span between starting time and block end as substracted time span.
                '(This should only be the case in first iteration if the passed starting time intersects an event)
                blockTime = nextBlock.End - startingDate
            Else
                'Substract time of block from total time difference.
                blockTime = nextBlock.End - nextBlock.Start
            End If
            
            'Decrease the searched timeDifference (will be output when no events are left to iterate over)
            timeDifference = timeDifference - blockTime
            
            'Update the starting time for the next run to search for the next events
            startingDate = nextBlock.End
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
    appointmentList As Outlook.Items, _
    Optional appointmentOnsetHours As Double = 0, _
    Optional appointmentOffsetHours As Double = 0) As Outlook.AppointmentItem
    
    'This function returns the next 'block'of appointments. A block is a bunch of appointments overlapping or nearly overlapping
    'with a threshhold. The threshold is a sum of the time which you need to analyze a previous event + the time you need to prepare for the next event.
    '
    'Input args:
    '  appointmentList:            List of all outlook events
    '  appointmentOnsetHours:      The hours one needs prior to the event to prepare
    '  apptoinmentOffsetHours:     The hours one needs after the event to analyze the event
    '
    'Output args:
    '  GetNextAppointmentBlock:    The 'blockified'bunch of events which are very close to each other or overlapping
    
    'Init output - init start before end otherwise errors will occur
    Set GetNextAppointmentBlock = Nothing
    Dim block As Outlook.AppointmentItem
        
    'Check args
    If appointmentList Is Nothing Then
        Exit Function
    End If
        
    'Start block detection
    Dim appointment As Outlook.AppointmentItem
            
    For Each appointment In appointmentList
        If block Is Nothing Then
            'Init the block in first loop run
            Set block = Outlook.CreateItem(olAppointmentItem)
            Set block = CalendarUtils.ChangeAppointmentTimeFixedDates(block, appointment.Start, appointment.End)
            'Set GetNextAppointmentBlock = block
        Else
            'Debug info
            'Debug.Print "Checking block against: '" + appointment.Subject + "'"
            
            If CalendarUtils.AppointmentsAreOverlapping(block, appointment, appointmentOnsetHours + appointmentOffsetHours) Then
                'Let the block grow with every match
                If block.End < appointment.End Then
                    'Update the block end time if the next item has a later ending time
                    block.End = appointment.End
                End If

            Else
                'The events are not overlapping. Stop cycling through them
                'Set GetNextAppointmentBlock = block
                GoTo t3aGetNextAppointmentBlockEnd
            End If
        End If
    Next appointment
    
t3aGetNextAppointmentBlockEnd:
    If Not block Is Nothing Then
        'Add the onset and offset time values (prepare time and follow-up time)
        'block.Start = block.Start - appointmentOnsetHours / 24
        'block.End = block.End + appointmentOffsetHours / 24
        'Set GetNextAppointmentBlock = block
        Set GetNextAppointmentBlock = ChangeAppointmentTimeFixedDates(block, block.Start - appointmentOnsetHours / 24, _
            block.End + appointmentOffsetHours / 24)
    End If
End Function



Function AppointmentsAreOverlapping(apptOne As Outlook.AppointmentItem, _
    apptTwo As Outlook.AppointmentItem, _
    Optional minDeltaHours As Double = 0#) As Boolean
    
    'This function tells you if two appointments are overlapping.
    '
    'Input args:
    '  apptOne:        firstAppointment
    '  apptTwo:        secondAppointment
    '  minDeltaHours:  Threshold. Events have at least to have a time span of 'minDeltaHours'in between them to not to be set as overlapping
    '
    'Output args:
    '  AppointmentsAreOverlapping:    True/False
    
    Dim hAppt As Outlook.AppointmentItem
    
    'Debug info
    'Debug.Print "Checking against: '" + apptOne.Subject + "'and '"; apptTwo.Subject + "'"
    
    If apptTwo.Start < apptOne.Start Then
        'Swap appointments to make apptOne the first event. This is the standard for the comparison
        Set hAppt = apptTwo
        Set apptTwo = apptOne
        Set apptOne = hAppt
    End If
    
    If apptTwo.Start < apptOne.End + minDeltaHours / 24 Then
        'Check whether the events are overlapping or close together with a certain time threshold. Convert to days
        AppointmentsAreOverlapping = True
    Else
        AppointmentsAreOverlapping = False
    End If
End Function



Function ChangeAppointmentTimeFixedDates(appt As Outlook.AppointmentItem, newStart As Date, newEnd As Date) As Outlook.AppointmentItem
    'This function sets fixed dates to a start and end point in time. Changing the start time also changes the end time of an event normally.
    'To avoid this call this function
    
    Dim changedAppt As Outlook.AppointmentItem

    Set changedAppt = appt
    changedAppt.Start = newStart
    changedAppt.End = newEnd
    
    Set ChangeAppointmentTimeFixedDates = changedAppt

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
    ByRef preWorkLeisure As Outlook.AppointmentItem, _
    ByRef postWorkLeisure As Outlook.AppointmentItem)
    
    'This function returns two events to block a day with leisure time to mark on which time spans a contributor is not working:
    '  (1) First event e.g. from 21/01/2019 0:00 to 21/01/2019 08:00 (time prior to start working)
    '  (2) Second event e.g. from 21/01/2019 6:00pm to 22/01/2019 00:00 (time after finishing work)
    '
    'Input args:
    '  refDay:             The day for which the appointment items should be calculated. You can pass any clock time, day date counts
    '  preWorkLeisure:     The returned pre work blocker appointment
    '  postWorkLeisure:    The returned post work blocker appointment
    
    'Init output
    Set preWorkLeisure = Outlook.CreateItem(olAppointmentItem)
    Set postWorkLeisure = Outlook.CreateItem(olAppointmentItem)
    
    'Check args
    'Do not check contributor name here - the standard values can also be returned if contributor name is empty if one whishes to do so.
    Dim workStart As Date
    Dim workEnd As Date
            
    refDay = CalendarUtils.GetStartOfDay(refDay)
    Call CalendarUtils.GetWorkingHours(contributor, refDay, workStart, workEnd)
    
    If workStart = workEnd Then
        'No working hours for the given day received - block the whole day
        Set preWorkLeisure = CalendarUtils.ChangeAppointmentTimeFixedDates(preWorkLeisure, _
            CalendarUtils.GetStartOfDay(refDay), _
            CalendarUtils.GetEndOfDay(refDay))
        Set postWorkLeisure = Nothing
    Else
        Set preWorkLeisure = CalendarUtils.ChangeAppointmentTimeFixedDates(preWorkLeisure, refDay, workStart)
        Set postWorkLeisure = CalendarUtils.ChangeAppointmentTimeFixedDates(postWorkLeisure, workEnd, refDay + 1)
    End If
End Function



Function MergeOrReturnEarliest(ByVal apptOne As Outlook.AppointmentItem, ByVal apptTwo As Outlook.AppointmentItem) As Outlook.AppointmentItem

    'This function checks two passed events for overlapping. If they are not overlapping the earliest event is returned, otherwise the whole
    'merged block is returned
    '
    'Input args:
    '  refDay:             The day for which the appointment items should be calculated. You can pass any clock time, day date counts
    '  preWorkLeisure:     The returned pre work blocker appointment
    '  postWorkLeisure:    The returned post work blocker appointment
    
    'Init output
    Set MergeOrReturnEarliest = Nothing
    
    'Check args
    If apptOne Is Nothing And apptTwo Is Nothing Then
        Exit Function
    End If
    
    Dim mergedEarliest As Outlook.AppointmentItem
    Set mergedEarliest = Outlook.CreateItem(olAppointmentItem)
        
    'If one event is nothing return the dates of the other one
    If apptOne Is Nothing Then
        Set MergeOrReturnEarliest = CalendarUtils.ChangeAppointmentTimeFixedDates(mergedEarliest, apptTwo.Start, apptTwo.End)
        Exit Function
    ElseIf apptTwo Is Nothing Then
        Set MergeOrReturnEarliest = CalendarUtils.ChangeAppointmentTimeFixedDates(mergedEarliest, apptOne.Start, apptOne.End)
        Exit Function
    End If
    
    'Two valid events were passed. Check overlapping and merge if necessary
    Dim earliestStart As Date
    Dim earliestEnd As Date
    Dim latestEnd As Date
    
    earliestStart = CalendarUtils.GetDateExtremum(apptOne.Start, apptTwo.Start, ceEarliest)
    
    If CalendarUtils.AppointmentsAreOverlapping(apptOne, apptTwo) Then
        latestEnd = CalendarUtils.GetDateExtremum(apptOne.End, apptTwo.End, ceLatest)
        Set mergedEarliest = CalendarUtils.ChangeAppointmentTimeFixedDates(mergedEarliest, earliestStart, latestEnd)
    Else
        earliestEnd = CalendarUtils.GetDateExtremum(apptOne.End, apptTwo.End, ceEarliest)
        Set mergedEarliest = CalendarUtils.ChangeAppointmentTimeFixedDates(mergedEarliest, earliestStart, earliestEnd)
    End If
    
    Set MergeOrReturnEarliest = mergedEarliest
End Function



Function GetDateExtremum(dateOne As Date, dateTwo As Date, Optional ceExtremum As DateExtremum) As Date
    
    'This function either returns the latest or earliest of the passed dates according to the selected extremum enum switch
    '
    'Input args:
    '  dateOne:
    '  dateTwo:
    '  ceExtremum: The enum switch so select the earliest or the latest date
    
    Select Case ceExtremum
        Case DateExtremum.ceEarliest
            GetDateExtremum = CDate(Base.Min(CDbl(dateOne), CDbl(dateTwo)))
        Case DateExtremum.ceLatest
            GetDateExtremum = CDate(Base.Max(CDbl(dateOne), CDbl(dateTwo)))
    End Select
    
    'working but ugly. ToDo: Delete if working
    '   If CDbl(dateOne) * ceExtremum > CDbl(dateTwo) * ceExtremum Then
    '       'Flip direction with enum var
    '       GetDateExtremum = dateTwo
    '   Else
    '       GetDateExtremum = dateOne
    '   End If
End Function



Function FilterAppointmentForMinEnd(ByVal appt As Outlook.AppointmentItem, _
    ByVal dateMinLimit As Date) As Outlook.AppointmentItem
    
    'This function returns an empty item if a passed appointment ends prior to a given date (only events occuring after the given date or overlapping
    'with it)
    '
    'Input args:
    '  appt:           The event
    '  dateMinLimit:   The given date used to filter
    
    'Init output
    Set FilterAppointmentForMinEnd = Nothing
    
    'Check args
    If appt Is Nothing Then
        Exit Function
    End If
    
    'If the appointment is below date limit - reject it
    If appt.End <= dateMinLimit Then
        Set FilterAppointmentForMinEnd = Nothing
    Else
        Set FilterAppointmentForMinEnd = appt
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

    'This function returns calendar items to a given mail address. These appointments are filtered by busy status and category
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
    
    Dim oNs As Outlook.Namespace
    Dim filteredItems As Outlook.Items
    Dim cal As Outlook.Folder
    Dim recp As Outlook.Recipient
    Dim restriction As String
    restriction = ""
    
    'Read the mail address of the current contributor
    Dim mail As String
    mail = SettingUtils.GetContributorMailSetting(contributor)
    
    'Try to resolve the recipient first
    Dim recpIsResolved As Boolean
    recpIsResolved = False
    
    If Not recpIsResolved And StrComp(mail, "") <> 0 Then
        Set recp = Outlook.Session.CreateRecipient(mail)
        recp.resolve
        recpIsResolved = recp.Resolved
    End If
    
    If Not recpIsResolved And StrComp(contributor, "") <> 0 Then
        'If mail is not resolved try to resolve contributor data via contributor name
        Set recp = Outlook.Session.CreateRecipient(contributor)
        recp.resolve
        recpIsResolved = recp.Resolved
    End If

    If recpIsResolved Then
        'The account of the contributor could be retrieved. Now collect items with multiple approaches
        
        Dim calendarIsResolved As Boolean: calendarIsResolved = False
        
        'Try to get default shared folder (error based approach)
        On Error Resume Next
        If Not calendarIsResolved Then
            Set cal = Outlook.Application.Session.GetSharedDefaultFolder(recp, olFolderCalendar)
            If Err.Number = 0 Then calendarIsResolved = True
            Err.Clear
        End If
        On Error GoTo 0

        If Not calendarIsResolved And CalendarUtils.IsFolderAvailable(oNs, recp.name) Then
            'Check folders of resolved recipient manually if 'GetSharedDefaultFolder' failed.
            
            Dim recpFolder As Outlook.Folder
            Set recpFolder = oNs.Folders(recp.name)
            
            Dim calName As String
            calName = SettingUtils.GetContributorCalIdSetting(contributor)
            If Not calendarIsResolved And CalendarUtils.IsFolderAvailable(recpFolder, calName) Then
                Set cal = recpFolder.Folders.item(calName)
                calendarIsResolved = True
            End If
        End If
        
        If calendarIsResolved Then
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
            
            'Filter the items according to their restrictions
            Set filteredItems = filteredItems.Restrict(restriction)
            Set GetCalItems = filteredItems
        Else
            Set GetCalItems = Nothing
        End If
    Else
        Set GetCalItems = Nothing
    End If
    
    'Debug info
    'Debug.Print "Count of filtered items is: " & filteredItems.Count
End Function



Function IsFolderAvailable(parent As Variant, inName As String) As Boolean
    'Check if outlook folders of an object are available or not by searching for their name
    '
    'Input args:
    '   parent: Can be Outlook.Folder or Outlook.Namespace
    
    'Init output
    IsFolderAvailable = False
    
    'Check input
    If parent Is Nothing Then Exit Function
    
    'Get the folders from parent
    Dim allCals As Outlook.Folders
    Set allCals = parent.Folders
    
    Dim cal As Outlook.Folder
    
    For Each cal In allCals
        'Search for the name
        If StrComp(cal.name, inName) = 0 Then
            IsFolderAvailable = True
            Exit Function
        End If
    Next cal
End Function



Function GetSelectedCalendarId(Optional ByRef folderPath As String) As String
    'Source: https://stackoverflow.com/questions/48789601/how-to-find-calendar-id
    
    'This function returns the id of a selected calendar folder as well as its path. Please do select the folder inside the explorer view
    '(calendar sheet with appointments) and not the list of calendars
    
    'Init output
    Dim id As String: id = ""
    folderPath = ""
    
    Dim fold As Outlook.Folder
    Set fold = Outlook.ActiveExplorer.CurrentFolder
    
    If Not fold Is Nothing Then
        If fold.DefaultItemType = OlItemType.olAppointmentItem Then
            'Check if folder holds appointment items (one can also select mail folders etc.)
            id = fold.EntryID
            GetSelectedCalendarId = id
            folderPath = fold.folderPath
        End If
    End If
    
    'Debug info
    'Debug.Print "Folder ID is: " & id
    'Debug.Print "Folder path is: " & folderPath
End Function
