Attribute VB_Name = "Constants"
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

'Debugging flag: Disables error catching
Public Const DEBUGGING_MODE As Double = False



'Planning sheet constants
Public Const PLANNING_SHEET_NAME As String = "Planning"
Public Const TASK_LIST_NAME As String = "TaskOverviewList"

Public Const TASK_ENTRY_HEADER As String = "Task no."
Public Const TASK_NAME_HEADER As String = "Task name"
Public Const INDICATOR_HEADER As String = "Indicator"
Public Const TASK_PRIORITY_HEADER As String = "Priority"
Public Const DUE_DATE_HEADER As String = "Due date"
Public Const TASK_FINISHED_ON_HEADER As String = "Finished on"
Public Const KANBAN_LIST_HEADER As String = "Kanban list"
Public Const COMMENT_HEADER As String = "Comment"
Public Const T_HASH_HEADER As String = "tHASH"

Public Const TASK_NAME_INITIAL As String = "<ENTER_NAME>"
Public Const TASK_ESTIMATE_INITIAL As String = "<ENTER_ESTIMATE>"
Public Const TASK_PRIO_INITIAL As Long = (2 ^ 15) - 1
Public Const CONTRIBUTOR_INITIAL As String = "Me"

Public Const KANBAN_LIST_TODO As String = "To do"
Public Const KANBAN_LIST_IN_PROGRESS As String = "In progress"
Public Const KANBAN_LIST_DONE As String = "Done"

Public Const TIME_ENTRY_HEADER As String = "Time entry no."
Public Const START_TIME_HEADER As String = "Start time"
Public Const END_TIME_HEADER As String = "End time"

Public Const TAG_REGEX As String = "%Tag *"
Public Const EBS_COLUMN_REGEX As String = "^EBS (\d\d)% ((time)|(date))$"



'Task sheet constants
Public Const TASK_SHEET_TEMPLATE_NAME As String = "Task sheet template"
Public Const TASK_SHEET_TIME_LIST_IDX As Integer = 1
Public Const TASK_SHEET_EBS_LIST_IDX As Integer = 2

Public Const VELOCITY_HEADER As String = "Current Velocity"
Public Const COMPARISON_ENTRY_HEADER As String = "Comparison entry no."
Public Const SINGLE_SUPPORT_POINT_HEADER As String = "Propability support point"
Public Const EBS_SELF_TIME_HEADER As String = "EBS self time"
Public Const ESTIMATE_QUOTIENT_HEADER As String = "EBS self time / Total time"
Public Const TIME_DELTA_HEADER As String = "Time delta in h"
Public Const TASK_SHEET_ACTION_ONE_HEADER As String = "Action1"
Public Const TASK_SHEET_ACTION_TWOO_HEADER As String = "Action2"
Public Const TASK_SHEET_ACTION_ONE_NAME As String = "SetPlainDelta()"
Public Const TASK_SHEET_ACTION_TWOO_NAME As String = "SetCalendarDelta()"

'EBS sheet
Public Const EBS_SHEET_PREFIX As String = "Sheduling"
Public Const EBS_SHEET_REGEX As String = EBS_SHEET_PREFIX + " \(.+\)"
Public Const EBS_SHEET_TEMPLATE_NAME As String = EBS_SHEET_PREFIX + " template"
Public Const EBS_MAIN_LIST_IDX As Integer = 1
Public Const EBS_RUNDATA_LIST_IDX As Integer = 2

Public Const EBS_VELO_CHART_INDEX = 1
Public Const EBS_PROP_CHART_INDEX = 2

Public Const EBS_PROP_CHART_MODE_HEADER As String = "Select ebs runs to show"
Public Const EBS_PROP_CHART_SCALING_HEADER As String = "Scaling"



'EBS main list
Public Const EBS_ENTRY_HEADER As String = "EBS entry no."
Public Const EBS_VELOCITY_POOL_HEADER As String = "[Velocity pool]"
Public Const EBS_RUN_DATE_HEADER As String = "EBS run date"
Public Const EBS_SHOW_POOL_HEADER As String = "Show pool?"
Public Const EBS_TIME_ESTIMATES_HEADER As String = "[Project time estimates]"
Public Const EBS_DATE_ESTIMATES_HEADER As String = "[Project date estimates]"
Public Const EBS_HASH_HEADER As String = "eHASH"

Public Const EBS_SHOW_POOL_TRUE As String = "Yes"
Public Const EBS_SHOW_POOL_FALSE As String = "No"

Public Const EBS_UPPER_RND_VELOCITY_LIMIT As Double = 0.2 'limit chosen to generate a uniform distribution for histo lower limit 0.2, upper limit 5.0, bar count = 8
Public Const EBS_LOWER_RND_VELOCITY_LIMIT As Double = 5 'limit chosen to generate a uniform distribution for histo lower limit 0.2, upper limit 5.0, bar count = 8

Public Const EBS_LOWER_HISTOGRAM_LIMIT As Double = 0.2  '20% estimate
Public Const EBS_UPPER_HISTOGRAM_LIMIT As Double = 5    '500% estimate
Public Const EBS_HISTOGRAM_BAR_COUNT As Integer = 9

Public Const EBS_TRACK_ENTRIES_DELTA_DAYS As Double = 2 'Days to wait before fetching finished tasks data for a new velocity pool.



'EBS run data list
Public Const EBS_RUNDATA_T_HASH_HEADER As String = "tHASH"
Public Const EBS_RUNDATA_PRIORITY_HEADER As String = "Logged Priority"
Public Const EBS_RUNDATA_TIME_SPENT_HEADER As String = "Time spent so far"
Public Const EBS_RUNDATA_ESTIMATES_POOL_HEADER As String = "[Monte Carlo estimate pool]"
Public Const EBS_RUNDATA_REMAINING_TIME_POOL_HEADER As String = "[Remaining time pool]"
Public Const EBS_RUNDATA_ACCUMULATED_TIME_POOL_HEADER As String = "[Accumulated time pool]"
Public Const EBS_RUNDATA_INTERPOLATED_TIME_HEADER As String = "[Interpolated time estimates]"
Public Const EBS_RUNDATA_INTERPOLATED_DATES_HEADER As String = "[Interpolated date estimates]"



'EBS propability constants
Public Const EBS_GENERATE_RND_VELOCITIES As Boolean = False
Public Const EBS_VELOCITY_POOL_SIZE As Integer = 200
Public Const EBS_VELOCITY_PICKS As Integer = 50



'Calendar constants
Public Const APPOINTMENT_ONSET_HOURS As Double = 1 / 4 '15min
Public Const APPOINTMENT_OFFSET_HOURS As Double = 1 / 12 '5min
Public Const WORKING_HOURS_START As Date = "08:00"
Public Const WORKING_HOURS_END As Date = "17:00"
Public Const STANDARD_GETTING_WORK_DONE_CATEGORY As String = "[EBS.Blocker]"
Public Const BUSY_AT_OPTIONAL_APPOINTMENTS As Boolean = False



'Shared constants
Public Const INDICATOR As String = "<current"
Public Const CONTRIBUTOR_HEADER As String = "Contributor"
Public Const TASK_TOTAL_TIME_HEADER As String = "Total time spent in h"
Public Const TASK_ESTIMATE_HEADER As String = "User time estimate"
Public Const EBS_SUPPORT_POINT_HEADER As String = "[Propability support points]"
Public Const TOTAL_TIME_HEADER As String = "Total time"

Public Const N_A As String = "N/A"
Public Const COUNTED_ENTRIES_FORMAT = "000000"
Public Const SERIALIZED_ARRAY_REGEX As String = "{*}"
Public Const INVALID_ENTRY_PLACEHOLDER As String = "<INVALID_ENTRY>"

Public Const HASH_REGEXP As String = "[\w]{18}"



'Settings constants
Public Const CONTRIBUTOR_LIST_NAME As String = "ContributorSettings"
Public Const SETTING_SHEET_NAME As String = "Settings"
Public Const MAIL_HEADER As String = "Mail"
Public Const CAL_FOLDER_HEADER As String = "CalendarFolder"
Public Const WORKING_DAYS_HEADER As String = "WorkingDays"
Public Const WORKING_HOURS_START_HEADER As String = "WorkingHoursStart"
Public Const WORKING_HOURS_END_HEADER As String = "WorkingHoursEnd"
Public Const GETTING_WORK_DONE_CAT_HEADER As String = "ExcludeCategories"
Public Const APPT_ONSET_HEADER As String = "AppointmentOnsetHours"
Public Const APPT_OFFSET_HEADER As String = "AppointmentOffsetHours"

Public Const SETTINGS_HIGHLIGHT_COLOR_HEADER As String = "HighlightColor"
Public Const SETTINGS_COMMON_COLOR_HEADER As String = "CommonColor"
Public Const SETTINGS_LIGHT_COLOR_HEADER As String = "LightColor"

Public Const CAL_ID_HEADER As String = "Calendar id"
Public Const CAL_PATH_HEADER As String = "Calendar Path"



'Array constants

'Support points for saved propabilities
Function EBS_SUPPORT_PROPABILITIES() As Double()
    EBS_SUPPORT_PROPABILITIES = Utils.CopyVarArrToDoubleArr(Array(0.05, 0.2, 0.35, 0.5, 0.65, 0.8, 0.95)) 'Array(0.1, 0.3, 0.5, 0.7, 0.9)
End Function



Function EBS_HIGHLIGHT_PROPABILITIES() As Double()
    EBS_HIGHLIGHT_PROPABILITIES = Utils.CopyVarArrToDoubleArr(Array(0.15, 0.5, 0.85))
End Function



Function STANDARD_WORKING_DAYS() As Double()
    'Days starting at sunday
    STANDARD_WORKING_DAYS = Utils.CopyVarArrToDoubleArr(Array(2, 3, 4, 5, 6))
End Function
