Attribute VB_Name = "EngrService"
'
' Release: 1.0
'
' Description:
'   This file contains the General declarations
Option Explicit


Public sgStartIn As String

'Merge Information
Public sgMergedLastDateRun As String    'Last Date the merger was checked
Public sgMergeLastTimeRun As String     'Last Time the merge was checked
Public sgMergedNextDateRun As String    'Next Date the the merge is to be checked
Public sgMergeNextTimeRun As String     'Next time that the merge is to be checked

Public sgSchdNextDateRun As String      'Next Date that Schedule is to be generated or the words 'After Automation'
Public sgSchdNextTimeRun As String      'Next time that Schedule is to be generated or the words 'After Automation'
Public sgSchdForDates As String         'Schedule Date(s) to be generated (1/2, 1/3)
Public sgSchdPurgeNextDateRun As String 'Next Date (or 'After Schedule') that Schedule and 'As Aired' is to be purged
Public sgSchdPurgeNextTimeRun As String 'Next Time (or 'After Schedule') that Schedule and 'As Aired' is to be purged

Public sgAutoNextDateRun As String      'Next Date that Automation is to be run or the words 'After Schedule'
Public sgAutoNextTimeRun As String      'Next time that Automation is to be run or the words 'After Schedule'
Public sgAutoForDates As String         'Automation Date(s) to be generated (1/2, 1/3)
Public sgAutoPurgeNextDateRun As String 'Next Date (or 'After Automation') that Schedule and 'As Aired' is to be purged
Public sgAutoPurgeNextTimeRun As String 'Next Time (or 'After Automation') that Schedule and 'As Aired' is to be purged

Public sgPurgeDate As String        'Purge Schedule and 'As Aired" prior to date specified

