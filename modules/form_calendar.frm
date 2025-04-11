VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_calendar 
   Caption         =   "Date Picker"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3495
   OleObjectBlob   =   "form_calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatePicker As Calendar

Private Sub btn_month_display_Click()
    
    Select Case DatePicker.viewLevel
        Case 1
            Call DatePicker.DisplayMonths
        Case 2
            Call DatePicker.DisplayYears
        Case 3
            Exit Sub
    End Select
    
End Sub

Private Sub UserForm_Initialize()
    Set DatePicker = New Calendar
End Sub

Private Sub btn_back_Click()
    'Call DatePicker.DecreaseViewDate
    Call DatePicker.ChangeViewDate(-1)
End Sub

Private Sub btn_forward_Click()
    'Call DatePicker.IncreaseViewDate
    Call DatePicker.ChangeViewDate(1)
End Sub







