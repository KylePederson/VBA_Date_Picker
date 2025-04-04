VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents SelectDate As MSForms.CommandButton
Attribute SelectDate.VB_VarHelpID = -1

Sub SelectDate_Click()
    Call form_calendar.DatePicker.PrintSelectedDate(SelectDate.Caption)
End Sub



