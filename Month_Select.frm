VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Month_Select 
   Caption         =   "Choose your path."
   ClientHeight    =   9000.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   OleObjectBlob   =   "Month_Select.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Month_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub

Sub OptimizeCode_End()

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.EnableEvents = EventState
Application.ScreenUpdating = True

End Sub
Private Sub CancelButt_Click()

Unload Me
End

End Sub

Private Sub SubmitButton_Click()
Dim summWB As Workbook

Set summWB = ActiveWorkbook

Call OptimizeCode_Begin
Application.DisplayAlerts = False

With summWB

If BoxYear.Value <> "" Then Sheets("Inspections").Cells(1, "AB").Value = BoxYear.Value


If JanButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
'This number format thing below formats the cell as text so one digit months have the preceding "0" so the macro can use the two digit month format to find the right file
.NumberFormat = "@"
.Value = "01"
End With
End If
If JanButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "D"
If JanButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "Jan"

If FebButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
.NumberFormat = "@"
.Value = "02"
End With
End If
If FebButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "E"
If FebButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "Feb"

If MarchButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
.NumberFormat = "@"
.Value = "03"
End With
End If
If MarchButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "F"
If MarchButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "March"

If AprilButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
.NumberFormat = "@"
.Value = "04"
End With
End If
If AprilButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "G"
If AprilButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "April"

If MayButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
.NumberFormat = "@"
.Value = "05"
End With
End If
If MayButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "H"
If MayButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "May"

If JuneButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
.NumberFormat = "@"
.Value = "06"
End With
End If
If JuneButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "I"
If JuneButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "June"

If JulyButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
.NumberFormat = "@"
.Value = "07"
End With
End If
If JulyButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "J"
If JulyButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "July"

If AugButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
.NumberFormat = "@"
.Value = "08"
End With
End If
If AugButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "K"
If AugButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "Aug"

If SeptButton = True Then
With summWB.Sheets("Inspections").Cells(1, "AC")
.NumberFormat = "@"
.Value = "09"
End With
End If
If SeptButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "L"
If SeptButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "Sept"

If OctButton = True Then Sheets("Inspections").Cells(1, "AC").Value = "10"
If OctButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "M"
If OctButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "Oct"

If NovButton = True Then Sheets("Inspections").Cells(1, "AC").Value = "11"
If NovButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "N"
If NovButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "Nov"

If DecButton = True Then Sheets("Inspections").Cells(1, "AC").Value = "12"
If DecButton = True Then Sheets("Inspections").Cells(1, "AG").Value = "O"
If DecButton = True Then Sheets("Inspections").Cells(1, "AE").Value = "Dec"

If InspButton = True Then Sheets("Inspections").Cells(1, "AD").Value = "Inspections"

If VehButton = True Then Sheets("Inspections").Cells(1, "AD").Value = "Vehicles"

If CitButton = True Then Sheets("Inspections").Cells(1, "AD").Value = "Citations"

If AllButton = True Then Sheets("Inspections").Cells(1, "AF").Value = "ALL SECTIONS"

If NButton = True Then Sheets("Inspections").Cells(1, "AF").Value = "NORTH"

If SButton = True Then Sheets("Inspections").Cells(1, "AF").Value = "SOUTH"

If BButton = True Then Sheets("Inspections").Cells(1, "AF").Value = "BORDER"

If SDBUtton = True Then Sheets("Inspections").Cells(1, "AF").Value = "SDCAPCD"

If STButton = True Then Sheets("Inspections").Cells(1, "O").Value = "STBES"

If BAButton = True Then Sheets("Inspections").Cells(1, "O").Value = "BAAQMD"

If DButton = True Then Sheets("Inspections").Cells(1, "O").Value = "DEES"

End With

Application.DisplayAlerts = True
Call OptimizeCode_End

Me.Hide

End Sub

Private Sub UserForm_Click()

End Sub
