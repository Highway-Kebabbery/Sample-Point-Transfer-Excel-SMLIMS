;This script is used to pull sample points from Microsoft Excel and enter them into Sample Manager LIMS.
;It's probably not necessary, but I took this as a chance to practice OOP.
;
;Pre-conditions for use (manual user setup):
;
;SM LIMS: Login to SM LIMS Prod.
;	
;Excel: Open Sampling Point List Edited. Columns 1-3 should be ID,
;	Name, and Desc. respectively. These columns should be fully
;	visible in Excel and on the left-most side of the screen.
;	Widths should be: 18.71, 55, and 96.29, respectively.
;	This script enters samples from the bottom going up. Scroll
;	down until none of the data	points to be entered are visible.
;	Click on a cell in column A and use the "up" arrow to
;	navigate up until the first data point to be entered is
;	visible at the top of the window. Of all the data points to
;	be entered, this should be the only	one visible.
;
;
;When the above tasks are complete, click "OK" to proceed.
;Click "Cancel" to exit.
;	   
;Version 1.0, 11-Apr-2023, Nathan Whisman
;_________________________________________________________________________________________________________
;_________________________________________________________________________________________________________


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn   ; Enable warnings to assist with detecting common errors.


InitializationMessageString =
(
This script is used to pull sample points from Microsoft Excel and enter them into Sample Manager LIMS.

Pre-conditions for use (manual user setup):

SM LIMS: Login to SM LIMS Prod.
	
Excel: Open "Sampling Point List Edited.xslx". Columns 1-3 should be ID,
	Name, and Desc. respectively. These columns should be fully
	visible in Excel and on the left-most side of the screen.
	Widths should be: 18.71, 55, and 96.29, respectively.
	This script enters samples from the bottom going up. Scroll
	down until none of the data points to be entered are visible.
	Click on a cell in column A and use the "up" arrow to
	navigate up until the first data point to be entered is
	visible at the top of the window. Of all the data points to
	be entered, this should be the only one visible.

When the above tasks are complete, click "OK" to proceed.
Click "Cancel" to exit.
	   
Version 1.0, 11-Apr-2023, Nathan Whisman
)

class NavToPage
{
	;Navigates from SampleManager LIMS - Explorer to SampleManager LIMS - Sample Point.
	navToSamplePointsFromHome(x, y, z)
	{
		WinWait, ahk_id 2492270 ;SampleManager LIMS - Explorer
		WinActivate, ahk_id 2492270 ;SampleManager LIMS - Explorer
		MouseMove, x, y 
		Click
		MouseMove, 0, 25,, R
		Sleep z
		MouseMove, 130, 125,, R
		Sleep z
		MouseMove, 160, 0,, R
		Sleep z
		Click
		Sleep z
	}
}

class RightClickMenuNav
{
	;Opens the "Sample Point - Add" window from within the "SampleManager LIMS - Sample Point" window.
	addSamplePoint(x, y, z)
	{
		WinWait, SampleManager LIMS - Sample Point
		WinActivate, SampleManager LIMS - Sample Point
		MouseMove, x, y
		Click, Right
		Sleep z
		MouseMove, 75, 15,, R
		Click
		Sleep z
		MouseMove, 100, 0,, R
		Click
	}
}

class CopyData
{
	;Navigates to Excel and copies a piece of data.
	grabDataFromExcel(x, y, z)
	{
		WinActivate, Sampling Point List Edited
		MouseMove, x, y
		Click, Left
		Send ^c
	}
}

class SamplePointAddWindowSequence
{
	;Waits for "Sample Point - Add" window to load, navigates to it, and pastes data.
	pasteData(x, y)
	{
		WinWait, Sample Point - Add
		WinActivate, Sample Point - Add
		MouseMove, x, y
		Click
		Send ^v
	}
	
	;The Sample ID field auto-populates the Sample Name field with unwanted data when it is filled.
	;Sending {Tab} forces this population in a controlled manner.
	sendTab()
	{
		Send {Tab}
	}
	
	;Selecting all and deleting clears the data previously auto-populated to Sample Name.
	clearCellPasteData(x, y)
	{
		WinWait, Sample Point - Add
		WinActivate, Sample Point - Add
		MouseMove, x, y
		Click
		Send ^a
		Send {Delete}
		Send ^v
	}

	;Adding "Group" in the "Sample Point - Add" window.
	addGroup(x, y, z)
	{
		WinActivate, Sample Point - Add
		MouseMove x, y
		Click
		Send Alachua
		Sleep z
		Send {Tab}
		Sleep z
	}

	;Clicks "OK" in the "Sample Point - Add" window.
	completeSamplePointAdd(x, y)
	{
		WinActivate, Sample Point - Add
		MouseMove x, y
		;Click
	}
}

class ChangeLineInExcel
{
	;Navigates to Excel and move up one line
	moveUpOneLineInExcel(x, y, z)
	{
		WinActivate, Sampling Point List Edited
		MouseMove, x, y
		Click, Left
		Send {Up}
		Sleep z
	}
}

MsgBox, 1, SM LIMS Sample Point Automation Script, % InitializationMessageString
IfMsgBox OK
	MsgBox ,, Instructions, Ctrl+Shift+S to begin, Ctrl+Shift+A to abort.
else
	ExitApp

navToSamplePointsPage := New NavToPage
samplePointAddMenuNav := New RightClickMenuNav
dataCopier := New CopyData
samplePointAdder := New SamplePointAddWindowSequence
lineChanger := New ChangeLineInExcel

^+a:: ExitApp
^+b:: BlockInput On                                               ;Only works if you run as Administrator.
^+s::
	navToSamplePointsPage.navToSamplePointsFromHome(105, 43, 500) ;Step 1 : Navigate to SampleManager LIMS - Sample Point.
	
	Loop 411
	{
	samplePointAddMenuNav.addSamplePoint(245, 250, 200),          ;Step 2 : Right-click and add Sample Point.
	dataCopier.grabDataFromExcel(65, 255, 0),                     ;Step 3 : Copy Sample ID from Excel.
	samplePointAdder.pasteData(155, 100),                         ;Step 4 : Paste Sample ID to "Sample Point - Add" window.
	samplePointAdder.sendTab(),                                   ;Step 5 : Send {Tab} to force Sample Name to auto-populate in a controlled manner.
	dataCopier.grabDataFromExcel(250, 255, 0),                    ;Step 6 : Copy Sample Name from Excel.
	samplePointAdder.clearCellPasteData(155, 143)                 ;Step 7 : Paste Sample Name to "Sample Point - Add" window.
	dataCopier.grabDataFromExcel(550, 255, 0)                     ;Step 8 : Copy Description from Excel.
	samplePointAdder.pasteData(155, 200)                          ;Step 9 : Paste Description to "Sample Point - Add" window.
	samplePointAdder.addGroup(200, 330, 50)                       ;Step 10: Select "Alachua" group from drop-down menu.
	samplePointAdder.completeSamplePointAdd(216, 572)             ;Step 11: Click "OK" in the "Sample Point - Add" window.
	lineChanger.moveUpOneLineInExcel(65, 255, 900)                ;Step 12: Reset the Excel workspace to begin a new line
	}

