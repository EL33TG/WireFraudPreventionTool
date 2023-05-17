'#############################################
' 		RP36 ACH Auto Approver Script		 '
'                                            '
' CREATED ON: 09/11/2013 BY: Brenden Boswell ' 
' SYSTEM: Passport                           '
'#############################################

Option Explicit
Dim total 'Total holder/checker
Dim count '# Personal Reminder: Remove variable after testing phase
count = 0
Dim Text 'Default
Dim ret 'Default
Dim state 'Check validation state
Dim iCheck ' Check for acknowledgement
Dim mlogin : mlogin = "Error: You will need to log on Passport before running this script" 'Login error message
Dim mload : mload = "Error: RP36 Did not load, do you have correct access? " 'Loading eror message
Dim mCont : mCont = "This script will automatically activate all ACH accounts located on the RP36 screen. Depending on the amount of accounts listed, this could take some time." & vbCrLF _ 
& vbCrLF & "Please reaffirm that you are wanting to perform this action."& vbCrLF _
& vbCrLF &"EMERGENCY STOP: Click RUN button in PASSPORT"&vbCrLF _
& vbCrLF & "Please click OK to proceed or CANCEL to stop"
Dim logn : logn = GetString (20, 3, 1)
Dim logn2 : logn2 = GetString (23, 2, 7)
Dim logErr
Const APP = "<pf4>" 'Key holder
Const BCK = "<pf3>" '^
Const CLR = "<clear>" '^
Const SCR1 = "rp36" '^
Const SUBM = "<enter>" '^
Const TB = "<tab>" '^
Const TXT1 = "ACH" 'Variable validator string

Sub ZMain() ' Main entry point

Call Validate() ' Validation checker Function

If (state) Then
		total = GetString (2, 78, 3)
	If IsNumeric(total) Then
		Else
			total = GetString (2, 79, 3)
	End If
	On Error Resume Next
		If (CInt(LTrim(total)) > 0) Then
				Call StartUpdate()
				Else
				'Do Nothing
				End if
	Else
	End If
	
End Sub

Sub StartUpdate() 'Perform action 
Dim screen
Dim sc1 : sc1 = "RP36"
Dim ErrC : ErrC = "NO PAYEES"

	Do Until (count = CInt(LTrim(total)) OR logErr = "NO PAYEES") ' Change to reflect Total when out of TESTING MODE (Ex. Do Until count = 2) This runs script only 2 times
	screen = GetString (1, 2, 4)
	If (screen = sc1 AND logErr <> ErrC) Then
		SendHostKeys TB : SendHostKeys TB : SendHostKeys TB
		SendHostKeys ("s" & SUBM)
		ret = WaitForHostUpdate(1)
		SendHostKeys APP : SendHostKeys APP : SendHostKeys BCK
		ret = WaitForHostUpdate(1)
		'Extra system wait is optional if needed, ret = WaitForHostUpdate(10)
		logErr = GetString (24, 2, 9)
		count = count + 1
		Else
		'Do Nothing
		Exit Do
	End If
		
	Loop
	MsgBox " (" & count & ") Out of " & "(" & total & ") Accounts Activated", 64, "Activation Complete!"
	'End Sub
End Sub

Sub Validate() ' Validate
Dim go
Dim value
Dim iMssg
Dim mRp : mRp = "RP36 - Access Error"
Dim mLg : mLg = "Log-in error"

On Error Resume Next 'Error handling
If (logn <> "=" AND logn2 <> "Command" AND Err = 0) Then
iMssg = "## ATTENTION ## - RP36 ACH Account Activation Script"
iCheck = MsgBox(mCont, vbOKCancel + 48, iMssg)

	if iCheck = vbCancel Then 
		'cancel button was pressed Exit
		Exit Sub
	End If
		SendHostKeys CLR : SendHostKeys SCR1 & SUBM ' Prepare screen
		bool = true
		ret = WaitForHostUpdate(10)
		value = GetString (1, 25, 3) 
		'logErr = GetString (24, 2, 9)
				
			If (RTrim(value) = TXT1 AND bool) Then
					state = true
				Else
					MsgBox mload, 48, mRp
			End If
	Else
		MsgBox mlogin, 16 , mLg
		Exit Sub
End If

	'End Sub
End Sub
