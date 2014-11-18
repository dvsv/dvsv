' ------------------------------------------------ List of Functions ------------------------------------------------------------
'Fn_Module
'Fn_Interface 
'Fn_Configuration
'Fn_CurrentKeyword
'Fn_TestCaseSelector
'Fn_GetPoint
'Fn_GetSize
'Fn_oPoint
'Fn_Label
'Fn_RadioButton
'Fn_CheckBox
'Fn_TextBox
'Fn_ComboBox
'Fn_Button
'------------------------------------------------------- End Of List -----------------------------------------------------------

'************************************************************************************************************************************
'Function Name : Fn_Module
'Purpose : Fetches a function name based upon the keyword in the module sheet
'************************************************************************************************************************************
Function Fn_Module ( strCurModule )

	Dim ModCompletionFlag

	If Environment("Driver") = True Then  
		 Set objWorkSheetDriver = Environment("objFlowWorkBook").Worksheets("Driver")
		 intDriverColCnt = objWorkSheetDriver.UsedRange.Columns.Count
		 intDriverRowCnt = objWorkSheetDriver.UsedRange.Rows.Count
' ********************************   Environment  varaible started storing  in Environment header ( Row 1 of Driver sheet is header and  Module row is Functional Name) ********************************
		 For driverRows = 2 to intDriverRowCnt
					If Trim(objWorkSheetDriver.Rows(driverRows).Columns(1).Value) = strCurModule Then
					For driverCols = 2 to intDriverColCnt
						Environment(Trim(objWorkSheetDriver.Rows(1).Columns(driverCols).Value)) =  Trim(objWorkSheetDriver.Rows(driverRows).Columns(driverCols).Value)
					Next
					Environment("Driver") = False
					Exit For
					End If
		 Next	
' ********************************   Environment  varaible Completes storing  in Environment header ( Row 1 of Driver sheet is header and  Module row is Functional Name) *****************************
	End If

	
	Set objWorkSheetModule = Environment("objFlowWorkBook").Worksheets(strCurModule)
' ************************************************   Iteration starts at row level ( Each row is a test case ) ***********************************************************************************************************************
	For cols = 1 to Environment( "intCurModColCnt" )
	 If Instr(1, Trim( LCase( objWorkSheetModule.Rows(1).Columns( cols ).Value ) ), "keyword", 1 ) > 0 Then
		 ModCompletionFlag = False
		 If Trim( LCase( Environment(objWorkSheetModule.Rows( 1 ).Columns(cols).Value ) ) ) = "yes" And Environment( "TestCaseFlag" ) = True Then   
			strFuncName = Environment("Driver"&Trim(objWorkSheetModule.Rows( 1 ).Columns(cols).Value)) 
			Call Fn_CurrentKeyword ( LCase(objWorkSheetModule.Rows( 1 ).Columns(cols).Value) )
			Execute("Call "& strFuncName &"( )") 
			ModCompletionFlag = True
			If ModCompletionFlag = False Then
				Exit For
			End If
		End If
	End If
	Next

'	Updating the results for Completed Test Case
	Call Fn_UpdateTestResult( )

End Function

'************************************************************************************************************************************
'Function Name : Fn_Interface 
'Purpose : Creates a HTML interface to provide necessary initial data for the application
'************************************************************************************************************************************

Function Fn_Interface (  )

    Dim intI
	Dim intPositionY
	Dim objCollection
	Dim objItems
	Dim oForm

	Set objCollection = CreateObject( "Scripting.Dictionary" )
	Set oForm = DOTNetFactory.CreateInstance("System.Windows.Forms.Form", "System.Windows.Forms")
	intPositionY = 25

	objCollection.Add "lblInputXML", Fn_Label( "sInputXML", "lblInputXML", 35, intPositionY )
	objCollection.Add "txtInputXML", Fn_TextBox( "C:\Automation\PPC Web Service Automation\InputData\Request_OPF-PPC.xml","txtInputXML",235,intPositionY )
	intPositionY = intPositionY + 25

'	objCollection.Add "lblRacfPwd", Fn_Label( "RACF Password", "lblAppPath", 35, intPositionY )
'	objCollection.Add "txtRacfPwd", Fn_TextBox( "","txtAppPath",235,intPositionY )
'	intPositionY = intPositionY + 25

	objCollection.Add "lblAppPath", Fn_Label( "Application Path", "lblAppPath", 35, intPositionY )
	objCollection.Add "txtAppPath", Fn_TextBox( "http://rdxzn0c:8282/ppcservice/services/ppcPort?wsdl","txtAppPath",235,intPositionY )
	intPositionY = intPositionY + 25
	objCollection.Add "lblDtName", Fn_Label( "Datatable Name", "lblDtName", 35, intPositionY )
	objCollection.Add "txtDtName", Fn_TextBox( "DataSheet.xls","txtDtName",235,intPositionY )
	intPositionY = intPositionY + 25
	objCollection.Add "lblRptFmt", Fn_Label( "Report Format", "lblRptFmt", 35, intPositionY )
	objCollection.Add "cmbRptFmt", Fn_ComboBox( "cmbRptFmt","HTML Report",235,intPositionY )
	intPositionY = intPositionY + 45
	objCollection.Add "lblModules", Fn_Label( "Select Modules to Perform Test", "lblModules", 35, intPositionY )
	intPositionY = intPositionY + 25
	objCollection.Add "RequestData", Fn_CheckBox( "RequestData ", "RequestData", 60, intPositionY )
	intPositionY = intPositionY + 25
'	objCollection.Add "ModuleName2", Fn_CheckBox( "ModuleName2 ", "ModuleName2", 60, intPositionY )
'	intPositionY = intPositionY + 25
'	objCollection.Add "ModuleName3", Fn_CheckBox( "ModuleName3 ", "ModuleName3", 60, intPositionY )
'	intPositionY = intPositionY + 40
	objCollection.Add "Submit", Fn_Button( "Submit", "Submit", 95, intPositionY )
	objCollection.Add "Reset", Fn_Button( "Reset", "Reset", 240, intPositionY )

	With oForm
		.Text = "EPH PPC Webservice Automation"
		.Height = 400
		.Width = 450
        .Location.X = 100
        .Location.Y = 100
        .Minimizebox = False
        .Maximizebox = False
	End With

	objItems = objCollection.Items
	For intI = 0 To objCollection.Count - 1
		oForm.Controls.Add objItems( intI )
	Next

	oForm.CancelButton = objCollection.Item( "Submit" )	
	oForm.Activate
	oForm.ShowDialog

	Environment("sInputXML") = objCollection.Item( "txtInputXML" ).Text 
	'Environment("RacfPwd") = Crypt.Encrypt(objCollection.Item( "lblRacfPwd" ).Text )
	Environment("AppPath") = objCollection.Item( "txtAppPath" ).Text 
	Environment("DTName") = objCollection.Item( "txtDtName" ).Text
	Environment("Output") = Trim( Replace( objCollection.Item( "cmbRptFmt" ).Text, "Report", "" ) )

	Environment("Modules") = ""
    If objCollection.Item( "RequestData" ).checked Then
		Environment("Modules") = Environment("Modules") & "RequestData" & "@@"
	End If
'	If objCollection.Item( "ModuleName2" ).checked Then
'		Environment("Modules") = Environment("Modules") & "ModuleName2" & "@@"
'	End If
'	If objCollection.Item( "ModuleName3" ).checked Then
'		Environment("Modules") = Environment("Modules") & "ModuleName3"
'	End If

	objCollection.RemoveAll
	Set objCollection = Nothing
	Set oForm = Nothing
	Set objWorkSheet = Nothing
	'SystemUtil.CloseProcessByName ("excel.exe")

End Function

'*************************************************************************************************************
' Function : Fn_Configuration
'*************************************************************************************************************
Function Fn_Configuration ( strCheck )

	If LCase(strCheck) = "setup" Then
'		Environment.LoadFromFile(Environment("TestLocation") &"\InputData\Environment.xml")
		Environment("DatatableLocation") = Environment("TestLocation") &"\InputData\"& Environment( "DTName" )
		Environment("Result_Location")  = Environment("TestLocation") &"\OutputFiles\"
		Environment("ScreenShot") = Environment("TestLocation") &"\OutputFiles\ScreenShot\"
	ElseIf LCase(strCheck) = "modulestart" Then
'		RepositoriesCollection.Add( Environment("TestLocation") &"\Repository\Repository.tsr" )
'		If Setting("IsInTestDirectorTest") Then
'			Call Fn_QCTestDetails ( "Create", QCFolderPath( ) &";"& QCUtil.CurrentTestSet.Name &";"& QCUtil.CurrentRun.TestId &";"& QCutil.CurrentTestSetTest.ID )
'		End If
	ElseIf LCase(strCheck) = "moduleend" Then
'		If Environment("Output") = "HTML" Then
			Call Fn_CreateHTML( "End", Environment( "strCurModule" ) )
'		End If
'		Call Fn_QCResultUpdate( )
'		Call Fn_QCInstanceCreation( )
		RepositoriesCollection.RemoveAll
	End If

End Function


'*************************************************************************************************************
' Function : Fn_CurrentKeyword
' Functionality : To Setup any any variables/prerequesited before executin the functions.
'*************************************************************************************************************
Function Fn_CurrentKeyword( CurrQueue )

			If CurrQueue = "keyword1" Then

			ElseIf CurrQueue = "keyword2" Then

			ElseIf CurrQueue = "keyword3" Then

			ElseIf CurrQueue = "keyword4" Then

			ElseIf CurrQueue = "keyword5" Then

			End If

End Function

'*************************************************************************************************************
' Function : Fn_TestCaseSelector
'*************************************************************************************************************
Function Fn_TestCaseSelector ( objWorkSheet, strModule )

	Dim intI
	Dim intJ
	Dim intPositionY
	Dim objCollection
	Dim oForm
	Dim dResult
	Dim intRowCnt
	Dim strTestCase
	Dim strDesc
	Dim arrTemp
	Dim strTemp

	Set objCollection = CreateObject( "Scripting.Dictionary" )
	Set oForm = DOTNetFactory.CreateInstance("System.Windows.Forms.Form", "System.Windows.Forms")
	Set dResult = DotNetFactory.CreateInstance("System.Windows.Forms.DialogResult", "System.Windows.Forms")

	intRowCnt = objWorkSheet.UsedRange.Rows.Count
	intPositionY = 25
	intJ = 0

	objCollection.Add "lblDesign1", Fn_Label( "*******************************************************************************************", "lblDesign1", 30, intPositionY )
	intPositionY = intPositionY + 20
	objCollection.Add "lblInstruction", Fn_Label( "Select the Test Cases to Perform Test", "lblInstruction", 135, intPositionY )
	intPositionY = intPositionY + 25
	objCollection.Add "lblDesign2", Fn_Label( "*******************************************************************************************", "lblDesign2", 30, intPositionY )
	intPositionY = intPositionY + 35
	objCollection.Add "lblTestCase", Fn_Label( "TestCases", "lblTestCase", 45, intPositionY )
	objCollection.Add "lblTestDesc", Fn_Label( "Description", "lblTestDesc", 175, intPositionY )
	intPositionY = intPositionY + 25

	With oForm
		.Text = strModule
		.Height = 450
		.Width = 500
		.Controls.Add objCollection.Item( "lblDesign1" )
		.Controls.Add objCollection.Item( "lblInstruction" )
		.Controls.Add objCollection.Item( "lblDesign2" )
		.Controls.Add objCollection.Item( "lblTestCase" )
		.Controls.Add objCollection.Item( "lblTestDesc" )
	End With

	For intI = 1 To intRowCnt - 1

		strTestCase = Trim( objWorkSheet.Rows( intI + 1 ).Columns( 1 ).Value )
		strDesc = Trim( objWorkSheet.Rows( intI + 1 ).Columns( 2 ).Value )
		If strTestCase <> "" Then
			objCollection.Add "chkTestCase"& intI, Fn_CheckBox( strTestCase, "TestCase"& i, 35, intPositionY )
			objCollection.Add "lblTestDesc"& intI, Fn_Label( strDesc, "lblTestDesc"& i, 175, intPositionY + 5 )
			intPositionY = intPositionY + 25
			oForm.Controls.Add objCollection.Item( "chkTestCase"& intI )
			oForm.Controls.Add objCollection.Item( "lblTestDesc"& intI )
			intJ = intJ + 1
		End If

	Next

	intPositionY = intPositionY + 30
	objCollection.Add "btSubmit", Fn_Button( "Submit", "btSubmit", 35, intPositionY )
	objCollection.Add "btSelectAll", Fn_Button( "SelectAll", "btSelectAll", 175, intPositionY )
	objCollection.Add "btDeSelectAll", Fn_Button( "Reset", "btDeSelectAll", 315, intPositionY )
	intPositionY = intPositionY + 40
	objCollection.Add "lblEnd", Fn_Label( ".", "lblEnd", 55, intPositionY )

	With oForm
		.Controls.Add objCollection.Item( "btSubmit" )
		.Controls.Add objCollection.Item( "btSelectAll" )
		.Controls.Add objCollection.Item( "btDeSelectAll" )
		.CancelButton = objCollection.Item( "btSubmit" )
	End With

	If CInt( oForm.Height ) > CInt( intPositionY ) Then
		oForm.Height = intPositionY + 90
	Else
		oForm.Controls.Add objCollection.Item( "lblEnd" )
		oForm.AutoScroll = True
	End If	

	oForm.Activate
	oForm.ShowDialog

	Do
		If oForm.DialogResult = dResult.Yes Then
			For intI = 1 To intJ	
				objCollection.Item( "chkTestCase"& intI ).Checked = True	
			Next
			oForm.Activate
			oForm.ShowDialog
		End If
	
		If oForm.DialogResult = dResult.No Then
			For intI = 1 To intJ	
				objCollection.Item( "chkTestCase"& intI ).Checked = False	
			Next
			oForm.Activate
			oForm.ShowDialog
		End If
	
		If oForm.DialogResult = dResult.Cancel Then
			Exit Do
		End If
	Loop

	strTemp = ""
	For intI = 1 To intJ
On error resume next
		If objCollection.Item( "chkTestCase"& intI ).Checked Then
			strTemp = strTemp &";"& objCollection.Item( "chkTestCase"& intI ).Text
		End If
On error goto 0
	Next
	Set objCollection = Nothing
	Set Fn_TestCaseSelector = CreateObject( "Scripting.Dictionary" )

	arrTemp = Split( strTemp, ";" )
	For intI=0 To UBound( arrTemp )
		If arrTemp( intI ) <> "" Then
			Fn_TestCaseSelector.Add "Case"& intI,  arrTemp( intI )
		End If
	Next

	Set oForm = Nothing
	Set dResult = Nothing

End Function

'*************************************************************************************************************
' Function : Fn_GetPoint
'*************************************************************************************************************
Function Fn_GetPoint ( x, y )

'	Create a POINT object with constructor int, int
	Set Fn_GetPoint = DotNetFactory("System.Drawing.Point","System.Drawing", x, y)

End Function

'*************************************************************************************************************
' Function : Fn_GetSize
'*************************************************************************************************************
Function Fn_GetSize( x, y )

'	Create a Size object with constructor int, int
	Set Fn_GetSize = DotNetFactory("System.Drawing.Size","System.Drawing", x, y)

End Function

'*************************************************************************************************************
' Function : Fn_oPoint
'*************************************************************************************************************
Function Fn_oPoint ( intX, intY )

    Set Fn_oPoint = DotNetFactory.CreateInstance("System.Drawing.Point", "System.Drawing", x, y)

	With Fn_oPoint
		.x = intX
		.y = intY
	End With

End Function

'*************************************************************************************************************
' Function : Fn_Label
'*************************************************************************************************************
Function Fn_Label( lblText, lblName, xx, yy )

	Set Fn_Label = DOTNetFactory.CreateInstance("System.Windows.Forms.Label", "System.Windows.Forms")

'	Label Properties
	With Fn_Label
		.Text = lblText
		.Name = lblName
		.Size = Fn_GetSize( ( Len( lblText ) + 5 )*7, 20)
		.Location = Fn_oPoint( xx, yy )
	End With

End Function

'*************************************************************************************************************
' Function : Fn_RadioButton
'*************************************************************************************************************
Function Fn_RadioButton( radText, radName, xx, yy )

	Set Fn_RadioButton = DOTNetFactory.CreateInstance("System.Windows.Forms.RadioButton", "System.Windows.Forms")

'	RadioButton Properties
	With Fn_RadioButton
		.Text = radText
		.Name = radName
		.Location = Fn_oPoint( xx, yy )
	End With

End Function

'*************************************************************************************************************
' Function : CheckBox
'*************************************************************************************************************
Function Fn_CheckBox ( chkText, chkName, xx, yy )

	Set Fn_CheckBox = DOTNetFactory.CreateInstance("System.Windows.Forms.CheckBox", "System.Windows.Forms")

'	CheckBox Properties
	With Fn_CheckBox
		.Text = chkText
		.Name = chkName
		.Size = Fn_GetSize( ( Len( chkText ) + 5 )*8, 20)
		.Location = Fn_oPoint( xx, yy )
	End With

End Function



'*************************************************************************************************************
' Function : Fn_EncodedTextBox
'*************************************************************************************************************
Function Fn_EncodedTextBox ( txtText, txtName, xx, yy )

	Set Fn_EncodedTextBox = DOTNetFactory.CreateInstance("System.Windows.Forms.TextBox", "System.Windows.Forms")

'	CheckBox Properties
	With Fn_EncodedTextBox
		.UseSystemPasswordChar = True
		.Text = txtText
		.Name = chkNAme
		.Width = 150
		.Location = Fn_oPoint( xx, yy )
	End With

End Function



'*************************************************************************************************************
' Function : Fn_TextBox
'*************************************************************************************************************
Function Fn_TextBox ( txtText, txtName, xx, yy )

	Set Fn_TextBox = DOTNetFactory.CreateInstance("System.Windows.Forms.TextBox", "System.Windows.Forms")

'	CheckBox Properties
	With Fn_TextBox
        .Text = txtText
		.Name = chkNAme
		.Width = 150
		.Location = Fn_oPoint( xx, yy )
	End With

End Function

'*************************************************************************************************************
' Function : ComboBox
'*************************************************************************************************************
Function Fn_ComboBox ( cmbName, strList, xx, yy )

	Dim strTemp
	Dim intI
	Set Fn_ComboBox = DotNetFactory("System.Windows.Forms.ComboBox","System.Windows.Forms")

	strTemp = Split( strList, ";" )	
	With Fn_ComboBox
		.Name = cmbName
		.TabIndex = 3
		.Width = 150
		
'		Clear all items in the combo box list
		.Items.Clear
		
'		Add items to the combo box list
		For intI = 0 To UBound ( strTemp )
			.Items.Add strTemp( intI )
		Next
		.SelectedIndex = 0
		.Location = Fn_oPoint( xx, yy )
	End with
	
End Function

'*************************************************************************************************************
' Function : Fn_Button
'*************************************************************************************************************
Function Fn_Button( btText, btName, xx, yy )

	Set Fn_Button = DotNetFactory.CreateInstance("System.Windows.Forms.Button", "System.Windows.Forms")
	Set dResult = DotNetFactory.CreateInstance("System.Windows.Forms.DialogResult", "System.Windows.Forms")
	
	With Fn_Button
		.Text = btText
		.Name = btName
		.Location = Fn_oPoint( xx, yy )
		.Width = 100
	End with

	If btText = "SelectAll" Then
		Fn_Button.DialogResult = dResult.Yes
	ElseIf btText = "Reset" Then
		Fn_Button.DialogResult = dResult.No
	ElseIf btText = "Submit" And btName = "btSubmit" Then
		Fn_Button.DialogResult = dResult.Cancel
	End If

End Function
'------------------------------------------------------- End Of Fucctions -----------------------------------------------------------
