' ------------------------------------------------ List of Functions ------------------------------------------------------------
'Fn_UpdateTestResult
'Fn_UpdateExcel
'Fn_CreateResultFile
'Fn_CreateExcel
'Fn_CreateHTML
'Fn_Create_Header
'Fn_TableName
'Fn_QCResultUpdate
'Fn_QCTestDetails
'Fn_QCFolderPath
'Fn_QCInstanceCreation
'Fn_QCLabInstanceCreate
'------------------------------------------------------- End Of List -----------------------------------------------------------

'*************************************************************************************************************
' Function : Fn_UpdateTestResult
' Functionality : To Update the test results of particular test case.
'*************************************************************************************************************
Function Fn_UpdateTestResult()

	Dim strResult

'  Condition to set the Test  Case Result  by the data obtained from the validation done on application
		If Environment("CurTestResult") = False Then
			Environment("Status") = "Fail"
			Environment("Fail_Count")  = Environment("Fail_Count")  + 1
			If Trim( Environment( "QC_TestName") ) <> "" Then
				Call FALMTestSetStatusChange( Environment( "QC_TestName"), Environment( "strQCLabTestSet"), "Failed" )
			End If
'			Call Fn_QCTestDetails ( "Update", Environment("TC_NO") &"-Failed;" )
		Else
			Environment("Status") = "Pass"
            Environment("Pass_Count") = Environment("Pass_Count") + 1
'			Call Fn_QCTestDetails ( "Update", Environment("TC_NO") &"-Passed;" )
			If Trim( Environment( "QC_TestName") ) <> "" Then
				Call FALMTestSetStatusChange( Environment( "QC_TestName"), Environment( "strQCLabTestSet"), "Passed" )
			End If
		End If
'  Test Results to be reported be given here with Semicolon as delimiter seperating the columns
		'strResult = "ResultValue1;ResultValu2;ResultValue3" 'ResultValue1;ResultValu2;...;ResultValueN

'	If Environment( "Output" ) = "Excel" Then
'		Call Fn_UpdateExcel( strResult )
'	ElseIf Environment( "Output" ) = "HTML" Then
'		Call Fn_TableName( strResult, Environment("Status"), Environment( "strCurModule" ) )
'	End If

End Function

'*************************************************************************************************************
' Function : Fn_UpdateExcel
' Functionality : To Update the test results of particular test case.
'*************************************************************************************************************
Function Fn_UpdateExcel( strResult )

	Dim intRowCnt, intColCnt, strTemp, intI
	Dim objExcel, objWorkBook, objWorkSheet

	Set objExcel = CreateObject("Excel.Application")
	Set objWorkBook = objExcel.Workbooks.Open ( Environment("Result_Location") & Environment("DTName") )
	Set objWorkSheet = objWorkBook.Worksheets( Environment("strCurModule") )
	strTemp = Split(strResult,";")

	intRowCnt = objWorkSheet.UsedRange.Rows.Count
	intColCnt = objWorkSheet.UsedRange.Columns.Count

	For intI = 0 To UBound( strTemp )
		objWorkSheet.Rows(intRowCnt + 1).Columns(intJ + 1).NumberFormat = "@"
		objWorkSheet.Rows(intRowCnt + 1).Columns(intI + 1).Value = Cstr( strTemp( intI ) )
		objWorkSheet.Rows(intRowCnt + 1).Columns(intJ + 1).Font.Size = 10
		objWorkSheet.Rows(intRowCnt + 1).Columns(intJ + 1).Font.Name = "Calibri"
		objWorkSheet.Rows(intRowCnt + 1).Columns(intJ + 1).BorderAround Weight = xlThin
	Next

	objWorkBook.Save
	objWorkBook.Close

	Set objExcel = Nothing
	Set objWorkBook = Nothing
	Set objWorkSheet = Nothing

End Function

'*************************************************************************************************************
' Function : Fn_CreateResultFile
' Functionality : To Create a Result of user's choice.
'*************************************************************************************************************
Function Fn_CreateResultFile ( )

	Dim strResult

	strResult = "Element Name;Baseline Value ;Actual Value; Result" ' HeaderValue1;HeaderValu2;...;HeaderValueN


	If Environment("Output") = "Excel" Then
		Call Fn_CreateExcel( strResult )
	ElseIf Environment("Output") = "HTML" Then
		Call Fn_CreateHTML( "Start", Environment( "strCurModule" ))
		Call Fn_TableName ( strResult, "Header", Environment( "strCurModule" ) )
	End If

End Function

'*************************************************************************************************************
' Function : Fn_CreateExcel
' Functionality : To create a excel file for result to be updated.
'*************************************************************************************************************
Function Fn_CreateExcel ( strResult )

    Dim objExcel, objWorkbook, objWorkSheet
	Dim intI, intJ, objFso, strTemp, intTemp
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objExcel = CreateObject("Excel.Application")

	strTemp = Split(strResult,";")
	If objFso.FileExists( Environment("Result_Location") & Environment("DTName")) Then
		Set objWorkbook = objExcel.Workbooks.Open ( Environment("Result_Location") & Environment("DTName") )
		intTemp = objWorkbook.Worksheets.Count
		For intI = 1 To intTemp
			strSheetName = objWorkbook.Worksheets(intI).Name
			If strSheetName = Environment("strCurModule") Then
				Set objWorkSheet = objWorkbook.Worksheets(strSheetName)
				For intJ = 0 To UBound(strTemp)					
					objWorkSheet.Rows(1).Columns(intJ + 1).Value = strTemp(intJ)
					objWorkSheet.Rows(1).Columns(intJ + 1).Font.Bold = True
					objWorkSheet.Rows(1).Columns(intJ + 1).Font.Size = 10
					objWorkSheet.Rows(1).Columns(intJ + 1).Font.Name = "Calibri"
					objWorkSheet.Rows(1).Columns(intJ + 1).Interior.Color = 49407
					objWorkSheet.Rows(1).Columns(intJ + 1).BorderAround Weight=xlThin
				Next
				strSheetFlag = True
				Exit For
			Else
				strSheetFlag = False
			End If
		Next

		If strSheetFlag = False Then
			Set objWorkSheet = objWorkbook.Worksheets.Add
			objWorkSheet.Name = Environment("strCurModule")				
			For intJ = 0 To UBound(strTemp)					
				objWorkSheet.Rows(1).Columns(intJ + 1).Value = strTemp(intJ)
				objWorkSheet.Rows(1).Columns(intJ + 1).Font.Bold = True
				objWorkSheet.Rows(1).Columns(intJ + 1).Font.Size = 10
				objWorkSheet.Rows(1).Columns(intJ + 1).Font.Name = "Calibri"
				objWorkSheet.Rows(1).Columns(intJ + 1).Interior.Color = 49407
				objWorkSheet.Rows(1).Columns(intJ + 1).BorderAround Weight=xlThin
			Next
		End If
		objWorkbook.Save
		objWorkbook.Close

	Else

        Set objWorkbook=objExcel.Workbooks.Add
		objWorkbook.Worksheets("Sheet1").Delete
		objWorkbook.Worksheets("Sheet2").Delete
		Set objWorkSheet = objWorkbook.Worksheets.Add
		objWorkSheet.Name = Environment("strCurModule")		
		For intJ = 0 To UBound(strTemp)					
			objWorkSheet.Rows(1).Columns(intJ + 1).Value = strTemp(intJ)
			objWorkSheet.Rows(1).Columns(intJ + 1).Font.Bold = True
			objWorkSheet.Rows(1).Columns(intJ + 1).Font.Size = 10
			objWorkSheet.Rows(1).Columns(intJ + 1).Font.Name = "Calibri"
			objWorkSheet.Rows(1).Columns(intJ + 1).Interior.Color = 49407
			objWorkSheet.Rows(1).Columns(intJ + 1).BorderAround Weight=xlThin
		Next
		objWorkbook.Worksheets("Sheet3").Delete
		objWorkbook.SaveAs( Environment("Result_Location") & Environment("DTName") )
		objWorkbook.Save
		objWorkbook.Close

	End If

	Set objWorkSheet = Nothing
	Set objWorkbook = Nothing
	Set objExcel = Nothing
	Set objFso = Nothing
	Environment ( "Fail_Count" ) = 0
	Environment ( "Pass_Count" ) = 0

End Function

'*************************************************************************************************************
'	Function						: 	Fn_CreateHTML
'	Description	                :	To create HTML report file for the current test run
'  	Input Argument(s)	 :	 strFlowStatus - indicates whether to write header or summary table
'                                               depending on the value passed 
'											   Start indicates start of the test and writes test details and output table header 
'												End indicates end of test and writes the summary table in the HTML report
' 	Called Function			:   Fn_Create_Header
'	Calling Function		:   Fn_CreateResult
'	Return Value(s)		  :	   None
'*************************************************************************************************************
Function Fn_CreateHTML( strFlowStatus, strResName )

	Dim objFSO, objExcel, objWorkBook, objWorkSheet
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0 'report writing
	Dim objFSO1, objGetFile, objOPFile    'report
	Dim TotTime, totalhrs, totalmins, totalsecs, TotalVal

	If LCase ( strFlowStatus )= "start" Then

		Environment ( "StartTime" ) = Timer
		Environment ( "Fail_Count" ) = 0
		Environment ( "Pass_Count" ) = 0

		Set objFSO1 = CreateObject ( "Scripting.FileSystemObject" )
'		Result file location
		Environment( strResName ) = ( Environment ( "Result_Location" ) ) &  strResName  & Replace ( date , "/" , "_" ) & "_" & Replace ( Time , ":" , "_" ) & ".html"

'		Creating the result file
		objFSO1.CreateTextFile Environment ( strResName ) 
		Set objGetFile = objFSO1.GetFile ( Environment ( strResName ) )

'		Openning the File to write TestDetails, ToolDetail, UserDetails
		Set objOPFile = objGetFile.OpenAsTextStream ( ForWriting , - 2 )
		objOPFile.writeLine( "<HTML><HEAD>" )
		objOPFile.writeLine( "<title>iGen Framework Report</title>" )
		objOPFile.writeLine( "</HEAD><BODY>" )
		objOPFile.writeLine( "<P><B><U><CENTER><FONT face=Verdana color=#0033CC  size=4>iGen Report</FONT></CENTER></U></B></P>" )
		objOPFile.writeLine( "<TABLE height=10 width='100%' borderColorLight=#008080 border=2>&nbsp" )
		objOPFile.writeLine( "<TR>" )
		objOPFile.writeLine( "<TD vAlign=Left  align=middle  width='3%' bgColor=#e1e1e1 height=20>" )
		objOPFile.writeLine( "<FONT face=Verdana color=#0033CC size=2><B>Execution Date :  " )
		objOPFile.writeLine( date ) 
		objOPFile.writeLine( "<BR><FONT face=Verdana color=#0033CC size=2><B>Module Name :  " )
		objOPFile.writeLine( Environment( "strCurModule" ) ) 
		objOPFile.writeLine( "<P align=Top>" )
		objOPFile.writeLine( "<TD vAlign=Left  align=middle width='3%' bgColor=#e1e1e1 height=20>" )
		objOPFile.writeLine( "<FONT face=Verdana color=#0033CC size=2><B>Executed By : " )
		objOPFile.writeLine(Environment( "UserName" ))
		objOPFile.writeLine("<BR>")
		objOPFile.writeLine( "<FONT face=Verdana color=#0033CC size=2><B>Tool Name: " )
		objOPFile.writeLine(Environment( "ProductName" )) 
		objOPFile.writeLine("<BR>")
		objOPFile.writeLine( "<FONT face=Verdana color=#0033CC size=2><B>Tool Version: " )
		objOPFile.writeLine(Environment( "ProductVer" )) 
		objOPFile.writeLine( "</TR>" )
		objOPFile.writeLine( "</B></FONT></TD>" )
		objOPFile.writeLine( "</Table>")

        objOPFile.Close
		Set objOPFile  = Nothing
		Set objFSO1 = Nothing
		Call Fn_Create_Header ( strResName )

	ElseIf LCase ( strFlowStatus ) = "end" Then

		Environment ( "EndTime" ) = Timer
		TotTime =    Environment ( "EndTime" ) -   Environment ( "StartTime" )

		Set objFSO1 = CreateObject( "Scripting.FileSystemObject" )
		Set objGetFile = objFSO1.GetFile ( Environment ( strResName ) )
		Set objOPFile = objGetFile.OpenAsTextStream ( ForAppending , - 2 )

'		To calculate time tin HH:MM:SS fromat
		totalhrs=int(TotTime/3600)
		totalmins=int(TotTime/60) - int(totalhrs*60)
		totalsecs=int(TotTime Mod 60)
		TotTime=totalhrs& " hrs " &totalmins& " mins "&totalsecs&" secs "

		objOPFile.writeLine( "</Table>")
		objOPFile.writeLine( "<TABLE  width='100%' >&nbsp")
		objOPFile.writeLine( "<P><B><U><Center><FONT face=Verdana color=#0033CC  size=4>Result</FONT></Centert></U></B></P>")
		objOPFile.writeLine( "</TABLE>")
		objOPFile.writeLine( "<TABLE height=10 width='50%' borderColorLight=#008080 border=2>")

		objOPFile.close
		Set objOPFile  = nothing
		Set objFSO1 = nothing

'		To Write the summary of test run
		TotalVal = CInt( Environment( "Pass_Count" ) ) + CInt( Environment( "Fail_Count" ) )
		Call Fn_TableName ( "Total Number of Test Cases Executed;"&TotalVal , "General" ,Environment( "strCurModule" )  )
		Call Fn_TableName ( "Total Number of Test Cases Passed;"&  Environment("Pass_Count") , "General" ,Environment( "strCurModule" )  )
		Call Fn_TableName ( "Total Number of Test Cases Failed;" & Environment("Fail_Count") , "Failed"   , Environment( "strCurModule" )  )
		Call Fn_TableName ( "Total Time taken for Execution of Test Cases;" & TotTime , "General" , Environment( "strCurModule" )  )
		Environment( "Result_Loc" ) = Environment( strResName )

	End If

End function


'*************************************************************************************************************
'	Function						: 	Fn_Create_Header
'	Description	                :	To create header in the HTML report
'  	Input Argument(s)	 :	 None
' 	Called Function			:   None
'	Calling Function		:   Fn_CreateHTML
'	Return Value(s)		  :	   None
'*************************************************************************************************************

Function Fn_Create_Header( strResName )

	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim objFSO1, objGetFile, objOPFile    'Report

	Set objFSO1 = CreateObject( "Scripting.FileSystemObject" )
	Set objGetFile = objFSO1.GetFile ( Environment( strResName ) )
	Set objOPFile = objGetFile.OpenAsTextStream ( ForAppending , - 2 )

	objOPFile.writeLine( "</Table>")
	objOPFile.writeLine( "<TABLE height=10 width='100%' borderColorLight=#008080 border=2>&nbsp" )
	objOPFile.close

	Set objOPFile  = Nothing
	Set objFSO1 = Nothing

End Function

'*************************************************************************************************************
'	Function						:	 Fn_TableName
'	Description	                :	To insert the result for every testcase
'  	Input Argument(s)	 :	   strTName -  holds all result/header values to be written into HTML file
'												strRName -  If value is "General" strTName is treated as result value
'																		If value is "Header" strTName is treated as column header 
' 	Called Function			:   N/A
'	Calling Function		:   Fn_CreateResult / Fn_CreateHTML
'	Return Value(s)		  :	   None
'*************************************************************************************************************

Function Fn_TableName( strTName , strRName, strResName )

		Dim objFSO1, objOPFile1, ColorVal, FailCondFlag, strName, strHName
		Const RED = "FF0000", GREEN = "009966"

'	   Create a File System Object
		Set objFSO1 = CreateObject( "Scripting.FileSystemObject" )
		Set objOPFile1 = objFSO1.OpenTextFile ( Environment( strResName ) , 8 , True)
'		Set the  Colour as Green
		ColorVal = GREEN
		FailCondFlag = False

		If Instr( 1 ,LCase(strRName) , "fail" , 1 ) > 0 Then
			ColorVal = RED
			FailCondFlag = True
		End If

		If strRName =  "Header"  Then
			strTName =  Replace ( strTName, ";", "::" )
		End If

			If strRName =  "TCName"  Then
				objOPFile1.writeLine( "<TR>" )
			 strName = Split ( strTName , ";" , - 1 , 1 )
			 objOPFile1.writeLine( "<TD align=center width='3%' bgColor=#ffffe1 height=30>" )
			 objOPFile1.writeLine( "<FONT face=Verdana color= Blue size=1><b> " & strName(0) & "</b></Font></TD>" )
				 objOPFile1.writeLine( "<TD colspan =3 align=center width='3%' bgColor=#ffffe1 height=30>" )
			 objOPFile1.writeLine( "<FONT face=Verdana color= Blue size=1><b> " & strName(1) & "</b></Font></TD>" )
			objOPFile1.writeLine( "</TR>" )
			Exit Function
		End If
'				If strRName =  "TCName"  Then
'			 objOPFile1.writeLine( " <TR Align=center align=middle width='4%' bgColor=#e1e1e1 height=60><FONT face=Verdana  color=#0033CC  size=2><B>" & strTName & "</B></FONT>" )
'			 		objOPFile1.writeLine( "</TR>" )
'			 Exit Function
'		End If

        strName = Split ( strTName , ";" , - 1 , 1 )
'		Split the values inorder to check whether the given value is header or column value (  :: will be used to indicate Column Header )
		strHName = Split ( strTName , "::" , - 1 , 1 )
'		Row Starts
		objOPFile1.writeLine( "<TR>" )
'        If :: is present then write the column Header

		If UBound ( strHName ) > 0 Then

			For i = 0 to UBound ( strHName )
                objOPFile1.writeLine( "<TD vAlign=center align=middle width='4%' bgColor=#e1e1e1 height=60><FONT face=Verdana  color=#0033CC  size=2><B>" & strHName(i) & "</B></FONT></TD>" )
    		Next
'			Write the Column Values

	   Else

			For i = 0 to UBound ( strName )
				objOPFile1.writeLine( "<TD align=center width='3%' bgColor=#ffffe1 height=30>" )
					If FailCondFlag = True Then
						objOPFile1.writeLine( "<FONT face=Verdana color= " & ColorVal & " size=1><b> <i> " & strName(i) & "</b></Font></TD>" )
					Else					
						objOPFile1.writeLine( "<FONT face=Verdana color= " & ColorVal & " size=1><b> " & strName(i) & "</b></Font></TD>" )
					End If				
			Next

		End If
		objOPFile1.writeLine( "</TR>" )

		Set objOPFile1  = nothing
		Set objFSO1 = nothing

End Function	
'************************************************************************************************************************************
'Function Name : Fn_QCResultUpdate 
'Purpose : To Upload Test Result in ALM.
'************************************************************************************************************************************
Function Fn_QCResultUpdate()

	Dim ObjCurrentTest
	Dim ObjAttch
	
	If Setting("IsInTestDirectorTest") Then
		Set ObjCurrentTest = QCUtil.CurrentTest
		Set ObjAttch = ObjCurrentTest.AddItem(Null)
		ObjAttch.FileName = Environment("Result_Loc")
		ObjAttch.Type = 1
		ObjAttch.Post
		ObjAttch.Refresh
	End If
	
	Set ObjAttch = Nothing
	Set ObjCurrentTest = Nothing

End Function

'************************************************************************************************************************************
'Function Name : Fn_QCFolderPath 
'Purpose : To Get the ALM Folder Structure
'************************************************************************************************************************************
Function Fn_QCFolderPath ( )

	Dim qcTest
	Dim qcfolder
	Dim qcfolderName
	Dim strPath

	If Setting("IsInTestDirectorTest") Then
		Set qcTest = QCUtil.CurrentTestSet
		Set qcfolder = qcTest.TestSetFolder
		qcfolderName = qcfolder.Name
		strPath = qcfolderName

		While qcfolderName <> "Root"
			Set qcfolder = qcfolder.Father
			qcfolderName = qcfolder.Name
			strPath = qcfolderName & "\" & strPath
		Wend

		QCFolderPath = strPath
	Else
		QCFolderPath = ""
	End If
	
	Set qcfolder = Nothing
	Set qcTest = Nothing

End Function

'************************************************************************************************************************************
'Function Name : Fn_QCTestDetails 
'Purpose : To Update Test Details.
'************************************************************************************************************************************
Function Fn_QCTestDetails( strType, strText )

	Dim wshShell
	Dim strFilePath
	Dim oFso
	Dim strqcDetails
	Dim strTestDetails
	Const ForReading = 1, ForWriting = 2, ForAppending = 8

	If Setting("IsInTestDirectorTest") Then
		Set wshShell = CreateObject("WScript.Shell")
		strFilePath = wshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\Temp.txt"
		Set wshShell = Nothing
		Set oFso = CreateObject( "Scripting.FileSystemObject" )
		If strType = "Create" Then
			Set oFile = oFso.OpenTextFile( strFilePath, ForWriting, True )
			oFile.WriteLine strText
		ElseIf strType = "Update" Then
			Set oFile = oFso.OpenTextFile( strFilePath, ForAppending, False )
			oFile.Write strText
		End If
		oFile.Close
	End If
	
	Set oFile = Nothing
    Set oFso = Nothing

End Function

'************************************************************************************************************************************
'Function Name : Fn_QCInstanceCreation 
'Purpose : To create instance of Test in TestLab of ALM
'************************************************************************************************************************************
Function Fn_QCInstanceCreation ( )

	Dim wshShell
	Dim strFilePath
	Dim oFso
	Dim strFileData
	Dim strTemp
	Dim strqcDetails
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	If Setting("IsInTestDirectorTest") Then
		Set wshShell = CreateObject("WScript.Shell")
		strFilePath = wshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\Temp.txt"
		Set wshShell = Nothing
		
		Set oFso = CreateObject("Scripting.FileSystemObject")
		If oFso.FileExists( strFilePath ) Then
			Set oFile = oFso.OpenTextFile( strFilePath, ForReading )
			If Not oFile.AtEndOfLine Then
				strFileData = oFile.ReadAll
				oFile.Close
				strTemp = Split( strFileData , Chr(10) )
				strqcDetails = Split( strTemp(0), ";" )
				Call QCLabInstanceCreate( Trim( strqcDetails(0) ), Trim( strqcDetails(1) ),  Trim( strqcDetails(2) ), Trim( strqcDetails(3) ), Trim( strTemp(1) ) )
				oFso.DeleteFile ( strFilePath )
			End If
			Set oFile = Nothing
		End If
	End If
	
	Set oFso = Nothing

End Function

'************************************************************************************************************************************
'Function Name : Fn_QCLabInstanceCreate 
'Purpose : To create instance of Test in TestLab of ALM
'************************************************************************************************************************************
Function Fn_QCLabInstanceCreate( strQCLabPath, strQCTestName, intCurrentTestID, intCurrentSetID, strTestStatusDetails )

	Dim qcConnect
	Dim qcTreeMgr
	Dim qcTestFactory
	Dim qcLabTreeMgr
	Dim qcLabFolder
	Dim qcTestSetFactory
	Dim qcFilter
	Dim qcLst
	Dim qcTestSet
	Dim qctestInstanceFctry
	Dim pTestInTestSetObj
	Dim intTotalCaseCount
	Dim intI
	Dim strTempTstDetails
	Dim strFlag
	Dim strTemp
	Dim strTestName
	Dim strqcTempStatus
	Dim pTestLst
	Dim pTestItem

	Set qcConnect = CreateObject("TDAPIOLE80.TDConnection")
	qcConnect.InitConnectionEx "https://almbnym.saas.hp.com/qcbin"
	qcConnect.Login "Suriya.Ravi", "Welcome"
	qcConnect.Connect "FMTS_TREASURY", "Trade_Banking_Systems"

	If (qcConnect.connected <> True) Then
'		MsgBox "QC Project Failed to Connect"
		WScript.Quit
	End If

	Set qcTreeMgr = qcConnect.TreeManager 
	Set qcTestFactory = qcConnect.TestFactory 
	Set qcLabTreeMgr = qcConnect.TestSetTreeManager
	Set qcLabFolder = qcLabTreeMgr.NodeByPath( strQCLabPath )
	Set qcTestSetFactory = qcLabFolder.TestSetFactory
	Set qcFilter = qcTestSetFactory.Filter
	qcFilter.Filter("CY_CYCLE") = strQCTestName

	Set qcLst = qcTestSetFactory.NewList(qcFilter.Text)
	If qcLst.Count = 0 Then 
		Set qcTestSet = qcTestSetFactory.AddItem( Null )
		qcTestSet.Field("CY_CYCLE") = strQCTestName
		qcTestSet.Post
	Else
		Set qcTestSet = qcLst.Item(1)
	End If

	strTempTstDetails = Split( strTestStatusDetails, ";" )
	strFlag = True
	Set qctestInstanceFctry = qcTestSet.TSTestFactory
	For intI = 0 To UBound( strTempTstDetails )
		If strTempTstDetails( intI ) <> "" And intI <> 0 Then	
		
			Set pTestInTestSetObj = qctestInstanceFctry.AddItem( intCurrentTestID )			
			strTemp = Split( strTempTstDetails( intI ), "-" )
			strTestName = Trim( CStr( strTemp( 0 ) ) )
			strqcTempStatus = Trim( CStr( strTemp( 1 ) ) )

			pTestInTestSetObj.Field( "TC_USER_TEMPLATE_01" ) = strTestName
			pTestInTestSetObj.Status = strqcTempStatus
			pTestInTestSetObj.Post
			AddTestToTestSet = True			
			Set pTestInTestSetObj = Nothing
			
		ElseIf strTempTstDetails( intI ) <> "" And intI = 0 Then			
			
			Set pTestItem = qctestInstanceFctry.Filter
			pTestItem.Filter("TC_TESTCYCL_ID") = CStr( intCurrentSetID )
			Set pTestLst = qctestInstanceFctry.NewList( pTestItem.Text )
			Set pTestInTestSetObj = pTestLst.Item( 1 )
			
			strTemp = Split( strTempTstDetails( intI ), "-" )
			strTestName = Trim( CStr( strTemp( 0 ) ) )
			strqcTempStatus = Trim( CStr( strTemp( 1 ) ) )
			pTestInTestSetObj.Field( "TC_USER_TEMPLATE_01" ) = strTestName
			pTestInTestSetObj.Status = strqcTempStatus 
			pTestInTestSetObj.Post
			
			Set pTestItem = Nothing
			Set pTestLst = Nothing
			Set pTestInTestSetObj = Nothing
			
		End If
	Next

	Set qctestInstanceFctry = Nothing
	Set qcTestSet = Nothing
	Set qcLst = Nothing
	Set qcFilter = Nothing
	Set qcTestSetFactory = Nothing
	Set qcLabFolder = Nothing
	Set qcLabTreeMgr = Nothing
	Set qcTestFactory = Nothing
	Set qcTreeMgr = Nothing
	Set qcConnect = Nothing

End Function

'*************************************************************************************************************************
'	Function name	        :	FALMTestSetStatusChange
'	Description	            :	To change the status of a particular testcase in a Testset 
'	Pre-Condition(s)	    :	ALM Lab folder should contain all Tests in particular Test Sets
'  	Input Argument(s)	    :   Testcase Status from the main function
'	Return Value(s)		    :	None
'	Created By		        :	Sukanya Ramanathan / Suriya Prakash Ravi
'	Created On		        :	22-Mar-2013
'*************************************************************************************************************************
'  function usage.

'	Call the function at the end of each test in QTP to update the status in ALM
'  		strTestName 		- Test case Name in ALM LAB folfer
'		strTestSetName		- Test Set Name under which test cases are available in ALM LAB folfer
'		strTestcaseStatus	- Test Case status ( Passed / Failed ) to be updated.
'*************************************************************************************************************************

Function FALMTestSetStatusChange(strTestName, strTestSetName, strTestcaseStatus)

	Dim qcConnect
	Dim qcTreeMgr
	Dim qcTestFactory
	Dim qcLabTreeMgr
	Dim qcLabFolder
	Dim qcTestSetFactory
	Dim qcFilter
	Dim qcLst
	Dim qcTestSet
	Dim qctestInstanceFctry
	Dim pTestInTestSetObj
	Dim intTotalCaseCount
	Dim intI
	Dim strTempTstDetails
	Dim strFlag
	Dim strTemp
	Dim strqcTempStatus
	Dim pTestLst
	Dim pTestItem

	strQCLabPath = Environment("strQCLabPath")
	UID = Environment("UID")
	Pwd = Environment("Pwd")
	Domain =  Environment("Domain")
	Project = Environment("Project")
	strTestStatusDetails = strTestName&"--"&strTestcaseStatus
	
	Set qcConnect = CreateObject("TDAPIOLE80.TDConnection")
	qcConnect.InitConnectionEx Environment("ALMURL")
	qcConnect.Login  UID, Pwd
	qcConnect.Connect Domain , Project

	If (qcConnect.connected <> True) Then		
		MsgBox "QC Project Failed to Connect"			
	Else	
		Set qcTreeMgr = qcConnect.TreeManager 
		Set qcTestFactory = qcConnect.TestFactory 
		Set qcLabTreeMgr = qcConnect.TestSetTreeManager
		Set qcLabFolder = qcLabTreeMgr.NodeByPath( strQCLabPath )
		Set qcTestSetFactory = qcLabFolder.TestSetFactory
		Set qcFilter = qcTestSetFactory.Filter
		qcFilter.Filter("CY_CYCLE") = "'"& strTestSetName &"'" 'strQCTestName
	
		Set qcLst = qcTestSetFactory.NewList(qcFilter.Text)
		Set qcTestSet = qcLst.Item(1)
	
		strTempTstDetails = Split( strTestStatusDetails, "--" )
		strFlag = True
		Set qctestInstanceFctry = qcTestSet.TSTestFactory
		Set pTestItem = qctestInstanceFctry.Filter 		
		pTestItem.Filter("TS_NAME") =  "'"& strTestName &"'"'Trim( CStr( strTempTstDetails( 0 ) ) )
		Set pTestLst = qctestInstanceFctry.NewList( pTestItem.Text )
		Set pTestInTestSetObj = pTestLst.Item( 1 )
		
		strqcTempStatus = Trim( CStr( strTempTstDetails( 1 ) ) )
		pTestInTestSetObj.Status = strTestcaseStatus 
		pTestInTestSetObj.Post
		
		Set pTestItem = Nothing
		Set pTestLst = Nothing
		Set pTestInTestSetObj = Nothing
				
		Set qctestInstanceFctry = Nothing
		Set qcTestSet = Nothing
		Set qcLst = Nothing
		Set qcFilter = Nothing
		Set qcTestSetFactory = Nothing
		Set qcLabFolder = Nothing
		Set qcLabTreeMgr = Nothing
		Set qcTestFactory = Nothing
		Set qcTreeMgr = Nothing
		Set qcConnect = Nothing		
	End If
		
End Function

'------------------------------------------------------- End Of Fucctions -----------------------------------------------------------
