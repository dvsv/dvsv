'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Project				::	EPH PPC Webservices Automation
'Module					::	Driver Module
'Description			::	  
'Tool					::	QTP  11.0
'Application			::	NA
'FrameWork				::	Modular Driven FrameWork
'Author					::	Sivasankar Karunagaran
'Functions Called		::	
'Last updated on		::	26 - Mar - 2013
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim objFlowExcel
Dim objFlowWorkBook
Dim objWorkSheet
Dim strCurModule
Dim intCurrRow
Dim TestLocation
Dim oFso
Dim arrModules
Dim strColName
Dim TestCaseFlag
Dim PerformTest
Dim intI

'Environment("sSavedXML") =  "C:\Documents and Settings\xecctfz.AMS\Desktop\EPH\PPC Web Service Automation\InputData\Request_OPF-PPC1.xml"
'  Following lines are coded to get the relative path of the test placed 
Set oFso = CreateObject( "Scripting.FileSystemObject" )
Environment( "TestLocation" ) = oFso.GetParentFolderName( Environment( "TestDir" ) )

'  Fn_Interface function is located Framework_Specific function file which is used to get the input from user on running the test ( HTML Popup will be displayed )
Call Fn_Interface (  )
'  Fn_Config  is located Framework_Specific function file which is used to import Object Repository and external environment variables defined in XML file
Call Fn_Configuration ( "SetUp" )

' Open only  if the Datatable exist in the location specified
If oFso.FileExists( Environment( "DatatableLocation" ) ) = True Then

' Create an object to open the  datatable
	Set objFlowExcel = CreateObject( "Excel.Application" )
	Environment("objFlowExcel") = objFlowExcel
	Set objFlowWorkBook = Environment( "objFlowExcel" ).Workbooks.Open ( Environment( "DatatableLocation" ) )
	Environment( "objFlowWorkBook" ) = objFlowWorkBook
	Environment("objFlowExcel").Visible = True

' IIterate the number of modules defined 
	arrModules = Split(Environment( "Modules" ), "@@", -1, 1)

' ************************************************ Iteration starts at module level  ( List of modules are available in interface ) ***********************************************************************************************
    For i = 0 to Ubound(arrModules)
		If Trim(arrModules ( i ) ) <> ""  Then	

			Call Fn_Configuration ( "ModuleStart" )
			Environment( "strCurModule" ) = Trim( arrModules ( i ) )

'  Fn_CreateResultFile is located  Generic VBscript  function file which is used to create the output file in HTML/Excel format 
			Call Fn_CreateResultFile( )

			Set objWorkSheet = objFlowWorkBook.Worksheets( Trim ( arrModules ( i ) ) ) 
			Environment("CurrModuleSheet") = objWorkSheet

'  Following line execute the macro defined in the datatable . This macro will set the number of test cases that needs to run on the module that is been executed.
			Set PerformTest = Fn_TestCaseSelector( objWorkSheet, Trim( arrModules( i ) ) )

'  Following lines will fetch the Used rows and columns count in the module sheet  and assigned them to an environemnt variable
			intCurModRowCnt = objWorkSheet.UsedRange.Rows.Count
			intCurModColCnt = objWorkSheet.UsedRange.Columns.Count
			Environment("intCurModRowCnt") = intCurModRowCnt
			Environment("intCurModColCnt") = intCurModColCnt 
			'-----------------------------------------------------\
				'Open the File and read the content
	
	''''''Loop the value recieved from Excel sheet to change the XML
	' Changing the value in the XML

		

	'Save the modified file as different file
	'Set fso = CreateObject("Scripting.FileSystemObject")



' ************************************************   Iteration starts at row level ( Each row is a test case ) ***********************************************************************************************************************
			For rows = 2 to intCurModRowCnt

				Set fso = CreateObject("Scripting.FileSystemObject")
				Set f = fso.OpenTextFile(Environment ("sInputXML" ), 1)
				InputFileContent =   f.ReadAll
				f.close

				Environment("Driver") = True
				Environment("TestCaseFlag") = True
				Environment("CurrentRow") = rows
				
				For intI = 1 To PerformTest.Count
					TestCaseFlag = False

' ************************************************   Environment  varaible started storing  in Environment header ( Row 1 of Data sheet is header and  Test case row is varaiable ) ************************************
					For cols = 1 to intCurModColCnt

' If the test case row is not required to run stop assigning environment variables for that test case
						If  Trim( objWorkSheet.Rows( rows ).Columns( cols ).Value ) <> Trim( PerformTest.Item( "Case"& intI ) ) And cols = 1 Then
							TestCaseFlag = False
							Exit For
						End If
						strColName = Trim( objWorkSheet.Rows( 1 ).Columns( cols ).Value )

						  InputFileContent=Replace(InputFileContent,"{"& strColName & "}",Trim( objWorkSheet.Rows( rows ).Columns( cols ).Value))
						If  objWorkSheet.Rows( rows ).Columns( cols ).Value  <> "" Then
							Environment( strColName ) = Trim( objWorkSheet.Rows( rows ).Columns( cols ).Value )
						Else 
							Environment( strColName ) = " "
	
						End If
						TestCaseFlag = True
'				
				Next
				
' ************************************************   Environment  varaible completed storing  in  Environment header ( Row 1 of Data sheet is header and  Test case row is varaiable ) *****************************

' Start executing the script for the cases specified as Run
					If TestCaseFlag = True Then
						Print "*****************************************" &  Environment("TC_NO") & "********************************************"
						
						TestDir = Replace ( Environment("TestDir")  , Environment("TestName"),"")
						 Environment("sSavedXML") =   TestDir & "OutputFiles\" &  Environment("TC_NO")& ".xml"
						Set fso = CreateObject("Scripting.FileSystemObject")
						Set f = fso.CreateTextFile(Environment("sSavedXML") ,True)

						RemoveHeader = Split ( Environment("RemoveHeader") , "||" ,-1 ,1)
						For RemoveCount = 0 to Ubound ( RemoveHeader)
						InputFileContent	= Replace (  InputFileContent ,"<ns:"&RemoveHeader( RemoveCount ) & ">" ,""  ,1)
						InputFileContent	= Replace (  InputFileContent ,"</ns:"&RemoveHeader( RemoveCount ) & ">" ,""  ,1)
						Next
						f.Write(InputFileContent)
						f.close
						set f= nothing
						set fso=nothing
						RemoveTags
' Fn_Module  is located Framework_Specific function file which actually taken the function name from the driver sheet and start executing them
						Call Fn_TableName(Environment("TC_NO")& ";" & Environment("Description")  , "TCName", Environment("strCurModule") )
						Environment("CurTestResult") = True
						Call Fn_Module ( Trim( arrModules ( i ) ) )
						Environment( "Driver" ) = False
					End If
				
				Next
				
			Next
' Fn_Configuration   is located Framework_Specific function file which will close all open connections
			Call Fn_Configuration ( "ModuleEnd" )

' ************************************************   Iteration ends at row level ( Each row is a test case ) **********************************************************************************************************************
		End If
		Set objWorkSheet = Nothing
		Set PerformTest = Nothing
	Next
' ************************************************ Iteration ends at module level  ( List of modules are available in interface ) ************************************************************************************************	
	objFlowWorkBook.Close

' Throw an error if the path speciified is not correct
Else
	Msgbox "Unable to find the Datatable in the path"& Chr( 10 ) & Environment( "DatatableLocation" ),,""& Chr( 10 ) &"Test Run Stopped"
End If

'Close all Objects created
Set objFlowWorkBook = Nothing
Set objFlowExcel = Nothing
Set oFso = Nothing
'----------------------------------------------------------------End Of Driver Script---------------------------------------------------------------------------------------







