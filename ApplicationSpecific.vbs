' ------------------------------------------------ List of Functions ------------------------------------------------------------
'List  of functions used below

'BuildXML
'ParseXML
'XMLCompare
'RemoveTags
'------------------------------------------------------- End Of List -----------------------------------------------------------

'*************************************************************************************************************
' Function : Build XML
' Functionality : <description >
'*************************************************************************************************************
Function BuildXML ( )

'PPCService = WebService("ppcService").PPCService( Trim( Environment("TRN")), Trim( Environment("PSD")),Trim( Environment("ClrChannel")),Trim( Environment("OBI")),Trim( Environment("OI")),Trim( Environment("sndrSWIFTBIC")),Trim( Environment("ST")),Trim( Environment("Xrate")), Trim( Environment("DBTBRC")),Trim( Environment("CDTBRC")),Trim( Environment("DrFl")),Trim( Environment("DCID")),Trim( Environment("DAcctNo")), Trim( Environment("DAccBranch")),Trim( Environment("DAccType")),Trim( Environment("DAccDDA")),Trim( Environment("DtrnAmt")),Trim( Environment("DTRNCur")),Trim( Environment("DAccEntity")),Trim( Environment("DAccLOB")) , Trim( Environment("CrFl")),Trim( Environment("CCID")),Trim( Environment("CAcctNo")),Trim( Environment("CAccBranch")),Trim( Environment("CAccType")),Trim( Environment("CtrnAmt")),Trim( Environment("CTRNCur")),Trim( Environment("CAccEntity")),Trim( Environment("CAccLOB")),Trim( Environment("msgType")),Trim( Environment("sndrRefer")),Trim( Environment("RelReference")),Trim( Environment("InstructCode")) , Trim( Environment("InterAmt")),Trim( Environment("InterCurrency")),Trim( Environment("valDate")),Trim( Environment("InstAmt")),Trim( Environment("InstCurrency")), Trim( Environment("OrderID")),Trim( Environment("OrderBIC")),Trim( Environment("OrderAdd1")),Trim( Environment("OrderAdd2")),Trim( Environment("OrderAdd3")), Trim( Environment("OrderAdd4")),Trim( Environment("InstID")),Trim( Environment("InstBIC")),Trim( Environment("InstAdd1")), Trim( Environment("InstAdd2")),Trim( Environment("InstAdd3")),Trim( Environment("InstAdd4")),Trim( Environment("corresID")),Trim( Environment("CorresBIC")),Trim( Environment("CorresAdd1")),Trim( Environment("CorresAdd2")), Trim( Environment("CorresAdd3")),Trim( Environment("CorresAdd4")),Trim( Environment("RecvID")),Trim( Environment("RecvBIC")),Trim( Environment("RecvAdd1")),Trim( Environment("RecvAdd2")),Trim( Environment("RecvAdd3")),Trim( Environment(("RecvAdd4")),Trim( Environment("MedID")),Trim( Environment(("MedBIC")),Trim( Environment("MedAdd1")),Trim( Environment("MedAdd2")),Trim( Environment("MedAdd3")),Trim( Environment("MedAdd4")),Trim( Environment("AcctID")),Trim( Environment("AcctBIC")),Trim( Environment("AcctAdd1")),Trim( Environment("AcctAdd2")),Trim( Environment("AcctAdd3")),Trim( Environment("AcctAdd4")),Trim( Environment("BeneID")),Trim( Environment("BeneBIC")),Trim( Environment("BeneAdd1")),Trim( Environment("BeneAdd2")),Trim( Environment("BeneAdd3")),Trim( Environment("BeneAdd4")),Trim( Environment("Remit1")),Trim( Environment("Remit2")),Trim( Environment("Remit3")),Trim( Environment("Remit4")),Trim( Environment("DtlCharge")),Trim( Environment("sndrAmt")),Trim( Environment("sndrCurrency")),"RecAmt")),Trim( Environment("RecCurrency")),Trim( Environment("sndrtorcv1")),Trim( Environment("sndrtorcv2")), Trim( Environment("sndrtorcv3")),Trim( Environment("sndrtorcv4")),Trim( Environment("sndrtorcv5")),Trim( Environment("sndrtorcv6")),Trim( Environment("msgRef")) )
PPCService = WebService("ppcService").PPCService( Trim( Environment("TRN")), Trim( Environment("PSD")),Trim( Environment("ClrChannel")),Trim( Environment("OBI")),Trim( Environment("OI")),Trim( Environment("sndrSWIFTBIC")),"",Trim( Environment("Xrate")), Trim( Environment("DBTBRC")),Trim( Environment("CDTBRC")),Trim( Environment("DrFl")),Trim( Environment("DCID")),Trim( Environment("DAcctNo")), Trim( Environment("DAccBranch")),Trim( Environment("DAccType")),Trim( Environment("DAccDDA")),Trim( Environment("DtrnAmt")),Trim( Environment("DTRNCur")),Trim( Environment("DAccEntity")),Trim( Environment("DAccLOB")) , Trim( Environment("CrFl")),Trim( Environment("CCID")),Trim( Environment("CAcctNo")),Trim( Environment("CAccBranch")),Trim( Environment("CAccType")),Trim( Environment("CtrnAmt")),Trim( Environment("CTRNCur")),Trim( Environment("CAccEntity")),Trim( Environment("CAccLOB")),Trim( Environment("msgType")),Trim( Environment("sndrRefer")),"",Trim( Environment("InstructCode")) , Trim( Environment("InterAmt")),Trim( Environment("InterCurrency")),Trim( Environment("valDate")),Trim( Environment("InstAmt")),Trim( Environment("InstCurrency")), Trim( Environment("OrderID")),Trim( Environment("OrderBIC")),Trim( Environment("OrderAdd1")),Trim( Environment("OrderAdd2")),Trim( Environment("OrderAdd3")), Trim( Environment("OrderAdd4")),Trim( Environment("InstID")),Trim( Environment("InstBIC")),Trim( Environment("InstAdd1")), Trim( Environment("InstAdd2")),Trim( Environment("InstAdd3")),Trim( Environment("InstAdd4")),Trim( Environment("corresID")),Trim( Environment("CorresBIC")),Trim( Environment("CorresAdd1")),Trim( Environment("CorresAdd2")), Trim( Environment("CorresAdd3")),Trim( Environment("CorresAdd4")),Trim( Environment("RecvID")),Trim( Environment("RecvBIC")),Trim( Environment("RecvAdd1")),Trim( Environment("RecvAdd2")),Trim( Environment("RecvAdd3")),Trim( Environment(("RecvAdd4")),Trim( Environment("MedID")),Trim( Environment(("MedBIC")),Trim( Environment("MedAdd1")),Trim( Environment("MedAdd2")),Trim( Environment("MedAdd3")),Trim( Environment("MedAdd4")),Trim( Environment("AcctID")),Trim( Environment("AcctBIC")),Trim( Environment("AcctAdd1")),Trim( Environment("AcctAdd2")),Trim( Environment("AcctAdd3")),Trim( Environment("AcctAdd4")),Trim( Environment("BeneID")),Trim( Environment("BeneBIC")),Trim( Environment("BeneAdd1")),Trim( Environment("BeneAdd2")),Trim( Environment("BeneAdd3")),Trim( Environment("BeneAdd4")),Trim( Environment("Remit1")),Trim( Environment("Remit2")),Trim( Environment("Remit3")),Trim( Environment("Remit4")),Trim( Environment("DtlCharge")),Trim( Environment("sndrAmt")),Trim( Environment("sndrCurrency")),"RecAmt")),Trim( Environment("RecCurrency")),Trim( Environment("sndrtorcv1")),Trim( Environment("sndrtorcv2")), Trim( Environment("sndrtorcv3")),Trim( Environment("sndrtorcv4")),Trim( Environment("sndrtorcv5")),Trim( Environment("sndrtorcv6")),Trim( Environment("msgRef")) )
WebService("ppcService").Check CheckPoint("PPCService")


End Function

'*************************************************************************************************************
' Function : Parse XML
' Functionality : 
'*************************************************************************************************************
Function PostXML ( )

'Environment("sSavedXML") =  "C:\Documents and Settings\xecctfz.AMS\Desktop\EPH\PPC Web Service Automation\InputData\Request_OPF-PPC1.xml"
sOutputXML = Replace ( Environment("sSavedXML") ,Environment("TC_NO"), Environment("TC_NO")& "_Output" )

Environment("Output")  = sOutputXML
'sOutputXML =  "C:\Documents and Settings\xecctfz.AMS\Desktop\EPH\PPC Web Service Automation\InputData\Request_OPF-PPC2.xml"
' Read the request file
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile( Environment("sSavedXML") , 1)
sInFileContent =   f.ReadAll
f.close
Set http = CreateObject("MSXML2.ServerXMLHTTP")
' Set the method, Webservice URL and asynchronous
'http.open "POST", "http://r7vpn0c:8181/ppcservice/services/ppcPort?wsdl" , False


StartTime0 = Timer
http.open "POST", 	Trim ( Environment("AppPath")  ) 
StartTime1 = Timer
http.send sInFileContent
StartTime2 = Timer
' Response Message Recieved
OutputResponse=http.ResponseText

EndTime = Timer
 TimeIt0 = EndTime - StartTime0
 TimeIt = EndTime - StartTime1
 TimeIt2 = EndTime - StartTime2
 Print "Response time = : " & TimeIt0
Print "Response time = : " & TimeIt
Print "Response time = : " & TimeIt2
' Save the Resposne File
 f.Close
Set f = fso.OpenTextFile(sOutputXML , 2, True)
f.Write(OutputResponse)
f.close
set f= nothing
set fso=nothing



End Function


'*************************************************************************************************************
' Function : XMLCompare
' Functionality : 
'*************************************************************************************************************
Function XMLCompare ( )

Set MyXMLReader = DotNetFactory.CreateInstance("System.Xml.XmlReader", "System.Xml")
Set oXML=MyXMLReader.Create(Environment("Output") )
Set objWorkSheet1 =  Environment ( "objFlowWorkBook" ).Worksheets("ResponseData"   )
Set oScript = CreateObject( "Scripting.Dictionary" )
CountVal = 1
While (oXML.Read())
    TagValue =  oXML.Value
	If   oXML.Depth.tostring() = 1 then
		NodeVal1 = oXML.Name
	ElseIf   oXML.Depth.tostring() = 2 then
		NodeVal2 = oXML.Name
		If  TagValue <>"" Then
	   '     Print "Name :" &  NodeVal1 & "_" &   oXML.Name   & "is " &  TagValue
		End If
	ElseIf   oXML.Depth.tostring() = 3 then
		NodeVal3 = oXML.Name
			If  TagValue <>"" Then
		'   	Print "Name :"&  NodeVal1 & "_" &  NodeVal2 & "_" &  oXML.Name   & "is " &  TagValue
			End If
		
	ElseIf   oXML.Depth.tostring() = 4 then
		NodeVal4 = oXML.Name
		If  TagValue <>"" Then
		   '	Print "Name :" &  NodeVal1 & "_" &  NodeVal2 & "_" &  NodeVal3 & "_" & oXML.Name   & "is " &  TagValue
			ExtValue  = ExtValue &  "Name :" &  NodeVal3 & "_" & oXML.Name   & "is " &  TagValue
		End If
			
	ElseIf   oXML.Depth.tostring() = 5 then
		NodeVal5 = oXML.Name
		If  TagValue <>"" Then
		   '	Print "Name :" &  NodeVal1 & "_" &  NodeVal2 & "_" &  NodeVal3 & "_" &  NodeVal4 & "_" & oXML.Name   & "is " &  TagValue
			ExtValue  = ExtValue &   "Name :" &  NodeVal3 & "_" &  NodeVal4 & "_" & oXML.Name   & "is " &  TagValue
		End If
	ElseIf   oXML.Depth.tostring() = 6 then
		NodeVal6 = oXML.Name
		If  TagValue <>"" Then
			'Print "Name :" &  NodeVal1 & "_" &  NodeVal2 & "_" &  NodeVal3 & "_" &  NodeVal4 & "_" & NodeVal5 & "_" & oXML.Name   & "is " &  TagValue
			ExtValue  = ExtValue &	 "Name :"  & NodeVal3 & "_" &  NodeVal4 & "_" & NodeVal5 & "_" & oXML.Name   & "is " &  TagValue
		End If
    End If
Wend

ActValue = Split ( ExtValue , "Name :" ,-1,1)
For i = 0 to ubound ( ActValue ) 
	If  ActValue (i)  <>  "" Then
		NameandValue  = Split ( ActValue (i)  , "_is" ,2,1)
		If  Lcase ( NameandValue(0)  ) =  "service_servicename"  Then
			ServiceNameVal =   NameandValue(1)
			IDCountVal = 0
			DataCountVal = 0
			WPCountVal = 0
		End If
		If  Instr  ( 1, Lcase ( NameandValue(0)  )  , "service_" ) > 0 Then 
			NameandValue(0) = Replace  ( NameandValue(0) , "Service_"  ,ServiceNameVal & "_" ,  1,1)
		End If

If Instr ( 1, Trim ( NameandValue(0) ) ,  "MatchedCriteria_FieldId" ) > 0 Then
NewMatch  = "MatchedCriteria" & IDCountVal & "_"
NameandValue(0) = Replace ( Trim ( NameandValue(0) ), "MatchedCriteria_" , NewMatch  )
IDCountVal = IDCountVal + 1
ElseIf  Instr ( 1, Trim ( NameandValue(0) ) ,  "MatchedCriteria_FieldData" ) > 0 Then
NewMatch  = "MatchedCriteria" & DataCountVal & "_"
NameandValue(0) = Replace ( Trim ( NameandValue(0) ), "MatchedCriteria_" , NewMatch  )
WPCountVal =  DataCountVal
DataCountVal =  DataCountVal  + 1

ElseIf Instr ( 1, Trim ( NameandValue(0) ) ,  "MatchedCriteria_WildcardPattern" ) > 0 Then
NewMatch  = "MatchedCriteria" & WPCountVal & "_"
NameandValue(0) = Replace ( Trim ( NameandValue(0) ), "MatchedCriteria_" , NewMatch  )
WPCountVal =  DataCountVal
End If
oScript.Add Trim ( NameandValue(0) ) , Trim( NameandValue(1) )
'Print NameandValue(0)
	End If
Next
oXML.Close()
Set oXML = Nothing
Set MyXMLReader = Nothing

intCurModColCn1 = objWorkSheet1.UsedRange.Columns.Count
' Iteration starts from 3rd column as 1st column is for test case number and second column is for Scenario name
For B = 3 to intCurModColCn1
	BaseColVal  = Trim( objWorkSheet1.Rows( Environment("CurrentRow")  ).Columns( B  ).Value )
	If BaseColVal <> "" Then
		BaseColHeadVal  = Trim( objWorkSheet1.Rows( 1  ).Columns( B  ).Value )
		If oScript.Exists( Trim( BaseColHeadVal )) Then
			ActColVal =  Trim ( oScript.Item( Trim( BaseColHeadVal ) )  )

			If  Trim ( BaseColVal )  = ActColVal   Then 
				Call Fn_TableName(BaseColHeadVal  & ";"& BaseColVal  & ";"& ActColVal    & ";"&  "Pass" , "Pass", Environment("strCurModule") )
			Else
				Call Fn_TableName(BaseColHeadVal  & ";"& BaseColVal  & ";"& ActColVal    & ";"&  "Fail" , "Fail", Environment("strCurModule") )
				Environment("CurTestResult") = False
			End If
			oScript.Item ( Trim( BaseColHeadVal ) )  = "Available"
		Else 
			Temp1 = Temp1 & BaseColHeadVal & vbcrlf
			Print  BaseColHeadVal & "not Exists"

		End If
	End If
Next 

i = oScript.Items 
k = oScript.Keys
Print " ********************* XML Values not defined in Excel as baseline***********************"
For x = 0 To oScript.Count-1
If i(x) <> "Available" Then
	If  Instr( 1 ,   k(x), "CurrentDateTime" ) > 0 Then

	Else
		Print k(x)
		NotAv = k(x) & k(x) & Vbcrlf
	End If
End If
Next
If NotAv <>"" Then
	Call Fn_TableName("XML Values present but  not defined in baseline ;Following fields should not contain values"  & ";"& NotAv    & ";"&  "Fail" , "Fail", Environment("strCurModule") )
	Environment("CurTestResult") = False
End If
If  Temp1 <> "" Then
	Call Fn_TableName( "XML Values not  present but   defined in baseline;Following fields should  contain values"  & ";"& Temp1    & ";"&  "Fail" , "Fail", Environment("strCurModule") )
	Environment("CurTestResult") = False
End If
Set xmlUtilObj = Nothing

End Function

'*************************************************************************************************************
' Function : RemoveTags
' Functionality : 
'*************************************************************************************************************

Function RemoveTags
	Const ForReading = 1
	Const ForWriting = 2
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile( Environment("sSavedXML")  , ForReading)
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		If InStr(strLine, "RemoveTag") = 0 Then
			strNewContents = strNewContents & strLine & vbCrLf
		End If
	Loop
	objFile.Close
	Set objFile = objFSO.OpenTextFile(Environment("sSavedXML")   , ForWriting)
	objFile.Write strNewContents
	objFile.Close
End Function
