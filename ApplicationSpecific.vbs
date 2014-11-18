' ------------------------------------------------ List of Functions ------------------------------------------------------------
'List  of functions used below

'BuildXML
'ParseXML
'XMLCompare
'------------------------------------------------------- End Of List -----------------------------------------------------------

'*************************************************************************************************************
' Function : Build XML
' Functionality : <description >
'*************************************************************************************************************
Function BuildXML ( )

PPCService = WebService("ppcService").PPCService( Trim( Environment("TRN")), Trim( Environment("PSD")),Trim( Environment("ClrChannel")),Trim( Environment("OBI")),Trim( Environment("OI")),Trim( Environment("sndrSWIFTBIC")),Trim( Environment("ST")),Trim( Environment("Xrate")), Trim( Environment("DBTBRC")),Trim( Environment("CDTBRC")),Trim( Environment("DrFl")),Trim( Environment("DCID")),Trim( Environment("DAcctNo")), Trim( Environment("DAccBranch")),Trim( Environment("DAccType")),Trim( Environment("DAccDDA")),Trim( Environment("DtrnAmt")),Trim( Environment("DTRNCur")),Trim( Environment("DAccEntity")),Trim( Environment("DAccLOB")) , Trim( Environment("CrFl")),Trim( Environment("CCID")),Trim( Environment("CAcctNo")),Trim( Environment("CAccBranch")),Trim( Environment("CAccType")),Trim( Environment("CtrnAmt")),Trim( Environment("CTRNCur")),Trim( Environment("CAccEntity")),Trim( Environment("CAccLOB")),Trim( Environment("msgType")),Trim( Environment("sndrRefer")),Trim( Environment("RelReference")),Trim( Environment("InstructCode")) , Trim( Environment("InterAmt")),Trim( Environment("InterCurrency")),Trim( Environment("valDate")),Trim( Environment("InstAmt")),Trim( Environment("InstCurrency")), Trim( Environment("OrderID")),Trim( Environment("OrderBIC")),Trim( Environment("OrderAdd1")),Trim( Environment("OrderAdd2")),Trim( Environment("OrderAdd3")), Trim( Environment("OrderAdd4")),Trim( Environment("InstID")),Trim( Environment("InstBIC")),Trim( Environment("InstAdd1")), Trim( Environment("InstAdd2")),Trim( Environment("InstAdd3")),Trim( Environment("InstAdd4")),Trim( Environment("corresID")),Trim( Environment("CorresBIC")),Trim( Environment("CorresAdd1")),Trim( Environment("CorresAdd2")), Trim( Environment("CorresAdd3")),Trim( Environment("CorresAdd4")),Trim( Environment("RecvID")),Trim( Environment("RecvBIC")),Trim( Environment("RecvAdd1")),Trim( Environment("RecvAdd2")),Trim( Environment("RecvAdd3")),Trim( Environment(("RecvAdd4")),Trim( Environment("MedID")),Trim( Environment(("MedBIC")),Trim( Environment("MedAdd1")),Trim( Environment("MedAdd2")),Trim( Environment("MedAdd3")),Trim( Environment("MedAdd4")),Trim( Environment("AcctID")),Trim( Environment("AcctBIC")),Trim( Environment("AcctAdd1")),Trim( Environment("AcctAdd2")),Trim( Environment("AcctAdd3")),Trim( Environment("AcctAdd4")),Trim( Environment("BeneID")),Trim( Environment("BeneBIC")),Trim( Environment("BeneAdd1")),Trim( Environment("BeneAdd2")),Trim( Environment("BeneAdd3")),Trim( Environment("BeneAdd4")),Trim( Environment("Remit1")),Trim( Environment("Remit2")),Trim( Environment("Remit3")),Trim( Environment("Remit4")),Trim( Environment("DtlCharge")),Trim( Environment("sndrAmt")),Trim( Environment("sndrCurrency")),"RecAmt")),Trim( Environment("RecCurrency")),Trim( Environment("sndrtorcv1")),Trim( Environment("sndrtorcv2")), Trim( Environment("sndrtorcv3")),Trim( Environment("sndrtorcv4")),Trim( Environment("sndrtorcv5")),Trim( Environment("sndrtorcv6")),Trim( Environment("msgRef")) )
'PPCService = WebService("ppcService").PPCService( Trim( Environment("TRN")), Trim( Environment("PSD")),Trim( Environment("ClrChannel")),Trim( Environment("OBI")),Trim( Environment("OI")),Trim( Environment("sndrSWIFTBIC")),"",Trim( Environment("Xrate")), Trim( Environment("DBTBRC")),Trim( Environment("CDTBRC")),Trim( Environment("DrFl")),Trim( Environment("DCID")),Trim( Environment("DAcctNo")), Trim( Environment("DAccBranch")),Trim( Environment("DAccType")),Trim( Environment("DAccDDA")),Trim( Environment("DtrnAmt")),Trim( Environment("DTRNCur")),Trim( Environment("DAccEntity")),Trim( Environment("DAccLOB")) , Trim( Environment("CrFl")),Trim( Environment("CCID")),Trim( Environment("CAcctNo")),Trim( Environment("CAccBranch")),Trim( Environment("CAccType")),Trim( Environment("CtrnAmt")),Trim( Environment("CTRNCur")),Trim( Environment("CAccEntity")),Trim( Environment("CAccLOB")),Trim( Environment("msgType")),Trim( Environment("sndrRefer")),"",Trim( Environment("InstructCode")) , Trim( Environment("InterAmt")),Trim( Environment("InterCurrency")),Trim( Environment("valDate")),Trim( Environment("InstAmt")),Trim( Environment("InstCurrency")), Trim( Environment("OrderID")),Trim( Environment("OrderBIC")),Trim( Environment("OrderAdd1")),Trim( Environment("OrderAdd2")),Trim( Environment("OrderAdd3")), Trim( Environment("OrderAdd4")),Trim( Environment("InstID")),Trim( Environment("InstBIC")),Trim( Environment("InstAdd1")), Trim( Environment("InstAdd2")),Trim( Environment("InstAdd3")),Trim( Environment("InstAdd4")),Trim( Environment("corresID")),Trim( Environment("CorresBIC")),Trim( Environment("CorresAdd1")),Trim( Environment("CorresAdd2")), Trim( Environment("CorresAdd3")),Trim( Environment("CorresAdd4")),Trim( Environment("RecvID")),Trim( Environment("RecvBIC")),Trim( Environment("RecvAdd1")),Trim( Environment("RecvAdd2")),Trim( Environment("RecvAdd3")),Trim( Environment(("RecvAdd4")),Trim( Environment("MedID")),Trim( Environment(("MedBIC")),Trim( Environment("MedAdd1")),Trim( Environment("MedAdd2")),Trim( Environment("MedAdd3")),Trim( Environment("MedAdd4")),Trim( Environment("AcctID")),Trim( Environment("AcctBIC")),Trim( Environment("AcctAdd1")),Trim( Environment("AcctAdd2")),Trim( Environment("AcctAdd3")),Trim( Environment("AcctAdd4")),Trim( Environment("BeneID")),Trim( Environment("BeneBIC")),Trim( Environment("BeneAdd1")),Trim( Environment("BeneAdd2")),Trim( Environment("BeneAdd3")),Trim( Environment("BeneAdd4")),Trim( Environment("Remit1")),Trim( Environment("Remit2")),Trim( Environment("Remit3")),Trim( Environment("Remit4")),Trim( Environment("DtlCharge")),Trim( Environment("sndrAmt")),Trim( Environment("sndrCurrency")),"RecAmt")),Trim( Environment("RecCurrency")),Trim( Environment("sndrtorcv1")),Trim( Environment("sndrtorcv2")), Trim( Environment("sndrtorcv3")),Trim( Environment("sndrtorcv4")),Trim( Environment("sndrtorcv5")),Trim( Environment("sndrtorcv6")),Trim( Environment("msgRef")) )
'WebService("ppcService").Check CheckPoint("PPCService")


End Function

'*************************************************************************************************************
' Function : Parse XML
' Functionality : 
'*************************************************************************************************************
Function ParseXML ( )

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(InputRequest, ForReading)
	sInFileContent =   f.ReadAll
	f.close
	Set http = CreateObject("MSXML2.ServerXMLHTTP")
	' Set the method, Webservice URL and asynchronous
	http.open "POST", sWebserviceURL,False
	' Set the Header
	http.setRequestHeader "content-type","application/xml"
	' Set the HTTPS Authentication encryped in base 64 format
	http.setRequestHeader "Authorization", "Basic KJIOJ"
	' Post the XML
	http.send sInFileContent
	' Response Message Recieved
	OutputResponse=http.ResponseText
	' Save the Resposne File
	Set f = fso.OpenTextFile(sOutputXML , ForWriting, True)
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
Set oXML=MyXMLReader.Create("C:\Documents and Settings\xecctfz.AMS\Desktop\EPH\PPC Web Service Automation\InputData\Request_OPF-PPC_SampleData.xml")
Count1 = 1
while (oXML.Read())
If oXML.NodeType = "XmlDeclaration" or oXML.NodeType = "Element"Then
	If Count1 = 1 Then
		nodeVal =  oXML.Name  & ";"  & oXML.Depth
		'nodeDepth = oXML.Depth
		Count1 = 0
	Else
		nodeVal = nodeVal & "--" & oXML.Name  & ";"  & oXML.Depth
	End If
end if
wend
 print nodeVal
oXML.Close()
Set oXML = Nothing
Set MyXMLReader = Nothing

ActualFileName = "C:\Documents and Settings\xecctfz.AMS\Desktop\EPH\PPC Web Service Automation\InputData\Request_OPF-PPC1.xml" 
ExpectedFileName = "C:\Documents and Settings\xecctfz.AMS\Desktop\EPH\PPC Web Service Automation\InputData\Request_OPF-PPC_SampleData.xml"
Set xmlUtilObj1 = XMLUtil.CreateXML
xmlUtilObj1.LoadFile ActualFileName
Set xmlUtilObj2 = XMLUtil.CreateXML
xmlUtilObj2.LoadFile ExpectedFileName
arrnodeVal =  split ( nodeVal , "--" , -1 )

For nodeValCount  = 0 to Ubound ( arrnodeVal) -1
	arrnodevalsplit  = split (  arrnodeVal ( nodeValCount )  , ";" ,-1)
	DepthVal =  arrnodevalsplit (1)

	If DepthVal =0  Then
		ZeroNodeVal =  arrnodevalsplit (0)
		PathBuild  = "/" &  arrnodevalsplit (0)
	ElseIf DepthVal =1  Then
		FirstNodeVal =  arrnodevalsplit (0)
		PathBuild  = "/" &  ZeroNodeVal & "/" & arrnodevalsplit (0)
    ElseIf DepthVal = 2  Then
		SecondNodeVal =  arrnodevalsplit (0)
			PathBuild  ="/" &  ZeroNodeVal &  "/" & FirstNodeVal  & "/"&  arrnodevalsplit (0)
	ElseIf DepthVal = 3  Then
		ThirdNodeVal =  arrnodevalsplit (0)
		 PathBuild  = "/" &  ZeroNodeVal & "/" & FirstNodeVal  &  "/" & SecondNodeVal  & "/"&  arrnodevalsplit (0)
	ElseIf DepthVal = 4 Then
		FourthNodeVal =  arrnodevalsplit (0)
		 PathBuild  ="/" &  ZeroNodeVal &  "/" & FirstNodeVal  &  "/" & SecondNodeVal  & "/"  & ThirdNodeVal  & "/"  &  arrnodevalsplit (0)
	End If
	'Set xmlChildList1 = xmlUtilObj1.ChildElementsByPath( "/soapenv:Envelope/soapenv:Body/ns:PPC_REQUEST/"& arrnodevalsplit ( 0 ) )
    'Set xmlChildList2 = xmlUtilObj2.ChildElementsByPath( "/soapenv:Envelope/soapenv:Body/ns:PPC_REQUEST/"& arrnodevalsplit ( 0 ) )


	Set xmlChildList1 = xmlUtilObj1.ChildElementsByPath( PathBuild )
    Set xmlChildList2 = xmlUtilObj2.ChildElementsByPath( PathBuild )
	' Traverse through the XML File to read ans display the contents
'	For i = 1 to xmlChildList1.count
'	If i <> 1 Then
	'ActualValue = ActualValue & "--" &  xmlChildList1.item(i).value
'	Else
	ActualValue =  xmlChildList1.item(1).value
'	End If
	'Next
'	For i = 1 to xmlChildList2.count
'		If i <> 1 Then
	'		ExpValue = ExpValue & "--" &  xmlChildList2.item(i).value
	'	Else
			ExpValue =  xmlChildList2.item(1).value
		
		'End If
	'Next

If  ActualValue =  ExpValue Then 
	
Call Fn_TableName(arrnodevalsplit (0) & ";"& ExpValue  & ";"& ActualValue    & ";"&  "Pass" , "Pass", Environment("strCurModule") )
Else
	
Call Fn_TableName(arrnodevalsplit (0) & ";"& ExpValue  & ";"& ActualValue    & ";"&  "Fail" , "Fail", Environment("strCurModule") )
End If

Next

'Destroy the XMLUtil Object
Set xmlUtilObj = Nothing



End Function

'TRN = "string (Autogenerated)"
'TRN,"string (Autogenerated)","string (Autogenerated)","string (Autogenerated)","string (Autogenerated)","string (Autogenerated)","string (Autogenerated)",0,True,"string (Autogenerated)","string (Autogenerated)",XMLWarehouse("ArrayOfDbtCrdt"),"string (Autogenerated)","string (Autogenerated)","string (Autogenerated)","string (Autogenerated)",XMLWarehouse("AmtCurr"),#03/21/2013 14:28:26#,XMLWarehouse("AmtCurr1"),XMLWarehouse("RequestAcctDetails"),XMLWarehouse("RequestAcctDetails1"),XMLWarehouse("RequestAcctDetails2"),XMLWarehouse("RequestAcctDetails3"),XMLWarehouse("RequestAcctDetails4"),XMLWarehouse("RequestAcctDetails5"),XMLWarehouse("RequestAcctDetails6"),XMLWarehouse("RemittanceInfo"),"string (Autogenerated)",XMLWarehouse("ArrayOfAmtCurr"),XMLWarehouse("AmtCurr2"),XMLWarehouse("SndrToRecvInfo"),"string (Autogenerated)",CurrentDateTime,OutChargeInd,RecvChrgPLAcc,Service)

