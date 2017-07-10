'
' Licensed Materials - Property of IBM
'
' (c) Copyright IBM Corporation 2008. All Rights Reserved. 
' * Note to U.S. Government Users Restricted Rights:  Use, duplication or 
' * disclosure restricted by GSA ADP Schedule Contract with IBM Corp.
' 
'
'#MAIN 
'
'  declare variables

	Dim http
	Dim doc
	Dim url, project, userid, passwd
	Dim query_service
	Dim update_service
	Dim factory_service
	Dim workItemId
	Dim workItemDoc
	Dim counter
	Dim parentID, childIDs(9)

'--------------
' Parameter definition
' CLM parameters
	url = "https://clm.example.com:9443/ccm"
	project = "TRADITIONAL"
	userid = "jazzadmin"
	passwd = "jazzadmin"
'-------------

	' Login to jazz application server
	Set http = JazzLogin(url, userid, passwd)

	' obtain service catalog from jazz root service.
	' obtain workitem factory (use default factory) 
	service_url = GetServicebyProjectName(http, url, project)
	factory_service = GetFactoryService(http, service_url)
	update_service = GetUpdateService(http, url)


	' Parent workitem create 
	attrString = "dcterms:type,task"
	attrString = attrString&","&"dcterms:title,This is parent"
	parentId = CreateWorkItem(http, factory_service, attrString)

	' Create 9 child workitems
	For counter = 1 to 9
		attrString = "dcterms:type,task"
		attrString = attrString&","&"dcterms:title,This is "&counter&" Child"
		childIds(counter-1) = CreateWorkItem(http, factory_service, attrString)
		

	Next
	
	' Then set child workitems to the parent
	call CreateParentChild(http, update_service, parentId, childIds)


	' Exit the script with return status 0 (zero)
	WScript.Quit 0

'#END MAIN


Public Function JazzLogin(url, userid, password)

	Dim jazzUserid
	Dim JazzPassword
	
	JazzUserid = "j_username=jazzadmin"
	JazzPassword = "j_password=jazzadmin"


	Set http = CreateObject("MSXML2.XMLHTTP")

	' login to jazz server specified in the parameter section.
	http.Open "GET", url&"/authenticated/identity", False
	http.Send

	http.Open "POST", url&"/authenticated/j_security_check", False
	http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	http.Send JazzUserid&"&"&JazzPassword

	Set JazzLogin = http


End Function


Public Function GetServicebyProjectName (http, url, projectName)

	Dim text
	Dim doc
	Dim xmldoc
	Dim element, elements
	

'	http.Open "GET", url&"/rootservices", False

	' Since Rational Team Concert workitem catalog service is fixed, skip rootservice 
	' checking. 
	http.Open "GET", url&"/oslc/workitems/catalog.xml", False
	http.Send

	Set doc = http.ResponseXML

	Set xmldoc = CreateObject("MSXML2.DOMDocument")

	xmldoc.loadXML(Http.ResponseText)
	
	'Obtain list of services from service provider for input project.
	set elements = xmldoc.getElementsByTagName("oslc_disc:ServiceProvider")
	
	For each element in elements
		If element.text = projectName Then
			
			set services = element.getElementsByTagName("oslc_disc:services")
			For each service in services

				' Parse service XML structure to get "rdf.resouce"
				' Attribute(0) is the resource value.
				service_url = service.attributes(0).nodeValue

			Next

		End If
	Next
	
	GetServicebyProjectName = service_url


End Function


Public Function GetQueryService (http, service_url)

	Dim url
	Dim doc
	Dim element, elements
	Dim service, services

'DebugPrint "GetQueryService[ENTER]"

	'Initialize url value
	url = ""
	

	http.Open "GET", service_url, False
	http.Send

	set doc = CreateObject("MSXML2.DOMDocument")
	
	doc.loadXML(http.responseText)
	

	set elements = doc.getElementsByTagName("oslc_cm:simpleQuery")

	For each element in elements

		If element.hasChildNodes then
		
			For each node in element.ChildNodes
			
				If node.nodeName = "oslc_cm:url" then
					
					url = node.text
					
				End If

			next
		
		End If
	
	Next

'DebugPrint "GetQueryService[EXIT]"

	GetQueryService = url


End Function

Public Function GetUpdateService (http, service_url)

	' [Note]. General method is to obtain resource URL for workitem.
	' But in this script, the procedure is skipped for efficiency.
	' resorce URL is known to work for RTC V3.0 and 4.0.
	
	GetUpdateService = service_url&"/resource/itemName/com.ibm.team.workitem.WorkItem"


End Function

Public Function GetFactoryService (http, service_url)

	Dim url
	Dim doc
	Dim element, elements
	Dim service, services

'DebugPrint "GetFactoryService[ENTER]"

	'Initialize url value
	url = ""
	

	http.Open "GET", service_url, False
	http.Send

	set doc = CreateObject("MSXML2.DOMDocument")
	
	doc.loadXML(http.responseText)
	

	set elements = doc.getElementsByTagName("oslc_cm:factory")

	' There are many factory services, select generic workitem creation factory
	' Currently, choosen method is to select factory which does not have calm:id
	' [TODO] elaborate more better way to find generic workitem creation factory
	For each element in elements

		' 
		If isNull( element.getAttribute("calm:id")) Then
		
			For each node in element.ChildNodes
			
				If node.nodeName = "oslc_cm:url" then

					url = node.text

					
				End If

			next
		
		End If
	
	Next


	GetFactoryService = url

'DebugPrint "GetFactoryService[EXIT]"

End Function

' QueryWorkItems requires two input 
'  query_url : base of query URL
'  query_string : specify oslc_where closure
'  result_string : specify colums of retrieval
'  Query language is defined in OSLC specification.
'  Sample expected string 
'  query = "dcterms:identifier=11"
'  result = "dcterms:identifier,dcterms:title,dcterms:description"
'
Public Function QueryWorkItems(query_url, query_string, result_string)

	Dim doc
	Dim attrArray, attrString
	Dim query, result
	Dim element, elements
	Dim collection

	' sample string to be explected.

	' Note that dcterms:identifier is always added as primary key of query

	If query_string = "" Then
		query="?"
	Else
		query = "?oslc.where="&query_string
	End If
	result = "&oslc.select="&"dcterms:identifier,"&result_string

	' Make query service to return XML.
	service = query_url&".xml"

	attrArray = split("dcterms:identifier,"&result_string, ",")
	
	' this is result set with key as workitem numer
	Set resultSet = CreateObject("Scripting.Dictionary")

	Http.Open "GET", service&query&result, False
	http.setRequestHeader "Content-Type", "application/xml"
	http.setRequestHeader "Accept", "application/xml"
	http.setRequestHeader "OSLC-Core-Version", "2.0"

	Http.Send

	set doc = Http.ResponseXML

	set elements = doc.getElementsByTagName("oslc_cm:ChangeRequest")
	
	For each element in elements
	
		Set collection = CreateObject("Scripting.Dictionary")

		For each attr in element.childNodes
		


			For i = 0 to UBound(attrArray)

				If attr.nodeName = attrArray(i) Then

					collection.Add attr.nodeName, attr.text


				End If

			Next
			


		Next
		
		key = collection.Keys

			
		'Add workitem number as key of collection
		attrString=collection.Item(key(0))
		For I = 1 to UBound(attrArray)

			attrString = attrString&","&collection.Item(key(I))

		Next		

		resultSet.Add collection.Item("dcterms:identifier"), attrString
		
		set collection = Nothing
		
	Next


	set QueryWorkItems = resultSet

End Function

Public Function DisplayResultSet (collection)

	'returns 0 if success (always)
	
	Dim key
	Dim i
	
' DebugPrint "DisplayResultSet[ENTER]"

	DisplayResultSet = 0
	count = 0

	key = collection.Keys
	
	
	For i = 0 to collection.Count - 1

		WScript.Echo key(i), collection.Item(key(I))

	Next



' DebugPrint "DisplayResultSet[EXIT]"

End Function

' input string must have the following format each are sperated by
' comma ","
'  <attribute>,<value>,<attribute>,<value>,....
' for example
'  dc:type,task,dc:title,this is title,dc:description,description
Public Function CreateWorkItemDocument (string)

'DebugPrint "CreateWorkItemDocument[ENTER]"

	Dim strArray
	Dim doc
	Dim xmltext
	Dim attrName, attrValue


	set doc = CreateObject("MSXML2.DOMDocument")
	
	' Create XML document according to spefification of OSLC-CM V2.0
	
	xmltext = "<?xml version=""1.0"" encoding=""UTF-8""?>"&vbCRLF
	xmltext = xmltext & "<rdf:RDF "  _
    	& " xmlns:oslc_pl=""http://open-services.net/ns/pl#""" _
		& " xmlns:rtc_ext=""http://jazz.net/xmlns/prod/jazz/rtc/ext/1.0/""" _
		& " xmlns:rtc_cm=""http://jazz.net/xmlns/prod/jazz/rtc/cm/1.0/""" _
		& " xmlns:dcterms=""http://purl.org/dc/terms/""" _
		& " xmlns:oslc_cmx=""http://open-services.net/ns/cm-x#""" _
		& " xmlns:acp=""http://jazz.net/ns/acp#""" _
		& " xmlns:oslc_cm=""http://open-services.net/ns/cm#""" _
		& " xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#""" _
		& " xmlns:oslc=""http://open-services.net/ns/core#"">" _
		& vbCRLF

	xmltext = xmltext & "<oslc_cm:ChangeRequest>" & vbCRLF


	strArray = split(string, ",")
	For I=0 to UBound (strArray)-1

		attrName = strArray(I)
		attrValue = strArray(I+1)

		' Check if this is rdf resource
		If InStr( attrValue, "rdf") = 0 Then
			xmltext = xmltext&"<"&attrName&">"&attrValue&"</"&attrName&">"&vbCrLf
		Else
			' This is RDF resource
			xmltext = xmltext&"<"&attrName&"  "&attrValue&"/>"&vbCrLf
		End If
		
		I = I + 1 'Increment i to skip


	Next

	xmltext = xmltext & "</oslc_cm:ChangeRequest>"&vbCRLF
	xmltext = xmltext & "</rdf:RDF>"

	doc.loadXML (xmltext)
	
	' clear collection
	set collection = Nothing 



	set CreateWorkItemDocument = doc

'DebugPrint "CreateWorkItemDocument [EXIT]"

End Function

Public Function GetWorkItemUrl(http, id)



End Function


Public Function UpdateWorkItem (http, update_service, workItemId, attrString)

	Dim doc  ' XLM dom object.
	Dim element, elements
	Dim workItemDoc

'DebugPrint "[Start]UpdateWorkItem"

	'return workitem ID upon sucessful update.


	update_url = update_service&"/"&workItemId

	Set workItemDoc = CreateWorkItemDocument(attrString)

	http.Open "PUT", update_url, False
	http.setRequestHeader "Content-Type", "application/xml"
	http.setRequestHeader "Accept", "application/xml"
	http.setRequestHeader "OSLC-Core-Version", "2.0"
	http.send(workItemDoc)

	' obtain response from workitem creation.
	set doc = CreateObject("MSXML2.DOMDocument")
	doc.loadXML(http.responseText)


	' When http status is not 200, then there is some kind of error happened.
	If http.status <> 200 Then
		workItemId = -1
		call PrintOslcErrorMessage(doc)
	End If

	set workItemDoc = Nothing
	set doc = Nothing

	UpdateWorkItem = workItemId
	

'DebugPrint "[End]UpdateWorkItem"

End Function



' CreateWorkItem function create new workitem using workitem document
'

Public Function CreateWorkItem (http, create_service, attrString)

	' return workitem ID created by this function, if it is failed, return -1
	'
	Dim workItemId   'workitem ID created by this function
	Dim create_url	' workitem creation url via factory url (create_service)
	Dim doc	'XML DOM object
	Dim workItemDoc

	' set workItemId to -1 (workitem creation has failed)
	workItemId = -1
	create_url = create_service

	Set workItemDoc = CreateWorkItemDocument(attrString)

	http.Open "POST", create_url, False
	http.setRequestHeader "Content-Type", "application/xml"
	http.setRequestHeader "Accept", "application/xml"
	http.setRequestHeader "OSLC-Core-Version", "2.0"
	http.send(workItemDoc)

	' obtain response from workitem creation.
	set doc = CreateObject("MSXML2.DOMDocument")
	doc.loadXML(http.responseText)

	' find created workitem ID	
	' http status 201 is return when workitem is created.
	If http.status = 201 Then
		set elements = doc.getElementsByTagName("oslc_cm:ChangeRequest")
		For each element in elements
			For each attr in element.childNodes
				If ( attr.nodeName = "dcterms:identifier" ) Then
					workItemId = attr.text
					Exit For
				End If
			Next
		Next
	Else  'error
		call PrintOslcErrorMessage(doc)
	End If

	' Free memory
	Set workItemDoc = Nothing
	Set doc = Nothing

	' set return value of this function as workitem ID created or -1 (fail)
	CreateWorkItem = workItemID


End Function


' CreateParentChild function create a link between workitems.
' It create parent/child relationship.
' childIDs are arrays of workitems.
Public Function CreateParentChild (http, service_url, parentID, childIDs)

	Dim I 	' Loop counter
	Dim childResource
	Dim attrString
	
	' Create resource URL


	childResource="rdf:resource="&""""&service_url&"/"&CStr(childIDs(0))&""""
	attrString = attrString&"rtc_cm:com.ibm.team.workitem.linktype.parentworkitem.children,"&childResource

	For i = 1 to UBound(childIDs)-1

		childResource="rdf:resource="&""""&service_url&"/"&CStr(childIDs(I))&""""
		attrString = attrString&","&"rtc_cm:com.ibm.team.workitem.linktype.parentworkitem.children,"&childResource

	Next

	
	' Update workitem.
	return = UpdateWorkItem (http, service_url, parentID, attrString)
	
	CreateParentChild = return
	


End Function

Public Sub PrintOslcErrorMessage (doc)

	Dim errorMsg
	Dim element, elements
	
	errorMsg = "Error!" ' This is default error message

	
	set elements = doc.getElementsByTagName("oslc:Error")
	
	For each element in elements
		For each attr in element.childNodes
			 errorMsg = attr.nodeName&":"&attr.text
			 WScript.Echo errorMsg
		Next
	Next
	
End Sub

Private Sub DebugPrint (string)

	Dim debug
	
	debug = true
	
	If debug = true Then
		WScript.Echo "[DEBUG]"&string
	End If

End Sub



' Convert date format from YYYY/MM/DD format into ISO date format
Public Function ConvertToISODate(date)
' 
 	Dim strDate
 	Dim strArray

	' if date is not planned, it returns zero time
	If date = "0:00:00" then
		strDate = ""
	Else
		strArray = split(date, "/")
 		yy = strArray(0)
 		mm = strArray(1)
 		dd = strArray(2)
		hr = "00"
 		min = "00"
 		sec = "00"
	 	strDate = yy&"-"&mm&"-"&dd&"T"&hr&":"&min&":"&sec&".000Z"
 	End If
	
	ConvertToISODate = strDate

End Function


