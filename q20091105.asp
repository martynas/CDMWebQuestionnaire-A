<%
	Set objAppl = Server.CreateObject("AstonRDLightWeightModel.Application")
	objAppl.OpenApplication "CDMPROD"
'	Set tmpInterview = objAppl.Class("STDQuestionary2Contact").LoadFrom(Array("ID","primaryContact","done"),Array(CStr(Trim(Request.QueryString("ID"))),CStr(Trim(Request.QueryString("CID"))),False))
	Set tmpInterview = objAppl.Class("STDQuestionary2Contact").LoadFrom(Array("ID"),Array(CStr(Trim(Request.QueryString("ID")))))
	If tmpInterview Is Nothing Then Response.Redirect "ThankYou.asp"
	Set tmpInterview = Nothing
	Set objAppl = Nothing

	Set objRoot = Server.CreateObject("AstonRDWebQuestionnaire.Root")	
	If objRoot.ProcessInterview("CDMPROD", Request.QueryString("ID"), Request, 1) Then Response.Redirect "ThankYou.asp"
	Set objInterview = objRoot.Interview
	
	' * ONE Pass Questionnaire - If an answer requires 
	'		complementing (variable type / comments), those
	'		will be shown together with the answer choices,
	'		i.e. no need to ask for complementing. Only
	'		"FIRST PASS" section will be shown.
	' * TWO Pass Questionnaire - If an answer requires 
	'		complementing (variable type / comments), those
	'		will not be shown with the answer choices, but if the 
	'		user selects an answer that requires complementing, 
	'		after pressing the "SUBMIT" button,	the "SECOND PASS" 
	'		section is shown to ask for complementing.
	blnOnePass = False ' Run ONE (True) or TWO (False) Pass Questionnaire 

	' Choose Layout File Below
%>
<!--#include file="INCLUDES\Page Layout.asp"-->
