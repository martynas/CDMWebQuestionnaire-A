<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=unicode">
<STYLE>
<!-- TD
{
    FONT-SIZE: 12px;
    COLOR: #000000;
    FONT-FAMILY: Courier New
}
-->
</STYLE>
<META content="MSHTML 5.50.4134.600" name=GENERATOR></HEAD>
<BODY>
<%
	Set objAppl = Server.CreateObject("AstonRDLightWeightModel.Application")
	objAppl.OpenApplication "CDMPROD"
	t = Timer
	Set objR = Server.CreateObject("AstonRDWebQuestionnaire.Root")
	Contents = Session("Questionnaire")
	If IsArray(Contents) Then
		objR.DePersist Contents
		objR.SetA objAppl
		Response.Write "<H3>DePersisting</H3>"
	Else
		Set objQuest = objAppl.Class("STDQuestionary").LoadFrom("code", "CS")
		objR.SetA objAppl
		objR.Test objQuest
		Response.Write "<H3>DB Loading</H3>"
	End If
	Set objQY = objR.Interview.Questionnaire 
%>
<TABLE cellSpacing=1 cellPadding=1 border=1>
  
  <TR>
    <TD COLSPAN="4"><STRONG>Questionnaire</STRONG></TD>
  </TR>
  <TR>
    <TD COLSPAN="4">Code = <%=objQY.Code%></TD>
  </TR>
  <TR>
    <TD COLSPAN="4">Heading = <%=objQY.Heading%></TD>
  </TR>
  <TR>
    <TD COLSPAN="4"><EM>Questions</EM></TD>
  </TR>
  <%For Each objQ In objQY.Questions%>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |---</TD>
    <TD COLSPAN="3"><STRONG>Question</STRONG></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Allow Comments = <%=objQ.AllowComments%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Allow Jump From = <%=objQ.AllowJumpFromThis%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Allow Jump To = <%=objQ.AllowJumpToThis%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Allow Multiple Answers = <%=objQ.AllowMultipleAnswers%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Long Text = <%=objQ.LongText%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Question Comments = <%=objQ.QuestionComments%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Save Answers = <%=objQ.SaveAnswers%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Sequence = <%=objQ.Sequence%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Short Text = <%=objQ.ShortText%></TD>
  </TR>
  <%
    Set objV = objQ.VariableType
    If objV Is Nothing Then
  %>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Variable Type - None</TD>
  </TR>
  <%Else%>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3">Variable Type</TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |---</TD>
    <TD COLSPAN="2"><STRONG>Variable Type</STRONG></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Code = <%=objV.Code%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">DBType = <%=objV.DBType%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Description = <%=objV.Description%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Heading = <%=objV.Heading%></TD>
  </TR>
  <%End If%>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="3"><EM>Answers</EM></TD>
  </TR>
  
  <%For Each objA In objQ.Answers%>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |---</TD>
    <TD COLSPAN="2"><STRONG>Answer</STRONG></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Allow Comments = <%=objA.AllowComments%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Long Text = <%=objA.LongText%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Sequence = <%=objA.Sequence%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Short Text = <%=objA.ShortText%></TD>
  </TR>
  <%
    Set objV = objA.VariableType
    If objV Is Nothing Then
  %>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Variable Type - None</TD>
  </TR>
  <%Else%>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD COLSPAN="2">Variable Type</TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |---</TD>
    <TD><STRONG>Variable Type</STRONG></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD>Code = <%=objV.Code%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD>DBType = <%=objV.DBType%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD>Description = <%=objV.Description%></TD>
  </TR>
  <TR>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD WIDTH="50">&nbsp;&nbsp; |</TD>
    <TD>Heading = <%=objV.Heading%></TD>
  </TR>
  <%End If%>
  <%Next%>
  <%Next%>
</TABLE>
<P><STRONG>Execution Time = <%=Timer - t%></STRONG></P>
<%
Response.Write IsArray(Session("Questionnaire"))
Session("Questionnaire") = objR.Persist
Response.Write IsArray(Session("Questionnaire"))
%>
<HR>
<%
	For Each objQ In objR.Interview.AllQuestions
		'With objR.Interview.AllQuestions.Item
		With objQ
		Response.Write "<LI><B>"
		Response.Write .Id 
		Response.Write " "
		Response.Write .Definition.Sequence
		Response.Write ". "
		Response.Write .Definition.ShortText
		Response.Write " </B>"
		Response.Write .HTMLComments("")
		for each objA in objQ.AllAnswers
			Response.Write "<BR>* "
			Response.Write objA.Id
			Response.Write " "
			Response.Write objA.Definition.Sequence
			Response.Write ". "
			Response.Write objA.Definition.ShortText
			Response.Write " "
			Response.Write objA.HTMLAnswer("")
			Response.Write vbcrlf
		Next
			Response.Write vbcrlf
			Response.Write vbcrlf
		End With
	Next 
%>
</BODY></HTML>
