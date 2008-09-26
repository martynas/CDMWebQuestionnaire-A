		<FORM METHOD="POST" ACTION="http://<%=Request.ServerVariables.Item("SERVER_NAME")%><%=Request.ServerVariables.Item("URL")%>?<%=Request.QueryString%>" id=form1 name=form1>
		<%=objInterview.HTMLRequiredHiddenInputs%>
		<TABLE WIDTH="100%" cellpadding="0" cellspacing="0" >
			<%For Each objQuestion In objInterview.CurrentQuestions%>
			<!-- START - SHOWN FOR EACH QUESTION -->
				<TR>
					<TD style="padding:9px"><!-- QUESTION HEADING -->
						<STRONG><%=objQuestion.Definition.Sequence%>. <%=objQuestion.Definition.ShortText%></STRONG>
						<%If objQuestion.Definition.LongText <> "" Then Response.Write "<BR>&nbsp;<BR>" & objQuestion.Definition.LongText%><BR>&nbsp;
					</TD>
				</TR>
				<%If Not objQuestion.RequiresComplement Then%>
					<!-- SHOW FIRST PASS - AKS QUESTION -->
					<!--#include file="Question Body FirstPass.asp"-->
				<%Else%>
					<!-- SHOW SECOND PASS - COMPLEMENT - ANSWER DEPENDENT -->
					<!--#include file="Question Body SecondPass.asp"-->
				<%End If%>
				<TR><TD WIDTH="100%" HEIGHT="1"><IMG SRC="IMAGES/hLINE1.GIF" WIDTH="100%" HEIGHT="1"></TD></TR>
				<TR><TD>&nbsp;</TD></TR>
				<!-- END - SHOWN FOR EACH QUESTION -->
			<%Next%>
			<TR>
				<TD style="padding:9px">
					<INPUT type="submit" value="Next -&gt;" style="WIDTH: 150px" id=submit1 name=submit1> &nbsp;&nbsp;&nbsp;
				</TD>
			</TR>
		</TABLE>
