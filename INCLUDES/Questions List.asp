		<%Icons = Array("icon_unanswered1.gif", "icon_answered1.gif", "icon_current1.gif", "icon_pointless.gif")%>
		<TABLE WIDTH="250" cellpadding="0" cellspacing="5">
			<TR>
				<TD COLSPAN="2"><STRONG>Questions:</STRONG><BR>&nbsp;</TD>
			</TR>
			<%For Each objQuestion In objInterview.AllQuestions%>
			<TR VALIGN="TOP">
				<TD WIDTH="40"><IMG SRC="IMAGES/<%=Icons(objQuestion.State)%>"> <%=objQuestion.Definition.Sequence%>. </TD>
				<TD><%=objQuestion.HTMLJumpToLinkOpen%><%=objQuestion.Definition.ShortText%><%=objQuestion.HTMLJumpToLinkClose%></TD>
			</TR>
			<%Next%>
		</TABLE>
