				<%' #### FIRST PASS - AKS QUESTION ####

				' **** START - ANSWER CHOICES SECTION ****
				If objQuestion.AllAnswers.Count > 0 Then
					' There are answer choices for this question - show them!
					If objQuestion.Definition.AllowMultipleAnswers Then
						Response.Write "<TR><TD style='padding-left:9px'><EM>Please select one or more answers</EM><BR>"
					Else
						Response.Write "<TR><TD style='padding-left:9px'><EM>Please select an answer</EM><BR>"
					End If
					For Each objAnswer In objQuestion.AllAnswers
						%><%=objAnswer.HTMLAnswer%> <%=objAnswer.Definition.Sequence%>. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%=objAnswer.Definition.ShortText%><BR><%
						
						If blnOnePass Then ' Run ONE or TWO Pass Questionnaire (See also DEMO.ASP)
							' **** START - ANSWER VARIABLE TYPE SECTION ****
							If Not objAnswer.Definition.VariableType Is Nothing Then
								' This Answer has a Variable Type
								Response.Write "<IMG SRC=""IMAGES/empty.gif"" WIDTH=""35"" HEIGHT=""1"">"
								Response.Write "<EM>" & objAnswer.Definition.VariableType.Description & "</EM><BR>"
								Response.Write "<IMG SRC=""IMAGES/empty.gif"" WIDTH=""35"" HEIGHT=""1"">"
								Response.Write objAnswer.HTMLVariableType & " " & objAnswer.Definition.VariableType.Heading & "<BR>"
							End If
							' **** END - ANSWER VARIABLE TYPE SECTION ****
			
							' **** START - ANSWER COMMENTS SECTION ****
							If objAnswer.Definition.AllowComments Then
								' This Answer Allows Commenting 
								Response.Write "<IMG SRC=""IMAGES/empty.gif"" WIDTH=""35"" HEIGHT=""1"">"
								Response.Write "<EM>If You choose <U>" & objAnswer.Definition.ShortText &  "</U>, Please, write comments in the box below!</EM><BR>"
								Response.Write "<IMG SRC=""IMAGES/empty.gif"" WIDTH=""35"" HEIGHT=""1"">"							
								Response.Write objAnswer.HTMLComments("COLS=""55"" ROWS=""4""") & "<BR>"
							End If
							' **** END - ANSWER COMMENTS SECTION ****						
						End If
					Next
					Response.Write "&nbsp;</TD></TR>"
				End If
				' **** END - ANSWER CHOICES SECTION ****
			
				' **** START - QUESTION VARIABLE TYPE SECTION ****
				If Not objQuestion.Definition.VariableType Is Nothing Then
					Response.Write "<TR><TD><EM>" & objQuestion.Definition.VariableType.Description & "</EM><BR>"
					Response.Write objQuestion.HTMLVariableType & " " & objQuestion.Definition.VariableType.Heading
					Response.Write "<BR>&nbsp;</TD></TR>"
				End If
				' **** END - QUESTION VARIABLE TYPE SECTION ****
			
				' **** START - QUESTION COMMENTS SECTION ****
				If objQuestion.Definition.AllowComments Then
					' This Question Allows Commenting 
					Response.Write "<TR><TD><EM>Please write your comments in the box below</EM><BR>"
					Response.Write objQuestion.HTMLComments("COLS=""60"" ROWS=""6""")
					Response.Write "<BR>&nbsp;</TD></TR>"
				End If
				' **** END - QUESTION COMMENTS SECTION ****
				%>
