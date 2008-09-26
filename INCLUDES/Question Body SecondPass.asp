				<%' #### SECOND PASS - ANSWER DEPENDENT ####
				
				For Each objAnswer In objQuestion.CheckedAnswers
					Response.Write "<TR><TD><EM>You selected the answer <U>"
					Response.Write objAnswer.Definition.Sequence & ". "
					Response.Write objAnswer.Definition.ShortText & "</U>. "
					Response.Write "Please supplement your answer:"
					
					' **** START - ANSWER VARIABLE TYPE SECTION ****
					If Not objAnswer.Definition.VariableType Is Nothing Then
						Response.Write "<TR><TD style='padding-left:9px'><EM>" & objAnswer.Definition.VariableType.Description & "</EM><BR>"
						Response.Write objAnswer.HTMLVariableType & " " & objAnswer.Definition.VariableType.Heading
						Response.Write "<BR>&nbsp;</TD></TR>"
					End If
					' **** END - ANSWER VARIABLE TYPE SECTION ****
			
					' **** START - ANSWER COMMENTS SECTION ****
					If objAnswer.Definition.AllowComments Then
						' This Answer Allows Commenting 
						Response.Write "<TR><TD style='padding-left:9px'><EM>Please write your comments in the box below </EM><BR>"
						Response.Write objAnswer.HTMLComments("COLS=""60"" ROWS=""6""")
						Response.Write "<BR>&nbsp;</TD></TR>"
					End If
					' **** END - ANSWER COMMENTS SECTION ****
				Next
				%>
