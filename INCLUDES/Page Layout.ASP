<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="style.css">
</HEAD>
<BODY>
<TABLE width="100%" height="100%" cellpadding="0" cellspacing="0">  
  <TR>
	<TD WIDTH="250" HEIGHT="110"><a href="/" class="logo"> Atrendia</a><!--<IMG SRC="IMAGES/atrendialogo.jpg"--></TD>
    <TD WIDTH="1" HEIGHT="110"><IMG SRC="IMAGES/VLINE1.GIF" HEIGHT="110" WIDTH="1"></TD>
    <TD WIDTH="100%" HEIGHT="110">
		<!-- THE HEADER (RIGHT TO THE LOGO) -->
		<STRONG><BIG><CENTER><%=objInterview.Questionnaire.Heading%><CENTER></BIG></STRONG><BR>
		<BIG><CENTER><b><big>Atrendia</big></b><CENTER></BIG>
	</TD>
  </TR>
  <TR>
    <TD WIDTH="250" HEIGHT="1"><IMG SRC="IMAGES/HLINE1.GIF" HEIGHT="1" WIDTH="250"></TD>
    <TD WIDTH="1" HEIGHT="1"><IMG SRC="IMAGES/CROSS1.GIF" HEIGHT="1" WIDTH="1"></TD>
    <TD WIDTH="100%" HEIGHT="1"><IMG SRC="IMAGES/HLINE1.GIF" HEIGHT="1" WIDTH="100%"></TD>
  </TR>
  <TR>
    <TD WIDTH="1" HEIGHT="100%" valign="top">
		<!-- THE QUESTION LIST (BELOW THE LOGO) - THIS IS OPTIONAL -->
		<!--#include file="Questions List.asp"-->
    </TD>
    <TD WIDTH="1" HEIGHT="100%"><IMG SRC="IMAGES/VLINE1.GIF" HEIGHT="100%" WIDTH="1"></TD>
    <TD WIDTH="100%" HEIGHT="100%" VALIGN="TOP" >
		<!-- THE MAIN BODY - QUESTIONS ARE ASKED HERE - REQUIRED -->
		<!--#include file="Questions Body.asp"-->
    </TD>
  </TR>
</TABLE>
</FORM>
</BODY>
</HTML>
