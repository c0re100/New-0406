<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Function computerChooses()
Dim randomNum
Dim choice
randomize
randomNum = int(rnd*15)+1
If randomNum = 1 OR randomNum = 3 OR randomNum = 7 OR randomNum = 8 OR randomNum = 15 OR randomNum = 12 Then
choice = "R"
ElseIf randomNum = 2 OR randomNum = 6 OR randomNum = 11 OR randomNum = 13 Then
choice = "S"
Else
choice = "P"
End If
computerChooses = choice
End Function
Sub determineWinner(playerChoice, computerChoice)
Const Rock = "R"
Const Scissor = "S"
Const Paper = "P"
Dim tempPlayer, tempComputer
If playerChoice = Rock Then
If computerChoice = Scissor Then
conn.Execute "update �Τ� set �Ȩ�=�Ȩ�+50 where id="&info(9)
%>
<title>�ŤM ���Y ��</title><body background="../../images/8.jpg">
<P><CENTER>
<IMG SRC="jhimg/shit.gif"><BR>
�A�����Y����F�p������Ťl�I�z�o��50��Ȥl�C</CENTER>
<%
End If
ElseIf playerChoice = Scissor Then
If computerChoice = Paper Then
conn.Execute  "update �Τ� set �Ȩ�=�Ȩ�+50 where id="&info(9)
set rs=nothing
conn.close
set conn=nothing
%>
<P><CENTER>
<IMG SRC="jhimg/jianz.gif"><BR>
�A���Ťl�ů}�F�p��������I�z�o��50��Ȥl�C</CENTER>
<%
End If
ElseIf playerChoice = Paper Then
If computerChoice = Rock Then
conn.Execute  "update �Τ� set �Ȩ�=�Ȩ�+50 where id="&info(9)
conn.close
set conn=nothing
set rs=nothing	
%>
<P><CENTER>
<IMG SRC="jhimg/bu.gif"><BR>
�A�����]��F�p��������Y�I�z�o��50��Ȥl�C</CENTER>
<%
End If
ElseIf playerChoice = computerChoice Then
%>
<p><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
�ݨӧA�̷Q��@�_�F�C</CENTER>
<%
End If
If computerChoice = Rock Then
If playerChoice = Scissor Then
conn.Execute  "update �Τ� set �Ȩ�=�Ȩ�-50 where id="&info(9)
conn.close
set conn=nothing
set rs=nothing
%>
<P><CENTER>
<IMG SRC="jhimg/shit.gif"><BR>
�p��������Y����F�A���Ťl�I�z�l��50��Ȥl�C</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
�ݨӧA�̷Q��@�_�F�C</CENTER>
<%
End If
ElseIf computerChoice = Scissor Then
If playerChoice = Paper Then
conn.Execute  "update �Τ� set �Ȩ�=�Ȩ�-50 where id="&info(9)
conn.close
set conn=nothing
set rs=nothing
%>
<P><CENTER>
<IMG SRC="jhimg/jianz.gif"><BR>
�p������Ťl�ů}�F�A�����I�z�l��50��Ȥl�C</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
�ݨӧA�̷Q��@�_�F�C</CENTER>
<%
End If
ElseIf computerChoice = Paper Then
If playerChoice = Rock Then
conn.Execute  "update �Τ� set �Ȩ�=�Ȩ�-50 where id="&info(9)
conn.close
set conn=nothing
set rs=nothing
%>
<P><CENTER>
<IMG SRC="jhimg/bu.gif"><BR>
�p��������]��F�A�����Y�I�z�l��50��Ȥl�C</CENTER>
<%
ElseIf playerChoice = computerChoice Then
%>
<P><CENTER>
<IMG SRC="jhimg/tie.gif"><BR>
�ݨӧA�̷Q��@�_�F�C</CENTER>
<%
End If
ElseIf computerChoice = playerChoice Then
%>
<P>
<CENTER>
<font color="#0000FF"><IMG SRC="jhimg/tie.gif"><BR>
�ݨӧA�̷Q��@�_�F�C</font>
</CENTER>
<%
End If
End Sub
Sub playGame()
%>
<P align="center"> <font color="#0000FF" size="5" face="����">�ŤM ���Y ��</font>
<table width="390" border="1" cellspacing="10" cellpadding="2" align="center" bgcolor="#336633">
<tr bgcolor="#FFCCCC">
<td width="178" height="193" bgcolor="#CCCCFF">
<div align="center"><font size="2">�� �� ��:</font> </div>
<form action="index.asp?action=winner" method="post">
<div align="center">
<table width="178" cellspacing="0" border="1" bgcolor="#CCCCCC" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr valign=top>
<td width="75">
<div align="center"><font size="2">�� �Y</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="R">
</div>
</td>
</tr>
<tr valign=top>
<td width="75">
<div align="center"><font size="2">�� �l</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="S">
</div>
</td>
</tr>
<tr valignn=top>
<td width="75">
<div align="center"><font size="2">��</font></div>
</td>
<td width="63">
<div align="center">
<input type="radio" name="playerSelect" value="P">
</div>
</td>
</tr>
</table>
<input type="submit" value="�} �l" name="submit">
</div>
</form>
<div align="center"> </div>
<p align="center"><font size="2"><a href="../betindex.asp">�����F</a></font><br>
</td>
<td width="211" bgcolor="#FFFFFF" height="193">
<div align="center"><font size="2">�m�Ťl�B���Y�B���n�C��</font><font color="#33CCCC"><br>
���A�ӧQ�I<br>
</font><img src="jhimg/bu.gif"><img src="jhimg/shit.gif"></div>
</td>
</tr>
</table>
<p align="center"><%
End Sub
Sub playAgain()
%>
<p>
<center>
�z�n�A���@���H
</center>
<br>
<center>
<a href="index.asp"><font size="2">�O��</font></a> <font size="2"><a href="../betindex.asp">�����F</a>
</font>
</center>
<H1 align="center"><BR>
</h1>
<CENTER>
</CENTER>
<%
End Sub
Sub endGame()
Response.Buffer = true
End Sub
Dim player, computer
Dim gameAction
gameAction = Request.QueryString("action")
Select Case gameAction
Case "winner"
player = Request.Form("playerSelect")
computer = computerChooses
		
determineWinner player, computer
Response.Write "<BR><BR>"
playAgain

Case "again"
playAgain
Case "gameover"
endGame
Case Else
playGame
End Select
%>
</body>
