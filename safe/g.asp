<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%><%
http = Request.ServerVariables("HTTP_REFERER")
username=Request.Form ("username")
if username="" then Response.End
http="http://"&Request.ServerVariables("SERVER_NAME")
if InStr(username,"=")<>0 or InStr(username,"`")<>0 or InStr(username,"'")<>0 or InStr(username," ")<>0 or InStr(username," ")<>0 or InStr(username,"'")<>0 or InStr(username,chr(34))<>0 or InStr(username,"\")<>0 or InStr(username,",")<>0 or InStr(username,"<")<>0 or InStr(username,">")<>0 then Response.end
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.Open ("select �m�W from �Τ� where �m�W='"&username&"'"),conn
if rs.BOF or rs.EOF then
Response.Write "�䤣�즹�Τ�W�I"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.End 
end if
pass=rs("�K�X")
email=rs("�H�c")
rs.Close
set rs=nothing
newpassword=""
for i=1 to len(Pass)
	newpassword=newpassword&"%"&cstr(hex(asc(mid(pass,i,1))))
next
pass=newpassword
mailbody="�m�W:<br>"&username&"<br> �K�X:<br>:"&pass&
'mailbody=mailbody&"g2.asp?username="&username&"&oldpass="&pass&"&email="&email
'mailbody="<a href='"&mailbody&"'>"&mailbody&"</a>"
HTML = "<html>"
HTML = HTML & "<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5> " 
HTML = HTML & "<title>Sending CDONTS Email Using HTML</title>"
HTML = HTML & "</head>"
HTML = HTML & "<body bgcolor=""FFFFFF"">"
HTML = HTML & "<p><font size=3>"
HTML = HTML & "�˷R���F�C����[<font color=red>"&username&"</font>]�B�͡A�z���K�X�v�򥻧�^�I�I<br>"
HTML = HTML & "...</font></p>"
HTML = HTML & MAILBODY
HTML = HTML & "<br></body>"
HTML = HTML & "</html>"
conn.Close
set conn=nothing
set mail=server.CreateObject("CDONTS.NewMail")
mail.To=email
mail.From="bear@asz.cn"
mail.Subject="�A���F�C���򪺱K�X�v�g���^�ӤF�I"
Mail.BodyFormat = 0 
Mail.MailFormat = 0 
mail.Body =HTML
mail.Send 
Response.Write "�l��o�e���\�I�I�p���ɶ�������l��A�Э��s�A���ӹL�I"
set mail=nothing
%>
