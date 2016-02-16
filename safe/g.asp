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
rs.Open ("select 姓名 from 用戶 where 姓名='"&username&"'"),conn
if rs.BOF or rs.EOF then
Response.Write "找不到此用戶名！"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.End 
end if
pass=rs("密碼")
email=rs("信箱")
rs.Close
set rs=nothing
newpassword=""
for i=1 to len(Pass)
	newpassword=newpassword&"%"&cstr(hex(asc(mid(pass,i,1))))
next
pass=newpassword
mailbody="姓名:<br>"&username&"<br> 密碼:<br>:"&pass&
'mailbody=mailbody&"g2.asp?username="&username&"&oldpass="&pass&"&email="&email
'mailbody="<a href='"&mailbody&"'>"&mailbody&"</a>"
HTML = "<html>"
HTML = HTML & "<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5> " 
HTML = HTML & "<title>Sending CDONTS Email Using HTML</title>"
HTML = HTML & "</head>"
HTML = HTML & "<body bgcolor=""FFFFFF"">"
HTML = HTML & "<p><font size=3>"
HTML = HTML & "親愛的靈劍江湖[<font color=red>"&username&"</font>]朋友，您的密碼己基本找回！！<br>"
HTML = HTML & "...</font></p>"
HTML = HTML & MAILBODY
HTML = HTML & "<br></body>"
HTML = HTML & "</html>"
conn.Close
set conn=nothing
set mail=server.CreateObject("CDONTS.NewMail")
mail.To=email
mail.From="bear@asz.cn"
mail.Subject="你的靈劍江湖的密碼己經飛回來了！"
Mail.BodyFormat = 0 
Mail.MailFormat = 0 
mail.Body =HTML
mail.Send 
Response.Write "郵件發送成功！！如長時間收不到郵件，請重新再次來過！"
set mail=nothing
%>
