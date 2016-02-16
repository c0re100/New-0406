<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
num=abs(trim(request.form("text")))
if not isnumeric(num) then Response.Redirect "../error.asp?id=54"
if num>1000000 then
 	Response.Write "<script language=JavaScript>{alert('一次會費值不能太大!');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 會員等級,會員費 from 用戶 where id="&info(9),conn
if rs("會員等級")=0 then
Response.Write "<script language=JavaScript>{alert('你不是會員請回吧！');location.href = '../help/info.asp?type=2&name=會員加入辦法';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("會員費")<num then
Response.Write "<script language=JavaScript>{alert('你有那麼多會費嗎？');location.href = 'hxiehy.asp';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
jbb=rs("會員費")-num
jb=int(num*100)
conn.execute("update 用戶 set 會員費="&jbb&",金幣=金幣+"&jb&",操作時間=now()  where id="&info(9))
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('會費轉換成金幣成功，請注意查收!');location.href='hxiehy.asp';</script>"
Response.End
%>
