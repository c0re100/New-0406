<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=request("name")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 用戶 where 姓名='" & name & "' and lastkick='" & my & "'",conn 
if rs.eof or rs.bof then
rs.close
conn.close
set rs=nothing	
set conn=nothing
Response.Write "<script Language=Javascript>alert('你還沒殺了這個人呢！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select 殺人傭金 from 殺手 where 姓名2='" & my & "' and 被殺者='" & name & "'" ,conn
shangjin=rs("殺人傭金")
conn.execute "delete * from 殺手 where 姓名2='" & my & "' and 被殺者='" & name & "'"
conn.execute "update 用戶 set lastkick='無' where 姓名='"&name&"'"
conn.execute "update 用戶 set 銀兩=銀兩+"&shangjin&" where id="&info(9)
rs.close
conn.close
set rs=nothing	
set conn=nothing
Response.Write "<script Language=Javascript>alert('恭喜你得到你應得的共有銀兩["&shangjin&"]');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>