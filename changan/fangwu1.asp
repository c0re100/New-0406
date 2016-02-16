<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
my=info(0)
if instr(action,"'") then response.end 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 等級,配偶 from 用戶 where id="&info(9),conn
'hy=rs("會員")
dj=rs("等級")
po=rs("配偶")
if dj<20 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('房源緊張，須等級大於20級以上才可以購買!');location.href = 'fangwu.asp';}</script>"

response.end
end if
if po="無" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你還沒有老婆，不能買房屋!');location.href = 'fangwu.asp';}</script>"

response.end
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%><!--#include file="data1.asp"--><%
sql="select * from 房屋 where 戶主='" & my & "' or 伴侶='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%><!--#include file="data1.asp"--><%
sql="select * from 房屋 where ID=" & id
rs=conn.execute(sql)
yin=rs("售價")
huzhu=rs("戶主")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩 from 用戶 where id="&info(9),conn
if rs("銀兩")<=yin and my=huzhu and huzhu<>"無" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('購買不成功，原因：你的銀兩不夠!');location.href = 'fangwu.asp';}</script>"

response.end
end if
     conn.execute "update 用戶 set 銀兩=銀兩-" & yin & " where id="&info(9)
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
 %><!--#include file="data1.asp"--><%
      sql="update 房屋 set 戶主='" & my & "',伴侶='" & po & "' where ID=" & id
	rs=conn.execute(sql)
	conn.close
	Response.Redirect "fangwu.asp"

else %> 
<script language=vbscript>
            MsgBox "您或您的伴侶已經買過房屋了。"
            location.href = "fangwu.asp"
                    </script>
<%
   rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
%>
