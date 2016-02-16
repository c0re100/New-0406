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
%>
<!--#include file="data1.asp"--><%
sql="select * from 房屋 where ID=" & id
rs=conn.execute(sql)
yin=rs("售價")
huzhu=rs("戶主")
blv=rs("伴侶")
zt=rs("狀態")
if rs("銀兩")<=0 or rs("伴侶銀兩")<=0 or rs("數量")<=0  then
conn.execute "update 房屋 set 戶主='無',伴侶='無',數量=0,銀兩=0,伴侶銀兩=0 where ID=" & id
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from 用戶 where id="&info(9)
set rs=conn.execute(sql)
if info(0)=huzhu  and zt="正常" then
      conn.execute "update 用戶 set 銀兩=銀兩+" & yin & "*1/5 where id="&info(9)
      %><!--#include file="data1.asp"--><%
      conn.execute "update 房屋 set 戶主='無',伴侶='無',數量=0,銀兩=0,伴侶銀兩=0 where ID=" & id
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你已經成功賣出！');location.href = 'fangwu.asp';</script>"
	Response.End
end if
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('交易不成功，原因：你不是這個房子的主人或是你的房子出了點狀況！');location.href = 'fangwu.asp';</script>"
	Response.End
%>
