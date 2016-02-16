<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="data1.asp"-->
<%set rs=conn.execute("select * from sheep where user='"&info(0)&"'")
if not rs.eof then
rs.close
conn.execute("delete from sheep where user='"&info(0)&"'")
conn.close
%>
<!--#include file="data.asp"-->
<%set rs=conn.execute("select * from 用戶 where 姓名='"&info(0)&"'")
money=rs("銀兩")+800
conn.execute("update 用戶 set 銀兩='"&money&"'  where 姓名='"&info(0)&"'")%>
你已經成功賣出了小羊並得到800塊錢！


<%else%>
你沒有小羊不能賣出
<%rs.close%>
<%end if%>