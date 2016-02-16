<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
cpass23=Request.Form("cpass23")
cpass222=int(Request.Form("cpass222"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 用戶 where 姓名='"&cpass23&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要查找的人我們找不到！請查看是否正確！');history.back();</script>"
else
aa=rs("等級")
bb=(aa-cpass222)*(aa-cpass222)*50
sql="update 用戶 set 等級=等級-"&cpass222&",allvalue="&bb&" where 姓名='"&cpass23&"'"
Set Rs=conn.Execute(sql)
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','減級操作')"
conn.Execute(sql)
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script language=javascript>alert('用戶：["&cpass23&"]減["&cpass222&"]級修改成功！');history.back();</script>"
end if
%> 