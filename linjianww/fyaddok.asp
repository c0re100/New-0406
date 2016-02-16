<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 物品買 WHERE 物品名='" & name & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
if clng(Request.Form("wupinll"))<0 then Response.Redirect "../error.asp?id=471"
if clng(Request.Form("wupintl"))<0 then Response.Redirect "../error.asp?id=471"
name=Request.Form("wupinname")
say=Request.Form("wupinsay")
a=clng(Request.Form("wupintl"))
b=clng(Request.Form("wupinfy"))
c=clng(Request.Form("wupingj"))
d=clng(Request.Form("wupinjg"))
e=clng(Request.Form("wupinnl"))
wupinll=(a+b)*400
set rs=server.createobject("adodb.recordset")
sql="select * from 物品買 where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("物品名")=name
rs("類型")=request.form("lx")
rs("體力")=a
rs("防御")=b
rs("攻擊")=c
rs("堅固度")=d
rs("內力")=e
rs("銀兩")=wupinll
rs("說明")=say
rs.update
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','物品增加')"
conn.Execute(sql)
rs.close
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=700"
%>
<%else%>
<script language=vbscript>
MsgBox "倉庫已經有這個物品了，請你添加點別的吧"
location.href = "javascript:history.back()"
</script>
<%end if%> 