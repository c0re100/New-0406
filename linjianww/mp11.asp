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
if info(2)<11 then Response.Redirect "../error.asp?id=900"
dim conn,rs,rsuser
owner=request.querystring("owner")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
select case request.querystring("action")
case "delete"
set rstmp=conn.execute("Select * From 門派 Where 門派='" & owner & "' ")
if rstmp.eof then
Response.Redirect "../error.asp?id=447"
else
conn.execute "Delete * From 門派 Where 門派= '" & owner & "' "
cname=rstmp("掌門")
conn.execute("Update 用戶 set 門派='遊俠',身份='弟子' where 門派='"&owner&"'")
Response.Redirect "../ok.asp?id=704"
end if
end select
set rs=nothing
%>
<script language=vbscript>
MsgBox "<%=response.write(str)%>"
location.href = "adminmpp.asp"
</script>