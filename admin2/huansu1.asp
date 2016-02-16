<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩,門派,性別,道德 from 用戶 where id="&info(9),conn
yl=rs("銀兩")
mp=rs("門派")
xb=rs("性別")
if yl<200000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="你的錢不夠，老實的出家吧，自己的路可是自己選的哦！"
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = 'huansu.asp';}</script>"
response.end
end if
if mp<>"出家" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="你都不是江湖的和尚或尼姑，哪裡來的流氓，我們不歡迎你哦，下山吧！"
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = 'huansu.asp';}</script>"
response.end
end if
yl1=yl-200000
dd=rs("道德")-300
mess="恭喜，你已經還俗了，自己的路自己走吧請退出後重新登陸"
if xb="男" then
conn.execute "update 用戶 set 銀兩='"&yl1&"',道德='"&dd&"',門派='遊俠',身份='無' where id="&info(9)
else
conn.execute "update 用戶 set 銀兩='"&yl1&"',道德='"&dd&"',門派='遊俠',身份='無' where id="&info(9)
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = '../exit.asp';}</script>"
response.end%>
<html>