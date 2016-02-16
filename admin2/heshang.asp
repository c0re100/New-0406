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
rs.open "select 銀兩,門派,身份,性別,道德 from 用戶 where id="&info(9),conn
yl=rs("銀兩")
mp=rs("門派")
sf=rs("身份")
xb=rs("性別")
if yl<200000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="你的錢不夠，沒辦法了，官府的新規定阿，回去好好想想辦法再來吧！"
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"

response.end
end if
if rs("性別")="女" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="這裡是和尚廟呀，你走錯啦！"
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"

response.end
end if
if mp="出家" then
conn.execute "update 用戶 set 銀兩=銀兩-"&yl&",門派='遊俠',身份='弟子' where id="&info(9)
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="恭喜，你已經還俗了，自己的路自己走吧請退出後重新登陸"
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"
response.end
end if
if mp="官府" or sf="掌門"  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="你是官府或是掌門不能出家！"
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"
response.end
end if
yl1=yl-200000
dd=rs("道德")-300
mess="恭喜，你已經成功出家了，以後的路自己走好吧請推出後重新登陸"
conn.execute "update 用戶 set 銀兩='"&yl1&"',道德='"&dd&"',門派='出家',身份='和尚' where id="&info(9)
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"
response.end%>