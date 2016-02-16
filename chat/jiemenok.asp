<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
yn=LCase(trim(request.querystring("yn")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
	Response.End 
end if
too=trim(Application("ljjh_tongmen"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派,同盟 FROM 用戶 WHERE id="&info(9),conn
menpai=rs("門派")
tongmen=rs("同盟")
rs.close
rs.open "select 門派,同盟,姓名 FROM 用戶 WHERE id="&id,conn
menpai1=rs("門派")
tongmen1=rs("同盟")
name=rs("姓名")
if info(0)<>trim(Application("ljjh_tongmen")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你想作什麼，人家也沒有說取你！');}</script>"
	Response.End
end if
if yn=1 then
	Response.Write "<script language=JavaScript>{alert('既然你這麼有誠意我同意結盟了！');}</script>"
if  Instr(1,tongmen,"|無|")=0  then
tongmen=tongmen+menpai1&"|"
else
tongmen="|"&menpai1&"|"
end if
if  Instr(1,tongmen1,"|無|")=0  then
tongmen1=tongmen1+menpai&"|"
else
tongmen1="|"&menpai&"|"
end if
	conn.execute "update 用戶 set 同盟='"& tongmen &"' where 門派='"&info(5)&"'"
	conn.execute "update 用戶 set 同盟='"& tongmen1 &"'  where 門派='"&menpai1&"'"

	jiemen="<font color=FFF0D0>["& menpai &"]</font>與<font color=FFF0D0>["& menpai1 &"]</font>締手結盟，正式結為秦晉之好！！"
   

else
	Response.Write "<script language=JavaScript>{alert('好好玩，別鬧了！');}</script>"
	Response.End

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFF0D0"
sd(197)="FFF0D0"
sd(198)="對"
sd(199)="<font color=FFF0D0>【幫派同盟】</font>"&jiemen
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Application("ljjh_tongmen")=""
%>
 