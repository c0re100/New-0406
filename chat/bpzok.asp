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
too=trim(Application("ljjh_bpz"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派 FROM 用戶 WHERE id="&info(9),conn
menpai=rs("門派")
rs.close
rs.open "select 門派,姓名 FROM 用戶 WHERE id="&id,conn
menpai1=rs("門派")
name=rs("姓名")
if info(0)<>trim(Application("ljjh_bpz")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('別亂點！');}</script>"
	Response.End
end if
rs.close
rs.open "select * FROM 幫派大戰 WHERE 敵對幫派='"&menpai&"'",conn
if not rs.bof or not rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你已經參加幫派大戰了！');}</script>"
	Response.End

end if
if yn=1 then
sql="insert into 幫派大戰 (主戰幫派,敵對幫派) values ('"&menpai1&"','"&menpai&"')"
conn.Execute(sql)
	Response.Write "<script language=JavaScript>{alert('開戰就開戰！');}</script>"
	jiemen="<font color=B0D0E0>["& menpai &"]</font>與<font color=B0D0E0>["& menpai1 &"]</font>達成協議，定於明晚20：00開始幫派大戰，請雙方門派弟子做好準備！！！"
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
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
sd(199)="<font color=B0D0E0>【幫派同盟】</font>"&jiemen
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Application("ljjh_bpz")=""
%>
 