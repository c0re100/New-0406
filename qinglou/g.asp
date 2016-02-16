<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select 操作時間,等級,銀兩,性別 from 用戶 where id="&info(9)
set rs=conn.execute(sql)
sj=DateDiff("s",rs("操作時間"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
dj=rs("等級")
yl=rs("銀兩")
xingbie=rs("性別")
if dj<3 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你還是江湖小輩，就想來這種地方!');location.href = 'xiaojie.asp';}</script>"
	response.end
else
if xingbie="女" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('怡紅院的姑娘不接待女的，請止步!');location.href = 'xiaojie.asp';}</script>"
	response.end
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
id=request("id")
sql="select 姓名,美貌度,介紹 from 名妓 where ID=" & id
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你有沒有搞錯呀，本院哪有這個姑娘呀，是不是想姑娘想瘋了!');location.href = 'xiaojie.asp';}</script>"
	response.end

else
mingji=rs("姓名")
meimao=rs("美貌度")
jieshao=rs("介紹")
if yl<meimao*1.5  then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('沒錢也想來這種高級的地方呀，請止步!');location.href = 'xiaojie.asp';}</script>"
	response.end
else
mess="<font color=FFD7F4>【"& info(0) &"】</font>在怡紅院與<font color=FFD7F4>"& mingji &"</font>JJ把酒暢飲，談天說地的聊了一個晚上。"
sql="update 用戶 set 內力=內力+"&meimao&"/2,體力=體力-500,操作時間=now() where id="&info(9)
conn.Execute sql
sql1="update 用戶 set 銀兩=銀兩-"&meimao&"*1.5 where id="&info(9)
conn.Execute sql1
sql2="update 用戶 set 銀兩=銀兩+"&meimao&"*1.0 where 姓名='"&jieshao&"' "
conn.Execute sql2
sql3="update 用戶 set 銀兩=銀兩+"&meimao&"*0.5 where 姓名='"&mingji&"' "
conn.Execute sql3
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end if
end if
end if

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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>小道消息:</font>"&mess&".." 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html><title>歌舞宴會</title><style type="text/css"><!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
-->
</style>
<body bgcolor="#660000">
<p>&nbsp;</p><table width="52%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center"><tr><td height="17" bgcolor="#996633" align="center">&nbsp;</td></tr><tr bgcolor="#66FF66"><td align="center" height="378" bgcolor="#FFCC66"><p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400"><param name=movie value="girl.swf"><param name=quality value=high><embed src="girl.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400"></embed></object><font> </font></p></td></tr><tr bgcolor="#0033CC"><td align="center" height="26" class="unnamed1" bgcolor="#FFCC66"><b></b></td></tr></table>
</body></html> 