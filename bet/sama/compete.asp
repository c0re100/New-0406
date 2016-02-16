<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
username=info(0)
dim horse(11)
horsecalc=0
horselist=""
dim horsepoint()
for i=0 to 11
redim preserve horsepoint(i)
horsetmp=Trim(Request.Form("horse"&i))
if horsetmp="" then
horsetmp=0
elseif not isnumeric(horsetmp) then
horsetmp=0
elseif horsetmp<0 then
horsetmp=0
elseif horsetmp>9999 then
horsetmp=9999
end if
horse(i)=horsetmp
horsepoint(i)=600
horselist=horselist&horsepoint(i)&","
horsecalc=horsecalc+horse(i)
next
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sqlstr="select id from 用戶 where id="&info(9) &" and 銀兩>="&horsecalc
rs.open sqlstr,conn
if rs.EOF or rs.BOF then Response.Redirect "../../error.asp?id=437"
rs.Close
set rs=nothing
randomize()
do while true
redim preserve horsepoint(i+1)
horsepoint(i)=horsepoint(i-12)-clng(rnd()*5)
horselist=horselist&horsepoint(i)&","
if horsepoint(i)<0 then
horsewin=i mod 12
exit do
end if
i=i+1
loop
conn.Execute "update 用戶 set 銀兩=銀兩+"&horse(horsewin)*10-horsecalc&" where id="&info(9)
conn.close
set conn=nothing
erase horsepoint
msg=msg&"<head><title></title><STYLE>"
for i=0 to 11
msg=msg&"#Layer"&i&" {POSITION: absolute;HEIGHT: 60px;WIDTH: 60px; Z-INDEX:3;left:600px;top="&i*30&"px}"
next
msg=msg&"</style><script language=javascript>var pos=0;var i=0;var posxlist='"&horselist&"';var posx=posxlist.split(/,/gi);function start(){for(i=0;i<12;i++){if(posx[pos+i]>=0){eval('document.all["&chr(34)&"Layer'+i+'"&chr(34)&"].style.left=posx[pos+i]');}else{win();return true;}}pos=pos+12;setTimeout('start()',10);}function win(){parent.betfrm.location.replace('win.asp?calc="&horsecalc&"&win="&horsewin&"&bet="&horse(horsewin)&"');}</script></head><body oncontextmenu=self.event.returnValue=false bgcolor='"&bgcolor&"' background='"&bgimage&"'"
if horsecalc<>0 then msg=msg&" onload=start()"
msg=msg&">"
for i=0 to 11
msg=msg&"<DIV id=Layer"&i&"><dd><IMG  src='horse.gif' ></dd></DIV>"
next
msg=msg&"<div style='POSITION: absolute;z-index:2'><table  border=1 bgcolor="&bgcolor&" width='100%' border=1>"
for i=0 to 11
msg=msg&"<tr><td width=32 height=18 align=right>"&i&"號</td><td width=632 height=18 bgcolor="&rgb((11-i)*20,(11+i)*7,i*20)&">&nbsp;</td><td> 賭金<input type=text readonly maxlength=4 size=4 value='"&horse(i)&"'></td></tr>"
next
msg=msg&"</table></div></body>"
erase horse
Response.Write msg
%>
<html><title>靈劍江湖賽馬程序！</title>
<body bgcolor="#CCCCCC">
<br>
</html>
