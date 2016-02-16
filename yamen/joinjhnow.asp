<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("ljjh_jm")=session("ljjh_jm")+1
if session("ljjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then
%>
<script language=vbscript>
alert ("你輸入的認證碼不正確，應該輸入:<%=regjm%>")
location.href = "javascript:history.back()"
</script>
<%
response.end
end if

name=LCase(trim(Request.form("name")))
sex=Request.form("sex")
psw=trim(Request.form("psw"))
pswc=trim(Request.Form("pswc"))
pswc2=trim(Request.Form("pswc2"))
e_mail=LCase(Request.form("e_mail"))
oicq=trim(Request.form("oicq"))
zhiye=LCase(trim(Request.form("zhiye")))
regyear=trim(Request.form("regyear"))
regmonth=trim(Request.form("regmonth"))
regday=trim(Request.form("regday"))
'ask=trim(LCase(Request.form("ask")))
'reply=trim(LCase(Request.form("reply")))
face=request("face")
name=trim(name)
name=server.HTMLEncode(name)
jsr=trim(request.form("jsr"))
if zhiye<>"戰士" and  zhiye<>"勇士" and zhiye<>"道士" then
 	Response.Write "<script language=JavaScript>{alert('嚴重警告，不要搞亂');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
if jsr=name then response.redirect "../error.asp?id=57"
if name="無" or name="未定" then response.redirect "../error.asp?id=130"
'if chuser(name) then Response.Redirect "../error.asp?id=120"
'if chuser(jsr) then Response.Redirect "../error.asp?id=60"
if len(oicq)<4 or len(oicq)>=11 then Response.Redirect "../error.asp?id=50"
if instr(e_mail,"@")=0 then Response.Redirect "../error.asp?id=51"
if len(pswc)<5 then Response.Redirect "../error.asp?id=52"
if len(pswc2)<5 then Response.Redirect "../error.asp?id=52"
if len(regyear)<>4 or len(regday)<>2 then Response.Redirect "../error.asp?id=67"
birthday=regyear&"-"&regmonth&"-"&regday
for i=1 to len(name)
usernamechr=mid(name,i,1)
if asc(usernamechr)>0 then  Response.Redirect "../error.asp?id=120"
next
'for i=1 to len(ask)
'usernamechr=mid(ask,i,1)
'if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=53"
'next
'for i=1 to len(reply)
'usernamechr=mid(reply,i,1)
'if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=53"
'next
'if ask=reply then Response.Redirect "../error.asp?id=66"
for i=1 to len(jsr)
usernamechr=mid(jsr,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=60"
next
if instr(name,"or")<>0 or instr(zhiye,"or")<>0 or instr(sex,"or")<>0 or instr(psw,"or")<>0 or instr(pswc,"or")<>0 or instr(pswc2,"or")<>0 or instr(email,"or")<>0 or instr(oicq,"or")<>0 or instr(ask,"or")<>0 or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
if instr(name,"=")<>0  or instr(zhiye,"=")<>0  or instr(sex,"=")<>0 or instr(psw,"=")<>0 or instr(pswc,"=")<>0 or instr(pswc2,"=")<>0 or instr(email,"=")<>0 or instr(oicq,"=")<>0 or instr(ask,"=")<>0  or instr(reply,"or")<>0 then Response.Redirect "../error.asp?id=54"
ip=Request.ServerVariables("REMOTE_ADDR")
if Instr(name,"朋友名字")>0 or Instr(name,"江湖總站")>0 or Instr(name,"站長")>0 or Instr(name,"網管")>0 or Instr(name,"爸")>0 or Instr(name,"嚴")>0 or Instr(name,"媽")>0 or Instr(name,"爸")>0 or Instr(name,"大家")>0 or Instr(name,"操")>0 then Response.Redirect "../error.asp?id=130"
if pswc<>psw then Response.Redirect "../error.asp?id=166"
if trim(request.form("Name"))="" or trim(request.form("psw"))="" or trim(request.form("pswc"))="" or  trim(request.form("pswc2"))="" or trim(Request.Form("e_mail"))="" or trim(request.form("oicq"))="" then Response.Redirect "../error.asp?id=56"
if trim(request.form("Name"))=trim(request.form("psw")) then Response.Redirect "../error.asp?id=129"
if left(name,3)="%20" OR InStr(name,"=")<>0 or InStr(name,"`")<>0 or InStr(name,"'")<>0 or InStr(name," ")<>0 or InStr(name," ")<>0 or InStr(name,"'")<>0 or InStr(name,chr(34))<>0 or InStr(name,"\")<>0 or InStr(name,",")<>0 or InStr(name,"<")<>0 or InStr(name,">")<>0 then Response.Redirect "../error.asp?id=120"
if left(zhiye,3)="%20" OR InStr(zhiye,"=")<>0 or InStr(zhiye,"`")<>0 or InStr(zhiye,"'")<>0 or InStr(zhiye," ")<>0 or InStr(zhiye," ")<>0 or InStr(zhiye,"'")<>0 or InStr(zhiye,chr(34))<>0 or InStr(zhiye,"\")<>0 or InStr(zhiye,",")<>0 or InStr(zhiye,"<")<>0 or InStr(zhiye,">")<>0 then Response.Redirect "../error.asp?id=120"
if left(jsr,3)="%20" OR InStr(jsr,"=")<>0 or InStr(jsr,"`")<>0 or InStr(jsr,"'")<>0 or InStr(jsr," ")<>0 or InStr(jsr," ")<>0 or InStr(jsr,"'")<>0 or InStr(jsr,chr(34))<>0 or InStr(jsr,"\")<>0 or InStr(jsr,",")<>0 or InStr(jsr,"<")<>0 or InStr(jsr,">")<>0 then Response.Redirect "../error.asp?id=120"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 用戶 WHERE 姓名='" & name & "'"
set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
if len(jsr)=0 then
jsr="無"
else
sql1="SELECT * FROM 用戶 WHERE 姓名='" & jsr & "'"
set Rs1=conn.Execute(sql1)
If Rs1.Bof OR Rs1.Eof then
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Redirect "../error.asp?id=58"
response.end
else
ip1=rs1("lastip")
if ip=ip1 then
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Redirect "../error.asp?id=59"
response.end
end if
end if
end if

temppass=StrReverse(left(psw&"godxtll,./",10))
templen=len(psw)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
psw=replace(mmpassword,"'","B")
pswc3=pswc2
temppass=StrReverse(left(pswc2&"godxtll,./",10))
templen=len(pswc2)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pswc2=replace(mmpassword,"'","B")
'沒有記錄 ?
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
userip=Request.ServerVariables("REMOTE_ADDR")
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
if sex="男" then
touxiang="../jhimg/boy00.gif"
else
touxiang="../jhimg/girl00.gif"
end if
sql="insert into 用戶(姓名,會員等級,會員日期,會員費,地區,生日,信箱,oicq,簽名,狀態,性別,年齡,名單頭像,頭像,配偶,密碼,密取,門派,身份,職業,師傅,師傅交錢,寶物,寶物修練,武功,武功加,內力,內力加,體力,體力加,攻擊,攻擊加,防御,防御加,道德,道德加,魅力,魅力加,知質,知質加,法力,法力點,保留1,保留2,保留,貼子數,領錢,等級,grade,times,allvalue,mvalue,regip,lasttime,lastip,lastkick,regtime,IP,登錄,保護,銀兩,存款,門派基金,結算日期,殺人數,kill,暴豆時間,泡豆點數,介紹人,操作時間,二婚,結婚時間,小孩,孩名,金幣,a1,a2,a3,a4,a5,a6,b1,b2,b3,b4,b5,b6,同盟) values('"&name&"',4,date(),5000000,'未知','"&birthday&"','"&e_mail&"','"&oicq&"','仲唔拉人,拉1人+100管','正常','"&sex&"','10','"& face &"','"&touxiang&"','無','"&psw&"','"&pswc2&"','遊俠','弟子','"&zhiye&"','無','無','無',0,200,0,30,0,4000,0,1500,1500,1300,1300,1300,1300,300,0,100,0,100,100,'保留','保留','保留',0," & SqlStr(sj) & ",1,1,1,50,0," & SqlStr(userip) & "," & SqlStr(sj) & "," & SqlStr(userip) & ",'無'," & SqlStr(sj) & ",'"&IP&"',now(),true,1000000,1000000,0,date(),0,'0',now(),0,'"& jsr &"',now(),'無',0,'無','無',1,0,0,0,0,0,0,0,0,0,0,0,0,'|無|')"
conn.execute(sql)
			
sql="SELECT * FROM 用戶 WHERE 姓名='" & name & "'"
Set Rs=conn.Execute(sql)
id=rs("id")
else
Response.Redirect "../error.asp?id=62"
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
sd(199)="<font color=FFD7F4>【新人加入】</font>["&name&"]大俠加入了靈劍江湖，大家對新人要多照顧啊!" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<title>靈劍江湖</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../css.css">
</head>
<body background="../images/8.jpg" leftmargin="0" topmargin="0">
<table border=1 align=center width=372 cellpadding="10" cellspacing="13" height="293" background="../images/12.jpg"> 
<tr valign="top" bgcolor="#FFFFFF"> <td height="226" background="../ljimage/downbg.jpg"> 
<div align="center"> <p><font size="3"><b>您成功加入靈劍江湖!</b></font><br> 
<p><br> <br> <br> </div><table width=100%> <tr> <td> <p align=center style='font-size:14;color:red'><font color="#FF6600">注冊系統ID<font color="#0000FF">:<%=id%></font></font><br> 
<font color="#FF6600">注冊用戶名<font color="#0000FF">:<%=name%></font><br> 注冊時密碼<font color="#0000FF">:<%=pswc%></font><BR>取回密碼<FONT COLOR="#0000FF">:<%=pswc3%></FONT><br> 
注冊聊天頭像：<img src="../ico/<%=face%>-2.gif" width='93' height='83'> <br> <font color="#FF0000"> 
請記住你的<FONT COLOR="#0000FF">取回密碼</FONT>這是你以後找回密碼必需的！</font></font></p><p align=center> 
<input type=button value=關閉 onClick='javascript:window.close()' name="button"> </p></td></tr> 
</table></td></tr> </table>
</body>
</html>