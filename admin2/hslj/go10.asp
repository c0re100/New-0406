<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")<>"hla" then
response.redirect "index.asp"
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 用戶 where id="&info(9)
Set Rs=conn.Execute(sql)
wgj=(rs("等級")*1280+3800)+rs("武功加")
wga=rs("武功")
if wga>=wgj then
sql="update 用戶 set 銀兩=銀兩+2000000,知質=知質+10000,武功='"& wgj &"' where id="&info(9)
conn.Execute(sql)
else
sql="update 用戶 set 銀兩=銀兩+2000000,知質=知質+10000,武功=武功+200000 where id="&info(9)
conn.Execute(sql)
end if
if rs("武功")>=wgj then 
sql="update 用戶 set 武功='"& wgj &"' where id="&info(9)
conn.Execute(sql)
end if
mess="<font color=FFCFCF>"&name&"</font>大俠勇闖華山絕頂，成功擊倒華山五大高手，打開了寶庫,取得了天下無敵的武林秘籍！"
set rs=nothing
conn.close
set conn=nothing
Session("iq")=""
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
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="對"
sd(199)="<font color=FFCFCF>華山傳來的消息:"&mess&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>

<html>

<head>
<title>你終於取到了獨孤求敗的寶藏</title>
<LINK href="../../chat/READONLY/Style.css" rel=stylesheet>
</head>
<body background="images/Bg.gif" oncontextmenu=self.event.returnValue=false >
<div align="center"> <table border="0" width="600"> <tr> <td width="100%"> <p align="left" style="line-height: 200%; margin: 20"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
終於到達獨孤求敗的寶藏了，可是當你用內力震開寶藏的門是，卻發現並沒有什麼天下無敵的武林秘籍，隻是山洞的石壁上龍飛鳳舞的寫著這樣一行字：</b></td></tr> 
<tr> <td width="100%"><img border="0" src="finish.gif"></td></tr> <tr> <td width="100%"> 
<p align="center" style="line-height: 200%; margin: 20"><b><font color="#000000" size="3">天下無敵的秘籍很簡單</font></b><font color="#000000" size="3">，<b>那就是勇氣和自信</b></font></p><p align="left" style="line-height: 200%; margin: 20"><b><font color="#FF0000">&nbsp;</font><font color="#000000">&nbsp;&nbsp; 
原來獨孤求敗終生未求得一敗的原因就是他的勇氣和自信呀！原來我們的勇氣和自信就是天下無敵的秘籍呀！（在山洞裡找到</font><font color="#FF0000">50000000</font><font color="#000000">兩銀子，受石壁上獨孤求敗真跡的啟發，資質增加</font><font color="#FF0000">10000</font><font color="#000000">點，武功增加了</font><font color="#FF0000">200000</font><font color="#000000">點）</font></b> 
</td></tr> </table></div>
</html>
