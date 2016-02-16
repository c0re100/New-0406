<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(3)<3 then
	Response.Write "<script Language=Javascript>alert('你還是江湖小輩，就想來這種地方！!');window.close();</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select 銀兩,性別 from 用戶 where ID=" & info(9)
set rs=conn.execute(sql)
yl=rs("銀兩") 
sex=rs("性別")
if sex="男" then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('迎花宮的帥哥不接待男的，請止步！!');window.close();</script>"
response.end
end if
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="jiu.asp"-->
<% id=request("id")
sql="select 姓名,美貌 from 帥男 where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
set rs=nothing
connt.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你有沒有搞錯呀，迎花宮哪有這個帥哥呀，是不是想人家想瘋了!');window.close();</script>"
response.end
end if
mingji=rs("姓名")
meimao=rs("美貌度")
nl=meimao*2
tl=meimao*0.1
qian=meimao*1.5
if yl<meimao*1.5  then
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('沒錢也想來這種高級的地方呀，請止步！!');window.close();</script>"
response.end
end if
sql="update 用戶 set 內力=內力+"&meimao&"*2,體力=體力-"&meimao&"*0.1 where 姓名='"&nickname&"' "
conn.execute sql
sql1="update 用戶 set 銀兩=銀兩-"&meimao&"*1.5 where 姓名='"&nickname&"' "
conn.execute sql1
connt.close
conn.close
set rs=nothing
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
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="對"
sd(199)="<font color=FFCFCF>【迎花宮消息】"&info(0)&"</font>與迎花宮的<font color=FFCFCF>"&mingji&"</font>傻獃了一晚上，內力增加<font color=FFCFCF>"&nl&"</font>,耗費體力<font color=FFCFCF>"&tl&"</font>，消費了白銀<font color=FFCFCF>"&qian&"兩</font>！！！"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

%>
<html>
<title>歌舞宴會</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
-->
</style>

<body bgcolor="#FFB366">
<p>&nbsp;</p>
<table width="52%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="17" bgcolor="#996633" align="center">&nbsp;</td>
  </tr>
  <tr bgcolor="#66FF66"> 
    <td align="center" height="378" bgcolor="#FFCC66"> 
      <p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400">
          <param name=movie value="girl.swf">
          <param name=quality value=high>
          <embed src="girl.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400">
          </embed> 
        </object><font> </font></p>
    </td>
  </tr>
  <tr bgcolor="#0033CC"> 
    <td align="center" height="26" class="unnamed1" bgcolor="#FFCC66"><b></b></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
