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
	Response.Write "<script Language=Javascript>alert('�A�٬O����p���A�N�Q�ӳo�ئa��I!');window.close();</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select �Ȩ�,�ʧO from �Τ� where ID=" & info(9)
set rs=conn.execute(sql)
yl=rs("�Ȩ�") 
sex=rs("�ʧO")
if sex="�k" then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���c���ӭ������ݨk���A�Ф�B�I!');window.close();</script>"
response.end
end if
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="jiu.asp"-->
<% id=request("id")
sql="select �m�W,���� from �Өk where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
set rs=nothing
connt.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A���S���d���r�A���c�����o�ӫӭ��r�A�O���O�Q�H�a�Q�ƤF!');window.close();</script>"
response.end
end if
mingji=rs("�m�W")
meimao=rs("������")
nl=meimao*2
tl=meimao*0.1
qian=meimao*1.5
if yl<meimao*1.5  then
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('�S���]�Q�ӳo�ذ��Ū��a��r�A�Ф�B�I!');window.close();</script>"
response.end
end if
sql="update �Τ� set ���O=���O+"&meimao&"*2,��O=��O-"&meimao&"*0.1 where �m�W='"&nickname&"' "
conn.execute sql
sql1="update �Τ� set �Ȩ�=�Ȩ�-"&meimao&"*1.5 where �m�W='"&nickname&"' "
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="��"
sd(199)="<font color=FFCFCF>�i���c�����j"&info(0)&"</font>�P���c��<font color=FFCFCF>"&mingji&"</font>���ʤF�@�ߤW�A���O�W�[<font color=FFCFCF>"&nl&"</font>,�ӶO��O<font color=FFCFCF>"&tl&"</font>�A���O�F�ջ�<font color=FFCFCF>"&qian&"��</font>�I�I�I"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

%>
<html>
<title>�q�R�b�|</title>
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
