<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select �ާ@�ɶ�,����,�Ȩ�,�ʧO from �Τ� where id="&info(9)
set rs=conn.execute(sql)
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if
dj=rs("����")
yl=rs("�Ȩ�")
xingbie=rs("�ʧO")
if dj<3 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�٬O����p���A�N�Q�ӳo�ئa��!');location.href = 'xiaojie.asp';}</script>"
	response.end
else
if xingbie="�k" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�ɬ��|���h�Q�����ݤk���A�Ф�B!');location.href = 'xiaojie.asp';}</script>"
	response.end
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
id=request("id")
sql="select �m�W,������,���� from �W�� where ID=" & id
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���S���d���r�A���|�����o�өh�Q�r�A�O���O�Q�h�Q�Q�ƤF!');location.href = 'xiaojie.asp';}</script>"
	response.end

else
mingji=rs("�m�W")
meimao=rs("������")
jieshao=rs("����")
if yl<meimao*1.5  then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�S���]�Q�ӳo�ذ��Ū��a��r�A�Ф�B!');location.href = 'xiaojie.asp';}</script>"
	response.end
else
mess="<font color=FFD7F4>�i"& info(0) &"�j</font>�b�ɬ��|�P<font color=FFD7F4>"& mingji &"</font>JJ��s�Z���A�ͤѻ��a����F�@�ӱߤW�C"
sql="update �Τ� set ���O=���O+"&meimao&"/2,��O=��O-500,�ާ@�ɶ�=now() where id="&info(9)
conn.Execute sql
sql1="update �Τ� set �Ȩ�=�Ȩ�-"&meimao&"*1.5 where id="&info(9)
conn.Execute sql1
sql2="update �Τ� set �Ȩ�=�Ȩ�+"&meimao&"*1.0 where �m�W='"&jieshao&"' "
conn.Execute sql2
sql3="update �Τ� set �Ȩ�=�Ȩ�+"&meimao&"*0.5 where �m�W='"&mingji&"' "
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�p�D����:</font>"&mess&".." 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html><title>�q�R�b�|</title><style type="text/css"><!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
-->
</style>
<body bgcolor="#660000">
<p>&nbsp;</p><table width="52%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center"><tr><td height="17" bgcolor="#996633" align="center">&nbsp;</td></tr><tr bgcolor="#66FF66"><td align="center" height="378" bgcolor="#FFCC66"><p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400"><param name=movie value="girl.swf"><param name=quality value=high><embed src="girl.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400"></embed></object><font> </font></p></td></tr><tr bgcolor="#0033CC"><td align="center" height="26" class="unnamed1" bgcolor="#FFCC66"><b></b></td></tr></table>
</body></html> 