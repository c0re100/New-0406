<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>

<!--#include file="../data1.asp"-->
<%
if trim(request.form("name"))="" or trim(request.form("type"))="" or trim(request.form("url"))="" or request.form("emptytype")="" then
%>
<script language="Vbscript">
msgbox"�ж�g����I",0,"����"
history.back
</script>

<%
conn.close
Response.End
end if
%>
<%

rs.open "select * from bookmark where user='"&info(0)&"' and name='"&request.form("name")&"' and url='http://"&request.form("url")&"'",conn,1,1
if not rs.bof then
%>
<script language="Vbscript">
msgbox"�������w�g�s�b�I",0,"����"
history.back
</script>
<%
rs.close
conn.close
Response.End
end if
%>
<!-- �P�_���}�O�_�X��óB�z -->
<%
url=trim(request.form("url"))
if left(url,7)<>"http://" then
url="http://"&url
end if
%>

<!-- �g�J -->
<%
rs.close
rs.open "select * from bookmark",conn,1,3
rs.addnew
rs.movelast
rs.update "user",info(0)
rs.update "type",request.form("type")
rs.update "emptytype",request.form("emptytype")
rs.update "name",request.form("name")
rs.update "url",url
rs.update "open",request.form("open")
rs.update "datenow",ccdate(date)
rs.update "number",0
rs.close
rs.open "select * from bookmark_count where user='"&info(0)&"'",conn,1,3
if rs.bof then
rs.addnew
rs.movelast
rs.update "user",info(0)
rs.update "count",0
rs.update "allcount",1
else
rs.update "count",rs("count")
rs.update "allcount",rs("allcount")+1
end if
rs.close
conn.close
%>
<script language="Vbscript">
msgbox"�z�w�g���\�a���J�F��ñ�I",0,"����"
history.back
</script>

<%
'=====================================================
'�ഫ���g�J�ƾڮw���榡�A�h���D�k�r��
'�եΡGputmeg(�r�Ŧ�)

function putmeg (inputmsg)
putmeg=replace(inputmsg,chr(13)&chr(10),"<br>")
putmeg=replace(putmeg," ","&nbsp;&nbsp;")
putmeg=replace(putmeg,"'","��")
putmeg=replace(putmeg,"<",".")

end function
'=====================================================


'=====================================================
'�ഫ������Φr�Ť���������榡�p�G2000-01-01
'�եΡGccdate(����ܶq�Τ���榡�r�Ŧ�)

function ccdate(sdate)
temp=cdate(sdate)
if len(month(temp))=1 then
m="0"&month(temp)
else
m=month(temp)
end if

if len(day(temp))=1 then
d="0"&day(temp)
else
d=day(temp)
end if
ccdate=year(temp)&"-"&m&"-"&d
end function
'=====================================================


%>