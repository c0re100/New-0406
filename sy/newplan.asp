<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if request.form("topic")="" then%>
<script language=vbscript>
MsgBox "��д��״ֽ����Ŀ��"
location.href = "javascript:history.back()"
</script>
<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=request.form("name")
play=request.form("play")
topic=request.form("topic")
topic=replace(topic,"'","")
topic=replace(topic,chr(34),"")
topic=Replace(topic,"<","")
topic=Replace(topic,">","")
topic=Replace(topic,"\x3c","")
topic=Replace(topic,"\x3e","")
topic=Replace(topic,"\074","")
topic=Replace(topic,"\74","")
topic=Replace(topic,"\75","")
topic=Replace(topic,"\76","")
topic=Replace(topic,"&lt","")
topic=Replace(topic,"&gt","")
topic=Replace(topic,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|asp|com|net|www|xajh|202|61|jh|����|or|261|����|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
topic=Replace(topic,bad(i),"**")
next
sjtext=request.form("text")
sjtext=replace(sjtext,"'","")
sjtext=replace(sjtext,chr(34),"")
sjtext=Replace(sjtext,"<","")
sjtext=Replace(sjtext,">","")
sjtext=Replace(sjtext,"\x3c","")
sjtext=Replace(sjtext,"\x3e","")
sjtext=Replace(sjtext,"\074","")
sjtext=Replace(sjtext,"\74","")
sjtext=Replace(sjtext,"\75","")
sjtext=Replace(sjtext,"\76","")
sjtext=Replace(sjtext,"&lt","")
sjtext=Replace(sjtext,"&gt","")
sjtext=Replace(sjtext,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|asp|com|net|www|xajh|202|61|jh|����|or|261|����|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
sjtext=Replace(sjtext,bad(i),"**")
next

set rs=server.createobject("adodb.recordset")
sql="select * from ��ԩ where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("����")=name
rs("����")=topic
rs("Ҫ��")=play
rs("״��")=sjtext
rs("ԭ��")=info(0)
rs.update
rs.close
conn.close
set conn=nothing
response.redirect("over.asp")
rs.close
set rs=nothing
end if
%>
<html></html> 