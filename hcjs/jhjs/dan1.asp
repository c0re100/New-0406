<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
MsgBox "�жi�J��ѫǦA�i�J��Q�ާ@�I�I"
window.close()
</script>
<%response.end
end if
%>
<%
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
id=lcase(trim(request("id")))
sl=int(abs(Request.form("sl")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
if InStr(sl,"or")<>0 or InStr(sl,"=")<>0 or InStr(sl,"`")<>0 or InStr(sl,"'")<>0 or InStr(sl," ")<>0 or InStr(sl," ")<>0 or InStr(sl,"'")<>0 or InStr(sl,chr(34))<>0 or InStr(sl,"\")<>0 or InStr(sl,",")<>0 or InStr(sl,"<")<>0 or InStr(sl,">")<>0 then Response.Redirect "../../error.asp?id=54"
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ާ@�ɶ� from �Τ� where id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
rs.close
rs.open "SELECT �Ȩ�,�ƶq,id FROM ���~ where ID=" & id & " and ����<>'�d��' and �˳�=false and �֦���="&info(9),conn
if rs.eof and rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S���Ӫ��~�Ϊ̸Ӫ��~���˳Ƥ��I');window.close();</script>"
	response.end
else
sl1=rs("�ƶq")
idd=rs("id")
if sl1>=sl then
conn.execute "delete * from ���~ where id="&idd
yin=int(rs("�Ȩ�")/10)
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+" & yin & ",�ާ@�ɶ�=now() where id="&info(9)
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ܽ椣���\�A�A�S������h���~�i��I');parent.history.go(-1);</script>"
	response.end

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "dan.asp"
end if
%>