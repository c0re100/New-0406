<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/liugrade")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
id=lcase(trim(request("id")))
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or inStr(id,"&")<>0 or inStr(id,"&&")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from ���~ where ���~�W='�y�P' and �ƶq>0 and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ɯŪZ���ݭn�y�P�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
idd=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-1 where id="&idd
rs.close
rs.open "select ����,����,����,���s from ���~ where id="&id&" and �ƶq>0 and �˳�=false and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���Z�����˳Ƶ۩O�A�Y�n�ɯŽиѰ��˳ơI');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
dj=rs("����")
gj=rs("����")
fy=rs("���s")
randomize timer
r=int(rnd*dj)
if r<8 then
if rs("����")="�k��" then
conn.execute "update ���~ set ����=����+5,����=����+500 where id="&id
end if
if rs("����")="����"  then
conn.execute "update ���~ set ����=����+5,���s=���s+500 where id="&id
end if
if  rs("����")="����" then
conn.execute "update ���~ set ����=����+5,���s=���s+200 where id="&id
end if
if  rs("����")="�Y��" then
conn.execute "update ���~ set ����=����+5,���s=���s+400 where id="&id
end if
if rs("����")="���}" then
conn.execute "update ���~ set ����=����+5,����=����+100,���s=���s+100 where id="&id
end if
if rs("����")="�˹�" then
conn.execute "update ���~ set ����=����+5,���s=���s+100 where id="&id
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ɯŪZ�����\');location.href = 'javascript:history.go(-1)';}</script>"
else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ɯŪZ������');location.href = 'javascript:history.go(-1)';}</script>"
end if%>