<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"jhmp")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
you=trim(request("you"))
if InStr(you,"oR")<>0 or InStr(you,"Or")<>0 or inStr(you,"&")<>0 or inStr(you,"&&")<>0 or InStr(you,"OR")<>0 or InStr(you,"or")<>0 or InStr(you,"=")<>0 or InStr(you,"`")<>0 or InStr(you,"'")<>0 or InStr(you," ")<>0 or InStr(you," ")<>0 or InStr(you,"'")<>0 or InStr(you,chr(34))<>0 or InStr(you,"\")<>0 or InStr(you,",")<>0 or InStr(you,"<")<>0 or InStr(you,">")<>0 then Response.Redirect "../error.asp?id=54"

if info(6)<>"�x��" then Response.Redirect "../error.asp?id=451"
if info(0)=you then
		Response.Write "<script language=JavaScript>{alert('�Y��ĵ�i�A���n�d��');location.href = 'javascript:history.back()';}</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

rs.open "Select ���� from �Τ� where ����='"&info(5)&"' and id="&info(9),conn
if rs("����")<>"�x��" then
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('�A���O�����x�����n�����I�I');window.close();</script>"
			response.end
end if
rs.close
rs.open "Select ���� from �Τ� where ����='"&info(5)&"' and �m�W='"&you&"'",conn
conn.execute "update �Τ� set ����='�̤l',grade=1 where �m�W='"&you&"'"
if rs.eof or rs.bof then
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('���H���O�������I�I');location.href = 'javascript:history.back()';</script>"
			response.end
			else
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('�ӧ̤l�������������I�I');location.href = 'javascript:history.back()';</script>"
			response.end
			end if
%>