<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
my=info(0)
if instr(action,"'") then response.end 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,�t�� from �Τ� where id="&info(9),conn
'hy=rs("�|��")
dj=rs("����")
po=rs("�t��")
if dj<20 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�з���i�A�����Ťj��20�ťH�W�~�i�H�ʶR!');location.href = 'fangwu.asp';}</script>"

response.end
end if
if po="�L" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�٨S���ѱC�A����R�Ы�!');location.href = 'fangwu.asp';}</script>"

response.end
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%><!--#include file="data1.asp"--><%
sql="select * from �Ы� where ��D='" & my & "' or ��Q='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%><!--#include file="data1.asp"--><%
sql="select * from �Ы� where ID=" & id
rs=conn.execute(sql)
yin=rs("���")
huzhu=rs("��D")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ� from �Τ� where id="&info(9),conn
if rs("�Ȩ�")<=yin and my=huzhu and huzhu<>"�L" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ʶR�����\�A��]�G�A���Ȩ⤣��!');location.href = 'fangwu.asp';}</script>"

response.end
end if
     conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin & " where id="&info(9)
     rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
 %><!--#include file="data1.asp"--><%
      sql="update �Ы� set ��D='" & my & "',��Q='" & po & "' where ID=" & id
	rs=conn.execute(sql)
	conn.close
	Response.Redirect "fangwu.asp"

else %> 
<script language=vbscript>
            MsgBox "�z�αz����Q�w�g�R�L�ЫΤF�C"
            location.href = "fangwu.asp"
                    </script>
<%
   rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
%>
