<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")

if info(4)=0 then 
aaao=0
else
aaao=1
end if

if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
a=trim(request.form("a"))
b=trim(request.form("b"))
c=trim(request.form("c"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id FROM �Τ� where �m�W='"&a&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�S���ӥΤ�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
idd=rs("id")
rs.close
rs.open "select id from ���~ where ���~�W='�y�P' and �֦���=" & idd,conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,�|��) values ('�y�P',"&idd&",'���~',0,1000,0,0,0,"&b&",10000000,111111,0,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+"&b&",�|��="&aaao&" where id="& id &" and �֦���="&idd
				
		        end if
rs.close
rs.open "select id from ���~ where ���~�W='�y�P�\' and �֦���=" & idd,conn
if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,�|��) values ('�y�P�\',"&idd&",'���~',0,1000,0,0,0,"&c&",10000000,2003,0,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="& id &" and �֦���="&idd
				
		        end if
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�y�P�B�y�P�\�o�e�ާ@')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
