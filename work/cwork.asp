<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><title>¾�~�ഫ</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../images/8.jpg">
<%my=info(0)
ziye=trim(request.form("jiu"))
if InStr(ziye,"or")<>0 or InStr(ziye,"'")<>0 or InStr(ziye,"`")<>0 or InStr(ziye,"=")<>0 or InStr(ziye,"-")<>0 or InStr(ziye,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');window.close();}</script>"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ¾�~,����,�Ȩ� from �Τ� where id="&info(9),conn
if left(rs("¾�~"),2)=ziye or left(rs("¾�~"),3)=ziye or left(rs("¾�~"),4)=ziye then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A��ӴN�O�o��¾�~�A���र��¾!');location.href = 'changework.asp';}</script>"
	response.end
end if
dj=rs("����")
if rs("�Ȩ�")<1000000  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾�ݭn100�U��¾�O!');location.href = 'changework.asp';}</script>"
	response.end
end if
if dj<50 and ziye="�ɲ��Ԥh"  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾�ɲ��Ԥh�԰����ť����j����50��!');location.href = 'changework.asp';}</script>"
	response.end
end if 

if ziye="���Z�Ԥh"  then 
if dj<250 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���Z�Ԥh�԰����ť����j����250��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P' and �ƶq>=5 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���Z�Ԥh�ݭn�y�P5��!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-5 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if 

if ziye="���ҾԤh"  then 
if dj<400 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���ҾԤh�԰����ť����j����400��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P�\' and �ƶq>=5 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���ҾԤh�ݭn�y�P�\5��!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-5 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if


if ziye="���s�Ԥh"  then 
if dj<650 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���s�Ԥh�԰����ť����j����650��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P' and �ƶq>=30 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���s�Ԥh�ݭn�y�P30��!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-30 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if

if dj<80 and ziye="�ʾԫi�h"  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾�ʾԫi�h�԰����ť����j����80��!');location.href = 'changework.asp';}</script>"
	response.end

end if

if ziye="��«i�h"  then 
if dj<280 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾��«i�h�԰����ť����j����280��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P' and �ƶq>=5 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾��«i�h�ݭn�y�P5��!');location.href = 'changework.asp';}</script>"	
Response.End 			
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-5 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if
if ziye="���s�i�h"  then 
if dj<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���s�i�h�԰����ť����j����500��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P�\' and �ƶq>=5 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���s�i�h�ݭn�y�P�\5��!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-5 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if
if ziye="����i�h"  then 
if dj<750 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾����i�h�԰����ť����j����750��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P' and �ƶq>=45 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾����i�h�ݭn�y�P45��!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-45 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if
if dj<90 and ziye="���D�h"  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���D�h�԰����ť����j����90��!');location.href = 'changework.asp';}</script>"
	response.end
end if
if ziye="���k�v"  then 
if dj<320 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���k�v�԰����ť����j����320��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P' and �ƶq>=5 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���k�v�ݭn�y�P5��!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-5 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if
if ziye="���u�H"  then 
if dj<550 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���u�H�԰����ť����j����550��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P�\' and �ƶq>=5 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���u�H�ݭn�y�P�\5��!');location.href = 'changework.asp';}</script>"
Response.End 				
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-5 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if
if ziye="���Ѯv"  then 
if dj<780 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���Ѯv�԰����ť����j����780��!');location.href = 'changework.asp';}</script>"
	response.end
else 
rs.close
rs.open "select id from ���~ where ���~�W='�y�P' and �ƶq>=65 and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
				rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��¾���u�H�ݭn�y�P65��!');location.href = 'changework.asp';}</script>"
Response.End 		
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-65 where id="& id &" and �֦���="&info(9)
				
		        end if
end if
end if

if ziye="�Ԥh" or ziye="�ɲ��Ԥh" or ziye="���Z�Ԥh" or ziye="���ҾԤh" or ziye="���s�Ԥh" or ziye="�i�h" or ziye="�ʾԫi�h" or ziye="��«i�h" or ziye="���s�i�h" or ziye="����i�h" or ziye="�D�h" or ziye="���D�h" or ziye="���k�v"  or ziye="���u�H" or ziye="���Ѯv" then
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-1000000,¾�~='"& ziye &"' where id="&info(9),conn


else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');window.close();}</script>"
Response.End 
end if
rs.close
	set rs=nothing
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z¾�~�ഫ���F�G"& ziye &"�u�@�A�I���T�w������e���f�I�I');window.close();}</script>"
Response.End 

%>