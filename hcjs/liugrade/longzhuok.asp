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
rs.open "select id from ���~ where ���~�W='�s�]' and �ƶq>0 and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('��}�Z���ݭn�s�]�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
idd=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-1 where id="&idd
rs.close
rs.open "select ����,����,����,���s,���~�W from ���~ where id="&id&" and �ƶq>0 and �˳�=false and �֦���="&info(9),conn
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
wuname=mid(rs("���~�W"),1,2)
randomize timer
r=int(rnd*dj*2)
if r<2 then
if rs("����")="�k��" then
select case wuname
case "�}�~"
ganliang=Replace(rs("���~�W"),"�}�~","�W�~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',����=����+2000 where id="&id
case "�W�~"
ganliang=Replace(rs("���~�W"),"�W�~","��~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',����=����+3000 where id="&id
case "��~"
ganliang=Replace(rs("���~�W"),"��~","���~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',����=����+4000 where id="&id
case "���~"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���Z���~��w�F�췥�����i�A��}�F�A�o���s�]�ڴN���A�O�ް�,�K�K�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update ���~ set ���~�W='�}�~'&���~�W,���s=���s+1000 where id="&id
end select
end if
if rs("����")="����"  then
select case wuname
case "�}�~"
ganliang=Replace(rs("���~�W"),"�}�~","�W�~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+2000 where id="&id
case "�W�~"
ganliang=Replace(rs("���~�W"),"�W�~","��~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+3000 where id="&id
case "��~"
ganliang=Replace(rs("���~�W"),"��~","���~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+4000 where id="&id
case "���~"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���Z���~��w�F�췥�����i�A��}�F�A�o���s�]�ڴN���A�O�ް�,�K�K�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update ���~ set ���~�W='�}�~'&���~�W,���s=���s+1000 where id="&id
end select
end if
if  rs("����")="����" then
select case wuname
case "�}�~"
ganliang=Replace(rs("���~�W"),"�}�~","�W�~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+1000 where id="&id
case "�W�~"
ganliang=Replace(rs("���~�W"),"�W�~","��~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+2000 where id="&id
case "��~"
ganliang=Replace(rs("���~�W"),"��~","���~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+3000 where id="&id
case "���~"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���Z���~��w�F�췥�����i�A��}�F�A�o���s�]�ڴN���A�O�ް�,�K�K�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update ���~ set ���~�W='�}�~'&���~�W,���s=���s+500 where id="&id
end select
end if
if  rs("����")="�Y��" then
select case wuname
case "�}�~"
ganliang=Replace(rs("���~�W"),"�}�~","�W�~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+600 where id="&id
case "�W�~"
ganliang=Replace(rs("���~�W"),"�W�~","��~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+800 where id="&id
case "��~"
ganliang=Replace(rs("���~�W"),"��~","���~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+900 where id="&id
case "���~"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���Z���~��w�F�췥�����i�A��}�F�A�o���s�]�ڴN���A�O�ް�,�K�K�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update ���~ set ���~�W='�}�~'&���~�W,���s=���s+600 where id="&id
end select
end if
if rs("����")="���}" then
select case wuname
case "�}�~"
ganliang=Replace(rs("���~�W"),"�}�~","�W�~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',����=����+600,���s=���s+600 where id="&id
case "�W�~"
ganliang=Replace(rs("���~�W"),"�W�~","��~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',����=����+700,���s=���s+700 where id="&id
case "��~"
ganliang=Replace(rs("���~�W"),"��~","���~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',����=����+800,���s=���s+800 where id="&id
case "���~"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���Z���~��w�F�췥�����i�A��}�F�A�o���s�]�ڴN���A�O�ް�,�K�K�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update ���~ set ���~�W='�}�~'&���~�W,����=����+500,���s=���s+500 where id="&id
end select
end if
if rs("����")="�˹�" then
select case wuname
case "�}�~"
ganliang=Replace(rs("���~�W"),"�}�~","�W�~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+500 where id="&id
case "�W�~"
ganliang=Replace(rs("���~�W"),"�W�~","��~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+600 where id="&id
case "��~"
ganliang=Replace(rs("���~�W"),"��~","���~")
conn.execute "update ���~ set ���~�W='"&ganliang&"',���s=���s+700 where id="&id
case "���~"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���Z���~��w�F�췥�����i�A��}�F�A�o���s�]�ڴN���A�O�ް�,�K�K�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update ���~ set ���~�W='�}�~'&���~�W,���s=���s+400 where id="&id
end select
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('��}�Z�����\');location.href = 'javascript:history.go(-1)';}</script>"
else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('��}�Z������');location.href = 'javascript:history.go(-1)';}</script>"
end if%>
