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
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select �m�W,������ from �W�� where ID=" & id
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�o�̨S���o��h�Q�r�H�I');location.href = 'xiaojie.asp';}</script>"
	response.end
end if
username=rs("�m�W")
meimao=rs("������")
rs.close
rs.open "SELECT ���O,�Z�\,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9)
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<2 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=2-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'xiaojie.asp';}</script>"
	Response.End
end if
if rs("���O")<0 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�S���O�F�r�H�I');location.href = 'xiaojie.asp';}</script>"
	response.end
end if
if rs("���O")<10000 and rs("�Z�\")<10000 then
mess=""&info(0)&"�Z�\�t�l,�Q�ϩh�Q["&username&"]�S�ϥX,�ۤv�ϦӳQ�ɬ��|���@�ìr���F�@�y,�y�O�U��"&meimao&"�I"
conn.execute "update �Τ� set �y�O=�y�O-'"&meimao&"' where id="&info(9)
else
mess=""&info(0)&"�Z�\���j�B���\�ϥX�@��h�Q["&username&"],�D�w�W��"&meimao&"�I!"
randomize timer
r=int(rnd*5)+1
if r=4 then
rs.close
rs.open "select id from ���~ where ���~�W='�ѥP���y' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,�|��) values ('�ѥP���y',"&info(9)&",'���~',0,0,0,0,0,1,200000,99009,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

	end if
mess=""&info(0)&"�Z�\���j�B���\�ϥX�@��h�Q["&username&"],�D�w�W��"&meimao&"�I,�x�����y�m�ѥP���y�n�@��<img src='../hcjs/jhjs/001/99009.gif' border='0'>"
end if
conn.execute "update �Τ� set ���O=���O-10000,�D�w=�D�w+5000,�ާ@�ɶ�=now() where id="&info(9)
conn.execute "delete from �W�� where �m�W='"& username &"'"

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
sd(199)="<font color=FFD7F4>�i�@�Ϧ�ʡj"&info(0)&"</font>�b�ɬ��|�@�ϵL�d�֤k�G<font color=FFD7F4>"&mess&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�@�Ϧ�ʵ����I');location.href = 'xiaojie.asp';}</script>"
	response.end
%> 