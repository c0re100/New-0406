<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();}</script>"
	Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/es")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if

id=lcase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
myname=info(9)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���~�W,����,���O,��O,���s,����,�Ȩ�,�ƶq,�G���,�֦���,sm,��T��,����,���� from ������� where id=" & id & " and �ƶq>0",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG��������èS���o�ت��~�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
wpname=rs("���~�W")
lx=rs("����")
nl=rs("���O")
tl=rs("��O")
fy=rs("���s")
gj=rs("����")
yin=rs("�Ȩ�")
sl=rs("�ƶq")
say=rs("����")
sm=rs("sm")
dj=rs("����")
jgd=rs("��T��")
esmoney=rs("�G���")
toname=rs("�֦���")
if toname=myname then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG���C�C���ܡA�ۤv�檺�F��ۤv�٭n�R�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "select �Ȩ�,�ާ@�ɶ� from �Τ� where id="&info(9)
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�藍�_�t�έ���A�е�["&s&"����]�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
if rs("�Ȩ�") < esmoney then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�d���F�A�ۤv�S���a�����I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
conn.Execute "update �Τ� set �Ȩ�=�Ȩ�-"& esmoney &",�ާ@�ɶ�=now() where id="&info(9)
conn.Execute "update �Τ� set �Ȩ�=�Ȩ�+"& esmoney &" where id="&toname
conn.execute "delete * from ������� where  id="&id
rs.close
rs.open "select id from ���~ where ���~�W='" & wpname & "' and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~ (���~�W,�֦���,����,���O,��O,����,���s,�ƶq,�Ȩ�,����,��T��,����,sm,�|��) values ('"&wpname&"',"& info(9) &",'"&lx& "',"& nl &","& tl &","& gj &","& fy &","& sl &","& yin &",'"& say &"',"& jgd &","& dj &","&sm&","&aaao&")"
else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+" & sl & ",�|��="&aaao&" where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ܡG�ʶR["& wpname &"]���\�I');location.href = 'esjy.asp';</script>"
response.end
%>
