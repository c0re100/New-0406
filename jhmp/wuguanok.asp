<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
my=info(0)
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 and money<>2000000 and money<>10000000 then 
	Response.Write "<script Language=Javascript>alert('�A�Q�@����H�I');location.href = 'javascript:history.back()';;</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ާ@�ɶ�,�Ȩ� from �Τ� where id="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<8 then
	s=8-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�藍�_�t�έ���A�е�["&s&"����]�A�ާ@�I');}</script>"
	Response.End
end if	
if rs("�Ȩ�")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�z���Ȩ⤣���r�A���M�O�v�{�A���L�S�����O�U�U���檺�I');location.href = 'javascript:history.back()';;</script>"
	response.end
end if
conn.execute "update �Τ� set �Z�\=�Z�\+"&int(money/1000)*30&",�Ȩ�=�Ȩ�-"& money &",�ާ@�ɶ�=now() where id="&info(9)
rs.close
rs.open "select ����,�Z�\,�Z�\�[ from �Τ� where id="&info(9),conn
wgj=(rs("����")*1280+3800)+rs("�Z�\�[")
if rs("�Z�\")>wgj then
conn.execute "update �Τ� set �Z�\=" & wgj & ",�ާ@�ɶ�=now() where id="&info(9)
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�P�v�Ŧb�K�Ǥ��W�߾ǲߡI�ש�Ϧۤv���Z�\�j�i,�Z�\�G+"& ((money/1000)*30) &"�I�A��O�Ȩ�G-"&money&"��I');location.href = 'javascript:history.back()';;</script>"
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../images/8.jpg">