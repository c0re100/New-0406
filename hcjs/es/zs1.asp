<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sl=abs(int(request.querystring("sl")))
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
id=lcase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/es")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();}</script>"
	Response.End 
end if
myname=info(0)
name=LCase(trim(Request.form("name")))
zy=LCase(trim(Request.form("zy")))
zy=replace(zy,"'","")
zy=replace(zy,chr(34),"")
zy=Replace(zy,"<","")
zy=Replace(zy,">","")
zy=Replace(zy,"\x3c","")
zy=Replace(zy,"\x3e","")
zy=Replace(zy,"\074","")
zy=Replace(zy,"\74","")
zy=Replace(zy,"\75","")
zy=Replace(zy,"\76","")
zy=Replace(zy,"&lt","")
zy=Replace(zy,"&gt","")
zy=Replace(zy,"\076","")
badstr="�侫|��|ȥ��|��ʺ|����|����|����|��|����|������|����|��|ɵB|����|����|���|����|����|����|����|����|����|����|����|����|����|غ��|��ȥ |���������|���������|���������|��Ƥ|��ͷ|��|�P|��|�H|����|��|��|����|����|����|����|����|��Ů|����|ʮ����|��ү|���|�Ҷ�|����|��|��|asp|com|net|www|xajh|202|61|jh|����|or|261|����|����"
bad=split(badstr,"|")
for i=0 to UBound(bad)
zy=Replace(zy,bad(i),"**")
next
if instr(zy,"or")<>0 or instr(zy,"'")<>0 or instr(name,"or")<>0 or instr(name,"'")<>0 then Response.Redirect "../../error.asp?id=54"
if name="" or zy="" or name=myname then
	Response.Write "<script Language=Javascript>alert('���ܡG�ذe�����\�A�ؤH�P�ۤv�m�W����ۦP�A�m�W�P�ب����ର�šI');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id,�ʧO from �Τ� where �m�W='"&name &"'",conn
nameid=rs("id")
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ذe�����\�A�ҿ�J�m�W:"& name &"�䤣��A�Э��s��J�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if info(1)=rs("�ʧO") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ذe�����\�A§�~�ذe�ȭ��󲧩ʤ����I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "select ���~�W,����,���O,��O,���s,����,�Ȩ�,����,sm,��T��,���� from ���~ where id=" & id & " and ����<>'�d��' and �ƶq>="&sl&" and �֦���="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ذe�����\�A�A�S���o�˪����~�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
wpname=rs("���~�W")
lx=rs("����")
nl=rs("���O")
tl=rs("��O")
fy=rs("���s")
gj=rs("����")
yin=rs("�Ȩ�")
say=rs("����")
sm=rs("sm")
dj=rs("����")
jgd=rs("��T��")
esj=0
conn.execute "update ���~ set �ƶq=�ƶq-"& sl &",�|��="&aaao&" where id=" & id
rs.close
rs.open "select id from ���~ where ���~�W='" & wpname & "' and �֦���="&nameid,conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~ (���~�W,�֦���,����,���O,��O,����,���s,�ƶq,�Ȩ�,����,��T��,����,sm,�|��) values ('"&wpname&"',"& nameid &",'"&lx& "',"& nl &","& tl &","& gj &","& fy &","& sl &","& yin &",'"& say &"',"& jgd &","& dj &","& sm &","&aaao&")"
else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+" & sl & ",�|��="&aaao&" where id="&id
end if

rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ܡG"& myname &"�ذe"& name &"���~�G"& wpname & sl &"���槹���I');location.href = 'wupin.asp';</script>"
response.end
%>

