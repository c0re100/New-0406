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
%>
<!--#include file="data1.asp"--><%
sql="select * from �Ы� where ID=" & id
rs=conn.execute(sql)
yin=rs("���")
huzhu=rs("��D")
blv=rs("��Q")
zt=rs("���A")
if rs("�Ȩ�")<=0 or rs("��Q�Ȩ�")<=0 or rs("�ƶq")<=0  then
conn.execute "update �Ы� set ��D='�L',��Q='�L',�ƶq=0,�Ȩ�=0,��Q�Ȩ�=0 where ID=" & id
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from �Τ� where id="&info(9)
set rs=conn.execute(sql)
if info(0)=huzhu  and zt="���`" then
      conn.execute "update �Τ� set �Ȩ�=�Ȩ�+" & yin & "*1/5 where id="&info(9)
      %><!--#include file="data1.asp"--><%
      conn.execute "update �Ы� set ��D='�L',��Q='�L',�ƶq=0,�Ȩ�=0,��Q�Ȩ�=0 where ID=" & id
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�w�g���\��X�I');location.href = 'fangwu.asp';</script>"
	Response.End
end if
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��������\�A��]�G�A���O�o�өФl���D�H�άO�A���Фl�X�F�I���p�I');location.href = 'fangwu.asp';</script>"
	Response.End
%>
