<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="data1.asp"-->
<%set rs=conn.execute("select * from sheep where user='"&info(0)&"'")
if not rs.eof then
rs.close
conn.execute("delete from sheep where user='"&info(0)&"'")
conn.close
%>
<!--#include file="data.asp"-->
<%set rs=conn.execute("select * from �Τ� where �m�W='"&info(0)&"'")
money=rs("�Ȩ�")+800
conn.execute("update �Τ� set �Ȩ�='"&money&"'  where �m�W='"&info(0)&"'")%>
�A�w�g���\��X�F�p�Ϩño��800�����I


<%else%>
�A�S���p�Ϥ����X
<%rs.close%>
<%end if%>