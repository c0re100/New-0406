<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>

<%Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("jiudian.asa")
Conn.Open connstr
sql="delete * from �b�|��"
set rs=Conn.execute(sql)
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>
