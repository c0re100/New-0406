<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("jiudian.asa")+";driver={microsoft access driver (*.mdb)};"
connt.open path
sql="delete * from �b�|�C��"
conn.execute(sql)
sql="delete * from �b�|��"
conn.execute(sql)
%>
�b�|��Ʈw�M�ŧ���