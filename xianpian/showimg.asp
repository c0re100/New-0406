<!--#include file="conn.asp"-->
<%
set rec=server.createobject("ADODB.recordset")

strsql="select photo from List where name='" & trim(request("id"))&"'"

rec.open strsql,conn,1,1

Response.ContentType = "image/*"

Response.BinaryWrite rec("photo").getChunk(7500000)

rec.close

set rec=nothing

%>
