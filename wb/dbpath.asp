<%
   dim conn   
   dim connstr
   on error resume next
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("ljjmwb.asa")
Conn.Open connstr
%>

