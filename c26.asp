<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="update 用戶 set 銀兩=int(銀兩),存款=int(存款) "
conn.execute(sql)
%>
恭喜你,體力內力各加了50000點