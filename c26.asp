<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="update �Τ� set �Ȩ�=int(�Ȩ�),�s��=int(�s��) "
conn.execute(sql)
%>
���ߧA,��O���O�U�[�F50000�I