<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="SELECT * FROM �Τ� WHERE id="&ljjh_jhid
              Set Rs=conn.Execute(sql)
              if rs("���O")>=rs("grade")*500000 or rs("��O")>=rs("grade")*500000 then %>
<script language=vbscript>
MsgBox "�A�𪺤Ӧh�F�A�p�߼���"
location.href = "xiaowu3.asp"
</script>
<%else
              sql="select * from ���~ where ����='���~' and �֦���='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("�ƶq")<=0 then 
			sql="delete * from ���~ where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu3.asp"
			else
			ti=rs("��O")
			sql="update ���~ set �ƶq=�ƶq-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update �Τ� set ��O=��O+" & ti & " where id="&ljjh_jhid
			set rs=conn.execute(sql)
                     conn.close
                     Response.Redirect "xiaowu3.asp"
                     response.end
    conn.close
    set rs=nothing
		end if
end if 
%>
<html><script language="JavaScript">                                                                  </script></html>