<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr           
sql="select * from ���~ where ����='���y' and �֦���='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("�ƶq")<=0 then 
			sql="delete * from ���~ where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu9.asp"
			else
                     dd=rs("���O")
			ml=rs("��O")
			sql="update ���~ set �ƶq=�ƶq-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update �Τ� set �D�w=�D�w+" & dd & " ,�y�O=�y�O+" & ml & " where �m�W='" & name & "'"
			set rs=conn.execute(sql)
                     conn.close
                     Response.Redirect "xiaowu9.asp"
                     response.end
                                   conn.close
                                   set rs=nothing
		end if
%>