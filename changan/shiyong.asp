<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr           
sql="select * from ���~ where ����='�åͥΫ~' and �֦���='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("�ƶq")<=0 then 
			sql="delete * from ���~ where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu4.asp"
			else
			ti=rs("��O")
			sql="update ���~ set �ƶq=�ƶq-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update �Τ� set �y�O=�y�O+" & ti & " where �m�W='" & name & "'"
			set rs=conn.execute(sql)
                     conn.close
                                      set rs=nothing
                     Response.Redirect "xiaowu4.asp"
                     response.end
                                    
		end if
%>