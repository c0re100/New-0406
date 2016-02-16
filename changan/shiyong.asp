<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr           
sql="select * from 珇 where 摸='矫ネノ珇' and 局Τ='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("计秖")<=0 then 
			sql="delete * from 珇 where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu4.asp"
			else
			ti=rs("砰")
			sql="update 珇 set 计秖=计秖-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update ノめ set 緔=緔+" & ti & " where ﹎='" & name & "'"
			set rs=conn.execute(sql)
                     conn.close
                                      set rs=nothing
                     Response.Redirect "xiaowu4.asp"
                     response.end
                                    
		end if
%>