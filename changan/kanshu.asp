<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr           
sql="select * from 珇 where 摸='膟' and 局Τ='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("计秖")<=0 then 
			sql="delete * from 珇 where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu9.asp"
			else
                     dd=rs("ず")
			ml=rs("砰")
			sql="update 珇 set 计秖=计秖-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update ノめ set 笵紈=笵紈+" & dd & " ,緔=緔+" & ml & " where ﹎='" & name & "'"
			set rs=conn.execute(sql)
                     conn.close
                     Response.Redirect "xiaowu9.asp"
                     response.end
                                   conn.close
                                   set rs=nothing
		end if
%>