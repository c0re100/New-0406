<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr           
sql="select * from 物品 where 類型='書籍' and 擁有者='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("數量")<=0 then 
			sql="delete * from 物品 where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu9.asp"
			else
                     dd=rs("內力")
			ml=rs("體力")
			sql="update 物品 set 數量=數量-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update 用戶 set 道德=道德+" & dd & " ,魅力=魅力+" & ml & " where 姓名='" & name & "'"
			set rs=conn.execute(sql)
                     conn.close
                     Response.Redirect "xiaowu9.asp"
                     response.end
                                   conn.close
                                   set rs=nothing
		end if
%>