<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr           
sql="select * from 物品 where 類型='衛生用品' and 擁有者='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("數量")<=0 then 
			sql="delete * from 物品 where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu4.asp"
			else
			ti=rs("體力")
			sql="update 物品 set 數量=數量-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update 用戶 set 魅力=魅力+" & ti & " where 姓名='" & name & "'"
			set rs=conn.execute(sql)
                     conn.close
                                      set rs=nothing
                     Response.Redirect "xiaowu4.asp"
                     response.end
                                    
		end if
%>