<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="SELECT * FROM 用戶 WHERE id="&ljjh_jhid
              Set Rs=conn.Execute(sql)
              if rs("內力")>=rs("grade")*500000 or rs("體力")>=rs("grade")*500000 then %>
<script language=vbscript>
MsgBox "你喫的太多了，小心撐著"
location.href = "xiaowu3.asp"
</script>
<%else
              sql="select * from 物品 where 類型='食品' and 擁有者='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("數量")<=0 then 
			sql="delete * from 物品 where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu3.asp"
			else
			ti=rs("體力")
			sql="update 物品 set 數量=數量-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update 用戶 set 體力=體力+" & ti & " where id="&ljjh_jhid
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