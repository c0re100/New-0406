<%'全部存點
function ff(fn1)
if info(2)<9 then
	Response.Write "<script language=JavaScript>{alert('操作只有9級以上官府才可以操作！');}</script>"
	Response.End
end if
ff="<script>parent.f3.location.href='savevalue.asp';</script>
%>