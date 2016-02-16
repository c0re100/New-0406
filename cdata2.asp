<% Response.Buffer = True 
Const JET_3X = 4
Response.write(CData2("news.asp") & "<br> ")
Response.Flush
Response.write(CData2("admin2/hslj/SCHOOL/message.asp") & "<br> ")
Response.Flush
Response.write(CData2("admin2/jr2000/db.asa") & "<br> ")
Response.Flush
Response.write(CData2("admin2/jtbc/guess.asa") & "<br> ")
Response.Flush
Response.write(CData2("admin2/SEA/GLOBAL.ASA") & "<br> ")
Response.Flush
Response.write(CData2("admin2/TAO/TAO.MDB") & "<br> ")
Response.Flush
Response.write(CData2("admin2/yhy1/yhy.mdb") & "<br> ")
Response.Flush
Response.write(CData2("bbs/bbs.asa") & "<br> ")
Response.Flush
Response.write(CData2("CHANGAN/fangzi.mdb") & "<br> ")
Response.Flush
Response.write(CData2("CHANGAN/user.mdb") & "<br> ")
Response.Flush
Response.write(CData2("chat/duju/DMJ.mdb") & "<br> ")
Response.Flush
Response.write(CData2("chat/duju/DPK.mdb") & "<br> ")
Response.Flush
Response.write(CData2("diary/data.asp") & "<br> ")
Response.Flush
Response.write(CData2("GUPIAO/gp.asp") & "<br> ")
Response.Flush
Response.write(CData2("hcjs/wulindahui/FIGHT.MDB") & "<br> ")
Response.Flush
Response.write(CData2("jhjd/jiudian.asa") & "<br> ")
Response.Flush
Response.write(CData2("JHQY/DB.MDB") & "<br> ")
Response.Flush
Response.write(CData2("jhshow/lfshop.mdb") & "<br> ")
Response.Flush
Response.write(CData2("linjiankkk/asz.asa") & "<br> ")
Response.Flush
Response.write(CData2("linjianww/inc/news.asp") & "<br> ")
Response.Flush
Response.write(CData2("myhome/SETUP/SETUP.asp") & "<br> ")
Response.Flush
Response.write(CData2("myhome/SHEEP/SETUP.mdb") & "<br> ")
Response.Flush
Response.write(CData2("myhome/SHEEPBOY/SETUP.MDB") & "<br> ")
Response.Flush
Response.write(CData2("myhome/sheepgirl/SETUP.MDB") & "<br> ")
Response.Flush
Response.write(CData2("shop/shop.asa") & "<br> ")
Response.Flush
Response.write(CData2("wb/ljjmwb.asa") & "<br> ")
Response.Flush
Response.write(CData2("XIANPIAN/cidu-net.asp") & "<br> ")
Response.Flush
Function CData2(dbPath)
dbPath = server.mappath(dbPath)
Dim fso, Engine, strDBPath
strDBPath = Left(dbPath, InStrRev(dbPath, "\"))
strDBPath = strDBPath & "temp.mdb"
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")
On Error Resume Next
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath, _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath
fso.CopyFile strDBPath, dbPath
fso.DeleteFile (strDBPath)
Set fso = Nothing
Set Engine = Nothing
   CData2 = "數據庫壓縮成功!"
Else
   CData2 = "數據庫" & dbPath & "不存在，壓縮失敗!"
End If
End Function
%>
<hr><br><div align="center">在線數據庫完成</div>