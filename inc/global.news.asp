<%
Set Conn = Server.CreateObject("ADODB.Connection")
jiutiandata="DBQ="+server.mappath("news.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open jiutiandata

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUsenews = 3

'---- CommandTypeEnum Values ----
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004

'系統參數設置
pagesize=10

'幫助設置
Public Arr_News_Type(12,2)

Arr_News_Type(1,1)="港龍管理公約"
Arr_News_Type(1,2)="0"

Arr_News_Type(2,1)="港龍會員辦法"
Arr_News_Type(2,2)="0"

Arr_News_Type(3,1)="港龍站長公告"
Arr_News_Type(3,2)="0"

Arr_News_Type(5,1)="港龍江湖幫助"
Arr_News_Type(5,2)="1"

Arr_News_Type(6,1)="入門篇"
Arr_News_Type(6,2)="10"

Arr_News_Type(7,1)="發財篇"
Arr_News_Type(7,2)="10"

Arr_News_Type(8,1)="成家篇"
Arr_News_Type(8,2)="10"

Arr_News_Type(9,1)="練武篇"
Arr_News_Type(9,2)="10"

Arr_News_Type(10,1)="聊天篇"
Arr_News_Type(10,2)="10"

Arr_News_Type(11,1)="動武篇"
Arr_News_Type(11,2)="10"

Arr_News_Type(12,1)="管理篇"
Arr_News_Type(12,2)="10"
%>