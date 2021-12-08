# $language = "VBScript"
# $interface = "1.0"

'============================================================================================='
'    全局参数区
'============================================================================================='
Dim OsInfo,g_strSpreadSheetPath,Connected,strPath,DeviceName
' OsInfo:      操作系统类型
' g_strSpreadSheetPath:  包含有需要保存设备列表的excel文件
' strPath：     保存配置文件的目录
' Connected:     连接状态
' DeviceName：    设备名称，主要用于设备配置文件保存的名称中


 '设备参数信息
Dim g_IP_COL, g_Factory_COL, g_USER_COL, g_PASS_COL, g_ACT_COL
Dim strIP, strFactory, strProtocol, strUsername, strPassword, strActive
 ' Convert Letter column indicators to numerical references
 g_IP_COL           = Asc("A") - 64
 g_Factory_COL      = Asc("B") - 64
 g_PROTO_COL        = Asc("C") - 64
 g_USER_COL         = Asc("D") - 64
 g_PASS_COL         = Asc("E") - 64
 g_ACT_COL          = Asc("F") - 64
 g_ACT_DATE_COL     = Asc("G") - 64
 g_ACT_RES_COL      = Asc("H") - 64


'============================================================================================='
'    模块函数（module）区
'============================================================================================='

'---------------------------------------------------------------------------------------
'
'函数SetDevList
'功能：用于选取包含固定格式的excel文件
'
'          A       |   B      |     C    |     D    |     E     |    F    |  G   |    H    |
'   +--------------+----------+----------+----------+-----------+---------+------+---------+
' 1 | IP Address   | Factory  | Protocol | Username | Password  | Active? | Date | Results |
'   +--------------+----------+----------+----------+-----------+---------+------+---------+
' 2 | 192.168.0.1  | huawei   | SSH v2   | admin    | p4$$w0rd  |   Yes   | 1/11 |  Succ   |
'   +--------------+----------+----------+----------+-----------+---------+------+---------+
' 3 | 192.168.0.2  | cisco    | Telnet   | root     | NtheCl33r |   No    | 1/11 |  Fail   |
'   +--------------+----------+----------+----------+-----------+---------+------+---------+
' 4 | 192.168.0.3  | h3c      | SSH v1   | root     | s4f3rN0w! |   Yes   | 1/11 |  Succ   |
'   +--------------+----------+----------+----------+-----------+---------+------+---------+
'---------------------------------------------------------------------------------------
Function SetDevList
 
      Dim Result
Result = ""
Dim IE : Set IE = CreateObject("InternetExplorer.Application")
With IE
.Visible = False
.Navigate("about:blank")
Do Until .ReadyState = 4 : Loop
With .Document
.Write "<html><body><input id='f' type='file'></body></html>"
With .All.f
.Focus
.Click
Result = .Value
End With
End With
.Quit
End With
Set IE = Nothing
ChooseFile = Result
End function


'---------------------------------------------------------------------------------------
'


'---------------------------------------------------------------------------------------
'
'函数selectsavepath
'功能：用于设置配置文件保存目录
'
'---------------------------------------------------------------------------------------
Function selectsavepath
 Const MY_COMPUTER = &H11&
 Const WINDOW_HANDLE = 0
 Const OPTIONS = 0

 Set objShell = CreateObject("Shell.Application")
 Set objFolder = objShell.BrowseForFolder (WINDOW_HANDLE, "Please select folder to save the configeration:", OPTIONS, strPath)
 If objFolder Is Nothing Then
  Set objShell = CreateObject("Shell.Application")
  Set objFolder = objShell.Namespace(MY_COMPUTER)
 End If
 Set objFolderItem = objFolder.Self
 strPath = objFolderItem.Path
End Function

'---------------------------------------------------------------------------------------
'
'函数setlog(DeviceName)
'功能：用于启动日志保存格式和文件名
'
'---------------------------------------------------------------------------------------
Function setlog(DeviceName)
 Dim CurrentTime
 CurrentTime=year(Now)&"-"&Month(Now)&"-"&day(Now)&"_"&Hour(Now)&"."&Minute(Now)&"."&Second(Now)
 If crt.Session.Logging=true Then
   crt.Session.Log False
   crt.Sleep 2000
 end If
 crt.session.LogFileName = strPath & "\" & DeviceName & CurrentTime & "_cfg.log"
 'msgbox(crt.session.LogFileName)
 crt.session.Log(true)
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ConnectDevice
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
On Error Resume Next
 crt.Sleep 5000
    g_strError = ""
    Err.Clear
    Connected = False
   
    Dim strCmd
    Select Case UCase(strProtocol)
        Case "TELNET"
            strCmd = "telnet " & strIP
      crt.Screen.Send strCmd & chr(13)
   crt.Screen.WaitForString("sername:")
   crt.Screen.Send strUsername & chr(13)
        Case "SSH V2"
   strCmd="ssh "&strUsername &"@"&strIP
      crt.Screen.Send strCmd & chr(13)
   'msgbox(strCmd)
        Case "SSH V1"
   strCmd="ssh -1 "&strUsername &"@"&strIP
      crt.Screen.Send strCmd & chr(13)
   'msgbox(strCmd)   
        Case Else
            ' Unsupported protocol
            g_strError = "Unsupported protocol: " & strProtocol
            Exit Function
    End Select
 crt.Sleep 5000
 crt.Screen.WaitForString("assword:")
 crt.Screen.Send strPassword & chr(13)
    If crt.Session.Connected <> True Then Exit Function
    crt.Screen.Synchronous = True

    Select Case strFactory
  Case "Huawei","H3C","Huawei-3Com"
   If crt.Screen.WaitForString(">",5) =True then
    Connected = True
   End if
  Case "cisco"
   If crt.Screen.WaitForString("#",5) =True then
    Connected = True
   End if
 End Select
 
    If Err.Number <> 0 Then
        g_strError = Err.Description
    End If

    On Error Goto 0
End Function ' End Function

'============================================================================================='
'    程序主函数（main）区
'============================================================================================='

'主函数
Sub Main
' 获取系统信息，采用不同的脚本获取设备列表文件
' GetOsInfo

' 选择需要对配置进行保存的设备列表，选择excel文件
 SetDevList
' 设置需要保存配置的目录
 selectsavepath

' 对选取的excel文件进行处理
 Dim g_objExcel
    Set g_objExcel = CreateObject("Excel.Application")

    Dim objWkBook
    Set objWkBook = g_objExcel.Workbooks.Open(g_strSpreadSheetPath)
   
    Dim objSheet
    Set objSheet = objWkBook.Sheets(1)

    Dim nRowIndex
    nRowIndex = 2

    Do
        strIP = Trim(objSheet.Cells(nRowIndex, g_IP_COL).Value)
        If strIP = "" Then Exit Do
        strActive = Trim(objSheet.Cells(nRowIndex, g_ACT_COL).Value)
        If LCase(strActive) = "yes" Then
      Dim bSuccess
            strfactory  = Trim(objSheet.Cells(nRowIndex, g_Factory_COL).Value)
            strProtocol = Trim(objSheet.Cells(nRowIndex, g_PROTO_COL).Value)
            strUsername = Trim(objSheet.Cells(nRowIndex, g_USER_COL).Value)
            strPassword = Trim(objSheet.Cells(nRowIndex, g_PASS_COL).Value)
            ConnectDevice

            If Connected<>true Then
                    objSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = "Failure: Unable to Connect"
            Else
    crt.Screen.Send vbcr
    Select Case strFactory
     Case "Huawei","H3C","Huawei-3Com"
      '获取设备名称定义
      DeviceName = crt.Screen.ReadString(">")
      DeviceName = right(DeviceName,Len(DeviceName)-InStrRev(DeviceName,"<"))
      '设置开始记录日志
      setlog(DeviceName&"_"&StrIP&"_")
      crt.Screen.Send "dis curr" & chr(13)
      Do
       Do while crt.Screen.WaitForString("More",3)
        crt.Screen.Send " "
       Loop
       crt.Screen.Send chr(13)
      Loop Until crt.Screen.WaitForString(DeviceName&">") = True
      'msgbox("总循环次数:"&i)
      crt.Screen.Send chr(13)
      If crt.Screen.WaitForString(DeviceName&">") = True Then
       'msgbox("已经完成对设备<" & DeviceName & ">的配置备份." &  "备份配置文件位置:" & crt.session.LogFileName)
       objSheet.Cells(nRowIndex, g_ACT_DATE_COL).Value = Now
       bSuccess=True
       crt.Screen.Send "quit" & chr(13)
       crt.Sleep 2000
      End if
     Case "cisco"
      DeviceName = crt.Screen.ReadString("#")
      Do While InStr(DeviceName,vbCrLf)<>0
       DeviceName = replace(DeviceName,vbCrLf," ")
      Loop
      Do While InStr(DeviceName,vbCr)<>0
       DeviceName = replace(DeviceName,vbCr," ")
      Loop
      DeviceName = trim(DeviceName)
      setlog(DeviceName&"_"&StrIP&"_")
      crt.Screen.Send "show running" & chr(13)
      Do
       Do while crt.Screen.WaitForString("More",3)
        crt.Screen.Send " "
       Loop
       crt.Screen.Send chr(13)
      Loop Until crt.Screen.WaitForString(DeviceName&"#") = True
      crt.Screen.Send chr(13)
      If crt.Screen.WaitForString(DeviceName&"#") = True Then
       objSheet.Cells(nRowIndex, g_ACT_DATE_COL).Value = Now
       bSuccess=True
       crt.Screen.Send "exit" & chr(13)
       crt.Sleep 2000
      Else
       msgbox("设备<" & DeviceName & ">的配置备份失败!")
      End If
    End Select
    
    
    '在excel文件中设置执行结果是成功还是失败，并记录执行时间
    Set objCell = objSheet.Cells(nRowIndex, g_ACT_RES_COL)
                If bSuccess Then
                    objCell.Value  = "Success!! Configuration file located:"&crt.session.LogFileName
    Else
     objCell.Value  = "Failure: Command failed. Matchindex = " & crt.Screen.MatchIndex
     '需要添加删除未备份完的备份文件。
                End If
            End If
        Else
            ' mark the skipped ones in the spreadsheet
   If objSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = "" Then
    objSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = "Skipped"
          ' We always record the date of action status
    objSheet.Cells(nRowIndex, g_ACT_DATE_COL).Value = Now 
   End if
        End If
               
        crt.session.Log(false)
  crt.screen.Synchronous = false
        nRowIndex = nRowIndex + 1
    Loop
   
    objWkBook.Save
    objWkBook.Close
    g_objExcel.Quit
    Set g_objExcel = Nothing
    msgbox("已完成全部设备配置备份!")
'    g_shell.Run Chr(34) & g_strSpreadSheetPath & Chr(34)


'    result = crt.Dialog.MessageBox("信息收集完毕，是否退出CRT?", "提示信息", ICON_QUESTION Or BUTTON_YESNO Or DEFBUTTON2)
'    If    result = IDYES Then
'        crt.quit
'    End If

End Sub