'---------������ҳ��http://www.fj543.cn ---------
Dim sv
Shell.Service.RunService "543webServer","543-Web-Server","WuShiChang's web server for asp"
Sub OnServiceStart()
Shell.Service.stop
Shell.Service.remove
Set sv=CreateObject("543web.HttpServer")
If sv.Create("",8999)=0 Then
 Set host=sv.AddHost("","\")
 host.EnableScript = true
 host.EnableWrite = false
 host.AddDefault "index.html"
 sv.Start
else
 Shell.MsgBox "�˿��ѱ�ռ�ã���ʹ�������˿ڣ�������ֹͣռ�øö˿ڵ����!","543web����������ʧ��"
 Shell.Quit 0
end if
Shell.Service.SetTimer 300
end sub
Sub OnServiceStop()
 sv.Close
End Sub
Sub OnServicePause()
 Shell.Service.Icon="D:\Program Files\543web������\skin\off.gif"
 Shell.MsgBox "����������ͣ����Ҫʱ�ǵô�Ŷ��","543web������"
 sv.Stop
End Sub
Sub OnServiceResume()
 OnServiceTimer
 sv.Start
End Sub
Sub OnServiceTimer
 ico1="D:\Program Files\543web������\skin\on1.gif"
 ico2="D:\Program Files\543web������\skin\on2.gif"
 if Shell.Service.icon=ico1 then Shell.Service.icon=ico2 else Shell.Service.icon=ico1
End sub
