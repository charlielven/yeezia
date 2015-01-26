'---------作者主页：http://www.fj543.cn ---------
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
 Shell.MsgBox "端口已被占用，请使用其它端口！或者先停止占用该端口的软件!","543web服务器启动失败"
 Shell.Quit 0
end if
Shell.Service.SetTimer 300
end sub
Sub OnServiceStop()
 sv.Close
End Sub
Sub OnServicePause()
 Shell.Service.Icon="D:\Program Files\543web服务器\skin\off.gif"
 Shell.MsgBox "服务器已暂停！需要时记得打开哦！","543web服务器"
 sv.Stop
End Sub
Sub OnServiceResume()
 OnServiceTimer
 sv.Start
End Sub
Sub OnServiceTimer
 ico1="D:\Program Files\543web服务器\skin\on1.gif"
 ico2="D:\Program Files\543web服务器\skin\on2.gif"
 if Shell.Service.icon=ico1 then Shell.Service.icon=ico2 else Shell.Service.icon=ico1
End sub
