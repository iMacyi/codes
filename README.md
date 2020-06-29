    option Explicit    
    dim wmi,proc,procs,proname,flag,WshShell    
    Do  
        proname="CQA.exe"  '这里是进程名称
    set wmi=getobject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")    
    set procs=wmi.execquery("select * from win32_process")    
      flag=true    
    for each proc in procs    
        if strcomp(proc.name,proname)=0 then    
          flag=false    
          exit for    
        end if    
    next    
      set wmi=nothing    
      if flag then    
        Set WshShell = Wscript.CreateObject("Wscript.Shell")    
        WshShell.Run ("C:\1\2\3.exe")    '这里是程序路径需要修改为您的程序路径
    end if    
      wscript.sleep 1000 '检测间隔时间，这里是1秒单位为ms 
    loop
