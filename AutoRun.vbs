option Explicit    
dim wmi,proc,procs,proname,flag,WshShell    
Do  
    proname="CQA.exe"  '�����ǽ�������
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
    WshShell.Run ("C:\1\2\3.exe")    '�����ǳ���·����Ҫ�޸�Ϊ���ĳ���·��
end if    
  wscript.sleep 1000 '�����ʱ�䣬������1�뵥λΪms 
loop