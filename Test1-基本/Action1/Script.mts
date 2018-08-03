'’     vbs   基于对象（老婆）语言 
'   1.  名字   ----小红     wscript.shell        ----一个对象  2.  功能  （打开功能   模拟键盘按键    获取路径。。。。）

'1.  打开软件
'   1.1 召唤对象   对象的赋值   Set      不是对象 赋值             a=8
    Set        wsh111=      createobject("wscript.shell")
	  wsh111.Run  "calc.exe"


'   2.输入用例（用户使用的条例）的数据  ---模拟键盘  按键
'  注意点    
'          特殊按键        ctrl 用 ^       shift  用+      Alt       用  %         称之为 占位符号  {} 表示它原来的含义
'  注意点2   -----键盘------慢        CPU  速度快  ---时间不同步
    wait  2
    wsh111.SendKeys   "3{+}5="


'   3.  预期结果 （人为猜测的结果）   实际结果（被测软件的结果）
'  获取实际结果
wait  2
wsh111.SendKeys   "^c"

'   获取的是  获取剪贴板（内存k-v区域   ）数据-htmlfile   对象名字

  Set        htm111=      createobject("htmlfile")

'   获取剪贴板中的数据
      result=  htm111.parentWindow.clipboarddata.getdata("text")
     '  &  拼接变量和常量
	  msgbox    "实际结果是"& result
       
     




'   4.   比对预期和实际结果  ----准则  两者一致  通过   不一致  bug、
   If    result="8"     Then
         msgbox  "测试通过"
	else
	      msgbox "加法功能存在bug"


   End If
