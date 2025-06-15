#Requires AutoHotkey v2.0

; 需要配置的全局变量

robotToken := "你的钉钉机器人的 Token"
checkInterval := 20  ; 每 x 次 B 键触发一次屏幕检查

; ======= 以下内容一般无需修改 =======

ListLines(false)

#Include ./lib/ImagePut.ahk
#include ./lib/OCR.ahk
#include ./lib/JSON.ahk

#UseHook
#MaxThreads 30
#MaxThreadsPerHotkey 1
#SingleInstance force

A_MenuMaskKey := "vkE8"

Persistent()

; 全局设置
SendMode("Input")
SetTitleMatchMode(2)
SetWorkingDir(A_ScriptDir)

; 提升脚本进程优先级
PID := DllCall("GetCurrentProcessId")
ProcessSetPriority("High", PID)

; 定义全局变量
antiIdleEnable := false
checkCounter := 0
voiceObject := ComObject("SAPI.SpVoice")
voiceObject.Volume := 100

; 定义窗口组
GroupAdd("TheDivision", "ahk_exe TheDivision2.exe")
GroupAdd("TheDivision", "Division")

DetectHiddenWindows(true)

; ============= 热键定义部分 ============

; 在指定窗口激活时生效的热键
#HotIf WinActive("ahk_group TheDivision")

; 切换自动防退出功能，默认是 Alt+1
!1::
{
  if (!WinExist("ahk_group TheDivision"))
  {
    return
  }

  global antiIdleEnable := !antiIdleEnable
  global checkCounter := 0

  if (antiIdleEnable)
  {
    Speak("启动自动挂机监控")
    AntiIdleTask()
    SetTimer(AntiIdleTask, 1000 * 15)
  }
  else
  {
    Speak("已停止自动挂机监控")
    SetTimer(AntiIdleTask, 0)
  }
}

#HotIf

; ====================================================

; 自定义函数定义

; 模拟按键按下和松开
; Key: 要按下的键名 (例如: "Space", "x")
; Delay: 按键按下的持续时间 (毫秒)
SendKey(Key, Delay := 60)
{
  Send("{" . Key . " down}")
  Sleep(Delay)
  Send("{" . Key . " up}")
}

; 使用 SAPI (语音应用程序接口) 朗读文本
Speak(Text)
{
  voiceObject.Speak(Text, 1)
}

AntiIdleTask()
{
  global antiIdleEnable, checkCounter, checkInterval

  ; 如果防退出功能未启用，则直接返回。
  if (!antiIdleEnable) {
    return
  }

  SendKey("b") ; 发送 B 键

  checkCounter += 1

  if (checkCounter >= checkInterval)
  {
    checkCounter := 0  ; 重置截图计数器

    Sleep(2000)
    CheckScreen()
  }
}

CheckScreen()
{
  now := FormatTime(, "yyyy-MM-dd HH:mm:ss")

  ocrResult := OCR.FromRect(0, 0, A_ScreenWidth, A_ScreenHeight)

  matchIndex := RegExMatch(ocrResult.Text, "DELTA-\d+")

  if (matchIndex > 0)
  {
    SendRobotText("全境封锁2 挂机监控`n`n`n" . ocrResult.Text . "`n`n`n" . now)
  }
}

SendRobotText(text)
{
  global robotToken

  web := ComObject("WinHttp.WinHttpRequest.5.1")
  web.Open("POST", "https://oapi.dingtalk.com/robot/send?access_token=" . robotToken)
  web.SetRequestHeader("Content-Type", "application/json")

  data := Map()

  data["msgtype"] := "text"
  data["text"] := Map()
  data["text"]["content"] := text

  web.Send(Jxon_Dump(data))
  web.WaitForResponse()

  return web.ResponseText
}
