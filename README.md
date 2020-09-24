<div align="center">

## Fun with MouseWheel


</div>

### Description

Just intercepting MouseWheel event with API. Make an empty project (standard exe) and paste code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[vViktor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vviktor.md)
**Level**          |Intermediate
**User Rating**    |4.9 (69 globes from 14 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vviktor-fun-with-mousewheel__1-56768/archive/master.zip)





### Source Code

```
Private Const PM_REMOVE = &H1
Private Type POINTAPI
 x As Long
 y As Long
End Type
Private Type Msg
 hWnd As Long
 Message As Long
 wParam As Long
 lParam As Long
 time As Long
 pt As POINTAPI
End Type
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean
Private Const WM_MOUSEWHEEL = 522
Private Sub ProcessMessages()
 Dim Message As Msg
 Do While Not bCancel
  WaitMessage 'Wait For message and...
  If PeekMessage(Message, Me.hWnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then '...when the mousewheel is used...
   If Message.wParam < 0 Then '...scroll up...
    Me.Top = Me.Top + 240
   Else '... or scroll down
    Me.Top = Me.Top - 240
   End If
  End If
  DoEvents
 Loop
End Sub
Private Sub Form_Load()
 Me.AutoRedraw = True
 Me.Print "Please use now mouse wheel to move this form."
 Me.Show
 ProcessMessages
End Sub
Private Sub Form_Unload(Cancel As Integer)
 bCancel = True
End Sub
```

