**MCL** Global VB Hotkey Control
===========================

To use:  Add the OCX to your file,
 Drop the VBHotkey control onto your form
 In the VBHotkey1_HotkeyPressed event add whatever code you wish to be triggered by the event
 Set the control's parameters at design time

 Now whenever your app is running (whether it is the foreground app or not) and the hotkey combination is triggered,
 your app. will execute the code in the event procedure.  When your application is terminated the hotkey will be unloaded.

Parameters:
  AltKey As Boolean       Whether the hotkey combination includes the ALT key
  ShiftKey As Boolean    Whether the hotkey combination includes the SHIFTkey
  CtrlKey As Boolean      Whether the hotkey combination includes the CONTROL key
  WinKey As Boolean     Whether the hotkey combination includes the WIN key
  VKey As KeyCodeConstants    The other key in the combination e.g vbKeyF12 for F12 etc

Notes:  This code is provided free of charge by Merrion Computing Ltd (http://www.merrioncomputing.com).  You are free to use 
   this code subject to acknowledgement in your applications given that you do not attempt to pass the code off as your own
   intelectual copyright.


Release History:
=============

18 June 2001 Release 1.0.3 (Beta) - Now supports multiple VBHotkey controls on each form safely