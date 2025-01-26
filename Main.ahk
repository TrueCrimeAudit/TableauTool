#Requires AutoHotkey v2.1-alpha.14
#SingleInstance Force
#Warn All, OutputDebug

#DllLoad "Msftedit.dll"

class Theme {
  static Colors := {
    Background: "0x202020", 
    ButtonBg: "0x282828",
    ButtonHover: "0x353535",
    ButtonPressed: "0x1D1D1D",
    Controls: "0x282828",
    Text: "0xFFFFFF",
    TitleBar: "0x151515",
    Accent: "0x383838"
}

  CustomDraw(ctrl, lParam) {
    static CDDS_PREPAINT := 0x1
    nmcd := Buffer(48)
    DllCall("RtlMoveMemory", "Ptr", nmcd, "Ptr", lParam, "Ptr", 48)

    if (NumGet(nmcd, 8, "UInt") = CDDS_PREPAINT) {
      isPressed := GetKeyState("LButton", "P")
      hdc := NumGet(nmcd, 16, "UPtr")

      color := isPressed ? Theme.Colors.ButtonPressed :
              ctrl.Focused ? Theme.Colors.ButtonHover :
              Theme.Colors.ButtonBg

      DllCall("gdi32\SetTextColor", "Ptr", hdc, "UInt", Theme.Colors.Text)
      DllCall("gdi32\SetBkColor", "Ptr", hdc, "UInt", color)
      brush := DllCall("gdi32\CreateSolidBrush", "UInt", color, "Ptr")
      DllCall("gdi32\SelectObject", "Ptr", hdc, "Ptr", brush)
      rc := Buffer(16)
      DllCall("GetClientRect", "Ptr", ctrl.hwnd, "Ptr", rc)
      DllCall("RoundRect", "Ptr", hdc, "Int", 0, "Int", 0, 
             "Int", NumGet(rc, 8, "Int"), "Int", NumGet(rc, 12, "Int"), 
             "Int", 5, "Int", 5)
      DllCall("gdi32\DeleteObject", "Ptr", brush)
      return 0x20
    }
    return 0
  }
}

TM()
class TM {
  d := GuiDim.Create()
  __New() {
    this.Directory := A_ScriptDir "\Tableau\Calculations"
    this.fileArray := []
    this.iniFile := "!!!Tableau.ini"
    this.lastWindow := 0
    this.editField := ""
    this.richEditHwnd := 0
    this.SetupDimensions()
    this.setupGui()
    this.LoadFiles()
    this.SetupHotkeys()
    this.ApplyDarkMode()
    OnMessage(0x0111, this.WM_COMMAND.Bind(this))
    
    pos := this.getScreenPosition()
    this.gui.Show(Format("x{} y{}", pos.x, pos.y))
  }
  SetupHotkeys() {
    HotIfWinActive("ahk_id " this.gui.Hwnd)
    Hotkey("Esc", this.closeGui.Bind(this))
    Hotkey("^s", this.saveEdit.Bind(this))
    Hotkey("WheelUp", (*) => this.handleWheel("Up"))
    Hotkey("WheelDown", (*) => this.handleWheel("Down"))
    Hotkey("e", (*) => this.editSelectedItem())
    HotIfWinExist("ahk_id " this.gui.Hwnd)
    Hotkey("Esc", this.closeGui.Bind(this))
    Hotkey("!Enter", this.sendText.Bind(this))
    Hotkey("^!c", this.copyText.Bind(this))
    Hotkey("!e", this.editSelectedItem.Bind(this))
    HotIfWinExist()
    Hotkey("!y", (*) => this.gui.Show())
  }
  WM_COMMAND(wParam, lParam, msg, hwnd) {
    if ((wParam >> 16) & 0xFFFF = 0x0300 && lParam = this.richEdit.Hwnd)
      this.UpdateFields()
  }
  SetupDimensions() {
    listWidth := 210
    editWidth := 400
    totalWidth := listWidth + (editWidth) + 40
    gui := {
      w: totalWidth,
      h: 530,
      pad: 10
    }
    this.d := {
      gui: gui,
      title: { 
        h: 30 },
      button: {
        w: 100,
        h: 30,
        bar: 30 + gui.pad * 2
      },
      list: {
        w: listWidth,
        x: gui.pad,
        y: 40
      },
      edit: {
        w: editWidth,
        h: 260,
        x: listWidth + gui.pad * 2
      },
      fields: {
        bw: 100,
        h: 30,
        checkWidth: 120
      }
    }
    this.d.list.h := gui.h - (this.d.title.h + this.d.button.bar) - gui.pad
    this.d.button.y := this.d.list.y + this.d.list.h - gui.pad
    this.d.edit.y := this.d.list.y
    this.d.fields.w := this.d.edit.w - this.d.button.w - gui.pad
  }

  GuiFormat(x, y, w, h, extraParams := "") {
    params := Format("x{} y{} w{} h{} Background{:X} c{:X}",
      x, y, w, h, Theme.Colors.Controls, Theme.Colors.Text)
    return extraParams ? params " " extraParams : params
  }

  setupGui() {
    d := this.d
    this.gui := Gui("+AlwaysOnTop -Caption +MinSize" d.gui.w "x" d.gui.h)
    this.gui.MarginX := 0
    this.gui.BackColor := Theme.Colors.Background
    this.gui.SetFont("s11 c" Format("{:X}", Theme.Colors.Text), "Segoe UI")

    this.titleBar := this.setupTitleBar("TextManager")
    this.textList := this.gui.AddListBox(this.GuiFormat(d.list.x, d.list.y, d.list.w, d.list.h))
    this.textList.OnEvent("DoubleClick", this.LoadFileContents.Bind(this))
    
    this.richEdit := RichEdit.Create(this.gui, this.GuiFormat(d.edit.x, d.edit.y, d.edit.w, d.edit.h))

    fdY := d.edit.y + d.edit.h + d.gui.pad
    fdW := d.edit.w - d.button.w - d.gui.pad
    this.fdBox := this.gui.AddEdit(this.GuiFormat(d.edit.x, fdY, fdW, 30, "+ReadOnly -VScroll"))
    
    this.bExtract := this.CreateStyledButton("Extract", d.edit.x + fdW + d.gui.pad, fdY, d.button.w, 30)
    this.bExtract.OnEvent("Click", this.ExtractFields.Bind(this))

    bY := fdY + 40
    pad := this.d.gui.pad
    editWidth := 150
    loop 3 {
      rowY := bY + (A_Index - 1) * (d.fields.h + pad)
      this.gui.AddEdit(this.GuiFormat(d.edit.x, rowY, editWidth - d.gui.pad, d.fields.h))
      this.gui.AddEdit(this.GuiFormat(d.edit.x + editWidth, rowY, editWidth - d.gui.pad, d.fields.h))
      this.CreateStyledButton("Replace", d.edit.x + (editWidth) * 2, rowY, d.button.w, d.fields.h)
    }

    this.editBtn := this.CreateStyledButton("Edit", d.gui.pad, d.button.y, d.button.w, d.button.h)
    this.newBtn := this.CreateStyledButton("New", d.gui.pad * 2 + d.button.w, d.button.y, d.button.w, d.button.h)
    this.saveBtn := this.CreateStyledButton("Save", d.edit.x + 300, d.button.y, d.button.w, d.button.h)
    
    this.setupButtonEvents()
  }
  setupTitleBar(text) {
    tb := this.gui.AddText(Format("x0 y0 w{} h30 Background{} Center 0x200 +0x8", 
                          this.d.gui.w, Format("{:X}", Theme.Colors.TitleBar)), text)
    tb.SetFont("s12 Bold c" Format("{:X}", Theme.Colors.Text))
    
    dragWindow(*) => PostMessage(0xA1, 2, , , "ahk_id " this.gui.Hwnd)
    tb.OnEvent("Click", dragWindow)
    tb.OnEvent("DoubleClick", this.closeGui.Bind(this))
    return tb ; Fixed return value
}
  FileGetName(path) {
    return RegExReplace(path, ".*\\")
  }
  OnRichEditChange(wParam, lParam, msg, hwnd) {
    if ((wParam >> 16) & 0xFFFF = 0x0300)
      this.UpdateFields()
  }
  UpdateFields(*) {
    text := RichEdit.GetText(this.richEdit)
    fields := []
    pos := 1
    while pos := RegExMatch(text, "\[([^\]]+)\]", &match, pos) {
      fields.Push(match[1])
      pos += match.Len
    }
    this.fdBox.Value := fields.Length ? "[" fields.Join("] [") "]" : ""
  }
  ExtractFields(*) {
    try {
      text := RichEdit.GetText(this.richEdit)
      if !text {
        Tooltip("No text found")
        return
      }
      
      ; Store the original selection
      origSel := this.richEdit.selection
      
      ; Find and highlight each field
      pos := 1
      fields := []
      while pos := RegExMatch(text, "\[([^\]]+)\]", &match, pos) {
        field := match[1]
        if !this.HasInArray(fields, field)
          fields.Push(field)
          
        ; Select and color the field
        this.richEdit.selection := [pos - 1, pos + match.Len - 1]
        rtf := "{\rtf{\colortbl;\red100\green220\blue150;}\cf1 " RichEdit.EscapeRTF(match[0]) "}"
        RichEdit.SetRTF(this.richEdit, rtf)
        
        pos += match.Len
      }
      
      ; Restore original selection
      this.richEdit.selection := origSel
      
      if fields.Length {
        output := "[" this.ArrayJoin(fields, "] [") "]"
        this.fdBox.Value := output
      } else {
        Tooltip("No fields found")
        this.fdBox.Value := ""  
      }
    } catch Error as e {
      Tooltip("Error: " e.Message, 2000)
    }
  }


  CreateStyledButton(text, x, y, w, h) {
    btn := this.gui.AddButton(this.GuiFormat(x, y, w, h), text)
    this.StyleButton(btn)
    return btn
  }

  StyleButton(btn) {
    static NM_CUSTOMDRAW := -12
    btn.Opt("+0x4000000") ; WS_CLIPSIBLINGS
    DllCall("uxtheme\SetWindowTheme", "Ptr", btn.hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
    
    static CDDS_PREPAINT := 0x1
    btn.OnNotify(NM_CUSTOMDRAW, this.HandleCustomDraw.Bind(this))
}

; Add this as a separate method in the class
HandleCustomDraw(ctrl, lParam) {
    static CDDS_PREPAINT := 0x1
    nmcd := Buffer(48)
    DllCall("RtlMoveMemory", "Ptr", nmcd, "Ptr", lParam, "Ptr", 48)

    if (NumGet(nmcd, 8, "UInt") = CDDS_PREPAINT) {
        isPressed := GetKeyState("LButton", "P")
        hdc := NumGet(nmcd, 16, "UPtr")
        
        color := isPressed ? Theme.Colors.ButtonPressed : 
                ctrl.Focused ? Theme.Colors.ButtonHover : 
                Theme.Colors.ButtonBg
                
        DllCall("gdi32\SetTextColor", "Ptr", hdc, "UInt", Theme.Colors.Text)
        DllCall("gdi32\SetBkColor", "Ptr", hdc, "UInt", color)
        
        brush := DllCall("gdi32\CreateSolidBrush", "UInt", color, "Ptr")
        DllCall("gdi32\SelectObject", "Ptr", hdc, "Ptr", brush)
        
        rc := Buffer(16)
        DllCall("GetClientRect", "Ptr", ctrl.hwnd, "Ptr", rc)
        DllCall("RoundRect", "Ptr", hdc, "Int", 0, "Int", 0, 
               "Int", NumGet(rc, 8, "Int"), "Int", NumGet(rc, 12, "Int"), 
               "Int", 5, "Int", 5)
               
        DllCall("gdi32\DeleteObject", "Ptr", brush)
        return 0x20 ; CDRF_NOTIFYITEMDRAW
    }
    return 0
}


  HasInArray(arr, needle) {
    for value in arr {
      if value = needle
        return true
    }
    return false
  }

  ArrayJoin(arr, delim) {
    str := ""
    for i, val in arr {
      str .= (i > 1 ? delim : "") val
    }
    return str
  }
  setupButtonEvents() {
    this.newBtn.OnEvent("Click", (*) => this.newItem())
    this.saveBtn.OnEvent("Click", (*) => this.saveEdit(this.textList.Value))
  }
  GetTxtFiles(dir) {
    files := []
    Loop Files dir "\*.txt" {
      files.Push(A_LoopFileName)
    }
    return files
  }
  LoadFiles() {
    filepath := A_ScriptDir "\Tableau\Calculations\*.txt"
    Loop Files filepath {
      this.fileArray.Push({
        name: RegExReplace(A_LoopFileName, "\.txt$")
        , path: A_LoopFilePath
      })
    }
    this.UpdateFileList()
  }
  UpdateFileList() {
    this.textList.Delete()
    for file in this.fileArray {
      this.textList.Add([file.name])
    }
  }
  LoadFileContents(*) {
    selectedIndex := this.textList.Value
    if (selectedIndex > 0) {
      selectedFile := this.fileArray[selectedIndex]
      file := FileOpen(selectedFile.path, "r")
      if IsObject(file) {
        this.gui.Submit(false)
        RichEdit.SetText(this.richEdit, file.Read())
        file.Close()
      }
    }
  }
  updateRichEdit() {
    try {
      if this.textList.Value {
        fullText := StrSplit(this.textList.Text, "|")[2]
        if fullText
          RichEdit.SetText(this.richEdit, fullText)
      }
    }
  }
  newItem(*) {
    items := ControlGetItems(this.textList)
    newIndex := items.Length + 1
    items.Push(newIndex ". New Item")
    this.textList.Delete()
    this.textList.Add(items)
    this.textList.Choose(newIndex)
    RichEdit.SetText(this.richEdit, "")
    this.editSelectedItem()
  }
  replaceItem(*) {
    selectedIndex := this.textList.Value
    if !selectedIndex
      return
    text := RichEdit.GetText(this.richEdit)
    if !text
      return
    IniWrite(text, this.iniFile, "Functions", "key" selectedIndex)
    this.UpdateFileList() ; Use existing UpdateFileList instead of loadTexts
  }
  ApplyDarkMode() {
    for ctrl in this.gui {
      DllCall("uxtheme\SetWindowTheme", "ptr", ctrl.hwnd, "ptr", StrPtr("DarkMode_Explorer"), "ptr", 0)
    }
    if VerCompare(A_OSVersion, "10.0.17763") >= 0 {
      attr := VerCompare(A_OSVersion, "10.0.18985") >= 0 ? 20 : 19
      DllCall("dwmapi\DwmSetWindowAttribute", "ptr", this.gui.hwnd, "int", attr, "int*", true, "int", 4)
    }
    uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
    for procId in [135, 136] {
      DllCall(DllCall("GetProcAddress", "ptr", uxtheme, "ptr", procId, "ptr"), "int", 2)
    }
  }
  MonitorRichEdit(wParam, lParam, msg, hwnd) {
    try {
      if this.richEdit && ((wParam >> 16) & 0xFFFF = 0x0300) && (lParam = this.richEdit.Hwnd)
        this.UpdateFields()
    }
  }
  getScreenPosition() {
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    loop MonitorGetCount() {
      MonitorGet(A_Index, &left, &top, &right, &bottom)
      if (mouseX >= left && mouseX <= right && mouseY >= top && mouseY <= bottom) {
        posX := mouseX + 250
        posY := mouseY - 125
        if (posX + 550 > right)
          posX := right - 550
        if (posY + 320 > bottom)
          posY := bottom - 320
        return { x: posX, y: posY }
      }
    }
    return { x: mouseX + 150, y: mouseY }
  }
  IsTextRunnerVisible() {
    try {
      if !this.gui || !this.textList
        return false
      return WinExist("ahk_id " this.gui.Hwnd)
    }
    return false
  }
  handleWheel(direction) {
    if (!this.IsTextRunnerVisible())
      return
    currValue := this.textList.Value
    items := ControlGetItems(this.textList)
    newValue := direction = "Up" ? currValue - 1 : currValue + 1
    if (newValue >= 1 && newValue <= items.Length)
      this.textList.Choose(newValue)
  }
  ScrollHandler(wParam, lParam, msg, hwnd) {
    if (!this.IsTextRunnerVisible())
      return
    CoordMode("Mouse", "Window")
    MouseGetPos(&mouseX, &mouseY, &mWin)
    if (mWin != this.gui.Hwnd)
      return
    direction := (wParam >> 16) > 0 ? "Up" : "Down"
    this.handleWheel(direction)
    return 0
  }
  cancelEdit(*) {
    if (this.editField) {
      this.editField.Destroy()
      this.editField := ""
    }
  }
  saveEdit(selectedItemIndex, *) {
    try {
      newText := RichEdit.GetText(this.richEdit)
      if !newText
        return
      items := ControlGetItems(this.textList)
      if (selectedItemIndex <= items.Length) {
        prefix := SubStr(items[selectedItemIndex], 1, InStr(items[selectedItemIndex], ".") + 1)
        items[selectedItemIndex] := prefix . newText
        this.textList.Delete()
        this.textList.Add(items)
        this.textList.Choose(selectedItemIndex)
        IniWrite(newText, this.iniFile, "key" selectedItemIndex)
      }
    } catch Error as e {
      MsgBox "Error saving to INI: " e.Message
    }
  }

  copyText(*) => (text := this.textList.Text) && (A_Clipboard := StrReplace(RegExReplace(text, "^\d+\.\s+"), "|", "`n"))

  sendText() {
    if !(text := this.textList.Text) || !this.lastWindow
      return
    WinActivate("ahk_id " this.lastWindow)
    Sleep(50)
    SendText(RegExReplace(text, "^\d+\.\s+"))
  }

  HandleEdit(editGui, selectedIndex) {
    saved := editGui.Submit()
    items := ControlGetItems(this.textList)
    prefix := SubStr(items[selectedIndex], 1, InStr(items[selectedIndex], ".") + 1)

    items[selectedIndex] := prefix . saved.NewValue
    this.textList.Delete()
    this.textList.Add(items)
    this.textList.Choose(selectedIndex)

    IniWrite(saved.NewValue, this.iniFile, "key" selectedIndex)
    editGui.Destroy()
  }

  editSelectedItem(*) {
    if !this.IsTextRunnerVisible() || !this.textList.Text
      return

    selectedItemIndex := this.textList.Value
    if !selectedItemIndex
      return

    this.expanded := true

    originalText := RegExReplace(this.textList.Text, "^\d+\.\s+")
    RichEdit.SetText(this.richEdit, originalText)
    this.richEdit.Focus()
  }

  showGui(*) {
    this.gui.Show()
  }

  closeGui(*) => this.gui.Hide()
  SetTheme(pszSubAppName, pszSubIdList := "") => (!DllCall("uxtheme\SetWindowTheme", "ptr", this.gui.hwnd, "ptr", StrPtr(pszSubAppName), "ptr", pszSubIdList ? StrPtr(pszSubIdList) : 0) ? true : false)

}

;#Region RichEdit
class RichEdit {
  static IID_ITextDocument := "{8CC497C0-A1DF-11CE-8098-00AA0047BE5D}"
  static MenuItems := ["Cut", "Copy", "Paste", "Delete", "", "Select All", "",
    "UPPERCASE", "lowercase", "TitleCase"]

  _Frozen := false
  _UndoSuspended := false
  _control := {}
  _EventMask := 0

  Text {
    get => StrReplace(this._control.Text, "`r")
    set => (this.Highlight(Value), Value)
  }

  selection[i := 0] {
    get => (
      this.SendMsg(0x434, 0, charrange := Buffer(8)),
      out := [NumGet(charrange, 0, "Int"), NumGet(charrange, 4, "Int")],
      i ? out[i] : out
    )
    set => (
      i ? (t := this.selection, t[i] := Value, Value := t) : "",
    NumPut("Int", Value[1], "Int", Value[2], charrange := Buffer(8)),
    this.SendMsg(0x437, 0, charrange),
    Value
    )
  }

  static FormatText(text) {
    colors := {
      white: "\red255\green255\blue255",
      blue: "\red150\green250\blue200",
      red: "\red255\green0\blue0",
      field: "\red100\green220\blue150"
    }

    rtf := "{\rtf{\colortbl;"
        . colors.white ";"
        . colors.blue ";"
        . colors.red ";"
        . colors.field ";"
        . "}\fs20"

    loop parse text, "`n", "`r" {
      line := A_LoopField
      if A_Index = 1
        rtf .= "\cf2 " this.EscapeRTF(line) "\line"
      else
        rtf .= "\cf1 " this.EscapeRTF(line) "\line"
    }
    rtf .= "}"
    return rtf
  }

  SelectedText {
    get {
      Selection := this.selection
      length := selection[2] - selection[1]
      b := Buffer((length + 1) * 2)
      if this.SendMsg(0x43E, 0, b) > length
        throw Error("Text larger than selection! Buffer overflow!")
      text := StrGet(b, length, "UTF-16")
      return StrReplace(text, "`r", "`n")
    }
    set {
      this.SendMsg(0xC2, 1, StrPtr(Value))
      this.Selection[1] -= StrLen(Value)
      return Value
    }
  }
  
  EventMask {
    get => this._EventMask
    set => (this._EventMask := Value, this.SendMsg(0x445, 0, Value), Value)
  }

  UndoSuspended {
    get => this._UndoSuspended
    set {
      try {
        if Value
          this.ITextDocument.Undo(-9999995)
        else
          this.ITextDocument.Undo(-9999994)
      }
      return this._UndoSuspended := !!Value
    }
  }

  Frozen {
    get => this._Frozen
    set {
      if (Value && !this._Frozen) {
        try
          this.ITextDocument.Freeze()
        catch
          this._control.Opt "-Redraw"
      } else if (!Value && this._Frozen) {
        try
          this.ITextDocument.Unfreeze()
        catch
          this._control.Opt "+Redraw"
      }
      return this._Frozen := !!Value
    }
  }

  Modified {
    get => this.SendMsg(0xB8, 0, 0)
    set => (this.SendMsg(0xB9, Value, 0), Value)
  }
  static Create(gui, options) {
    static WM_VSCROLL := 0x0115
    static COLOR_SCROLLBAR := 0
    static COLOR_BTNFACE := 15

    ; Add thin border style and remove the default border
    control := gui.AddCustom("ClassRichEdit50W +0x5031b1c4 +E0x20000 +Wrap -Border +Theme " options)

    ; Set darker background color
    SendMessage(0x0443, 0, 0x202020, control)  
    ; Reduce border width/padding
    SendMessage(0x044D, 4, 0x202020, control)

    DllCall("uxtheme\SetWindowTheme", "Ptr", control.hwnd, "Str", "DarkMode_CFD", "Ptr", 0)
    static WM_SYSCOLORCHANGE := 0x0015
    control.OnMessage(WM_SYSCOLORCHANGE, (ctrl, *) => (
        this.SetSysColor(COLOR_SCROLLBAR, 0x262626),  ; Darker scrollbar color
        this.SetSysColor(COLOR_BTNFACE, 0x262626)     ; Darker button face color
    ))

    PostMessage(WM_SYSCOLORCHANGE, 0, 0, control)

    bufpIRichEditOle := Buffer(A_PtrSize, 0)
    SendMessage(0x43C, 0, bufpIRichEditOle, control)
    pIRichEditOle := NumGet(bufpIRichEditOle, "UPtr")
    IRichEditOle := ComValue(9, pIRichEditOle, 1)
    pITextDocument := ComObjQuery(IRichEditOle, this.IID_ITextDocument)
    control.ITextDocument := ComValue(9, pITextDocument, 1)

    this.menu := Menu()
    for Index, Entry in this.MenuItems
      (entry == "") ? this.menu.Add() : this.menu.Add(Entry, (*) => this.RightClickMenu.Bind(this))

    return control
  }

  RightClickMenu(ItemName, ItemPos, MenuName) {
    if (ItemName == "Cut")
      A_Clipboard := this.SelectedText, this.SelectedText := ""
    else if (ItemName == "Copy")
      A_Clipboard := this.SelectedText
    else if (ItemName == "Paste")
      this.SelectedText := A_Clipboard
    else if (ItemName == "Delete")
      this.SelectedText := ""
    else if (ItemName == "Select All")
      this.Selection := [0, -1]
    else if (ItemName == "UPPERCASE")
      this.SelectedText := Format("{:U}", this.SelectedText)
    else if (ItemName == "lowercase")
      this.SelectedText := Format("{:L}", this.SelectedText)
    else if (ItemName == "TitleCase")
      this.SelectedText := Format("{:T}", this.SelectedText)
  }

  static GetText(control) {
    buf := Buffer(32768)
    SendMessage(0x000D, 32768, buf.Ptr, control)
    return StrGet(buf)
  }

  static SetText(control, text) {
    text := StrReplace(text, "|", "`n")
    this.SetRTF(control, this.FormatText(text))
    SendMessage(0x044D, 6, 0, control)
  }

  static SetRTF(control, rtf) {
    buf := Buffer(StrPut(rtf, "UTF-8"))
    StrPut(rtf, buf, "UTF-8")
    settextex := Buffer(8, 0)
    NumPut("UInt", 0, "UInt", 1200, settextex)
    SendMessage(0x461, settextex, buf, control)
  }

  static EscapeRTF(text) {
    return RegExReplace(text, "([{}])", "\\$1")
  }

  static SetSysColor(index, color) {
    DllCall("SetSysColors", "Int", 1, "Int*", index, "UInt*", color)
    return true
  }

  SendMsg(msg, wParam, lParam) {
    return SendMessage(msg, wParam, lParam, this._control.Hwnd)
  }
}

class GuiDim {
  static p := 10
  static Create() {
    return {
      p: this.p,
      gui: this.Gui(),
      section: this.Sections(),
      size: this.Sizes()
    }
  }
  static Gui() {
    return { w: 800, h: 600 }
  }
  static Sections() {
    return {
      left: { x: this.p, w: 250 },
      right: { x: this.p * 2 + 250, w: 500 }
    }
  }
  static Sizes() {
    return {
      w: { xs: 75, s: 100, m: 150, l: 200, xl: 250 },
      h: { xs: 23, s: 30, m: 40, l: 50, xl: 60 }
    }
  }
  static Format(x, y, w, h, opts := "") {
    return Format("x{} y{} w{} h{} {} Background0x2b2b2b cWhite",
      x, y, w, h, opts ? " " opts : "")
  }
}
