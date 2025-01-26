
#Requires AutoHotkey v2.1-alpha.14
#SingleInstance Force

EditorGui()

class EditorGui {
	__New() {
			this.gui := Gui()
			this.gui.Title := "AHK Code Editor"
			this.gui.OnEvent("Close", (*) => ExitApp())
			
			this.settings := {
					Font: {
							Typeface: "Consolas",
							Size: 10,
							Bold: false
					},
					FGColor: 0xEDEDED,
					BGColor: 0x1E1E1E,
					TabSize: 4,
					WordWrap: true,
					HighlightDelay: 200,
					UseHighlighter: true,
					Highlighter: HighlightAHK,
					Colors: {
							Comments: 0x57A64A,
							Functions: 0x569CD6,
							Keywords: 0x569CD6,
							Strings: 0xD69D85,
							Numbers: 0xB5CEA8,
							Punctuation: 0xD4D4D4,
							Flow: 0x569CD6,
							Commands: 0x569CD6,
							A_Builtins: 0x5F9EA0,
							Directives: 0xC586C0
					}
			}

			this.CreateControls()
			this.SetupHotkeys()
	}

	CreateControls() {
			this.code := RichCode(this.gui, this.settings, "w800 h600")
			
			testCode := '
			(
    // YearDone
    IF DATETRUNC('year', [Order Date]) = DATEADD('year', -1, DATETRUNC('year', TODAY())) THEN
        (SUM([Sales]) - LOOKUP(SUM([Sales]), -1)) / ABS(LOOKUP(SUM([Sales]), -1))
    ELSE
        NULL
    END
			)'
			
			this.code.Text := SubStr(testCode, 2)
			this.gui.Show()
	}

	SetupHotkeys() {
			HotKey("^s", (*) => this.SaveFile())
			HotKey("^o", (*) => this.OpenFile())
	}

	SaveFile(*) {
			if !(file := FileSelect("S"))
					return
			try {
					FileOpen(file, "w").Write(this.code.Text)
					this.gui.Title := "AHK Code Editor - " file
			}
	}

	OpenFile(*) {
			if !(file := FileSelect())
					return
			try {
					this.code.Text := FileRead(file)
					this.gui.Title := "AHK Code Editor - " file
			}
	}
}

class RichCode
{
	#DllLoad "msftedit.dll"
	static IID_ITextDocument := "{8CC497C0-A1DF-11CE-8098-00AA0047BE5D}"
	static MenuItems := ["Cut", "Copy", "Paste", "Delete", "", "Select All", ""
		, "UPPERCASE", "lowercase", "TitleCase"]

	_Frozen := False

	/** @type {Gui.Custom} the underlying control */
	_control := {}

	Settings := {}

	gutter := { Hwnd: 0 }

	; --- Static Methods ---

	static BGRFromRGB(RGB) => RGB >> 16 & 0xFF | RGB & 0xFF00 | RGB << 16 & 0xFF0000

	; --- Properties ---

	Text {
		get => StrReplace(this._control.Text, "`r")
		set => (this.Highlight(Value), Value)
	}
	
	; TODO: reserve and reuse memory
	selection[i := 0] {
		get => (
			this.SendMsg(0x434, 0, charrange := Buffer(8)), ; EM_EXGETSEL
			out := [NumGet(charrange, 0, "Int"), NumGet(charrange, 4, "Int")],
			i ? out[i] : out
		)

		set => (
			i ? (t := this.selection, t[i] := Value, Value := t) : "",
			NumPut("Int", Value[1], "Int", Value[2], charrange := Buffer(8)),
			this.SendMsg(0x437, 0, charrange), ; EM_EXSETSEL
			Value
		)
	}

	SelectedText {
		get {
			Selection := this.selection
			length := selection[2] - selection[1]
			b := Buffer((length + 1) * 2)
			if this.SendMsg(0x43E, 0, b) > length ; EM_GETSELTEXT
				throw Error("Text larger than selection! Buffer overflow!")
			text := StrGet(b, length, "UTF-16")
			return StrReplace(text, "`r", "`n")
		}

		set {
			this.SendMsg(0xC2, 1, StrPtr(Value)) ; EM_REPLACESEL
			this.Selection[1] -= StrLen(Value)
			return Value
		}
	}

	EventMask {
		get => this._EventMask

		set {
			this._EventMask := Value
			this.SendMsg(0x445, 0, Value) ; EM_SETEVENTMASK
			return Value
		}
	}

	_UndoSuspended := false
	UndoSuspended {
		get {
			return this._UndoSuspended
		}

		set {
			try { ; ITextDocument is not implemented in WINE
				if Value
					this.ITextDocument.Undo(-9999995) ; tomSuspend
				else
					this.ITextDocument.Undo(-9999994) ; tomResume
			}
			return this._UndoSuspended := !!Value
		}
	}

	Frozen {
		get => this._Frozen

		set {
			if (Value && !this._Frozen)
			{
				try ; ITextDocument is not implemented in WINE
					this.ITextDocument.Freeze()
				catch
					this._control.Opt "-Redraw"
			}
			else if (!Value && this._Frozen)
			{
				try ; ITextDocument is not implemented in WINE
					this.ITextDocument.Unfreeze()
				catch
					this._control.Opt "+Redraw"
			}
			return this._Frozen := !!Value
		}
	}

	Modified {
		get {
			return this.SendMsg(0xB8, 0, 0) ; EM_GETMODIFY
		}

		set {
			this.SendMsg(0xB9, Value, 0) ; EM_SETMODIFY
			return Value
		}
	}

	; --- Construction, Destruction, Meta-Functions ---

	__New(gui, Settings, Options := "")
	{
		this.__Set := this.___Set
		this.Settings := Settings
		FGColor := RichCode.BGRFromRGB(Settings.FGColor)
		BGColor := RichCode.BGRFromRGB(Settings.BGColor)

		this._control := gui.AddCustom("ClassRichEdit50W +0x5031b1c4 +E0x20000 " Options)

		; Enable WordWrap in RichEdit control ("WordWrap" : true)
		if this.Settings.HasOwnProp("WordWrap")
			this.SendMsg(0x448, 0, 0)

		; Register for WM_COMMAND and WM_NOTIFY events
		; NOTE: this prevents garbage collection of
		; the class until the control is destroyed
		this.EventMask := 1 ; ENM_CHANGE
		this._control.OnCommand 0x300, this.CtrlChanged.Bind(this)

		; Set background color
		this.SendMsg(0x443, 0, BGColor) ; EM_SETBKGNDCOLOR

		; Set character format
		f := settings.font
		cf2 := Buffer(116, 0)
		NumPut("UInt", 116, cf2, 0)          ; cbSize      = sizeof(CF2)
		NumPut("UInt", 0xE << 28, cf2, 4)    ; dwMask      = CFM_COLOR|CFM_FACE|CFM_SIZE
		NumPut("UInt", f.Size * 20, cf2, 12) ; yHeight     = twips
		NumPut("UInt", fgColor, cf2, 20) ; crTextColor = 0xBBGGRR
		StrPut(f.Typeface, cf2.Ptr + 26, 32, "UTF-16") ; szFaceName = TCHAR
		SendMessage(0x444, 0, cf2, this.Hwnd) ; EM_SETCHARFORMAT

		; Set tab size to 4 for non-highlighted code
		tabStops := Buffer(4)
		NumPut("UInt", Settings.TabSize * 4, tabStops)
		this.SendMsg(0x0CB, 1, tabStops) ; EM_SETTABSTOPS

		; Change text limit from 32,767 to max
		this.SendMsg(0x435, 0, -1) ; EM_EXLIMITTEXT

		; Bind for keyboard events
		; Use a pointer to prevent reference loop
		this.OnMessageBound := this.OnMessage.Bind(this)
		OnMessage(0x100, this.OnMessageBound) ; WM_KEYDOWN
		OnMessage(0x205, this.OnMessageBound) ; WM_RBUTTONUP

		; Bind the highlighter
		this.HighlightBound := this.Highlight.Bind(this)

		; Create the right click menu
		this.menu := Menu()
		for Index, Entry in RichCode.MenuItems
			(entry == "") ? this.menu.Add() : this.menu.Add(Entry, (*) => this.RightClickMenu.Bind(this))

		; Get the ITextDocument object
		bufpIRichEditOle := Buffer(A_PtrSize, 0)
		this.SendMsg(0x43C, 0, bufpIRichEditOle) ; EM_GETOLEINTERFACE
		this.pIRichEditOle := NumGet(bufpIRichEditOle, "UPtr")
		this.IRichEditOle := ComValue(9, this.pIRichEditOle, 1)
		; ObjAddRef(this.pIRichEditOle)
		this.pITextDocument := ComObjQuery(this.IRichEditOle, RichCode.IID_ITextDocument)
		this.ITextDocument := ComValue(9, this.pITextDocument, 1)
		; ObjAddRef(this.pITextDocument)
	}

	RightClickMenu(ItemName, ItemPos, MenuName)
	{
		if (ItemName == "Cut")
			Clipboard := this.SelectedText, this.SelectedText := ""
		else if (ItemName == "Copy")
			Clipboard := this.SelectedText
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

	__Delete()
	{
		; Release the ITextDocument object
		this.ITextDocument := unset, ObjRelease(this.pITextDocument)
		this.IRichEditOle := unset, ObjRelease(this.pIRichEditOle)

		; Release the OnMessage handlers
		OnMessage(0x100, this.OnMessageBound, 0) ; WM_KEYDOWN
		OnMessage(0x205, this.OnMessageBound, 0) ; WM_RBUTTONUP

		; Destroy the right click menu
		this.menu := unset
	}

	__Call(Name, Params) => this._control.%Name%(Params*)
	__Get(Name, Params) => this._control.%Name%[Params*]
	___Set(Name, Params, Value) {
		try {
			this._control.%Name%[Params*] := Value
		} catch Any as e {
			e2 := Error(, -1)
			e.What := e2.What
			e.Line := e2.Line
			e.File := e2.File
			throw e
		}
	}

	; --- Event Handlers ---

	OnMessage(wParam, lParam, Msg, hWnd)
	{
		if (hWnd != this._control.hWnd)
			return

		if (Msg == 0x100) ; WM_KEYDOWN
		{
			if (wParam == GetKeyVK("Tab"))
			{
				; Indentation
				Selection := this.Selection
				if GetKeyState("Shift")
					this.IndentSelection(True) ; Reverse
				else if (Selection[2] - Selection[1]) ; Something is selected
					this.IndentSelection()
				else
				{
					; TODO: Trim to size needed to reach next TabSize
					this.SelectedText := this.Settings.Indent
					this.Selection[1] := this.Selection[2] ; Place cursor after
				}
				return False
			}
			else if (wParam == GetKeyVK("Escape")) ; Normally closes the window
				return False
			else if (wParam == GetKeyVK("v") && GetKeyState("Ctrl"))
			{
				this.SelectedText := A_Clipboard ; Strips formatting
				this.Selection[1] := this.Selection[2] ; Place cursor after
				return False
			}
		}
		else if (Msg == 0x205) ; WM_RBUTTONUP
		{
			this.menu.Show()
			return False
		}
	}

	CtrlChanged(control)
	{
		; Delay until the user is finished changing the document
		SetTimer this.HighlightBound, -Abs(this.Settings.HighlightDelay)
	}

	; --- Methods ---

	; First parameter is taken as a replacement Value
	; Variadic form is used to detect when a parameter is given,
	; regardless of content
	Highlight(NewVal := unset)
	{
		if !(this.Settings.UseHighlighter && this.Settings.Highlighter) {
			if IsSet(NewVal)
				this._control.Text := NewVal
			return
		}

		; Freeze the control while it is being modified, stop change event
		; generation, suspend the undo buffer, buffer any input events
		PrevFrozen := this.Frozen, this.Frozen := True
		PrevEventMask := this.EventMask, this.EventMask := 0 ; ENM_NONE
		PrevUndoSuspended := this.UndoSuspended, this.UndoSuspended := True
		PrevCritical := Critical(1000)

		; Run the highlighter
		Highlighter := this.Settings.Highlighter
		if !IsSet(NewVal)
			NewVal := this.text
		RTF := Highlighter(this.Settings, &NewVal)

		; "TRichEdit suspend/resume undo function"
		; https://stackoverflow.com/a/21206620


		; Save the rich text to a UTF-8 buffer
		buf := Buffer(StrPut(RTF, "UTF-8"))
		StrPut(RTF, buf, "UTF-8")

		; Set up the necessary structs
		zoom := Buffer(8, 0) ; Zoom Level
		point := Buffer(8, 0) ; Scroll Pos
		charrange := Buffer(8, 0) ; Selection
		settextex := Buffer(8, 0) ; SetText settings
		NumPut("UInt", 1, settextex) ; flags = ST_KEEPUNDO

		; Save the scroll and cursor positions, update the text,
		; then restore the scroll and cursor positions
		MODIFY := this.SendMsg(0xB8, 0, 0)    ; EM_GETMODIFY
		this.SendMsg(0x4E0, ZOOM.ptr, ZOOM.ptr + 4)   ; EM_GETZOOM
		this.SendMsg(0x4DD, 0, POINT)        ; EM_GETSCROLLPOS
		this.SendMsg(0x434, 0, CHARRANGE)    ; EM_EXGETSEL
		this.SendMsg(0x461, SETTEXTEX, Buf) ; EM_SETTEXTEX
		this.SendMsg(0x437, 0, CHARRANGE)    ; EM_EXSETSEL
		this.SendMsg(0x4DE, 0, POINT)        ; EM_SETSCROLLPOS
		this.SendMsg(0x4E1, NumGet(ZOOM, "UInt")
			, NumGet(ZOOM, 4, "UInt"))        ; EM_SETZOOM
		this.SendMsg(0xB9, MODIFY, 0)         ; EM_SETMODIFY

		; Restore previous settings
		Critical PrevCritical
		this.UndoSuspended := PrevUndoSuspended
		this.EventMask := PrevEventMask
		this.Frozen := PrevFrozen
	}

	IndentSelection(Reverse := False, Indent := unset) {
		; Freeze the control while it is being modified, stop change event
		; generation, buffer any input events
		PrevFrozen := this.Frozen
		this.Frozen := True
		PrevEventMask := this.EventMask
		this.EventMask := 0 ; ENM_NONE
		PrevCritical := Critical(1000)

		if !IsSet(Indent)
			Indent := this.Settings.Indent
		IndentLen := StrLen(Indent)

		; Select back to the start of the first line
		sel := this.selection
		top := this.SendMsg(0x436, 0, sel[1]) ; EM_EXLINEFROMCHAR
		bottom := this.SendMsg(0x436, 0, sel[2]) ; EM_EXLINEFROMCHAR
		this.Selection := [
			this.SendMsg(0xBB, top, 0), ; EM_LINEINDEX
			this.SendMsg(0xBB, bottom + 1, 0) - 1 ; EM_LINEINDEX
		]

		; TODO: Insert newlines using SetSel/ReplaceSel to avoid having to call
		; the highlighter again
		Text := this.SelectedText
		out := ""
		if Reverse { ; Remove indentation appropriately
			loop parse text, "`n", "`r" {
				if InStr(A_LoopField, Indent) == 1
					Out .= "`n" SubStr(A_LoopField, 1 + IndentLen)
				else
					Out .= "`n" A_LoopField
			}
		} else { ; Add indentation appropriately
			loop parse Text, "`n", "`r"
				Out .= "`n" Indent . A_LoopField
		}
		this.SelectedText := SubStr(Out, 2)

		this.Highlight()

		; Restore previous settings
		Critical PrevCritical
		this.EventMask := PrevEventMask

		; When content changes cause the horizontal scrollbar to disappear,
		; unfreezing causes the scrollbar to jump. To solve this, jump back
		; after unfreezing. This will cause a flicker when that edge case
		; occurs, but it's better than the alternative.
		point := Buffer(8, 0)
		this.SendMsg(0x4DD, 0, POINT) ; EM_GETSCROLLPOS
		this.Frozen := PrevFrozen
		this.SendMsg(0x4DE, 0, POINT) ; EM_SETSCROLLPOS
	}

	; --- Helper/Convenience Methods ---

	SendMsg(Msg, wParam, lParam) =>
		SendMessage(msg, wParam, lParam, this._control.Hwnd)
}


class HighlightAHK {
	static needle := (
			"ims)"
			"((?:^|\s)//[^\n]+)" ; Comments
			"|(\[[^\]\n]*\])"    ; Fields
	)

	static Call(Settings, &Code) {
			local FoundPos, Match, rtf, Pos
			GenHighlighterCache(Settings)
			Map := Settings.Cache.ColorMap

			rtf := ""
			Pos := 1

			while FoundPos := RegExMatch(Code, this.needle, &Match, Pos) {
					rtf .= (
							"\cf" Map.Plain " "
							EscapeRTF(SubStr(Code, Pos, FoundPos - Pos))
							"\cf" (
									Match.1 != "" && Map.Comments ||
									Match.2 != "" && Map.Strings ||
									Map.Plain
							) " "
							EscapeRTF(Match.0)
					), Pos := FoundPos + Match.Len
			}

			return Settings.Cache.RTFHeader . rtf
					. "\cf" Map.Plain " " EscapeRTF(SubStr(Code, Pos)) "\`n}"
	}
}


EscapeRTF(Code)
{
	for Char in ["\", "{", "}", "`n"]
		Code := StrReplace(Code, Char, "\" Char)
	return StrReplace(StrReplace(Code, "`t", "\tab "), "`r")
}


GenHighlighterCache(Settings)
{
	if Settings.HasOwnProp("Cache")
		return
	Cache := Settings.Cache := {}
	
	
	; --- Process Colors ---
	Cache.Colors := Settings.Colors.Clone()
	
	; Inherit from the Settings array's base
	BaseSettings := Settings
	while (BaseSettings := BaseSettings.Base)
		if BaseSettings.HasProp("Colors")
			for Name, Color in BaseSettings.Colors.OwnProps()
				if !Cache.Colors.HasProp(Name)
					Cache.Colors.%Name% := Color
	
	; Include the color of plain text
	if !Cache.Colors.HasOwnProp("Plain")
		Cache.Colors.Plain := Settings.FGColor
	
	; Create a Name->Index map of the colors
	Cache.ColorMap := {}
	for Name, Color in Cache.Colors.OwnProps()
		Cache.ColorMap.%Name% := A_Index
	
	
	; --- Generate the RTF headers ---
	RTF := "{\urtf"
	
	; Color Table
	RTF .= "{\colortbl;"
	for Name, Color in Cache.Colors.OwnProps()
	{
		RTF .= "\red"   Color>>16 & 0xFF
		RTF .= "\green" Color>>8  & 0xFF
		RTF .= "\blue"  Color     & 0xFF ";"
	}
	RTF .= "}"
	
	; Font Table
	if Settings.Font
	{
		FontTable .= "{\fonttbl{\f0\fmodern\fcharset0 "
		FontTable .= Settings.Font.Typeface
		FontTable .= ";}}"
		RTF .= "\fs" Settings.Font.Size * 2 ; Font size (half-points)
		if Settings.Font.Bold
			RTF .= "\b"
	}
	
	; Tab size (twips)
	RTF .= "\deftab" GetCharWidthTwips(Settings.Font) * Settings.TabSize
	
	Cache.RTFHeader := RTF
}


GetCharWidthTwips(Font)
{
	static Cache := Map()
	
	if Cache.Has(Font.Typeface "_" Font.Size "_" Font.Bold)
		return Cache[Font.Typeface "_" font.Size "_" Font.Bold]
	
	; Calculate parameters of CreateFont
	Height := -Round(Font.Size*A_ScreenDPI/72)
	Weight := 400+300*(!!Font.Bold)
	Face := Font.Typeface
	
	; Get the width of "x"
	hDC := DllCall("GetDC", "UPtr", 0)
	hFont := DllCall("CreateFont"
	, "Int", Height ; _In_ int     nHeight,
	, "Int", 0      ; _In_ int     nWidth,
	, "Int", 0      ; _In_ int     nEscapement,
	, "Int", 0      ; _In_ int     nOrientation,
	, "Int", Weight ; _In_ int     fnWeight,
	, "UInt", 0     ; _In_ DWORD   fdwItalic,
	, "UInt", 0     ; _In_ DWORD   fdwUnderline,
	, "UInt", 0     ; _In_ DWORD   fdwStrikeOut,
	, "UInt", 0     ; _In_ DWORD   fdwCharSet, (ANSI_CHARSET)
	, "UInt", 0     ; _In_ DWORD   fdwOutputPrecision, (OUT_DEFAULT_PRECIS)
	, "UInt", 0     ; _In_ DWORD   fdwClipPrecision, (CLIP_DEFAULT_PRECIS)
	, "UInt", 0     ; _In_ DWORD   fdwQuality, (DEFAULT_QUALITY)
	, "UInt", 0     ; _In_ DWORD   fdwPitchAndFamily, (FF_DONTCARE|DEFAULT_PITCH)
	, "Str", Face   ; _In_ LPCTSTR lpszFace
	, "UPtr")
	hObj := DllCall("SelectObject", "UPtr", hDC, "UPtr", hFont, "UPtr")
	size := Buffer(8, 0)
	DllCall("GetTextExtentPoint32", "UPtr", hDC, "Str", "x", "Int", 1, "Ptr", SIZE)
	DllCall("SelectObject", "UPtr", hDC, "UPtr", hObj, "UPtr")
	DllCall("DeleteObject", "UPtr", hFont)
	DllCall("ReleaseDC", "UPtr", 0, "UPtr", hDC)
	
	; Convert to twpis
	Twips := Round(NumGet(size, 0, "UInt")*1440/A_ScreenDPI)
	Cache[Font.Typeface "_" Font.Size "_" Font.Bold] := Twips
	return Twips
}
