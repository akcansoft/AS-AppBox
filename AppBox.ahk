/*
AS AppBox
-----------------------------
This script serves as an application manager that helps organize and manage your frequently used programs.
Main features:
- Add/remove applications to the list
- Run selected applications
- Filter applications by name
- Save/load application lists
- Drag and drop support
- Context menu for quick actions
- Multiple view modes (list/icon)
- Open file locations
-----------------------------
v1.0 R16
25/04/2025
-----------------------------
Mesut Akcan
makcan@gmail.com
youtube.com/mesutakcan
mesutakcan.blogspot.com
github.com/akcansoft
-----------------------------
*/

#Requires AutoHotkey v2.0
#SingleInstance Force

appName := "AS AppBox"
appVer := "1.0"
settingsFile := A_ScriptDir "\settings.ini"
listFile := A_ScriptDir "\applist.txt"
appListChanged := false
iconLoadIndex := 0
isLargeView := false
; Create ImageLists for icons
IL_ID1 := IL_Create(, , false) ; small
IL_ID2 := IL_Create(, , true) ; large

; Menu Text
menuTxt := {
  add: "&Add`tIns",
  remove: "Re&move selected`tDel",
  run: "&Run application`tEnter",
  open: "&Open file location",
  copy: "&Copy full path",
  save: "&Save list`tCtrl+S",
  clear: "&Clear list",
  select: "&Select all`tCtrl+A",
  deselect: "&Deselect all`tEsc",
  exit: "&Exit`tAlt+F4",
  about: "&About",
  file: "&File",
  help: "&Help"
}

; GUI Controls
mainGui := Gui("+Resize +MinSize400x300", appName)
mainGui.Opt("+OwnDialogs")
; Filter Controls
mainGui.Add("Text", "y+10", "Filter:")
txtFilter := mainGui.Add("Edit", "x+5 yp-3 w200")
btnClearFilter := mainGui.Add("Button", "x+5", "❌")
btnClearFilter.OnEvent("Click", (*) => (txtFilter.Value := "", FilterList()))
txtFilter.OnEvent("Change", FilterList)

; Toolbar Buttons
mainGui.SetFont("s10")
buttonHeight := 30
btnAdd := mainGui.Add("Button", "x10 y+10 h" buttonHeight, "➕ &Add")
btnAdd.OnEvent("Click", AddToList)
btnRemove := mainGui.Add("Button", "Disabled x+5 h" buttonHeight, Chr(0x2796) " &Remove")
btnRemove.OnEvent("Click", RemoveSelected)
btnRun := mainGui.Add("Button", "Disabled x+5 h" buttonHeight, "🚀 &Run")
btnRun.OnEvent("Click", RunSelected)
btnSave := mainGui.Add("Button", "Disabled x+5 h" buttonHeight, "💾 &Save")
btnSave.OnEvent("Click", SaveList)
mainGui.Add("Button", "x+5 h" buttonHeight, "🔃 &Switch view").OnEvent("Click", SwitchView)
mainGui.Add("Button", "x+5 h" buttonHeight, "⏻ &Exit").OnEvent("Click", g1Close)

; ListView & StatusBar
mainGui.SetFont("s9")
LV1 := mainGui.Add("ListView", "x10 y+5 r15 w380 Grid", ["File name", "Full path"])
LV1.OnEvent("DoubleClick", RunSelected)
LV1.OnEvent("ItemSelect", UpdateUIState)
LV1.OnEvent("ContextMenu", ShowContextMenu)
sbMain := mainGui.Add("StatusBar")

mainGui.OnEvent("Size", GuiResize)
mainGui.OnEvent("Close", g1Close)
mainGui.OnEvent("DropFiles", DropHandler)

; Create Menu
fileMenu := Menu()
AddMenuItems(fileMenu)
fileMenu.Add() ; Separator
fileMenu.Add(menuTxt.exit, g1Close)
rcMenu := Menu() ; Context menu
AddMenuItems(rcMenu)
helpMenu := Menu()
helpMenu.Add(menuTxt.about, About)
mnuBar := MenuBar()
mnuBar.Add(menuTxt.file, fileMenu)
mnuBar.Add(menuTxt.help, helpMenu)
mainGui.MenuBar := mnuBar

; Hotkeys
#HotIf WinActive("ahk_id " mainGui.Hwnd)
Insert:: AddToList()
^s:: (btnSave.Enabled ? SaveList() : SoundBeep())
^a:: SelectAll(true)
Esc:: SelectAll(false)
Del:: (btnRemove.Enabled ? RemoveSelected() : SoundBeep())
#HotIf

; Show GUI
mainGui.Show("w550 h600")
LoadList()
SwitchView()

; --- FUNCTIONS ---
; Add menu items to the context menu and the main menu
AddMenuItems(mnu) {
  global menuTxt
  ; Define menu items with their properties
  menuItems := [
    [menuTxt.add, AddToList, true],
    [menuTxt.remove, RemoveSelected],
    [""], ; Separator
    [menuTxt.run, RunSelected],
    [menuTxt.open, OpenFileLocation],
    [menuTxt.copy, CopyFullPath],
    [menuTxt.save, SaveList],
    [menuTxt.clear, ClearList],
    [""], ; Separator
    [menuTxt.select, (*) => SelectAll(true)],
    [menuTxt.deselect, (*) => SelectAll(false)]
  ]

  ; Add items to menu
  for item in menuItems {
    if (item[1] = "") ; Separator
      mnu.Add()
    else {
      mnu.Add(item[1], item[2])
      if (!item.Has(3)) ; If not initially enabled
        mnu.Disable(item[1])
    }
  }
}

GuiResize(thisGui, minMax, width, height) {
  if (minMax = -1)
    return
  LV1.GetPos(&lvX, &lvY, &lvW, &lvH)
  newW := width - 20, newH := height - lvY - 30
  LV1.Move(, , newW, newH)
  LV1.ModifyCol(1, newW * 0.3)
  LV1.ModifyCol(2, newW * 0.7 - 10)
}

; Filter Logic
FilterList(*) {
  global appList, iconLoadIndex
  filter := txtFilter.Value
  LV1.Delete()
  LV1.Opt("-Redraw")
  for path in appList {
    if (filter = "" || InStr(path, filter, false)) {
      SplitPath(path, &name)
      LV1.Add(, name, path)
    }
  }
  LV1.Opt("+Redraw")
  iconLoadIndex := 0
  SetTimer(LoadIconsStep, 10)
  UpdateUIState()
  btnClearFilter.Enabled := (filter != "")
}

; Select icon size according to view mode when loading icons into ListView
LoadIconsStep() {
  global iconLoadIndex, isLargeView
  rowCount := LV1.GetCount()
  if (++iconLoadIndex > rowCount) {
    SetTimer(LoadIconsStep, 0)
    LV1.Redraw()
    return
  }
  path := LV1.GetText(iconLoadIndex, 2)
  iconIdx := GetFileIcon(path, isLargeView)
  if iconIdx
    LV1.Modify(iconLoadIndex, "Icon" iconIdx)
}

; Add to List
AddToList(*) {
  global appListChanged
  files := FileSelect("M3", , "Select Application", "All Files (*.*)")
  if !files.Length
    return

  added := 0
  for file in files
    added += AddToAppList(file) ? 1 : 0

  if added {
    appListChanged := true
    UpdateUIState()
  }
  skipped := files.Length - added
  sbMain.SetText(Format("Added: {1}, Duplicates: {2}", added, skipped))
  FilterList()
}

AddToAppList(path) {
  global appList
  ; Duplicate check
  for p in appList
    if (p = path)
      return false

  appList.Push(path)
  return true
}

GetFileIcon(path, large := false) {
  static iconCacheSmall := Map(), iconCacheLarge := Map()
  getIconIdx(path, imagelist, flags, cache) {
    if cache.Has(path)
      return cache[path]
    sfi_size := A_PtrSize + 688
    sfi := Buffer(sfi_size)
    if DllCall("Shell32\SHGetFileInfoW", "Str", path, "UInt", 0, "Ptr", sfi, "UInt", sfi_size, "UInt", flags) {
      hIcon := NumGet(sfi, 0, "Ptr")
      iconIdx := DllCall("ImageList_ReplaceIcon", "Ptr", imagelist, "Int", -1, "Ptr", hIcon) + 1
      DllCall("DestroyIcon", "Ptr", hIcon)
      cache[path] := iconIdx
      return iconIdx
    }
    cache[path] := 0
    return 0
  }
  ; small icon
  if !iconCacheSmall.Has(path)
    getIconIdx(path, IL_ID1, 0x101, iconCacheSmall) ; SHGFI_ICON | SHGFI_SMALLICON
  ; large icon
  if !iconCacheLarge.Has(path)
    getIconIdx(path, IL_ID2, 0x100, iconCacheLarge) ; SHGFI_ICON (large)
  return large ? iconCacheLarge[path] : iconCacheSmall[path]
}

FindRowInListView(name, path) {
  Loop LV1.GetCount()
    if (LV1.GetText(A_Index, 1) = name && LV1.GetText(A_Index, 2) = path)
      return A_Index
  return 0
}

; Remove Selected
RemoveSelected(*) {
  global appList, appListChanged
  selectedRows := []
  row := 0
  while (row := LV1.GetNext(row))
    selectedRows.InsertAt(1, row) ; Reverse order

  if !selectedRows.Length {
    sbMain.SetText("No items selected")
    return
  }

  changed := false
  for row in selectedRows {
    path := LV1.GetText(row, 2)
    idx := 0
    for i, p in appList
      if (p = path) {
        idx := i
        break
      }
    if idx {
      appList.RemoveAt(idx)
      changed := true
    }
  }

  if changed {
    appListChanged := true
    UpdateUIState()
  }

  sbMain.SetText("Removed " selectedRows.Length " item(s)")
  FilterList()
}

RunSelected(*) {
  row := 0, ok := 0, fail := 0
  while row := LV1.GetNext(row) {
    path := LV1.GetText(row, 2)
    name := LV1.GetText(row, 1)
    try {
      RunWait(path)
      ok++
      sbMain.SetText("Running: " name)
    } catch as err {
      fail++
      sbMain.SetText("Failed to run: " name)
      MsgBox("Failed to run: " path "`nError: " err.Message, "Error", "Icon!")
    }
  }
  if (ok > 0 || fail > 0)
    sbMain.SetText(Format("Launched: {1}, Failed: {2}", ok, fail))
  else
    sbMain.SetText("No items selected")
}

; Copy Full Path
CopyFullPath(*) {
  selected := [] ; Array to store selected paths
  row := 0
  while (row := LV1.GetNext(row))
    selected.Push(LV1.GetText(row, 2))

  if !selected.Length {
    sbMain.SetText("No items selected")
    return
  }

  A_Clipboard := selected.Length = 1 ? selected[1] : Join("`n", selected)
  sbMain.SetText("Copied " selected.Length " path(s) to clipboard")
}

; Join Function
; Joins an array of strings with a specified separator.
Join(sep, arr) {
  if !arr.Length
    return ""
  str := arr[1]
  Loop arr.Length - 1
    str .= sep arr[A_Index + 1]
  return str
}

; Save List
SaveList(*) {
  global appList, appListChanged
  FileDelete(listFile)
  for path in appList
    FileAppend(path "`n", listFile)
  appListChanged := false
  UpdateUIState()
  sbMain.SetText("Saved " appList.Length " application(s)")
}

; Load List
LoadList(*) {
  global appList := [], appListChanged
  if !FileExist(listFile) {
    sbMain.SetText("No list file found")
    FilterList()
    return
  }
  lines := StrSplit(FileRead(listFile), "`n", "`r")
  loaded := 0
  for line in lines {
    path := Trim(line)
    if (path != "") {
      appList.Push(path)
      loaded++
    }
  }
  appListChanged := false
  UpdateUIState()
  sbMain.SetText(loaded ? "Loaded " loaded " application(s)" : "No valid apps in list")
  FilterList()
}

UpdateUIState(*) {
  global appListChanged, menuTxt
  hasSelected := LV1.GetCount("Selected") > 0
  hasItems := LV1.GetCount() > 0

  ; Update buttons
  btnRemove.Enabled := hasSelected
  btnRun.Enabled := hasSelected
  btnSave.Enabled := appListChanged && hasItems

  ; Update menu items
  menuItems := Map(
    menuTxt.save, appListChanged && hasItems,
    menuTxt.remove, hasSelected,
    menuTxt.run, hasSelected,
    menuTxt.copy, hasSelected,
    menuTxt.clear, hasItems,
    menuTxt.select, hasItems,
    menuTxt.deselect, hasItems,
    menuTxt.open, hasSelected
  )

  ; Apply states to both menus
  for mText, enabled in menuItems {
    for menu in [fileMenu, rcMenu] {
      try {
        enabled ? menu.Enable(mText) : menu.Disable(mText)
      }
    }
  }
}

; GUI Close
g1Close(*) {
  global appListChanged
  if appListChanged {
    result := MsgBox("The list has been modified. Do you want to save changes?", "Save Changes", "YesNo Icon?")
    if (result = "Yes")
      SaveList()
  }
  ExitApp()
}

; Context menu for the ListView
ShowContextMenu(LV1, Item, IsRightClick, X, Y) {
  rcMenu.Show(X, Y)
}

; Clears the ListView and updates the buttons.
ClearList(*) {
  global appList, appListChanged
  if (LV1.GetCount() = 0) {
    sbMain.SetText("List is already empty")
    return
  }

  result := MsgBox("All items in the list will be deleted. Are you sure?", "Confirmation", "OKCancel Icon? Default2 T5")
  if (result != "OK") {
    sbMain.SetText("Operation cancelled")
    return
  }

  LV1.Delete()
  appList := []
  appListChanged := true
  UpdateUIState()
  sbMain.SetText("List cleared")
}

SelectAll(select) { ; Select all or deselect all items in the ListView.
  if (select && LV1.GetCount() = 0) {
    sbMain.SetText("No items to select")
    return
  }

  LV1.Modify(0, select ? "Select" : "-Select")
  sbMain.SetText(select ? "All items selected" : "Selection cleared")
  UpdateUIState()
}

SwitchView(*) {
  static s := 0
  s := Mod(++s, 4)
  isLargeView := (s & 1) ; 1 veya 3 ise büyük ikon
  LV1.Opt(s < 2 ? "+Report" : "+Icon")
  LV1.SetImageList(isLargeView ? IL_ID2 : IL_ID1, s < 2)
}

DropHandler(gObj, gCtrlObj, FileArray, X, Y) {
  global appListChanged
  if (!FileArray.Length) {
    sbMain.SetText("No valid files dropped")
    return
  }

  added := 0
  for file in FileArray
    added += AddToAppList(file) ? 1 : 0

  if (added) {
    appListChanged := true
    UpdateUIState()
    sbMain.SetText("Added " added " file(s) by drag & drop")
  } else
    sbMain.SetText("No new files added (possibly duplicates)")

  FilterList()
}

OpenFileLocation(*) {
  row := LV1.GetNext()
  if !row {
    sbMain.SetText("No item selected")
    return
  }
  path := LV1.GetText(row, 2)
  if !FileExist(path) {
    sbMain.SetText("File not found")
    MsgBox("File not found: " path, "Error", "Icon!")
    return
  }
  Run('explorer.exe /select,"' path '"')
  sbMain.SetText("Opened file location")
}

About(*) {
MsgBox(Format("
(
{1} v{2}

©2025
Mesut Akcan
makcan@gmail.com

akcansoft.blogspot.com
mesutakcan.blogspot.com
github.com/akcansoft
youtube.com/mesutakcan
)", appName, appVer), "About", "Owner" mainGui.Hwnd)
}
