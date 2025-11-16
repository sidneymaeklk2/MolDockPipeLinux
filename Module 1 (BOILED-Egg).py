#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

; ------------------ Globals ------------------
global running := false
global mainGui, pasteEdit, btnClip, initDelayEdit, keyDelayEdit
global advanceDDL, skipBlankCB, testModeCB, btnStart, btnStop
global colModeCB, colCountDDL

; ------------------ GUI ------------------
mainGui := Gui("+AlwaysOnTop", "Excel → Web Grade Typer")
mainGui.SetFont("s10", "Segoe UI")
mainGui.AddText(, "1) Copy cells from Excel  2) Paste below  3) Click first web cell  4) Start")

pasteEdit := mainGui.AddEdit("w650 r14")

btnClip := mainGui.AddButton("w130", "From Clipboard")
btnClip.OnEvent("Click", (*) => pasteEdit.Value := A_Clipboard)

mainGui.AddText("x+10 yp+2", "Initial delay (s):")
initDelayEdit := mainGui.AddEdit("w60"), initDelayEdit.Value := 3

mainGui.AddText("x+10 yp", "Key delay (ms):")
keyDelayEdit := mainGui.AddEdit("w60"), keyDelayEdit.Value := 250

mainGui.AddText("xm y+8", "Advance key (single-column):")
advanceDDL := mainGui.AddDropDownList("w150", ["Enter","Tab","Down","Right"])
advanceDDL.Choose(1)

; NEW multi-column controls
colModeCB := mainGui.AddCheckbox("xm y+8", "Multi-column mode (column-wise)")
colModeCB.Value := 0

mainGui.AddText("x+10 yp+2", "Columns:")
colCountDDL := mainGui.AddDropDownList("w60", ["1","2","3","4","5"])
colCountDDL.Choose(1)

skipBlankCB := mainGui.AddCheckbox("xm y+8", "Skip blank cells")
skipBlankCB.Value := 1

testModeCB := mainGui.AddCheckbox("x+10", "Test mode (don't type)")
testModeCB.Value := 0

btnStart := mainGui.AddButton("xm y+8 w160 Default", "Start (F9)")
btnStop  := mainGui.AddButton("x+10 w160", "Stop (F10)")
btnStart.OnEvent("Click", Start)
btnStop.OnEvent("Click", Stop)

mainGui.Show()

Hotkey "F9", Start
Hotkey "F10", Stop
Hotkey "Esc", Stop


; ------------------ MAIN START ------------------
Start(*) {
    global running, pasteEdit, initDelayEdit, keyDelayEdit
    global advanceDDL, skipBlankCB, testModeCB
    global colModeCB, colCountDDL

    if running
        return

    txt := pasteEdit.Value
    if (txt = "") {
        MsgBox "Paste some Excel cells first.", "GradeTyper", "Icon!"
        return
    }

    running   := true
    initDelay := initDelayEdit.Value + 0
    keyDelay  := keyDelayEdit.Value + 0
    advKey    := advanceDDL.Text
    skipBlank := !!skipBlankCB.Value
    tmode     := !!testModeCB.Value
    multiCol  := !!colModeCB.Value
    colCount  := colCountDDL.Text + 0

    ToolTip "Starting in " initDelay " s..." . "`nClick the FIRST cell in the website."
    Sleep initDelay * 1000
    ToolTip

    if !multiCol {
        ; ------------------ SINGLE-COLUMN MODE ------------------
        values := ParseExcelBlockFlat(txt, skipBlank)

        for v in values {
            if !running
                break

            if (v = "" && skipBlank) {
                SendAdvance(advKey, keyDelay, tmode)
                continue
            }

            if tmode {
                ToolTip "Would type: " v
                Sleep 350
            } else {
                SendText v
                Sleep keyDelay
                SendAdvance(advKey, keyDelay)
            }
        }

        if !tmode && (advKey != "Enter") {
            Send "{Enter}"
            Sleep keyDelay
        }

    } else {
        ; ------------------ MULTI-COLUMN MODE (COLUMN-WISE) ------------------
        matrix := ParseExcelBlock2D(txt)
        rowCount := matrix.Length

        if (rowCount = 0) {
            MsgBox "No rows detected in pasted data.", "GradeTyper", "Icon!"
            running := false
            return
        }

        maxCols := 0
        for row in matrix
            if row.Length > maxCols
                maxCols := row.Length

        if (colCount > maxCols)
            colCount := maxCols

        ; -------- Column by Column Typing --------
        colLoop:
        Loop colCount {
            colIdx := A_Index

            ; --- Type DOWN through this column ---
            Loop rowCount {
                if !running
                    break colLoop

                rIdx := A_Index
                row  := matrix[rIdx]

                val := ""
                if (colIdx <= row.Length)
                    val := Trim(row[colIdx])

                if (val = "" && skipBlank) {
                    if tmode {
                        ToolTip "Skip blank (row " rIdx ", col " colIdx ")"
                        Sleep 250
                    } else {
                        Send "{Enter}"    ; skip but still move down
                        Sleep keyDelay
                    }
                    continue
                }

                if tmode {
                    ToolTip "Would type: " val " (row " rIdx ", col " colIdx ")"
                    Sleep 300
                } else {
                    SendText val
                    Sleep keyDelay
                    Send "{Enter}"      ; ENTER = commit + auto-down
                    Sleep keyDelay
                }
            }

            if !running
                break

            ; ---------------- SAFE RESET ----------------
            ; After last ENTER cursor is at Row N+1
            ; Reset to Row 1:
            Loop rowCount {          ; Up × rowCount (always safe)
                Send "{Up}"
                Sleep keyDelay
            }

            ; Move to next column
            if colIdx < colCount {
                Send "{Right}"
                Sleep keyDelay
            }
        }
    }

    running := false
    SoundBeep 1000, 120
    ToolTip "Done."
    SetTimer(() => ToolTip(), -800)
}


; ------------------ STOP ------------------
Stop(*) {
    global running
    running := false
    ToolTip "Stopped."
    SetTimer(() => ToolTip()), -600
}


; ------------------ PARSERS ------------------
ParseExcelBlockFlat(txt, skipBlank := true) {
    arr := []
    rows := StrSplit(RTrim(txt, "`r`n"), "`n")

    for , row in rows {
        row := StrReplace(row, "`r")
        cells := StrSplit(row, A_Tab)
        for , cell in cells {
            val := Trim(cell)
            if (val = "" && skipBlank)
                arr.Push("")
            else
                arr.Push(val)
        }
    }
    return arr
}

ParseExcelBlock2D(txt) {
    mat := []
    rows := StrSplit(RTrim(txt, "`r`n"), "`n")

    for , rowText in rows {
        rowText := StrReplace(rowText, "`r")
        cells := StrSplit(rowText, A_Tab)
        for i, cell in cells
            cells[i] := Trim(cell)
        mat.Push(cells)
    }
    return mat
}


; ------------------ ADVANCE (Single-Column Mode) ------------------
SendAdvance(key, dly, tmode := false) {
    if tmode
        return
    switch key {
        case "Enter": Send "{Enter}"
        case "Tab":   Send "{Tab}"
        case "Down":  Send "{Down}"
        case "Right": Send "{Right}"
        default:      Send "{Enter}"
    }
    Sleep dly
}
