-- Usage: osascript excel_recalc.applescript /abs/path/file.xlsx "Sheet!A1" ...
on run argv
    set filePath to item 1 of argv
    set n to count of argv
    set outLines to {}
    tell application "Microsoft Excel"
        set wb to open workbook workbook file name (POSIX file filePath)
        calculate
    end tell
    repeat with i from 2 to n
        set theRef to item i of argv
        set AppleScript's text item delimiters to "!"
        set sheetName to text item 1 of theRef
        set cellRef to text item 2 of theRef
        set AppleScript's text item delimiters to ""
        tell application "Microsoft Excel"
            set v to value of range cellRef of worksheet sheetName of wb
        end tell
        set end of outLines to (theRef & tab & (v as text))
    end repeat
    tell application "Microsoft Excel" to close wb saving no
    set AppleScript's text item delimiters to linefeed
    set outText to outLines as text
    set AppleScript's text item delimiters to ""
    return outText
end run
