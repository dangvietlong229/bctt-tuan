on run argv
    set workbookPath to item 1 of argv
    set worksheetName to item 2 of argv
    set chartNumber to (item 3 of argv) as integer
    set pptPath to item 4 of argv
    set slideNumber to (item 5 of argv) as integer
    set targetShapeName to item 6 of argv

    tell application "Microsoft Excel"
        activate
        set wb to open workbook workbook file name workbookPath update links never read only true
        set ws to worksheet worksheetName of wb
        copy picture chart object chartNumber of ws appearance screen format bitmap
        close wb saving no
    end tell

    delay 1

    tell application "Microsoft PowerPoint"
        open pptPath
        delay 5
        activate
        set pres to active presentation
        set targetSlide to slide slideNumber of pres
        delete shape targetShapeName of targetSlide

        set activeWindow to active window
        set slide of view of activeWindow to targetSlide
        paste object (view of activeWindow)
        set newShape to last shape of targetSlide

        -- Fixed frame of the Vingroup chart in the approved report template.
        set lock aspect ratio of newShape to false
        set left position of newShape to 610.8094488188976
        set top of newShape to 289.5663779527559
        set width of newShape to 334.5296062992126
        set height of newShape to 226.68590551181103
        set name of newShape to targetShapeName
        save pres
        close pres saving yes
    end tell
end run
