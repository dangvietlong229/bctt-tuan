on run argv
    set imagePath to item 1 of argv
    set pptPath to item 2 of argv
    set slideNumber to (item 3 of argv) as integer
    set targetShapeName to item 4 of argv

    set imageData to read (POSIX file imagePath) as «class PNGf»
    set the clipboard to imageData
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
        set lock aspect ratio of newShape to false
        set left position of newShape to 610.8095275590551
        set top of newShape to 60.136062992126
        set width of newShape to 334.5295275590551
        set height of newShape to 229.4303149606299
        set name of newShape to targetShapeName
        save pres
        close pres saving yes
    end tell
end run
