/*
HTML Stats Display - Standalone script for live honey/backpack display
Uses the exact same detection methods as background.ahk and StatMonitor.ahk
*/

#SingleInstance Force
#NoTrayIcon
#MaxThreads 255

#Include "%A_ScriptDir%\..\lib"
#Include "Gdip_All.ahk"
#Include "Gdip_ImageSearch.ahk"
#Include "Roblox.ahk"
#Include "nowUnix.ahk"

OnError (e, mode) => (mode = "Return") ? -1 : 0
SetWorkingDir A_ScriptDir "\.."
CoordMode "Pixel", "Screen"
DetectHiddenWindows 1

pToken := Gdip_Startup()
bitmaps := Map(), bitmaps.CaseSense := 0
#Include "%A_ScriptDir%\..\nm_image_assets\offset\bitmaps.ahk"

; Global variables used by Roblox.ahk functions
global windowX := 0, windowY := 0, windowWidth := 0, windowHeight := 0

; OCR initialization - check if OCR is available (from StatMonitor.ahk)
ocr_enabled := 1
ocr_language := ""
for k,v in Map("Windows.Globalization.Language","{9B0252AC-0C27-44F8-B792-9793FB66C63E}", "Windows.Graphics.Imaging.BitmapDecoder","{438CCB26-BCEF-4E95-BAD6-23A822E58D01}", "Windows.Media.Ocr.OcrEngine","{5BFFA85A-3384-3540-9940-699120D428A8}")
{
    hString := Buffer(8), DllCall("Combase.dll\WindowsCreateString", "WStr", k, "UInt", StrLen(k), "Ptr", hString)
    GUID := Buffer(16), DllCall("ole32\CLSIDFromString", "WStr", v, "Ptr", GUID)
    result := DllCall("Combase.dll\RoGetActivationFactory", "Ptr", hString, "Ptr", GUID, "Ptr*", &pFactory:=0)
    DllCall("Combase.dll\WindowsDeleteString", "Ptr", hString)
    if (result != 0)
    {
        ocr_enabled := 0
        break
    }
}
if (ocr_enabled = 1)
{
    try {
        list := ocr("ShowAvailableLanguages")
        for lang in ["en", "de", "es", "fr", "pt", "it", "nl", "ja", "ko", "zh"]
        {
            Loop Parse list, "`n"
            {
                if (InStr(A_LoopField, lang) = 1)
                {
                    ocr_language := A_LoopField
                    break 2
                }
            }
        }
        if (ocr_language = "")
            ocr_language := SubStr(list, 1, InStr(list, "`n")-1)
        ; Final fallback - use FirstFromAvailableLanguages
        if (ocr_language = "")
            ocr_language := "FirstFromAvailableLanguages"
    } catch {
        ocr_enabled := 0
    }
}

; Settings
backpackUpdateInterval := 1000  ; 1 second
honeyUpdateInterval := 5000     ; 5 seconds

; State
currentHoney := 0
currentBackpack := 0
statsFile := A_WorkingDir "\feed\stat_monitor_stats.js"
backpackSamples := []  ; Rolling average filter for smoother backpack readings
nectarValues := Map(), nectarValues.CaseSense := 0

nectarBitmaps := Map(), nectarBitmaps.CaseSense := 0
nectarBitmaps["comforting"] := Gdip_CreateBitmap(3,1)
pGraphics := Gdip_GraphicsFromImage(nectarBitmaps["comforting"]), Gdip_GraphicsClear(pGraphics, 0xff7e9eb3), Gdip_DeleteGraphics(pGraphics)
nectarBitmaps["motivating"] := Gdip_CreateBitmap(3,1)
pGraphics := Gdip_GraphicsFromImage(nectarBitmaps["motivating"]), Gdip_GraphicsClear(pGraphics, 0xff937db3), Gdip_DeleteGraphics(pGraphics)
nectarBitmaps["satisfying"] := Gdip_CreateBitmap(3,1)
pGraphics := Gdip_GraphicsFromImage(nectarBitmaps["satisfying"]), Gdip_GraphicsClear(pGraphics, 0xffb398a7), Gdip_DeleteGraphics(pGraphics)
nectarBitmaps["refreshing"] := Gdip_CreateBitmap(3,1)
pGraphics := Gdip_GraphicsFromImage(nectarBitmaps["refreshing"]), Gdip_GraphicsClear(pGraphics, 0xff78b375), Gdip_DeleteGraphics(pGraphics)
nectarBitmaps["invigorating"] := Gdip_CreateBitmap(3,1)
pGraphics := Gdip_GraphicsFromImage(nectarBitmaps["invigorating"]), Gdip_GraphicsClear(pGraphics, 0xffb35951), Gdip_DeleteGraphics(pGraphics)

for v in ["comforting","motivating","satisfying","refreshing","invigorating"]
    nectarValues[v] := Map()

buffValues := Map(), buffValues.CaseSense := 0
for v in ["redboost","whiteboost","blueboost","haste","focus","bombcombo","balloonaura","inspire","reindeerfetch","honeymark","pollenmark","popstar","melody","bear","babylove","jbshare","guiding"]
    buffValues[v] := Map()

buffCharacters := Map()
buffCharacters[0] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAAKCAAAAAC2kKDSAAAAAnRSTlMAAHaTzTgAAAA9SURBVHgBATIAzf8BAADzAAAA8wAAAAAAAAAA8wAAAAIAAAAAAgAAAAACAAAAAAAAAAAAAADzAAABAADzAIAxBMg7bpCUAAAAAElFTkSuQmCC")
buffCharacters[1] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAIAAAAMCAAAAABt1zOIAAAAAnRSTlMAAHaTzTgAAAACYktHRAD/h4/MvwAAABZJREFUeAFjYPjM+JmBgeEzEwMDLgQAWo0C7U3u8hAAAAAASUVORK5CYII=")
buffCharacters[2] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAALCAAAAAB9zHN3AAAAAnRSTlMAAHaTzTgAAABCSURBVHgBATcAyP8BAPMAAADzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPMAAADzAAAA8wAAAPMAAAAB8wAAAAIAAAAAtc8GqohTl5oAAAAASUVORK5CYII=")
buffCharacters[3] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAAKCAAAAAC2kKDSAAAAAnRSTlMAAHaTzTgAAAA9SURBVHgBATIAzf8BAPMAAAAAAAAAAAAAAAAAAAAAAAAAAADzAAAAAAAAAAAAAAAAAAAAAPMAAAABAPMAAFILA8/B68+8AAAAAElFTkSuQmCC")
buffCharacters[4] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAAGCAAAAADBUmCpAAAAAnRSTlMAAHaTzTgAAAApSURBVHgBAR4A4f8AAAAA8wAAAAAAAAAA8wAAAPMAAALzAAAAAfMAAABBtgTDARckPAAAAABJRU5ErkJggg==")
buffCharacters[5] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAALCAAAAAB9zHN3AAAAAnRSTlMAAHaTzTgAAABCSURBVHgBATcAyP8B8wAAAAIAAAAAAPMAAAACAAAAAAHzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHzAAAAgmID1KbRt+YAAAAASUVORK5CYII=")
buffCharacters[6] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAAJCAAAAAAwBNJ8AAAAAnRSTlMAAHaTzTgAAAA4SURBVHgBAS0A0v8AAAAA8wAAAPMAAADzAAACAAAAAAEA8wAAAPPzAAAA8wAAAAAA8wAAAQAA8wC5oAiQ09KYngAAAABJRU5ErkJggg==")
buffCharacters[7] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAAMCAAAAABgyUPPAAAAAnRSTlMAAHaTzTgAAABHSURBVHgBATwAw/8B8wAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA8wIAAAAAAgAAAABDdgHu70cIeQAAAABJRU5ErkJggg==")
buffCharacters[8] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAAKCAAAAAC2kKDSAAAAAnRSTlMAAHaTzTgAAAA9SURBVHgBATIAzf8BAADzAAAA8wAAAgAAAAABAPMAAAEAAPMAAADzAAAAAAAAAADzAAAAAADzAAABAADzALv5B59oKTe0AAAAAElFTkSuQmCC")
buffCharacters[9] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAQAAAAKCAAAAAC2kKDSAAAAAnRSTlMAAHaTzTgAAAA9SURBVHgBATIAzf8BAADzAAAA8wAAAPMAAAAAAPMAAAEAAPMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA87TcBbXcfy3eAAAAAElFTkSuQmCC")

buffBitmaps := Map(), buffBitmaps.CaseSense := 0
buffBitmaps["pBMHaste"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMHaste"]), Gdip_GraphicsClear(pGraphics, 0xfff0f0f0), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMBoost"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMBoost"]), Gdip_GraphicsClear(pGraphics, 0xff90ff8e), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMFocus"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMFocus"]), Gdip_GraphicsClear(pGraphics, 0xff22ff06), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMBombCombo"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMBombCombo"]), Gdip_GraphicsClear(pGraphics, 0xff272727), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMBalloonAura"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMBalloonAura"]), Gdip_GraphicsClear(pGraphics, 0xfffafd38), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMJBShare"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMJBShare"]), Gdip_GraphicsClear(pGraphics, 0xfff9ccff), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMBabyLove"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMBabyLove"]), Gdip_GraphicsClear(pGraphics, 0xff8de4f3), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMInspire"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMInspire"]), Gdip_GraphicsClear(pGraphics, 0xfff4ef14), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMReindeerFetch"] := Gdip_CreateBitmap(5,1)
pGraphics := Gdip_GraphicsFromImage(buffBitmaps["pBMReindeerFetch"]), Gdip_GraphicsClear(pGraphics, 0xffcc2c2c), Gdip_DeleteGraphics(pGraphics)
buffBitmaps["pBMHoneyMark"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAkAAAAEBAMAAACuIQj9AAAAMFBMVEUcJhYXKxsXKxwZLx0xNRc0YDI1YTM4ZzZ3axp8cBs9cTueih2vlx7WtiDYtyHsxyJxibSYAAAAI0lEQVR4AQEYAOf/AKqqcwvwAKqqUZ7wAKqqUY3wAKqqYkzwjf0MCuMjsQoAAAAASUVORK5CYII=")
buffBitmaps["pBMPollenMark"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAoAAAAFCAMAAABLuo1aAAAAQlBMVEUPHBYRHRcUIhgUJRoaMB0pMiQcNSAuNyYfOSI5QCwnSChdYD81YzQ2ZDQ5ajg8bjo8cDo9cTuEglSknWS9tHLk1YZKij78AAAAQklEQVR4AQE3AMj/ABERERERDggCCxQAERERERERDAYABQAREREREREPCgMBABEREREREAoDBxIAERERERENBAkTFUoXAq+Dil5HAAAAAElFTkSuQmCC")
buffBitmaps["pBMGuiding"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAwAAAACCAMAAABboc2lAAAAOVBMVEWPf02QgE6RgE+SgU/SuHDTunHUunHhxnjhx3niyHrjyXvky3zn0oPp1ITq1obq2Yju4o7u44/v5JDO0m0EAAAAJUlEQVR4AQEaAOX/ABIQDAgEAwEECQ0REgASDgoHBQACBgcLDxIMQwDt+rZJwwAAAABJRU5ErkJggg==")
buffBitmaps["pBMBearBrown"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAAwAAAABBAMAAAAYxVIKAAAAD1BMVEUwLi1STEihfVWzpZbQvKTt7OCuAAAAEklEQVR4AQEHAPj/ACJDEAE0IgLvAM1oKEJeAAAAAElFTkSuQmCC")
buffBitmaps["pBMBearBlack"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAA4AAAABBAMAAAAcMII3AAAAFVBMVEUwLi1TTD9lbHNmbXN5enW5oXHQuYJDhTsuAAAAE0lEQVR4AQEIAPf/ACNGUQAVZDIFbwFmjB55HwAAAABJRU5ErkJggg==")
buffBitmaps["pBMBearPanda"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAABAAAAABBAMAAAAlVzNsAAAAGFBMVEUwLi1VU1G9u7m/vLXAvbbPzcXg3dfq6OXkYMPeAAAAFElEQVR4AQEJAPb/AENWchABJ2U0CO4B3TmcTKkAAAAASUVORK5CYII=")
buffBitmaps["pBMBearPolar"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAA4AAAABBAMAAAAcMII3AAAAElBMVEUwLi1JSUqOlZy0vMbY2dnc3NtuftTJAAAAE0lEQVR4AQEIAPf/AFVDIQASNFUFhQFVdZ1AegAAAABJRU5ErkJggg==")
buffBitmaps["pBMBearGummy"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAA4AAAABBAMAAAAcMII3AAAAFVBMVEWYprGDrKWisd+hst+ctNtFyJ4xz5uqDngAAAAAE0lEQVR4AQEIAPf/ACNAFWZRBDIFqwFmOuySwwAAAABJRU5ErkJggg==")
buffBitmaps["pBMBearScience"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAAA4AAAABBAMAAAAcMII3AAAAFVBMVEUwLi1TTD+zjUy0jky8l1W5oXHevny+g95vAAAAE0lEQVR4AQEIAPf/ACNGUQAVZDIFbwFmjB55HwAAAABJRU5ErkJggg==")
buffBitmaps["pBMBearMother"] := Gdip_BitmapFromBase64("iVBORw0KGgoAAAANSUhEUgAAABAAAAABBAMAAAAlVzNsAAAAJFBMVEVBNRlDNxtTRid8b0avoG69r22+sG7Qw4PRw4Te0Jbk153m2Z5VNHxxAAAAFElEQVR4AQEJAPb/AFVouTECSnZVDPsCv+2QpmwAAAAASUVORK5CYII=")

; Start timers
SetTimer(UpdateBackpack, backpackUpdateInterval)
SetTimer(UpdateHoney, honeyUpdateInterval)
SetTimer(UpdateNectars, 1000)
SetTimer(UpdateBuffs, 1000)
SetTimer(CheckMacroRunning, 5000)

; Initial update
UpdateBackpack()
UpdateHoney()
UpdateNectars()
UpdateBuffs()

Persistent

;===========================================
; CHECK IF MAIN MACRO IS STILL RUNNING
;===========================================
CheckMacroRunning() {
    DetectHiddenWindows 1
    if !WinExist("natro_macro ahk_class AutoHotkey")
        ExitApp
    DetectHiddenWindows 0
}

;===========================================
; BACKPACK DETECTION - from background.ahk (5% increments)
;===========================================
UpdateBackpack() {
    global currentBackpack, currentHoney, statsFile, backpackSamples
    global windowX, windowY, windowWidth, windowHeight

    hwnd := GetRobloxHWND()
    if !hwnd
        return

    GetRobloxClientPos(hwnd)
    offsetY := GetYOffset(hwnd)

    if !(windowHeight >= 500)
        return

    backpackColor := PixelGetColor(windowX + windowWidth//2 + 59 + 3, windowY + offsetY + 6)
    BackpackPercent := 0

    if ((backpackColor & 0xFF0000) <= 0x690000) { ; <= 50%
        if ((backpackColor & 0xFF0000) <= 0x4B0000) { ; <= 25%
            if ((backpackColor & 0xFF0000) <= 0x420000) { ; <= 10%
                if ((backpackColor & 0xFF0000 <= 0x410000) && (backpackColor & 0x00FFFF <= 0x00FF80) && (backpackColor & 0x00FFFF > 0x00FF86))
                    BackpackPercent := 0
                else if ((backpackColor & 0xFF0000 > 0x410000) && (backpackColor & 0x00FFFF <= 0x00FF80) && (backpackColor & 0x00FFFF > 0x00FC85))
                    BackpackPercent := 5
                else
                    BackpackPercent := 0
            } else { ; > 10%
                if ((backpackColor & 0xFF0000) <= 0x470000) { ; <= 20%
                    if ((backpackColor & 0xFF0000 <= 0x440000) && (backpackColor & 0x00FFFF <= 0x00FE85) && (backpackColor & 0x00FFFF > 0x00F984))
                        BackpackPercent := 10
                    else if ((backpackColor & 0xFF0000 > 0x440000) && (backpackColor & 0x00FFFF <= 0x00FB84) && (backpackColor & 0x00FFFF > 0x00F582))
                        BackpackPercent := 15
                    else
                        BackpackPercent := 0
                } else if ((backpackColor & 0xFF0000 > 0x470000) && (backpackColor & 0x00FFFF <= 0x00F782) && (backpackColor & 0x00FFFF > 0x00F080))
                    BackpackPercent := 20
                else
                    BackpackPercent := 0
            }
        } else { ; > 25%
            if ((backpackColor & 0xFF0000) <= 0x5B0000) { ; <= 40%
                if ((backpackColor & 0xFF0000 <= 0x4F0000) && (backpackColor & 0x00FFFF <= 0x00F280) && (backpackColor & 0x00FFFF > 0x00EA7D))
                    BackpackPercent := 25
                else { ; > 30%
                    if ((backpackColor & 0xFF0000 <= 0x550000) && (backpackColor & 0x00FFFF <= 0x00EC7D) && (backpackColor & 0x00FFFF > 0x00E37A))
                        BackpackPercent := 30
                    else if ((backpackColor & 0xFF0000 > 0x550000) && (backpackColor & 0x00FFFF <= 0x00E57A) && (backpackColor & 0x00FFFF > 0x00DA76))
                        BackpackPercent := 35
                    else
                        BackpackPercent := 0
                }
            } else { ; > 40%
                if ((backpackColor & 0xFF0000 <= 0x620000) && (backpackColor & 0x00FFFF <= 0x00DC76) && (backpackColor & 0x00FFFF > 0x00D072))
                    BackpackPercent := 40
                else if ((backpackColor & 0xFF0000 > 0x620000) && (backpackColor & 0x00FFFF <= 0x00D272) && (backpackColor & 0x00FFFF > 0x00C66D))
                    BackpackPercent := 45
                else
                    BackpackPercent := 0
            }
        }
    } else { ; > 50%
        if ((backpackColor & 0xFF0000) <= 0x9C0000) { ; <= 75%
            if ((backpackColor & 0xFF0000) <= 0x850000) { ; <= 65%
                if ((backpackColor & 0xFF0000) <= 0x7B0000) { ; <= 60%
                    if ((backpackColor & 0xFF0000 <= 0x720000) && (backpackColor & 0x00FFFF <= 0x00C86D) && (backpackColor & 0x00FFFF > 0x00BA68))
                        BackpackPercent := 50
                    else if ((backpackColor & 0xFF0000 > 0x720000) && (backpackColor & 0x00FFFF <= 0x00BC68) && (backpackColor & 0x00FFFF > 0x00AD62))
                        BackpackPercent := 55
                    else
                        BackpackPercent := 0
                } else if ((backpackColor & 0xFF0000 > 0x7B0000) && (backpackColor & 0x00FFFF <= 0x00AF62) && (backpackColor & 0x00FFFF > 0x009E5C))
                    BackpackPercent := 60
                else
                    BackpackPercent := 0
            } else { ; > 65%
                if ((backpackColor & 0xFF0000 <= 0x900000) && (backpackColor & 0x00FFFF <= 0x00A05C) && (backpackColor & 0x00FFFF > 0x008F55))
                    BackpackPercent := 65
                else if ((backpackColor & 0xFF0000 > 0x900000) && (backpackColor & 0x00FFFF <= 0x009155) && (backpackColor & 0x00FFFF > 0x007E4E))
                    BackpackPercent := 70
                else
                    BackpackPercent := 0
            }
        } else { ; > 75%
            if ((backpackColor & 0xFF0000) <= 0xC40000) { ; <= 90%
                if ((backpackColor & 0xFF0000 <= 0xA90000) && (backpackColor & 0x00FFFF <= 0x00804E) && (backpackColor & 0x00FFFF > 0x006C46))
                    BackpackPercent := 75
                else { ; > 80%
                    if ((backpackColor & 0xFF0000 <= 0xB60000) && (backpackColor & 0x00FFFF <= 0x006E46) && (backpackColor & 0x00FFFF > 0x005A3F))
                        BackpackPercent := 80
                    else if ((backpackColor & 0xFF0000 > 0xB60000) && (backpackColor & 0x00FFFF <= 0x005D3F) && (backpackColor & 0x00FFFF > 0x004637))
                        BackpackPercent := 85
                    else
                        BackpackPercent := 0
                }
            } else { ; > 90%
                if ((backpackColor & 0xFF0000 <= 0xD30000) && (backpackColor & 0x00FFFF <= 0x004A37) && (backpackColor & 0x00FFFF > 0x00322E))
                    BackpackPercent := 90
                else { ; > 95%
                    if ((backpackColor = 0xF70017) || ((backpackColor & 0xFF0000 >= 0xE00000) && (backpackColor & 0x00FFFF <= 0x002427) && (backpackColor & 0x00FFFF > 0x001000)))
                        BackpackPercent := 100
                    else if ((backpackColor & 0x00FFFF <= 0x00342E))
                        BackpackPercent := 95
                    else
                        BackpackPercent := 0
                }
            }
        }
    }

    ; Apply rolling average filter (6 samples) for smoother readings
    global backpackSamples
    backpackSamples.InsertAt(1, BackpackPercent)
    if (backpackSamples.Length > 6)
        backpackSamples.Pop()

    ; Calculate average
    sum := 0
    for val in backpackSamples
        sum += val
    currentBackpack := Round(sum / backpackSamples.Length)

    WriteStats()
}

;===========================================
; HONEY DETECTION - exact copy from StatMonitor.ahk
;===========================================
UpdateHoney() {
    global currentHoney, currentBackpack, statsFile
    global windowX, windowY, windowWidth, windowHeight
    global ocr_enabled, ocr_language

    ; Skip if OCR is not available
    if (ocr_enabled != 1)
        return

    try {
        ; Check roblox window exists
        hwnd := GetRobloxHWND()
        GetRobloxClientPos(hwnd), offsetY := GetYOffset(hwnd)
        if !(windowHeight >= 500)
            return

        ; Initialise array to store detected values and get bitmap and effect ready
        detected := Map()
        pBM := Gdip_BitmapFromScreen(windowX + windowWidth//2 - 241 "|" windowY + offsetY "|140|36")
        pEffect := Gdip_CreateEffect(5, -80, 30)

        ; Detect honey, enlarge image if necessary
        Loop 25 {
            i := A_Index
            Loop 2 {
                pBMNew := Gdip_ResizeBitmap(pBM, ((A_Index = 1) ? (250 + i * 20) : (750 - i * 20)), 36 + i * 4, 2)
                Gdip_BitmapApplyEffect(pBMNew, pEffect)
                hBM := Gdip_CreateHBITMAPFromBitmap(pBMNew)
                Gdip_DisposeImage(pBMNew)
                pIRandomAccessStream := HBitmapToRandomAccessStream(hBM)
                DllCall("DeleteObject", "Ptr", hBM)
                try detected[v := ((StrLen((n := RegExReplace(StrReplace(StrReplace(StrReplace(StrReplace(ocr(pIRandomAccessStream, ocr_language), "o", "0"), "i", "1"), "l", "1"), "a", "4"), "\D"))) > 0) ? n : 0)] := detected.Has(v) ? [detected[v][1]+1, detected[v][2] " " i . A_Index] : [1, i . A_Index]
            }
        }

        ; Clean up
        Gdip_DisposeImage(pBM), Gdip_DisposeEffect(pEffect)
        DllCall("psapi.dll\EmptyWorkingSet", "UInt", -1)

        ; Evaluate current honey
        current_honey := 0
        for k, v in detected
            if ((v[1] > 2) && (k > current_honey))
                current_honey := k

        if current_honey
            currentHoney := current_honey

        WriteStats()
    } catch {
        ; OCR failed, skip this update
    }
}

;===========================================
; NECTAR DETECTION - based on StatMonitor.ahk
;===========================================
UpdateNectars() {
    global nectarValues, nectarBitmaps
    global windowX, windowY, windowWidth, windowHeight

    time_value := (60*A_Min+A_Sec)//6
    i := (time_value = 0) ? 600 : time_value

    hwnd := GetRobloxHWND()
    if !hwnd
        return

    GetRobloxClientPos(hwnd)
    offsetY := GetYOffset(hwnd)
    if !(windowHeight >= 500)
        return

    pBMArea := Gdip_BitmapFromScreen(windowX "|" windowY+offsetY+30 "|" windowWidth "|50")

    for v in ["comforting","motivating","satisfying","refreshing","invigorating"]
    {
        if (Gdip_ImageSearch(pBMArea, nectarBitmaps[v], &list, , 30, , , , , 6) != 1)
        {
            nectarValues[v][i] := 0
            continue
        }

        x := SubStr(list, 1, InStr(list, ",")-1)

        if (Gdip_ImageSearch(pBMArea, nectarBitmaps[v], &list, x, 6, x+38, 44) != 1)
        {
            nectarValues[v][i] := 0
            continue
        }

        y := SubStr(list, InStr(list, ",")+1)
        nectarValues[v][i] := String(Round(Min((44 - y) / 38, 1) * 100, 0))
    }

    Gdip_DisposeImage(pBMArea)
    WriteStats()
}

;===========================================
; BUFF DETECTION - based on StatMonitor.ahk
;===========================================
UpdateBuffs() {
    global buffValues

    time_value := (60*A_Min+A_Sec)//6
    i := (time_value = 0) ? 600 : time_value

    DetectBuffs(i)
    UpdatePopstar(i)
    WriteStats()
}

UpdatePopstar(i) {
    global buffValues
    global windowX, windowY, windowWidth, windowHeight

    try
        result := ImageSearch(&FoundX, &FoundY, windowX + windowWidth//2 - 275, windowY + 3*windowHeight//4, windowX + windowWidth//2 + 275, windowY + windowHeight, "*30 nm_image_assets\\popstar_counter.png")
    catch
        return

    buffValues["popstar"][i] := (result = 1) ? 1 : 0
}

DetectBuffs(i) {
    global buffValues, buffCharacters, buffBitmaps
    global windowX, windowY, windowWidth, windowHeight

    hwnd := GetRobloxHWND()
    if !hwnd
        return

    GetRobloxClientPos(hwnd)
    offsetY := GetYOffset(hwnd)
    if !(windowHeight >= 500)
    {
        for k, v in buffValues
            v[i] := 0
        return
    }

    pBMArea := Gdip_BitmapFromScreen(windowX "|" windowY+offsetY+30 "|" windowWidth "|50")

    ; basic on/off
    for v in ["jbshare","babylove","guiding"]
        buffValues[v][i] := (Gdip_ImageSearch(pBMArea, buffBitmaps["pBM" v], , , 30, , , (v = "guiding") ? 10 : 0, , 7) = 1)

    ; bear morphs
    buffValues["bear"][i] := 0
    for v in ["Brown","Black","Panda","Polar","Gummy","Science","Mother"]
    {
        if (Gdip_ImageSearch(pBMArea, buffBitmaps["pBMBear" v], , , 43, , 45, 8, , 2) = 1)
        {
            buffValues["bear"][i] := 1
            break
        }
    }

    ; basic x1-x10
    for v in ["focus","bombcombo","balloonaura","honeymark","pollenmark","reindeerfetch"]
    {
        if (Gdip_ImageSearch(pBMArea, buffBitmaps["pBM" v], &list, , InStr(v, "mark") ? 20 : 30, , 50, InStr(v, "mark") ? 6 : 0, , 7) != 1)
        {
            buffValues[v][i] := 0
            continue
        }

        x := SubStr(list, 1, InStr(list, ",")-1)

        Loop 9
        {
            if (Gdip_ImageSearch(pBMArea, buffCharacters[10-A_Index], , x-20, 15, x, 50) = 1)
            {
                buffValues[v][i] := (A_Index = 9) ? 10 : 10 - A_Index
                break
            }
            if (A_Index = 9)
                buffValues[v][i] := 1
        }
    }

    ; melody / haste
    x := 0
    Loop 3
    {
        if (Gdip_ImageSearch(pBMArea, buffBitmaps["pBMHaste"], &list, x, 30, , , , , 6) != 1)
            break

        x := SubStr(list, 1, InStr(list, ",")-1)

        if ((s := Gdip_ImageSearch(pBMArea, buffBitmaps["pBMMelody"], , x+2, 15, x+34, 40, 12)) != 1)
        {
            if !buffValues["haste"].Has(i)
            {
                Loop 9
                {
                    if (Gdip_ImageSearch(pBMArea, buffCharacters[10-A_Index], , x+6, 15, x+44, 50) = 1)
                    {
                        buffValues["haste"][i] := (A_Index = 9) ? 10 : 10 - A_Index
                        break
                    }
                    if (A_Index = 9)
                        buffValues["haste"][i] := 1
                }
            }
        }
        else if (s = 1)
            buffValues["melody"][i] := 1

        x += 44
    }
    for v in ["melody","haste"]
        if !buffValues[v].Has(i)
            buffValues[v][i] := 0

    ; colour boost x1-x10
    x := windowWidth
    Loop 3
    {
        if (Gdip_ImageSearch(pBMArea, buffBitmaps["pBMBoost"], &list, , 30, x, , , , 7) != 1)
            break

        x := SubStr(list, 1, InStr(list, ",")-1)
        y := SubStr(list, InStr(list, ",")+1)

        pBMPxRed := Gdip_CreateBitmap(1,2), pBMPxBlue := Gdip_CreateBitmap(1,2)
        pGRed := Gdip_GraphicsFromImage(pBMPxRed), pGBlue := Gdip_GraphicsFromImage(pBMPxBlue)
        Gdip_GraphicsClear(pGRed, 0xffe46156), Gdip_GraphicsClear(pGBlue, 0xff56a4e4)
        Gdip_DeleteGraphics(pGRed), Gdip_DeleteGraphics(pGBlue)
        v := (Gdip_ImageSearch(pBMArea, pBMPxRed, , x-30, 15, x-4, 34, 20, , , 2) = 2) ? "redboost"
            : (Gdip_ImageSearch(pBMArea, pBMPxBlue, , x-30, 15, x-4, 34, 20, , , 2) = 2) ? "blueboost"
            : "whiteboost"
        Gdip_DisposeImage(pBMPxRed), Gdip_DisposeImage(pBMPxBlue)

        Loop 9
        {
            if Gdip_ImageSearch(pBMArea, buffCharacters[10-A_Index], , x-20, 15, x, 50)
            {
                buffValues[v][i] := (A_Index = 9) ? 10 : 10 - A_Index
                break
            }
            if (A_Index = 9)
                buffValues[v][i] := 1
        }

        x -= 2*y-53
    }
    for v in ["redboost","blueboost","whiteboost"]
        if !buffValues[v].Has(i)
            buffValues[v][i] := 0

    ; inspire (2 digit)
    if (Gdip_ImageSearch(pBMArea, buffBitmaps["pBMInspire"], &list, , 20, , , 0, , 7) != 1)
    {
        buffValues["inspire"][i] := 0
    }
    else
    {
        x := SubStr(list, 1, InStr(list, ",")-1)
        (digits := Map()).Default := ""

        Loop 10
        {
            n := 10-A_Index
            if ((n = 1) || (n = 3))
                continue
            Gdip_ImageSearch(pBMArea, buffCharacters[n], &list:="", x-20, 15, x, 50, 1, , 5, 5, , "`n")
            Loop Parse list, "`n"
                if (A_Index & 1)
                    digits[Integer(A_LoopField)] := n
        }

        for m,n in [1,3]
        {
            Gdip_ImageSearch(pBMArea, buffCharacters[n], &list:="", x-20, 15, x, 50, 1, , 5, 5, , "`n")
            Loop Parse list, "`n"
            {
                if (A_Index & 1)
                {
                    if (((n = 1) && (digits[A_LoopField - 5] = 4)) || ((n = 3) && (digits[A_LoopField - 1] = 8)))
                        continue
                    digits[Integer(A_LoopField)] := n
                }
            }
        }

        num := ""
        for x,y in digits
            num .= y

        buffValues["inspire"][i] := num ? Min(num, 50) : 1
    }

    Gdip_DisposeImage(pBMArea)
}

;===========================================
; WRITE TO JS FILE
;===========================================
WriteStats() {
    global currentHoney, currentBackpack, statsFile
    global nectarValues
    global buffValues

    try {
        try FileDelete(statsFile)
        f := FileOpen(statsFile, "w", "UTF-8")
        planterItems := ""
        Loop 3 {
            idx := A_Index
            planterName := IniRead("settings\nm_config.ini", "Planters", "PlanterName" idx, "None")
            planterField := IniRead("settings\nm_config.ini", "Planters", "PlanterField" idx, "None")
            planterNectar := IniRead("settings\nm_config.ini", "Planters", "PlanterNectar" idx, "None")
            planterHarvestTime := IniRead("settings\nm_config.ini", "Planters", "PlanterHarvestTime" idx, 0)
            if (planterHarvestTime = "")
                planterHarvestTime := 0
            item := '{"name": "' . JsonEscape(planterName) . '", "field": "' . JsonEscape(planterField) . '", "nectar": "' . JsonEscape(planterNectar) . '", "harvestTime": ' . (planterHarvestTime + 0) . '}'
            planterItems .= (A_Index > 1 ? ", " : "") . item
        }
        nectarItems := ""
        unix_now := nowUnix()
        Loop 5
        {
            j := ["comforting","motivating","satisfying","refreshing","invigorating"][A_Index]
            nectar_value := 0
            Loop 601
            {
                if (nectarValues[j].Has(601-A_Index) && (nectarValues[j][601-A_Index] > 0))
                {
                    nectar_value := nectarValues[j][601-A_Index]
                    break
                }
            }

            projected_value := 0
            Loop 3
            {
                planterName := IniRead("settings\nm_config.ini", "Planters", "PlanterName" A_Index, "None")
                planterNectar := IniRead("settings\nm_config.ini", "Planters", "PlanterNectar" A_Index, "None")
                planterEstPercent := IniRead("settings\nm_config.ini", "Planters", "PlanterEstPercent" A_Index, 0)
                planterHarvestTime := IniRead("settings\nm_config.ini", "Planters", "PlanterHarvestTime" A_Index, 0)
                if (planterName = "None")
                    continue
                if (StrLower(planterNectar) = j)
                    projected_value += (planterEstPercent - Max(planterHarvestTime - unix_now, 0)/864)
            }
            projected_value := Max(Min(projected_value, 100-nectar_value), 0)

            item := '{"name": "' . j . '", "value": ' . (nectar_value + 0) . ', "projected": ' . Round(projected_value) . '}'
            nectarItems .= (A_Index > 1 ? ", " : "") . item
        }

        buffItems := ""
        buffMax := Map(
            "redboost", 10, "whiteboost", 10, "blueboost", 10,
            "haste", 10, "focus", 10, "bombcombo", 10, "balloonaura", 10,
            "inspire", 50, "reindeerfetch", 10, "honeymark", 10, "pollenmark", 10,
            "popstar", 1, "melody", 1, "bear", 1, "babylove", 1, "jbshare", 1, "guiding", 1
        )

        time_value := (60*A_Min+A_Sec)//6
        i := (time_value = 0) ? 600 : time_value
        for k, name in ["redboost","whiteboost","blueboost","haste","focus","bombcombo","balloonaura","inspire","reindeerfetch","honeymark","pollenmark","popstar","melody","bear","babylove","jbshare","guiding"]
        {
            value := buffValues[name].Has(i) ? buffValues[name][i] : 0
            maxVal := buffMax[name]
            percent := (maxVal > 0) ? Round(Min((value/maxVal)*100, 100)) : 0
            item := '{"name": "' . name . '", "value": ' . (value + 0) . ', "percent": ' . percent . '}'
            buffItems .= (k > 1 ? ", " : "") . item
        }

        f.Write('var statData = {"honey": ' . currentHoney . ', "backpack": ' . currentBackpack . ', "planters": [' . planterItems . '], "nectars": [' . nectarItems . '], "buffs": [' . buffItems . ']};')
        f.Close()
    }
}

JsonEscape(value) {
    value := StrReplace(value, "\", "\\")
    value := StrReplace(value, '"', '\"')
    value := StrReplace(value, "`r", "\r")
    value := StrReplace(value, "`n", "\n")
    return value
}

;===========================================
; OCR FUNCTIONS - exact copy from StatMonitor.ahk
;===========================================
HBitmapToRandomAccessStream(hBitmap) {
    static IID_IRandomAccessStream := "{905A0FE1-BC53-11DF-8C49-001E4FC686DA}"
         , IID_IPicture            := "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
         , PICTYPE_BITMAP := 1
         , BSOS_DEFAULT   := 0
         , sz := 8 + A_PtrSize * 2

    DllCall("Ole32\CreateStreamOnHGlobal", "Ptr", 0, "UInt", true, "PtrP", &pIStream:=0, "UInt")

    PICTDESC := Buffer(sz, 0)
    NumPut("uint", sz, "uint", PICTYPE_BITMAP, "ptr", hBitmap, PICTDESC)

    riid := CLSIDFromString(IID_IPicture)
    DllCall("OleAut32\OleCreatePictureIndirect", "Ptr", PICTDESC, "Ptr", riid, "UInt", false, "PtrP", &pIPicture:=0, "UInt")
    ComCall(15, pIPicture, "Ptr", pIStream, "UInt", true, "UIntP", &size:=0, "UInt")
    riid := CLSIDFromString(IID_IRandomAccessStream)
    DllCall("ShCore\CreateRandomAccessStreamOverStream", "Ptr", pIStream, "UInt", BSOS_DEFAULT, "Ptr", riid, "PtrP", &pIRandomAccessStream:=0, "UInt")
    ObjRelease(pIPicture)
    ObjRelease(pIStream)
    Return pIRandomAccessStream
}

CLSIDFromString(IID, &CLSID?) {
    CLSID := Buffer(16)
    if res := DllCall("ole32\CLSIDFromString", "WStr", IID, "Ptr", CLSID, "UInt")
        throw Error("CLSIDFromString failed. Error: " . Format("{:#x}", res))
    Return CLSID
}

ocr(file, lang := "FirstFromAvailableLanguages") {
    static OcrEngineStatics, OcrEngine, MaxDimension, LanguageFactory, Language, CurrentLanguage:="", BitmapDecoderStatics, GlobalizationPreferencesStatics
    if !IsSet(OcrEngineStatics) {
        CreateClass("Windows.Globalization.Language", ILanguageFactory := "{9B0252AC-0C27-44F8-B792-9793FB66C63E}", &LanguageFactory)
        CreateClass("Windows.Graphics.Imaging.BitmapDecoder", IBitmapDecoderStatics := "{438CCB26-BCEF-4E95-BAD6-23A822E58D01}", &BitmapDecoderStatics)
        CreateClass("Windows.Media.Ocr.OcrEngine", IOcrEngineStatics := "{5BFFA85A-3384-3540-9940-699120D428A8}", &OcrEngineStatics)
        ComCall(6, OcrEngineStatics, "uint*", &MaxDimension:=0)
    }
    text := ""
    if (file = "ShowAvailableLanguages") {
        if !IsSet(GlobalizationPreferencesStatics)
            CreateClass("Windows.System.UserProfile.GlobalizationPreferences", IGlobalizationPreferencesStatics := "{01BF4326-ED37-4E96-B0E9-C1340D1EA158}", &GlobalizationPreferencesStatics)
        ComCall(9, GlobalizationPreferencesStatics, "ptr*", &LanguageList:=0)
        ComCall(7, LanguageList, "int*", &count:=0)
        loop count {
            ComCall(6, LanguageList, "int", A_Index-1, "ptr*", &hString:=0)
            ComCall(6, LanguageFactory, "ptr", hString, "ptr*", &LanguageTest:=0)
            ComCall(8, OcrEngineStatics, "ptr", LanguageTest, "int*", &bool:=0)
            if (bool = 1) {
                ComCall(6, LanguageTest, "ptr*", &hText:=0)
                b := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hText, "uint*", &length:=0, "ptr")
                text .= StrGet(b, "UTF-16") "`n"
            }
            ObjRelease(LanguageTest)
        }
        ObjRelease(LanguageList)
        return text
    }
    if (lang != CurrentLanguage) or (lang = "FirstFromAvailableLanguages") {
        if IsSet(OcrEngine) {
            ObjRelease(OcrEngine)
            if (CurrentLanguage != "FirstFromAvailableLanguages")
                ObjRelease(Language)
        }
        if (lang = "FirstFromAvailableLanguages")
            ComCall(10, OcrEngineStatics, "ptr*", OcrEngine)
        else {
            CreateHString(lang, &hString)
            ComCall(6, LanguageFactory, "ptr", hString, "ptr*", &Language:=0)
            DeleteHString(hString)
            ComCall(9, OcrEngineStatics, "ptr", Language, "ptr*", &OcrEngine:=0)
        }
        if (OcrEngine = 0) {
            return ""
        }
        CurrentLanguage := lang
    }
    IRandomAccessStream := file
    ComCall(14, BitmapDecoderStatics, "ptr", IRandomAccessStream, "ptr*", &BitmapDecoder:=0)
    WaitForAsync(&BitmapDecoder)
    BitmapFrame := ComObjQuery(BitmapDecoder, IBitmapFrame := "{72A49A1C-8081-438D-91BC-94ECFC8185C6}")
    ComCall(12, BitmapFrame, "uint*", &width:=0)
    ComCall(13, BitmapFrame, "uint*", &height:=0)
    if (width > MaxDimension) or (height > MaxDimension)
        return ""
    BitmapFrameWithSoftwareBitmap := ComObjQuery(BitmapDecoder, IBitmapFrameWithSoftwareBitmap := "{FE287C9A-420C-4963-87AD-691436E08383}")
    ComCall(6, BitmapFrameWithSoftwareBitmap, "ptr*", &SoftwareBitmap:=0)
    WaitForAsync(&SoftwareBitmap)
    ComCall(6, OcrEngine, "ptr", SoftwareBitmap, "ptr*", &OcrResult:=0)
    WaitForAsync(&OcrResult)
    ComCall(6, OcrResult, "ptr*", &LinesList:=0)
    ComCall(7, LinesList, "int*", &count:=0)
    loop count {
        ComCall(6, LinesList, "int", A_Index-1, "ptr*", &OcrLine:=0)
        ComCall(7, OcrLine, "ptr*", &hText:=0)
        buf := DllCall("Combase.dll\WindowsGetStringRawBuffer", "ptr", hText, "uint*", &length:=0, "ptr")
        text .= StrGet(buf, "UTF-16") "`n"
        ObjRelease(OcrLine)
    }
    Close := ComObjQuery(IRandomAccessStream, IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}")
    ComCall(6, Close)
    Close := ComObjQuery(SoftwareBitmap, IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}")
    ComCall(6, Close)
    ObjRelease(IRandomAccessStream)
    ObjRelease(BitmapDecoder)
    ObjRelease(SoftwareBitmap)
    ObjRelease(OcrResult)
    ObjRelease(LinesList)
    return text
}

CreateClass(str, interface, &Class) {
    CreateHString(str, &hString)
    GUID := CLSIDFromString(interface)
    result := DllCall("Combase.dll\RoGetActivationFactory", "ptr", hString, "ptr", GUID, "ptr*", &Class:=0)
    if (result != 0)
        throw Error("RoGetActivationFactory failed: " result)
    DeleteHString(hString)
}

CreateHString(str, &hString) {
    DllCall("Combase.dll\WindowsCreateString", "wstr", str, "uint", StrLen(str), "ptr*", &hString:=0)
}

DeleteHString(hString) {
    DllCall("Combase.dll\WindowsDeleteString", "ptr", hString)
}

WaitForAsync(&Object) {
    AsyncInfo := ComObjQuery(Object, IAsyncInfo := "{00000036-0000-0000-C000-000000000046}")
    loop {
        ComCall(7, AsyncInfo, "uint*", &status:=0)
        if (status != 0) {
            if (status != 1) {
                throw Error("AsyncInfo status error")
            }
            break
        }
        sleep 10
    }
    ComCall(8, Object, "ptr*", &ObjectResult:=0)
    ObjRelease(Object)
    Object := ObjectResult
}
