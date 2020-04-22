Set uftApp = CreateObject("QuickTest.Application")
uftApp.Launch
uftApp.visible = True

uftApp.open"C:\Users\xinlu\Documents\Unified Functional Testing\GUITest4",True
set uftTest=uftApp.Test
uftTest.Run
wScript.sleep 100000
uftTest.close
uftApp.quit
set uftApp = nothing