Dim objuft

Set objuft=createObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.open("C:\Users\sfjbs\Desktop\UFT\Framework\KeywordDrivenFramework\Driver\Driver")
objuft.Test.Run
objuft.Test.Close
objuft.quit
set objuft=nothing