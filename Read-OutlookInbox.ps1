
 Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
 $olDefaultFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
 $outlook = New-Object -comobject Outlook.Application
 $namespace = $outlook.GetNameSpace("MAPI") 
 $folders = $namespace.Folders.Item(2) #May vary, try to change number(0-99) if object results empty.
 $folder = $folders.Folders.Item("Inbox")
 $sender= $folder.Items |  Select -ExpandProperty SenderName -Last 1 
 $subject = $folder.Items | Select -ExpandProperty Subject -Last 1  
 
 $speaker = New-Object -ComObject SAPI.SPVoice
 $a = $speaker.voice= $speaker.GetVoices().item(2)
 $speaker.speak("You have email from $sender with title $subject")