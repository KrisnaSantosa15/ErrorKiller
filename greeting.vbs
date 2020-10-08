
dim speaks,speech,username,currentTime

username =CreateObject("WScript.Network").Username
speaks = "hello mr.krisna?how are you today?? Today is"
currentTime = FormatDateTime(time(),1)

Set speech=CreateObject("sapi.SpVoice")

set speech.voice = speech.GetVoices("gender=female").item(0)

speech.Speak username
speech.Speak speaks
speech.Speak currentTime