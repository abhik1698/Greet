dim speaks,speech,dys,msg1,msg2,ss

Set speech=CreateObject("sapi.spvoice")
Set speech.Voice = speech.GetVoices.Item(0)

i=hour(time)
ye= year(Year(date))
mo= Month( Month (date))
da= (day(date))

if (mo & "/" & da)= "m/d"	then
	ss = "I think you should be asleep by now Sur, What in the Fuck are you Trying to do at this time"
end if
speech.Speak ss
If i<6 Then
	speaks="I think you should be asleep by now Sur, What in the Fuck are you Trying to do at this time"
ElseIf i<12 Then
	speaks="Good Morning Sur, Have a Good daiy"
End If
If i>16 Then
	speaks="Good Evening sur!"
ElseIf i>12 Then
	speaks="Good AfterNoon Sur, Have a Good daiy"
End If


speech.Speak speaks

d=weekday(date)

Select Case d
	Case 1
		dys="It's a Sunday"
	Case 2
		dys="For god's sayk, It's a Monday"
	Case 3
		dys="It's a Tuesday"
	Case 4
		dys="Well, It's a Wednesday"
	Case 5
		dys="Thuuu, It's a Thurdsay"
	Case 6
		dys="It's a freakin Friday"
	case else
		dys="Its Saturday by the way"
End Select

speech.Speak dys

'msg1="Now I am Playing a SOng which suits you. Seriously trust me it is like a song which is sung for you"

'speech.Speak msg1

'CreateObject("WScript.Shell").Run("""E:music\\Raajakumara.mp3""")



'CreateObject("Wscript.Shell").Run "wmplayer /play /close ""D:\Games\Sounds_CounterStrike\"&day(date)&".wav""", 0, False