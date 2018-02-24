Dim Message, ALIN
Dim coins
Dim game
Dim cards

Randomize

Function RandomWithinRange(min, max)
    RandomWithinRange = Int((max - min + 1) * Rnd() + min)
End Function

Function RandItemFromArray(arr)
    RandItemFromArray= arr(RandomWithinRange(LBound(arr), UBound(arr)))
End Function

    cards = Array("ace of spades", "ace of diamonds", "ace of clubs", "ace of hearts", "two of spades", "two of diamonds", "two of clubs", "two of hearts", "three of spades", "three of diamonds", "three of clubs", "three of hearts", "four of spades", "four of diamonds", "four of clubs", "four of hearts", "five of spades", "five of diamonds", "five of clubs", "five of hearts", "six of spades", "six of diamonds", "six of clubs", "six of hearts", "seven of spades", "seven of diamonds", "seven of clubs", "seven of hearts", "eight of spades", "eight of diamonds", "eight of clubs", "eight of hearts", "nine of spades", "nine of diamonds", "nine of clubs", "nine of hearts", "ten of spades", "ten of diamonds", "ten of clubs", "ten of hearts", "king of spades", "king of diamonds", "king of clubs", "king of hearts", "queen of spades", "queen of diamonds", "queen of clubs", "queen of hearts", "jack of spades", "jack of diamonds", "jack of clubs", "jack of hearts")

Randomize

Function RandomWithinRange(min, max)
    RandomWithinRange = Int((max - min + 1) * Rnd() + min)
End Function

Function RandItemFromArray(arr)
    RandItemFromArray= arr(RandomWithinRange(LBound(arr), UBound(arr)))
End Function

    coins = Array("Heads", "Tails", "Ok, It's Heads", "Ok, It's Tails")

Randomize

Function RandomWithinRange(min, max)
    RandomWithinRange = Int((max - min + 1) * Rnd() + min)
End Function

Function RandItemFromArray(arr)
    RandItemFromArray= arr(RandomWithinRange(LBound(arr), UBound(arr)))
End Function

    game = Array("Rock", "Paper", "Scissors")

Start=MsgBox("welcome to ALIN 2016 v1.6! the ALIN program was designed by and is property of DroneTech under the direction of Thinkit.inc. You are not permitted to modify ALIN in any way. press ok to continue.",1+0,"ALIN �copyright Thinkit.inc 2016 all rights reserved")

If Start=vbOk then User=InputBox("Please tell me your name","ALIN")
Set Speak=CreateObject("sapi.spvoice")
If User="matthew zenn" then Speak.Speak "greetings master, today is " & Date & ", and the time is " & Time else if Start=vbOk then Speak.Speak "greetings "+User+"....today is " & Date & "...and the time is " & Time & ".....allow me to introduce myself ...i'm ALIN, the advanced learning interface network" else speak.speak "ok, good bye!"

If Start=vbOk then speak.speak "how is your day"
If Start=vbOk then x=InputBox ("'Good' or 'Bad'", "ALIN")

If x= "good" then speak.speak "that's great!"
If x= "good" then speak.speak "what do you need"
If x= "good" then v=InputBox ("what do you want me to do!", "ALIN")

If v= "tell me a joke" then speak.speak "what do virtual assistance do in their free time?...nothing, because virtual assistance don't get free time"
If v= "tell me the time" then speak.speak (Time)
If v= "tell me the date" then speak.speak (Date)
If v= "guess my age" then a=InputBox("what year where you born in","ALIN")
If v= "roll a dice" then Dice= Int((6 - 1+1) * rnd + 1)
If v= "flip a coin" then speak.speak RandItemFromArray(coins)
If v= "rock paper scissors" then speak.speak "ok, 1, 2, 3, i chose " & RandItemFromArray(game)
If v= "pick a card" then speak.speak "ok, i chose the " & RandItemFromArray(cards)

If v= "roll a dice" then speak.speak "rolling .............................................. it landed on " & Dice
If v= "guess my age" then speak.speak Year(date)-a

If v= "calculator" then speak.speak "what's the equation"
If v= "calculator" then num1=inputbox("Number 1","ALIN")
If v= "calculator" then symbol=inputbox("equation symbol","ALIN")
If v= "calculator" then num2=inputbox("Number 2","ALIN")

If symbol="+" then speak.speak num1 + num2
If symbol="-" then speak.speak num1 - num2
If symbol="*" then speak.speak num1 * num2
If symbol="/" then speak.speak num1 / num2

If x= "bad" then speak.speak "i hope your day gets better"