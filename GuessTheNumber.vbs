dim luck
min =1
max=9
randomize
dim ran
ran = Int((max-min+1)*Rnd+min)
maxChances=0
Do 
 luck =InputBox("try your luck") 
 
 if ((luck-ran)=0) then 
    MsgBox "you Won" 
	Exit Do
 end if
 maxChances=maxChances+1
 if maxChances = 3 then 
    MsgBox "you Lose" 
	Exit Do
 end if
Loop While luck <> ran
