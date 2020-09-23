P L E A S E   V O T E   F O R   T H I S   C O D E.

This code is great for mp3 players and card shufflers.  You just enter the range for the 
numbers let it generate the sequence and then keep calling the random numbers from a function. 
When it reaches the end of the sequence it will create a new one.

You have to enter the lower and upper bounds for the sequence.

There are two functions which you have to use:
   kRandomize (LowerBound, UpperBound) - this will generate a random sequence.

   kRnd - this will call the random numbers without repeats.

This code is only good if you want to generate a sequence with a range of less than 5000.  
On my P333, which is so low on drive space that it runs like a P200, it can generate a range of 5000 in 13 secs.
It can generate upto 1200 in less than a second.  If you do need a huge range but don't need the whole sequence 
straight away then you could add DOEVENTS in all the loops and let it create the sequence while the user is using your
program.  It gets really slow towards the end when it is trying to find the last few numbers which havn't been used.  If 
anyone out there does improve the code or finds a way of speeding it up then please let me know at: kg50c@yahoo.com

P L E A S E   V O T E   F O R   T H I S   C O D E.