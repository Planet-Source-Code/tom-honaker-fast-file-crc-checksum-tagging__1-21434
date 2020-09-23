

WHO WROTE THIS...
-----------------------------------------------------------------------------
Hi, I'm Tom Honaker. I'm a freelance webmaster and crazed old-school coder.
I do most of my coding in Visual Studio 6 Enterprise Edition, in VB 
specifically. I also code in C/C++ and dabble in OO Pascal (Delphi) although 
I never got really into Delphi. 



CREDITS...
-----------------------------------------------------------------------------
I'm flinging mad props to these people, places, and things...

 - Fredrik Qvarfort, for his killer CRC checksum generator class, clsCRC.
   Without this single chunk of code the CRC verification code presented
   here would not have been possible.

 - The guy known to me only as "Detonate." He posts a good bit of good
   code on Planet Source Code, and seems to know what he's doing. In this
   case, it was one of his code submissions that gave me the idea to make
   this verification system.

 - Planet Source Code, THE best Visual Basic sourcecode website. Go here
   to visit:

   http://www.planet-source-code.com/vb/

   Both of the above guys can be looked up on Planet Source Code's search
   function.

I'm hereby acknowledging the Copyright on the CRC checksum generator class, 
which is Copyright (C) 2000 by Fredrik Qvarfort. Reproduced here with 
permission. If you use this code, please DO NOT MODIFY ANY COPYRIGHT 
NOTICES or we'll have to do bad things to you in your sleep.



WHAT IS IT...
-----------------------------------------------------------------------------
Lightning-quick CRC32 checksum generation is impossible to do in Visual
Basic. Yes, you can do CRC32 generation and several people have, but it's
a slow process in VB. 

Then a fellow named Fredrik Qvarfort comes along and posts his code on
generating 16- and 32-bit CRC checksums using a curious mix of precompiled
assembler and VB code. This code is CRAZY kind of fast, 50+ megabytes a
second! The actual CRC generation code is done strictly in assembler, so
it's fast to the point of unholy and evil, and he skillfully blended it
right into VB.

I'd seen an idea on CRC-protecting an executable posted on Planet Source
Code by a fellow calling himself "Detonate." His code was very well done
and worked nicely, but as it was an all-VB approach the speed limitations
of VB were evident. Detonate did a LOT to speed-optimize his code but
in the end he created a selective-CRC system to permit spot-checksumming 
critical portions of a file, which helped offset the speed problems. (I
don't know if that was the motivation for his creating the selective
version, but I'd imagine that had something to do with it. I haven't
asked him so I don't know for sure.)

So, I took the CRC checksum code from Mr. Qvarfort and the idea of an 
executable-file checksumming program from Detonate, threw both into a 
mixer with a scoop of VB code, and poured out a small, tight, FAST CRC
checksum system for executables that combined Detonate's ideas with the
ASM speed Mr. Qvarfort created.



WHAT FILE TYPES DOES IT WORK ON...
-----------------------------------------------------------------------------
I've used it on ActiveX controls and DLLS as well as standalone executables
and all have worked fine.

However, if you are using an executable compressor you may find that CRC 
tagging won't work, depending on the type of compressor. My experiments with
UPX 1.07 have shown to me that UPX-compressed executables cannot be reliably
tagged with a CRC that will match the executable. Not sure why that is, but
I'm looking into it.



HOW BIG A FILE CAN IT HANDLE...
-----------------------------------------------------------------------------
I've been using it on controls, DLLs, and EXEs of varying sizes up to a 
half-megabyte or so in size, and the execution time of the CRC verification
code is so quick it's not perceptible in most cases. I'd expect that one 
could in practical terms get away with using the code on files up to 2 to 3
megabytes without having a ridiculous pause in execution as a result.

As always, though, your mileage may vary. ;-)



HOW DO I USE IT...
-----------------------------------------------------------------------------
It's simple - open the project "exeCRC.vbp" in the "Executable CRC-Checksum 
Tagger" directory and look it over. Then compile it into an EXE when you're
happy that it's not infested with malicious code.

Use the compiled EXE to "tag" files you want to CRC-check. The program will
calculate the CRC for the file and append the CRC as a 4-byte "tag" after 
the end of the file.

In an executable you want to add CRC file checking to, you need to add the
clsCRC.cls class and CheckCRC.bas module, and call a single function in
CheckCRC.bas from your program. That's it. If the CRC isn't what the tag
says it should be, the file was changed.



FINAL NOTES...
-----------------------------------------------------------------------------
This code is about as strightforward as it can be, and is quite heavily 
commented, so most reasonably expert VB programmers should be able to 
implement it immediately and modify it easily. If you modify the CRC 
checksum generation code, please be sure to let Fredrik Qvarfort know by
E-mailing him from Planet Source Code. And, please let ME know if you 
modify my code portions (contact me the same way)



-----------------------------------------------------------------------------
END OF README.TXT
