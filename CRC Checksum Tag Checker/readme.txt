
IMPORTANT NOTE...
-----------------------------------------------------------------------------
Please note that this is the code that is to be inserted into a project that 
will be CRC-tagged and protected. It will not compile into a standalone 
executable as it is presented here.



TO USE THIS CODE...
-----------------------------------------------------------------------------
1. Add the class clsCRC.cls and module CheckCRC.bas to your project.

2. Edit this line in the VerifyCRC() function in CheckCRC.bas if necessary:

    TFile$ = App.Path & "\" & App.EXEName & ".exe"

   If you are using this code in another type of file, change the extenion 
   from ".exe." to whatever your project's compiled form will be (".ocx"
   for ActiveX controls, etc.)

3. Call VerifyCRC() from your program and check the result returned.
   A return value of True means the CRC tag matched the actual CRC checksum
   for the file. A return value of False means that either the file has no 
   CRC tag (in which case, tag it with the CRC Tagger project's compiled
   executable) or the file has been altered and there is a CRC mismatch.



CHECKING OTHER FILES...
-----------------------------------------------------------------------------
If you want to use the code to check other files, change the TFile$ 
reference mentioned above to point to the file and make sure you "tag" it 
before you run the compiled project. Simple.
