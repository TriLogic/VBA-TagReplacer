# VBA-TagReplacer
A library for complex text replacement rendered in VBA for applications.

This VBA library can be used in MS Excel, MS Access, MS Word, VB5, VB6 to perform complex text replacement
operations. 

1) Tags are formatted like so: "Hello World! It is ${Temp} degrees outside."
   Where '${Temp}' will be replaced with whatever the replacement value for 'Temp' would be.
2) Tags may contain other tags: "Hello World! It is ${TempIn${Units}} degrees ${Units}." 
   In this case the tag '${Units}' might replaced with the value 'F' or 'C'
   and the resulting ${TempInC} or ${TempInF} used to retrieve the corresponding temperature value according to the
   termerature scale desired.
   
The code was originally written in MSAccess and designed to "compile" a pattern to be used when iterating through 
recordsets.

~Enjoy
