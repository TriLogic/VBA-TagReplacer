# VBA-TagReplacer
A library for complex text replacement rendered in VBA for applications.

This VBA library can be used in MS Excel, MS Access, MS Word, VB5, VB6 to perform complex text replacement.

Tags are formatted like so: "Hello World! It is ${Temp} degrees outside."

In this example the value of the tag '${Temp}' would be replaced with whatever the replacement value for 'Temp' is. The resulting text might be: 

   "Hello World! It is 86 degrees outside."

Tags may contain other tags: "Hello World! The temperature outside is ${TempIn${Units}}${Units}." 
   
In this case the tag '${Units}' might replaced with the value 'F' or 'C' and the resulting ${TempInC} or ${TempInF} used to retrieve the corresponding temperature value according to the temperature scale desired with the final resulting being:

   "Hello World! The temperature outside is 78F."
   
The code was originally written in MSAccess and designed to "compile" a pattern to be used when iterating through 
recordsets.

~Enjoy
