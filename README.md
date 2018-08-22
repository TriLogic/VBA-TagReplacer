# VBA-TagReplacer
A library for complex text replacement rendered in VBA for applications.

This VBA library can be used in MS Excel, MS Access, or MS Word (perhaps others) to perform complex text replacement
operations. 

1) Tags are formatted like so: ${abc} - Where 'abc' is the string key for which a replacement value is to be supplied.
2) Tags may contain other tags: ${abc${def}} - In this case the tag '${def}' is first resolved replaced with its replacement value '123'
   resulting in ${abc123} which in turn is replaced with its replacement value.
   
In the example provided the text "abc${def}ghi${jkl${mno}pqr${stu${vw}x}y}z"
is ultimately replaced with      "abc@{DEF}ghi[jkl[mno]pqr[stu[vw]x]y]z" using an example only TagSource.

The code is written to "compile" a pattern to be used repeatedly

Enjoy!
~Andrew
