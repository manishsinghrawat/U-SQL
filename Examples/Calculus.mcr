c:\biblio.mdb
Differentiate
init|i
init|j
init|k
init|l
init|m
input|i|Enter the function|Differentiate
input|j|Enter the value at which differential is to be calculated|Differentiate
set|j|#j#
asg|k|#i#
asg|l|#i#
replace|k|x|(#j#+0.0005)
replace|l|x|(#j#)
set|m|(#k#-#l#)/0.0005
msg|The Differential is #m#.|Sum
break
Integrate
init|func
init|lower
init|upper
init|sum
init|temp
init|m
input|func|Enter the function|Integrate
input|lower|Enter the Lower Limit|Integrate
input|upper|Enter the Upper Limit|Integrate
set|sum|0
set|lower|#lower#
set|upper|#upper#
for|m|#lower#|#upper#|0.04
asg|temp|#func#
replace|temp|x|#m#
set|temp|#temp#
set|sum|#sum#+(#temp#)*0.04
next
msg|The Integral is #sum#.|Sum
break
function Calculator
init|i
init|j
init|k
init|l
init|m
input|i|Enter the function|Function Calculator
input|j|Enter the value of function|Function Calculator
set|j|#j#
asg|k|#i#
replace|k|x|(#j#)
set|m|#k#
msg|Value of function is #m#.|Sum
break
Plotter
init|i
init|a
init|b
init|k
init|l
init|m
init|n
clrg
input|i|Enter the function|Function Calculator
input|a|Enter the lower limit of domain|Function Calculator
set|a|#a#
input|b|Enter the upper limit of domain|Function Calculator
set|b|#b#
newdata|manish|65535
for|n|#a#|#b#|0.04
asg|k|#i#
replace|k|x|(#n#)
set|m|#k#
entdata|manish|#n#|#m#
next
plot|manish
showgraph
break
Differential Plotter
init|i
init|a
init|b
init|k
init|l
init|m
init|n
clrg
input|i|Enter the function|Function Calculator
input|a|Enter the lower limit of domain|Function Calculator
set|a|#a#
input|b|Enter the upper limit of domain|Function Calculator
set|b|#b#
newdata|manish|65535
for|n|#a#|#b#|0.04
asg|k|#i#
asg|l|#i#
replace|k|x|(#n#+0.0005)
replace|l|x|(#n#)
set|m|(#k#-#l#)/0.0005
entdata|manish|#n#|#m#
next
plot|manish
showgraph
break