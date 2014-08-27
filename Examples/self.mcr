c:\biblio.mdb
calculate
init|i
init|j
input|i|Input Something|input
if|#i#=1
for|j|0|1|1
msg|you entered 1|what
next
elseif|#i#=2
msg|you entered 2|what
elseif|#i#=3
msg|you entered 3|what
elseif|#i#=4
msg|you entered 4|what
elseif|#i#=5
msg|you entered 5|what
else
msg|you entered something|what
end if
msg|The sum of numbers is #j#|Sum
break