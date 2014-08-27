c:\biblio.mdb
assrank
init|a
init|i
crt|temp|select * from student order by marks desc
crt|temp1|select * from student
cardinality|temp|a
for|i|1|#a#|1
editit|temp|#i#|3|#i#
next
sql|drop table student
crt|student|select temp1.name,temp1.marks,temp.rank from temp1,temp where temp1.name=temp.name
sql|drop table temp
sql|drop table temp1
show|student
break

