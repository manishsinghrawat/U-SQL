c:\biblio.mdb
Create table
edit|select authors.author from authors,titleauthor where authors.au_id=titleauthor.au_id 
break
Insert Records
init|a
init|b
init|c
input|a|Enter Author ID|Entry
input|b|Enter author name|Entry
input|c|Enter Year Born|Entry
sql|insert into authors values(#a#,'#b#',#c#)
show|authors
break