# Useful Excel Forumlas

## Convert 8 and 10 digit OS grid refs to 6 digits in excel with a formula
an excel formula to convert 8 and 10 digit grid references to 6 digit

Works for formulas in a single string with no spaces, eg

NS7486320333

becomes

NS 748 203

=LEFT(A1,2)&" "&MID(A1,3,(((LEN(A1)-2)/2)+2)/2)&" "&MID(A1,(((LEN(A1)-2)/2)+2)+1,(((LEN(A1)-2)/2)+2)/2)

## Top 5 
Where a list of sites is in cells Sheet1!B2:B33, and the corresponding cells in column C have the numbers of things at each site, this will return the sites with the biggest number of things. The ROW(A1) returns the firs biggest, so dragging down a column will then give you the second biggest in the cell beneath, etc.

=INDEX(Sheet1!$B$2:$B$33, MATCH(LARGE(Sheet1!C$2:C$33, ROW($A1)),Sheet1!C$2:C$33, 0))
