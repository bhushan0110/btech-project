pattern =""
string=""
ans=[]
length=0
cnt=0

req = str(input("Enter required Pattern:"))
l = int(len(req))

i=0
while i<l:
  string+='#'
  i+=1


print("Enter time sequence")
print("Press * to end:")
while 1:
  tmp=input("Next value:")
  if str(tmp)=='*':
    break
  length +=1
  pattern+=str(tmp)
  string+=str(tmp)
  string=string[1:l+10]

  if string==req:
    ans.append(length)
    cnt+=1

print("Count of occurance of Pattern:"+str(cnt))
print("Index of occurance of pattern:"+str(ans))