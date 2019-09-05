import math

primes = [2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, 53, 59, 61, 67, 71, 73, 79, 83, 89, 97]

mymin = 0
mymax = len(primes) - 1
target = 67 
iteration = 0

while mymin <= mymax:
  iteration = iteration + 1
  myresult = math.floor((mymin + mymax) / 2)
  print(str(myresult))
  if primes[myresult] < target:
    mymin = myresult + 1
  elif primes[myresult] > target:
    mymax = myresult - 1
  else:
    print(primes[myresult])
    break
  print(" no target match")