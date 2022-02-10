import math
def function2(x):
    return "%.58f" %(float(p*(1-(p**2)*(x**2)))/float((1-(p*x)+(p**2)*(x**2))**2));
pler=[.8,.95,.75]
for p in pler:
    print(50*"="+" p="+str(float(p))+" "+50*"=")
    xler=[0,float(1+math.sqrt((4*p)-3))/float(2*p),float(1-math.sqrt((4*p)-3))/float(2*p)]
    for i in range(len(xler)):
        print("x"+str(i+1)+": "+str(xler[i])+" ,\t\tfunction result:"+str(function2(xler[i])))
