import numpy as np


a = 10
b = 20
c = np.arange(1,1000)
# j = np.arange(0,len(c))
# mask = j < 10

print (c[(c < 10) & (c > 4)])