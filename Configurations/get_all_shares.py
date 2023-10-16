from nsetools import Nse    
from nsepython import *
nse = Nse()
all_stock_codes = nse.get_stock_codes()
print(all_stock_codes)

print('###############################')

a = fnolist()
#print(a[1])
#for item in a:
#    print(item,',',item)