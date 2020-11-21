#期末余额
from decimal import *
getcontext().prec = 28
def qimo(row):
    return Decimal(row[-1].value)