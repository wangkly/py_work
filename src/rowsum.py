#每行期间借方，期间贷方之和
from decimal import *
getcontext().prec = 28
def rowSum(row):
    return Decimal(row[-3].value) + Decimal(row[-2].value)