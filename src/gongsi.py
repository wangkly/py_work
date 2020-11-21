# 公司代码筛选
def gongsi_filter(index,value):
    def inner_filter(n):
        name = n[index]
        return name.value == value
    return inner_filter


