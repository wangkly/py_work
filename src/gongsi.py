# 公司代码筛选
def gongsi_filter(index,value):
    def inner_filter(n):
        if value == '-1':
            return True
            
        name = n[index]
        return name.value == value
    return inner_filter


