def outer_filter(optins,index):
    def inner_filter(n):
        cell = n[index]
        return cell.value in optins
    return inner_filter        


