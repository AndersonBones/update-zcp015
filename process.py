import math
def process_status(num, full_process=False):
    if full_process == True:
        size = 21
        porcentagem = num*100/size
        return f'{math.trunc(porcentagem)}%'
    else:
        size = 10
        porcentagem = num*100/size
        return f'{math.trunc(porcentagem)}%'