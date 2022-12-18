def calcCategory(mark):
    if mark >= 90 or mark == 100:
        return 'A+'
    elif mark >= 80 or mark == 89:
        return 'A'
    elif mark >= 70 or mark == 79:
        return 'B'
    elif mark >= 60 or mark == 69:
        return 'C'
    elif mark >= 50 or mark == 59:
        return 'D'
    elif mark >= 40 or mark == 49:
        return 'E'
    elif mark >= 30 or mark == 39:
        return 'F'
    elif mark >= 20 or mark == 29:
        return 'G'
    elif mark >= 0 or mark == 19:
        return 'U'

def assignMarkComment(mark):
    # print(type(mark))
    if mark >= 80:
        return 'Excellent'
    elif mark >= 70 or mark == 79:
        return 'Very good'
    elif mark >= 60 or mark == 69:
        return 'Good'
    elif mark >= 50 or mark == 59:
        return 'Fair'
    if mark >= 0 or mark == 49:
        return 'Below average'
    elif not isinstance(mark, float) or not isinstance(mark, int):
        return ''
    
# print(assignMarkComment(float(95)))