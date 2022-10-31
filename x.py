def round_to(n, precision):
    correction = 0.5 if n >= 0 else -0.5
    return int( n/precision+correction ) * precision

def myround(n):
    return round_to(n, 0.5)
print(myround(3.74))


print("Test")