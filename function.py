def lenWell(x1, y1, x2, y2):
    lenght = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
    return lenght

def lenWell_gorizont(x1, y1, x2, y2, xx2, yy2):
    L1 = lenWell(x1, y1, x2, y2)
    L2 = lenWell(x1, y1, xx2, yy2)
    L = lenWell(x2, y2, xx2, yy2) # длина ГС
    if (L1 * L1 > L2 * L2 + L * L) or (L2 * L2 > L1 * L1 + L * L):
        P = min(L1, L2)
    else:
        if x2 == xx2:
            x3 = x2
            y3 = y1
        elif y2 == yy2:
            x3 = x1
            y3 = y2
        else:
            A = yy2 - y2
            B = x2 - xx2
            C = -1 * x2 * (yy2 - y2) + y2 * (xx2 - x2)
            x3 = (x2 * ((yy2 - y2)**2) + x1 * ((xx2 - x2)**2) + (xx2 - x2) * (yy2 - y2) * (y1 - y2)) / ((yy2 - y2)**2 + (xx2 - x2)**2)
            y3 = (xx2 - x2) * (x1 - x3) / (yy2 - y2) + y1

            x3 = (B * x1 / A - C / B - y1) * A * B / (A * A + B * B)
            y3 = B * x3 / A + y1 - B * x1 / A

        P = lenWell(x1, y1, x3, y3)

    return P
