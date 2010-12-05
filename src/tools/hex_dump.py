def hex_dump(data):
    hexi = ascii = ''
    j = 0
    
    for i in range(len(data)):
        hexi += ' %02x' % ord(data[i])
        ascii += data[i] if ord(data[i]) >= 32 else '.'

        if (j == 7):
            hexi += ' '
            ascii += ' '

        j += 1
        if j == 16 or i == len(data) - 1:
            print "%04x %-49s  %s" % (i + 1, hexi, ascii)

            hexi = ascii = ''
            j = 0
