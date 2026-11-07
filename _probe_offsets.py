import struct
samples = [
"1e426807801609cf09085654504d4f202020007feea6018d0480053a0201",
"26622907801657ff090841494334454d202000908814025c0419ddda41eb109f42c0032b0201",
"1c276107800b814f080853454a3437312020007d00fe015504c00700",
"1768190780079c4908085655424b48202020000080fc00",
]
for s in samples:
    b=bytes.fromhex(s)
    print('\nlen',len(b),'first',b[0])
    for i,v in enumerate(b):
        print(f"{i:02d}:{v:02x}", end=' ')
    print()
    # try float decode on every 4-byte aligned chunk after 16
    for i in range(0,len(b)-3):
        chunk=b[i:i+4]
        be=struct.unpack('>f',chunk)[0]
        le=struct.unpack('<f',chunk)[0]
        if -180<=be<=180:
            if abs(be)>1:
                pass
        # print likely lat lon around india
        if -90<=be<=90 and -90<=le<=90:
            continue
    if len(b)>=34:
        for i in range(len(b)-4):
            be=struct.unpack('>f', b[i:i+4])[0]
            if 5<=be<=40:
                print('be candidate',i,be)
            be2=struct.unpack('>f', b[i+4:i+8])[0] if i+8<=len(b) else None
            if be2 is not None and 60<=be2<=100 and 5<=be<=40:
                print('pair',i,be,be2)
