import io, struct
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_voices\andrew.mp3"
b = io.open(p, "rb").read()
print("size", len(b), "head", b[:4].hex())
# skip ID3
i = 0
if b[:3] == b"ID3":
    sz = ((b[6] & 0x7f) << 21) | ((b[7] & 0x7f) << 14) | ((b[8] & 0x7f) << 7) | (b[9] & 0x7f)
    i = 10 + sz
h = b[i:i+4]
print("frame hdr", h.hex())
br = [0,32,40,48,56,64,80,96,112,128,160,192,224,256,320,0]
sr = [44100,48000,32000,0]
bits = struct.unpack(">I", h)[0]
ver = (bits >> 19) & 3
layer = (bits >> 17) & 3
bi = (bits >> 12) & 15
si = (bits >> 10) & 3
mode = (bits >> 6) & 3
print("ver", ver, "layer", layer, "kbps", br[bi], "sr", sr[si], "mode(3=mono)", mode)
