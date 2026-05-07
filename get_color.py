from PIL import Image
from collections import Counter

img = Image.open('avla_logo.png').convert('RGB')
w, h = img.size

samples = []
for x in range(5, min(60, w)):
    for y in range(5, min(40, h)):
        r, g, b = img.getpixel((x, y))
        samples.append((r, g, b))

common = Counter(samples).most_common(5)
for c, n in common:
    print(f'RGB{c} = #{c[0]:02X}{c[1]:02X}{c[2]:02X}  (n={n})')
