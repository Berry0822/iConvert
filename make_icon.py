import math
from PIL import Image, ImageDraw

S = 256
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

# --- rounded gradient background ---
bg = Image.new("RGBA", (S, S), (0, 0, 0, 0))
gd = ImageDraw.Draw(bg)
top = (37, 99, 235)      # #2563eb
bot = (29, 78, 216)      # #1d4ed8
for y in range(S):
    t = y / S
    c = (int(top[0] + (bot[0]-top[0])*t),
         int(top[1] + (bot[1]-top[1])*t),
         int(top[2] + (bot[2]-top[2])*t), 255)
    gd.line([(0, y), (S, y)], fill=c)
mask = Image.new("L", (S, S), 0)
ImageDraw.Draw(mask).rounded_rectangle([0, 0, S-1, S-1], radius=56, fill=255)
img.paste(bg, (0, 0), mask)

d = ImageDraw.Draw(img)
cx = cy = S // 2
R = 66
w = 24
white = (255, 255, 255, 255)

def head(angle_deg, pointing_deg, size=30):
    """Filled triangle arrowhead at point on circle, aimed along 'pointing'."""
    a = math.radians(angle_deg)
    px, py = cx + R*math.cos(a), cy + R*math.sin(a)
    p = math.radians(pointing_deg)
    dx, dy = math.cos(p), math.sin(p)          # forward
    nx, ny = -dy, dx                            # normal
    tip = (px + dx*size, py + dy*size)
    b1 = (px + nx*size*0.8, py + ny*size*0.8)
    b2 = (px - nx*size*0.8, py - ny*size*0.8)
    d.polygon([tip, b1, b2], fill=white)

# Top arc (clockwise), arrowhead at right end pointing down-ish
d.arc([cx-R, cy-R, cx+R, cy+R], start=185, end=330, fill=white, width=w)
head(330, 330+90)   # tangent (clockwise) = angle+90

# Bottom arc (clockwise), arrowhead at left end pointing up-ish
d.arc([cx-R, cy-R, cx+R, cy+R], start=5, end=150, fill=white, width=w)
head(150, 150+90)

# Save PNG preview + multi-size ICO
img.save("/sessions/upbeat-brave-curie/mnt/outputs/iConvert/icon_preview.png")
img.save("/sessions/upbeat-brave-curie/mnt/outputs/iConvert/icon.ico",
         sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])
print("icon.ico + icon_preview.png written")
