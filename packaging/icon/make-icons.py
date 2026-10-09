#!/usr/bin/env python3
# Tegner AulaSyncs ikon (et kalenderblad med lektioner på blå baggrund) og skriver
#   AulaSync.ico  (Windows: exe, installationsprogram, vinduer)
#   AulaSync.icns (Mac: AulaSync.app, macOS 15)
#   AulaSync.icon (Mac: Apples nye format til macOS 26 og nyere; make-dmg.sh oversætter det med actool)
#   AulaSync.png  (1024 px, til README og forhåndsvisning)
#   tray/         (systembakken på Windows: tray*.ico i 16-64 px; menulinjen på Mac: menubar*.png, skabelon-ikoner)
#   wizard/       (det lille billede i installations- og afinstallationsprogrammet)
# Kør igen efter ændringer:  python3 packaging/icon/make-icons.py   (kræver Pillow)
import json
from pathlib import Path
from PIL import Image, ImageChops, ImageDraw, ImageFilter

HERE = Path(__file__).parent
BLUE_TOP, BLUE_BOTTOM = (74, 138, 244), (36, 88, 179)        # accentfarven #2f6fde i en let gradient
HEADER = (33, 83, 159)                                        # StaffInk #21539f
BARS = [((204, 221, 248), (33, 83, 159)),                     # medarbejder (StaffInk #21539f)
        ((201, 233, 212), (31, 107, 61)),                     # klasse (ClassInk #1f6b3d)
        ((231, 212, 245), (107, 53, 144))]                    # lokale (RoomInk #6b3590)
S = 4                                                         # tegnes i 4x og skaleres ned (kantudglatning)


def gradient(size, top, bottom):
    img = Image.new('RGB', (1, size))
    for y in range(size):
        t = y / (size - 1)
        img.putpixel((0, y), tuple(round(a + (b - a) * t) for a, b in zip(top, bottom)))
    return img.resize((size, size))


def rr(d, box, radius, **kw):
    # Afrundet rektangel på hele pixels; Pillow tegner sømme ved kommatal.
    x0, y0, x1, y1 = (round(v) for v in box)
    d.rounded_rectangle((x0, y0, x1, y1), round(min(radius, (x1 - x0) / 2, (y1 - y0) / 2)), **kw)


def draw(px, tile, radius, detail, background=True, dot=False):
    """Ikon på px x px. tile = (x, y, w) for den blå flade; detail: 2 = fuld, 1 = forenklet, 0 = mindst.
    background=False: kun kalenderbladet på gennemsigtig baggrund (lag til AulaSync.icon)."""
    n = px * S
    img = Image.new('RGBA', (n, n), (0, 0, 0, 0))
    tx, ty, tw = (v * S for v in tile)
    r = radius * S
    if background:
        mask = Image.new('L', (n, n), 0)
        rr(ImageDraw.Draw(mask), (tx, ty, tx + tw, ty + tw), r, fill=255)
        if detail == 2:  # svag skygge under fladen (Mac-stil)
            shadow = Image.new('RGBA', (n, n), (0, 0, 0, 0))
            rr(ImageDraw.Draw(shadow), (tx, ty + tw * 0.012, tx + tw, ty + tw * 1.012), r, fill=(0, 0, 0, 90))
            img = Image.alpha_composite(img, shadow.filter(ImageFilter.GaussianBlur(tw * 0.02)))
        bg = gradient(n, BLUE_TOP, BLUE_BOTTOM).convert('RGBA')
        img.paste(bg, (0, 0), mask)
    d = ImageDraw.Draw(img)

    # Kalenderbladet
    u = tw / 100                                              # enhed: 1 % af fladen
    px0, py0, pw, ph = tx + 21 * u, ty + 25 * u, 58 * u, 54 * u
    pr = 7 * u
    rr(d, (px0, py0, px0 + pw, py0 + ph), pr, fill=(255, 255, 255))
    hh = 13 * u if detail else 15 * u
    rr(d, (px0, py0, px0 + pw, py0 + hh + pr), pr, fill=HEADER)
    d.rectangle(tuple(round(v) for v in (px0, py0 + hh, px0 + pw, py0 + hh + pr)), fill=(255, 255, 255))
    # Ringe
    for cx in (px0 + 16 * u, px0 + pw - 16 * u):
        w = 5.5 * u if detail else 7 * u
        rr(d, (cx - w / 2, py0 - 7 * u, cx + w / 2, py0 + 6 * u), w / 2, fill=(255, 255, 255),
                            outline=HEADER if detail else None, width=max(1, round(1.4 * u)) if detail else 0)
    if dot:  # systembakke "opdaterer": en prik midt i bladet i stedet for lektionerne
        cx, cy, r = px0 + pw / 2, py0 + hh + (ph - hh) / 2 + 2 * u, 10 * u
        d.ellipse(tuple(round(v) for v in (cx - r, cy - r, cx + r, cy + r)), fill=HEADER)
        return img.resize((px, px), Image.LANCZOS)
    # Lektioner
    rows = BARS if detail else BARS[:2]
    top = py0 + hh + (7 * u if detail else 8 * u)
    gap = (ph - hh - (12 * u if detail else 14 * u)) / len(rows)
    for i, (fill, ink) in enumerate(rows):
        y = top + i * gap
        bh = gap * (0.62 if detail else 0.7)
        x0, x1 = px0 + 8 * u, px0 + pw - 8 * u
        if detail == 2:
            x1 = px0 + pw - (8 + (0, 14, 6)[i]) * u
        rr(d, (x0, y, x1, y + bh), bh * 0.28, fill=fill)
        rr(d, (x0, y, x0 + max(bh * 0.4, 2.6 * u), y + bh), bh * 0.28, fill=ink)
    return img.resize((px, px), Image.LANCZOS)


def windows(px, dot=False):
    # Windows: fladen fylder næsten hele ikonet; små størrelser forenkles og holder sig til hele pixels. dot: systembakkens
    # "opdaterer" (en prik i stedet for lektionerne) på præcis samme flade.
    if px <= 24:
        return draw(px, (0, 0, px), px * 0.2, 0, dot=dot)
    if px <= 48:
        return draw(px, (0, 0, px), px * 0.2, 1, dot=dot)
    m = px * 0.03
    return draw(px, (m, m, px - 2 * m), px * 0.19, 2 if px >= 128 else 1, dot=dot)


def mac(px):
    # Mac: Big Sur-gitteret, 824/1024 af lærredet med hjørneradius 185/1024 og plads til skyggen.
    t = px * 824 / 1024
    return draw(px, ((px - t) / 2, px * 92 / 1024, t), px * 185 / 1024, 2 if px >= 64 else 1)


AMBER = (178, 106, 0)                                         # "kræver handling" (#b26a00)
TRAY_SIZES = [16, 20, 24, 28, 32, 36, 40, 48, 56, 64]         # systembakken: 16 px gange 100-400 % skalering
WIZARD_SIZES = [58, 77, 97, 116, 124, 143, 159]               # Inno Setup 6.6's lille billede ved 100-250 % (dokumentationen)


def attention(img):
    # Prik i øverste højre hjørne med en gennemsigtig kant, så den skilles fra fladen; resten af ikonet røres ikke.
    px = img.width
    n, r, gap = px * S, px * 0.21 * S, max(1, round(px / 16)) * S
    hole = Image.new('L', (n, n), 0)
    ImageDraw.Draw(hole).ellipse((n - 2 * r - gap, -gap, n + gap, 2 * r + gap), fill=255)
    dot = Image.new('RGBA', (n, n), (0, 0, 0, 0))
    ImageDraw.Draw(dot).ellipse((n - 2 * r, 0, n, 2 * r), fill=AMBER)
    out = img.copy()
    out.putalpha(ImageChops.subtract(img.getchannel('A'), hole.resize((px, px), Image.LANCZOS)))
    return Image.alpha_composite(out, dot.resize((px, px), Image.LANCZOS))


def tray(px, state):
    # Windows: præcis exe-ikonet (windows(px)); "opdaterer" bytter lektionerne ud med en prik på samme flade.
    img = windows(px, dot=(state == 'updating'))
    return attention(img) if state == 'attention' else img


def menubar(px, state):
    # Mac: samme kalenderblad som skabelon-ikon (kun alfa tæller; macOS farver det sort/hvidt).
    n, u = px * S, px * S / 100
    img = Image.new('RGBA', (n, n), (0, 0, 0, 0))
    d, K, C = ImageDraw.Draw(img), (0, 0, 0, 255), (0, 0, 0, 0)
    x0, y0, x1, y1, sw, hh = 8 * u, 16 * u, 92 * u, 92 * u, 8 * u, 22 * u
    rr(d, (x0, y0, x1, y1), 12 * u, fill=K)                    # bladet med fyldt hoved
    rr(d, (x0 + sw, y0 + hh, x1 - sw, y1 - sw), 6 * u, fill=C)  # hul krop
    for cx in (30 * u, 70 * u):                                 # ringe med luft omkring
        rr(d, (cx - 7 * u, 4 * u, cx + 7 * u, y0 + 10 * u), 7 * u, fill=C)
        rr(d, (cx - 4 * u, 6 * u, cx + 4 * u, y0 + 8 * u), 4 * u, fill=K)
    if state == 'updating':
        r, cy = 11 * u, (y0 + hh + y1 - sw) / 2
        d.ellipse((50 * u - r, cy - r, 50 * u + r, cy + r), fill=K)
    else:
        for i, w in enumerate((56, 40)):
            y = y0 + hh + 12 * u + i * 18 * u
            rr(d, (x0 + sw + 8 * u, y, x0 + sw + (8 + w) * u, y + 9 * u), 4 * u, fill=K)
    if state == 'attention':
        r, g = 16 * u, 7 * u
        d.ellipse((n - 2 * r - g, -g, n + g, 2 * r + g), fill=C)
        d.ellipse((n - 2 * r, 0, n, 2 * r), fill=K)
    return img.resize((px, px), Image.LANCZOS)


# AulaSync.icon: Icon Composer-format med den blå flade som fyld og kalenderbladet som eneste lag (fylder lærredet, så
# macOS selv lægger det på sin flade).
ICON_JSON = {
    'fill': {'automatic-gradient': 'srgb:%.5f,%.5f,%.5f,1.00000' % tuple(c / 255 for c in (47, 111, 222))},
    'groups': [{
        'layers': [{'name': 'kalender', 'image-name': 'kalender.png', 'glass': False}],
        'shadow': {'kind': 'neutral', 'opacity': 0.5},
        'translucency': {'enabled': False, 'value': 0},
        'specular': False,
    }],
    'supported-platforms': {'squares': 'shared'},
}


if __name__ == '__main__':
    # Windows' størrelser (100-400 % skalering) plus 128; vinduerne bruger op til 128, Stifinder 256.
    ico_sizes = [16, 20, 24, 30, 32, 36, 40, 48, 60, 64, 72, 80, 96, 128, 256]
    ico = [windows(s) for s in ico_sizes]
    ico[-1].save(HERE / 'AulaSync.ico', sizes=[(s, s) for s in ico_sizes], append_images=ico[:-1])
    mac(1024).save(HERE / 'AulaSync.icns', append_images=[mac(s) for s in (16, 32, 64, 128, 256, 512)])
    mac(1024).save(HERE / 'AulaSync.png')
    icon = HERE / 'AulaSync.icon'
    (icon / 'Assets').mkdir(parents=True, exist_ok=True)
    draw(1024, (0, 0, 1024), 0, 2, background=False).save(icon / 'Assets' / 'kalender.png')
    (icon / 'icon.json').write_text(json.dumps(ICON_JSON, indent=2) + '\n')
    tray_dir = HERE / 'tray'
    tray_dir.mkdir(exist_ok=True)
    for state, suffix in (('normal', ''), ('updating', '-updating'), ('attention', '-attention')):
        frames = [tray(s, state) for s in TRAY_SIZES]
        frames[-1].save(tray_dir / f'tray{suffix}.ico', sizes=[(s, s) for s in TRAY_SIZES], append_images=frames[:-1])
        menubar(36, state).save(tray_dir / f'menubar{suffix}.png')   # 18 punkter på Retina
    # Det lille billede øverst til højre i installations- og afinstallationsprogrammet (Inno Setup: 58 px ved 100 %).
    wizard_dir = HERE / 'wizard'
    wizard_dir.mkdir(exist_ok=True)
    for px in WIZARD_SIZES:
        windows(px).save(wizard_dir / f'AulaSync-{px}.png')
    print('ok:', ', '.join(p.name for p in sorted(HERE.glob('AulaSync.*'))))
