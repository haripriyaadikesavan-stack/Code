"""Wireframe PNG of a dashboard's zones (main layout) to review placement."""
import sys, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from lxml import etree

def fm(z):
    s = z.find("zone-style")
    return {f.get("attr"): f.get("value") for f in s.findall("format")} if s is not None else {}

def draw(d, out):
    size = d.find("size"); W = int(size.get("maxwidth")); H = int(size.get("maxheight"))
    fig = plt.figure(figsize=(W/100, H/100), dpi=60)
    ax = fig.add_axes([0,0,1,1]); ax.set_xlim(0,W); ax.set_ylim(H,0); ax.axis("off")
    top = d.find("zones").findall("zone")
    def rect(z, floating):
        x,y,w,h = [int(z.get(k))/100000*(W if k in "xw" else H) for k in "xywh"]
        t = z.get("type-v2") or ("sheet" if z.get("name") else "?")
        f = fm(z); bg = f.get("background-color")
        if t.startswith("layout"):
            if bg: ax.add_patch(Rectangle((x,y),w,h,fc=bg,ec="none"))
            return
        col = {"sheet":"#ffffff","text":bg or "#fff8c0","paramctrl":"#dbe9ff","filter":"#ffe0e0","color":"#f0e0ff","dashboard-object":"#ffd9a0","title":bg or "#eeeeee","empty":bg or "#f8f8f8"}.get(t,"#cccccc")
        ax.add_patch(Rectangle((x,y),w,h,fc=col if t!="sheet" else (bg or "#ffffff"),ec="red" if floating else "#555555",lw=1.2 if floating else 0.6, alpha=0.9))
        label = z.get("name") or ""
        if t == "text":
            label = "".join(r.text or "" for r in z.iter("run")).strip()[:50]
        elif t in ("paramctrl","filter","color"):
            label = t+": "+(z.get("param") or "").split("].[")[-1][:30]
        if label and w > 30 and h > 10:
            ax.text(x+3, y+3, label, fontsize=7, va="top", ha="left", clip_on=True, color="white" if bg=="#1f7a4d" else "black")
    for z in top[0].iter("zone"):
        if z is not top[0]: rect(z, False)
        else:
            bg = fm(z).get("background-color")
            if bg: ax.add_patch(Rectangle((0,0),W,H,fc=bg,ec="none"))
    for zz in top[1:]:
        for z in zz.iter("zone"): rect(z, True)
    fig.savefig(out); plt.close(fig)

r = etree.parse(sys.argv[1]).getroot()
os.makedirs(sys.argv[2], exist_ok=True)
names = sys.argv[3:] or [d.get("name") for d in r.find("dashboards")]
for d in r.find("dashboards"):
    if d.get("name") in names:
        draw(d, os.path.join(sys.argv[2], d.get("name").replace("/","_")+".png"))
