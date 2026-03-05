"""
Steps Recorder — MS Steps Recorder Replacement
Screenshots on every mouse click, with annotated click highlights,
exported to a self-contained HTML report with editable captions.

Requirements:
    pip install pynput pillow

Usage:
    python steps_recorder.py
    Press ENTER to start, click around, press ENTER again to stop.
    The HTML report opens in your browser — add captions there and save.
"""

import base64
import io
import os
import sys
import threading
import time
from datetime import datetime

try:
    from pynput import mouse
    from PIL import Image, ImageDraw, ImageGrab
except ImportError:
    print("Missing dependencies. Please run:")
    print("    pip install pynput pillow")
    sys.exit(1)

# ── Config ────────────────────────────────────────────────────────────────────
HIGHLIGHT_RADIUS   = 22
HIGHLIGHT_COLOR    = (255, 60, 60, 190)
HIGHLIGHT_BORDER   = (255, 255, 255, 230)
HIGHLIGHT_BORDER_W = 3
SCREENSHOT_DELAY   = 0.10
OUTPUT_DIR         = "."
# ─────────────────────────────────────────────────────────────────────────────

steps = []
recording = False
_listener = None
_lock = threading.Lock()


def take_screenshot(x, y):
    time.sleep(SCREENSHOT_DELAY)
    try:
        shot = ImageGrab.grab()
    except Exception as e:
        print(f"  [!] Screenshot failed: {e}")
        return ""

    img = shot.convert("RGBA")
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    r, bw = HIGHLIGHT_RADIUS, HIGHLIGHT_BORDER_W
    draw.ellipse([x-r-bw, y-r-bw, x+r+bw, y+r+bw], fill=HIGHLIGHT_BORDER)
    draw.ellipse([x-r, y-r, x+r, y+r], fill=HIGHLIGHT_COLOR)
    draw.line([x-r+5, y, x+r-5, y], fill=(255, 255, 255, 220), width=2)
    draw.line([x, y-r+5, x, y+r-5], fill=(255, 255, 255, 220), width=2)

    result = Image.alpha_composite(img, overlay).convert("RGB")
    buf = io.BytesIO()
    result.save(buf, format="PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


def on_click(x, y, button, pressed):
    global steps, recording
    if not recording or not pressed:
        return

    step_num = len(steps) + 1
    btn_name = "Right-click" if button == mouse.Button.right else "Click"
    print(f"  Step {step_num}: {btn_name} at ({int(x)}, {int(y)})")

    b64 = take_screenshot(int(x), int(y))

    with _lock:
        steps.append({
            "num":       step_num,
            "action":    btn_name,
            "x":         int(x),
            "y":         int(y),
            "timestamp": datetime.now().strftime("%H:%M:%S"),
            "image_b64": b64,
        })


def build_html(steps, title):
    date_str = datetime.now().strftime("%A, %d %B %Y  %H:%M:%S")

    cards = ""
    for s in steps:
        img_html = (
            f'<img src="data:image/png;base64,{s["image_b64"]}" alt="Step {s["num"]}" loading="lazy">'
            if s["image_b64"] else
            '<div class="no-img">Screenshot unavailable</div>'
        )

        cards += f"""
    <div class="card" id="step-{s['num']}">
      <div class="card-head">
        <span class="badge">Step {s['num']}</span>
        <span class="action">{s['action']} at ({s['x']}, {s['y']})</span>
        <span class="ts">⏱ {s['timestamp']}</span>
      </div>
      <div class="card-body">{img_html}</div>
      <div class="card-desc">
        <textarea class="desc-input" data-step="{s['num']}" placeholder="Add a caption for this step…" rows="2"></textarea>
      </div>
    </div>"""

    empty = '<div class="empty"><h2>No steps were recorded.</h2></div>' if not steps else ""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#0f1117;color:#e2e8f0;min-height:100vh}}

.hdr{{background:linear-gradient(135deg,#1a1d2e,#16213e,#0f3460);border-bottom:1px solid #2d3748;padding:28px 40px;position:sticky;top:0;z-index:100;box-shadow:0 4px 24px rgba(0,0,0,.4)}}
.hdr-inner{{max-width:1100px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap}}
.hdr h1{{font-size:1.5rem;font-weight:700;background:linear-gradient(90deg,#63b3ed,#a78bfa);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}}
.hdr p{{font-size:.8rem;color:#718096;margin-top:4px}}
.hdr-right{{display:flex;align-items:center;gap:20px}}
.stat-val{{font-size:2rem;font-weight:700;color:#63b3ed;line-height:1}}
.stat-lbl{{font-size:.7rem;color:#718096;text-transform:uppercase;letter-spacing:.08em;margin-top:2px;text-align:center}}

.save-btn{{background:linear-gradient(135deg,#3182ce,#7c3aed);color:#fff;border:none;padding:10px 22px;border-radius:8px;font-size:.85rem;font-weight:600;cursor:pointer;transition:opacity .2s;white-space:nowrap}}
.save-btn:hover{{opacity:.85}}
.save-btn:disabled{{opacity:.45;cursor:not-allowed}}
.save-notice{{font-size:.75rem;color:#68d391;margin-top:5px;text-align:center;min-height:1em}}

main{{max-width:1100px;margin:0 auto;padding:36px 24px 80px}}

.card{{background:#1a1d2e;border:1px solid #2d3748;border-radius:12px;overflow:hidden;margin-bottom:24px;transition:border-color .2s}}
.card:hover{{border-color:#4a5568}}
.card-head{{display:flex;align-items:center;gap:12px;padding:13px 18px;background:#16213e;border-bottom:1px solid #2d3748;flex-wrap:wrap}}
.badge{{background:linear-gradient(135deg,#3182ce,#7c3aed);color:#fff;font-size:.72rem;font-weight:700;padding:3px 10px;border-radius:20px;letter-spacing:.05em;white-space:nowrap}}
.action{{font-size:.86rem;font-weight:600;color:#e2e8f0;flex:1}}
.ts{{font-size:.76rem;color:#718096;white-space:nowrap}}
.card-body{{padding:14px;background:#0f1117}}
.card-body img{{width:100%;border-radius:6px;border:1px solid #2d3748;display:block;cursor:zoom-in;transition:opacity .2s}}
.card-body img:hover{{opacity:.9}}
.no-img{{padding:40px;text-align:center;color:#4a5568;font-size:.85rem}}

.card-desc{{padding:12px 14px;background:#1a1d2e;border-top:1px solid #2d3748}}
.desc-input{{width:100%;background:#0f1117;border:1px solid #2d3748;border-radius:6px;color:#e2e8f0;font-size:.84rem;font-family:inherit;padding:9px 12px;resize:vertical;transition:border-color .2s;line-height:1.5}}
.desc-input:focus{{outline:none;border-color:#63b3ed}}
.desc-input::placeholder{{color:#4a5568}}

.empty{{text-align:center;padding:100px 20px;color:#4a5568}}
.empty h2{{color:#718096}}

.lb{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.93);z-index:999;align-items:center;justify-content:center;cursor:zoom-out}}
.lb.on{{display:flex}}
.lb img{{max-width:95vw;max-height:95vh;border-radius:8px;box-shadow:0 0 60px rgba(0,0,0,.8)}}
.lb-x{{position:fixed;top:18px;right:26px;font-size:2rem;color:#fff;cursor:pointer;opacity:.7;line-height:1}}
.lb-x:hover{{opacity:1}}

footer{{text-align:center;padding:20px;font-size:.72rem;color:#4a5568;border-top:1px solid #1a1d2e}}
</style>
</head>
<body>

<header class="hdr">
  <div class="hdr-inner">
    <div>
      <h1>📋 {title}</h1>
      <p>Recorded on {date_str}</p>
    </div>
    <div class="hdr-right">
      <div>
        <div class="stat-val">{len(steps)}</div>
        <div class="stat-lbl">Steps</div>
      </div>
      <div style="text-align:center">
        <button class="save-btn" id="save-btn" onclick="saveReport()">💾 Save Report</button>
        <div class="save-notice" id="save-notice"></div>
      </div>
    </div>
  </div>
</header>

<main>
  {cards or empty}
</main>
<footer>Generated by Steps Recorder &nbsp;·&nbsp; {date_str}</footer>

<div class="lb" id="lb">
  <span class="lb-x" id="lb-x">✕</span>
  <img id="lb-img" src="" alt="">
</div>

<script>
// ── Lightbox ──
const lb=document.getElementById('lb'),lbImg=document.getElementById('lb-img');
document.querySelectorAll('.card-body img').forEach(i=>i.addEventListener('click',()=>{{lbImg.src=i.src;lb.classList.add('on')}}));
lb.addEventListener('click',e=>{{if(e.target!==lbImg)lb.classList.remove('on')}});
document.getElementById('lb-x').addEventListener('click',()=>lb.classList.remove('on'));
document.addEventListener('keydown',e=>{{if(e.key==='Escape')lb.classList.remove('on')}});

// ── Save report with captions baked in ──
async function saveReport() {{
  const btn = document.getElementById('save-btn');
  const notice = document.getElementById('save-notice');
  btn.disabled = true;
  notice.textContent = '';

  // Bake current textarea values into the HTML before serialising
  document.querySelectorAll('.desc-input').forEach(ta => {{
    ta.setAttribute('value', ta.value);         // not used by textarea, but harmless
    ta.innerHTML = escapeHtml(ta.value);         // this IS what gets serialised
  }});

  const html = '<!DOCTYPE html>\\n' + document.documentElement.outerHTML;

  try {{
    const handle = await window.showSaveFilePicker({{
      suggestedName: '{title}'.replace(/[^a-z0-9]/gi,'_') + '.html',
      types: [{{ description: 'HTML File', accept: {{'text/html': ['.html']}} }}]
    }});
    const writable = await handle.createWritable();
    await writable.write(new Blob([html], {{type:'text/html'}}));
    await writable.close();
    notice.textContent = '✔ Saved!';
  }} catch(e) {{
    if (e.name !== 'AbortError') {{
      // Fallback for browsers without File System Access API
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([html], {{type:'text/html'}}));
      a.download = '{title}'.replace(/[^a-z0-9]/gi,'_') + '_captioned.html';
      a.click();
      notice.textContent = '✔ Downloaded!';
    }}
  }}

  setTimeout(() => notice.textContent = '', 3000);
  btn.disabled = false;
}}

function escapeHtml(s) {{
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}
</script>
</body>
</html>"""


def main():
    global recording, _listener

    print("=" * 54)
    print("  Steps Recorder  —  MS Steps Recorder Replacement")
    print("=" * 54)
    print()
    print("  • Screenshots taken on every mouse click")
    print("  • Add captions in the HTML report afterwards")
    print()

    title = input("  Session title (or ENTER for default): ").strip()
    if not title:
        title = f"Steps Recording - {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    input("\n  Press ENTER to START recording... ")

    recording = True
    _listener = mouse.Listener(on_click=on_click)
    _listener.start()
    print("\n  ● Recording — switch to the window you want to document.")
    print("    Press ENTER here when done.\n")
    input()

    recording = False
    _listener.stop()
    print(f"\n  ■ Stopped. {len(steps)} step(s) captured.")

    if not steps:
        print("  Nothing to save. Exiting.")
        return

    html = build_html(steps, title)
    safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in title)[:48].strip()
    fname = f"steps_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe}.html".replace(" ", "_")
    out = os.path.join(OUTPUT_DIR, fname)

    with open(out, "w", encoding="utf-8") as f:
        f.write(html)

    abs_path = os.path.abspath(out)
    print(f"\n  ✔ Report saved: {abs_path}\n")

    try:
        import webbrowser
        webbrowser.open(f"file://{abs_path}")
        print("  Opened in your default browser.")
    except Exception:
        pass


if __name__ == "__main__":
    main()