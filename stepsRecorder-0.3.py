"""
Steps Recorder — MS Steps Recorder Replacement

Requirements:
    pip install pynput pillow

Usage:
    python steps_recorder.py
    - Press F9 (anywhere) to START/STOP recording
    - Click around to capture steps
    - Edit title, captions, reorder, delete in the HTML report
    - Save or Export to PDF from the report
"""

import base64
import io
import os
import sys
import threading
import time
from datetime import datetime

try:
    from pynput import mouse, keyboard
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
DEBOUNCE_MS        = 500
TOGGLE_KEY         = keyboard.Key.f9
OUTPUT_DIR         = "."
# ─────────────────────────────────────────────────────────────────────────────

steps = []
recording = False
_mouse_listener = None
_kb_listener    = None
_lock           = threading.Lock()
_last_click_time = 0.0


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
    global steps, recording, _last_click_time
    if not recording or not pressed:
        return

    now = time.time()
    if (now - _last_click_time) * 1000 < DEBOUNCE_MS:
        return
    _last_click_time = now

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


def on_key_press(key):
    global recording, _mouse_listener
    if key != TOGGLE_KEY:
        return

    if not recording:
        recording = True
        _mouse_listener = mouse.Listener(on_click=on_click)
        _mouse_listener.start()
        print("\n  ● Recording STARTED (press F9 to stop)...")
    else:
        recording = False
        if _mouse_listener:
            _mouse_listener.stop()
        print(f"\n  ■ Recording STOPPED. {len(steps)} step(s) captured so far.")
        print("  Press F9 to resume, or ENTER here to finish and generate report.")


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
    <div class="card" draggable="true" id="step-{s['num']}" data-step="{s['num']}">
      <div class="card-head">
        <span class="drag-handle" title="Drag to reorder">⠿</span>
        <span class="badge">Step {s['num']}</span>
        <span class="action">{s['action']} at ({s['x']}, {s['y']})</span>
        <span class="ts">⏱ {s['timestamp']}</span>
        <button class="delete-btn" title="Delete this step" onclick="deleteStep(this)">✕</button>
      </div>
      <div class="card-body">{img_html}</div>
      <div class="card-desc">
        <textarea class="desc-input" placeholder="Add a caption for this step…" rows="2"></textarea>
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

/* ── Header ── */
.hdr{{background:linear-gradient(135deg,#1a1d2e,#16213e,#0f3460);border-bottom:1px solid #2d3748;padding:28px 40px;position:sticky;top:0;z-index:100;box-shadow:0 4px 24px rgba(0,0,0,.4)}}
.hdr-inner{{max-width:1100px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap}}
.hdr h1{{font-size:1.5rem;font-weight:700}}
#report-title{{background:linear-gradient(90deg,#63b3ed,#a78bfa);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;border-bottom:1px dashed #4a5568;padding-bottom:1px;outline:none;cursor:text;transition:border-color .2s}}
#report-title:hover{{border-color:#63b3ed}}
#report-title:focus{{border-color:#a78bfa}}
.hdr p{{font-size:.8rem;color:#718096;margin-top:4px}}
.hdr-right{{display:flex;align-items:center;gap:16px;flex-wrap:wrap}}
.stat-val{{font-size:2rem;font-weight:700;color:#63b3ed;line-height:1}}
.stat-lbl{{font-size:.7rem;color:#718096;text-transform:uppercase;letter-spacing:.08em;margin-top:2px;text-align:center}}

.btn{{border:none;padding:10px 18px;border-radius:8px;font-size:.82rem;font-weight:600;cursor:pointer;transition:opacity .2s;white-space:nowrap;color:#fff}}
.btn:hover{{opacity:.85}}
.btn:disabled{{opacity:.4;cursor:not-allowed}}
.btn-save{{background:linear-gradient(135deg,#3182ce,#7c3aed)}}
.btn-pdf{{background:linear-gradient(135deg,#2f855a,#276749)}}
.btn-group{{display:flex;flex-direction:column;align-items:center;gap:4px}}
.save-notice{{font-size:.72rem;color:#68d391;min-height:1em;text-align:center}}

/* ── Layout ── */
main{{max-width:1100px;margin:0 auto;padding:36px 24px 80px}}

/* ── Cards ── */
.card{{background:#1a1d2e;border:1px solid #2d3748;border-radius:12px;overflow:hidden;margin-bottom:24px;transition:border-color .2s,opacity .2s,box-shadow .2s}}
.card:hover{{border-color:#4a5568}}
.card.dragging{{opacity:.4;box-shadow:0 0 0 2px #63b3ed}}
.card.drag-over{{border-color:#a78bfa;box-shadow:0 -3px 0 0 #a78bfa}}

.card-head{{display:flex;align-items:center;gap:10px;padding:13px 18px;background:#16213e;border-bottom:1px solid #2d3748;flex-wrap:wrap}}
.drag-handle{{font-size:1.1rem;color:#4a5568;cursor:grab;padding:0 4px;user-select:none;line-height:1}}
.drag-handle:active{{cursor:grabbing}}
.badge{{background:linear-gradient(135deg,#3182ce,#7c3aed);color:#fff;font-size:.72rem;font-weight:700;padding:3px 10px;border-radius:20px;letter-spacing:.05em;white-space:nowrap}}
.action{{font-size:.86rem;font-weight:600;color:#e2e8f0;flex:1}}
.ts{{font-size:.76rem;color:#718096;white-space:nowrap}}
.delete-btn{{background:none;border:none;color:#fc8181;font-size:1rem;cursor:pointer;padding:2px 6px;border-radius:4px;line-height:1;transition:background .15s;margin-left:4px}}
.delete-btn:hover{{background:rgba(252,129,129,.15)}}

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

/* ── Lightbox ── */
.lb{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.93);z-index:999;align-items:center;justify-content:center;cursor:zoom-out}}
.lb.on{{display:flex}}
.lb img{{max-width:95vw;max-height:95vh;border-radius:8px;box-shadow:0 0 60px rgba(0,0,0,.8)}}
.lb-x{{position:fixed;top:18px;right:26px;font-size:2rem;color:#fff;cursor:pointer;opacity:.7;line-height:1}}
.lb-x:hover{{opacity:1}}

footer{{text-align:center;padding:20px;font-size:.72rem;color:#4a5568;border-top:1px solid #1a1d2e}}

/* ── Print / PDF ── */
@media print {{
  body{{background:#fff;color:#000}}
  .hdr{{position:static;background:#fff;border-bottom:2px solid #000;box-shadow:none;padding:20px}}
  .hdr h1{{font-size:1.4rem}}
  #report-title{{background:none;-webkit-text-fill-color:#000;color:#000;border-bottom:none}}
  .hdr p{{color:#555}}
  .hdr-right,.drag-handle,.delete-btn,.save-notice{{display:none!important}}
  main{{padding:20px 0}}
  .card{{border:1px solid #ccc;border-radius:6px;margin-bottom:20px;break-inside:avoid;page-break-inside:avoid;box-shadow:none}}
  .card-head{{background:#f0f0f0;border-bottom:1px solid #ccc;padding:10px 14px}}
  .badge{{background:#333;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
  .action{{color:#000}}
  .ts{{color:#555}}
  .card-body{{background:#fff;padding:10px}}
  .card-body img{{border:1px solid #ddd}}
  .card-desc{{background:#fafafa;border-top:1px solid #ddd;padding:10px 14px}}
  .desc-input{{background:#fafafa;border:none;color:#000;font-size:.85rem;resize:none;width:100%}}
  .lb,.btn{{display:none!important}}
  footer{{color:#999;border-top:1px solid #ddd}}
}}
</style>
</head>
<body>

<header class="hdr">
  <div class="hdr-inner">
    <div>
      <h1>📋 <span id="report-title" contenteditable="true" spellcheck="false" title="Click to edit title">{title}</span></h1>
      <p>Recorded on {date_str}</p>
    </div>
    <div class="hdr-right">
      <div>
        <div class="stat-val" id="step-count">{len(steps)}</div>
        <div class="stat-lbl">Steps</div>
      </div>
      <div class="btn-group">
        <button class="btn btn-save" id="save-btn" onclick="saveReport()">💾 Save Report</button>
        <div class="save-notice" id="save-notice"></div>
      </div>
      <button class="btn btn-pdf" onclick="window.print()">🖨 Export PDF</button>
    </div>
  </div>
</header>

<main id="steps-container">
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
function bindImgClicks(){{
  document.querySelectorAll('.card-body img').forEach(i=>{{
    i.onclick=()=>{{lbImg.src=i.src;lb.classList.add('on')}};
  }});
}}
bindImgClicks();
lb.addEventListener('click',e=>{{if(e.target!==lbImg)lb.classList.remove('on')}});
document.getElementById('lb-x').onclick=()=>lb.classList.remove('on');
document.addEventListener('keydown',e=>{{if(e.key==='Escape')lb.classList.remove('on')}});

// ── Step counter + renumber ──
function updateCount(){{
  document.getElementById('step-count').textContent=document.querySelectorAll('.card').length;
}}
function renumber(){{
  document.querySelectorAll('.card').forEach((card,i)=>{{
    card.querySelector('.badge').textContent='Step '+(i+1);
  }});
  updateCount();
}}

// ── Delete ──
function deleteStep(btn){{
  const card=btn.closest('.card');
  card.style.transition='opacity .25s,max-height .35s';
  card.style.opacity='0';
  card.style.maxHeight=card.offsetHeight+'px';
  setTimeout(()=>{{card.style.maxHeight='0';card.style.marginBottom='0';card.style.overflow='hidden';}},50);
  setTimeout(()=>{{card.remove();renumber();}},380);
}}

// ── Drag & Drop ──
const container=document.getElementById('steps-container');
let dragSrc=null;

function addDragListeners(card){{
  card.addEventListener('dragstart',e=>{{
    dragSrc=card;
    setTimeout(()=>card.classList.add('dragging'),0);
    e.dataTransfer.effectAllowed='move';
  }});
  card.addEventListener('dragend',()=>{{
    card.classList.remove('dragging');
    document.querySelectorAll('.card').forEach(c=>c.classList.remove('drag-over'));
    renumber();
  }});
  card.addEventListener('dragover',e=>{{
    e.preventDefault();
    e.dataTransfer.dropEffect='move';
    if(card!==dragSrc){{
      document.querySelectorAll('.card').forEach(c=>c.classList.remove('drag-over'));
      card.classList.add('drag-over');
    }}
  }});
  card.addEventListener('drop',e=>{{
    e.preventDefault();
    if(dragSrc&&dragSrc!==card){{
      const cards=[...container.querySelectorAll('.card')];
      const fromIdx=cards.indexOf(dragSrc);
      const toIdx=cards.indexOf(card);
      if(fromIdx<toIdx)card.after(dragSrc);
      else card.before(dragSrc);
    }}
    card.classList.remove('drag-over');
  }});
}}
document.querySelectorAll('.card').forEach(addDragListeners);

// ── Save report ──
async function saveReport(){{
  const btn=document.getElementById('save-btn');
  const notice=document.getElementById('save-notice');
  btn.disabled=true;

  // Bake title
  const titleEl=document.getElementById('report-title');
  titleEl.textContent=titleEl.textContent;

  // Bake captions
  document.querySelectorAll('.desc-input').forEach(ta=>{{
    ta.innerHTML=escapeHtml(ta.value);
  }});

  const html='<!DOCTYPE html>\\n'+document.documentElement.outerHTML;

  try{{
    const handle=await window.showSaveFilePicker({{
      suggestedName:titleEl.textContent.replace(/[^a-z0-9]/gi,'_')+'.html',
      types:[{{description:'HTML File',accept:{{'text/html':['.html']}}}}]
    }});
    const w=await handle.createWritable();
    await w.write(new Blob([html],{{type:'text/html'}}));
    await w.close();
    notice.textContent='✔ Saved!';
  }}catch(e){{
    if(e.name!=='AbortError'){{
      const a=document.createElement('a');
      a.href=URL.createObjectURL(new Blob([html],{{type:'text/html'}}));
      a.download=titleEl.textContent.replace(/[^a-z0-9]/gi,'_')+'_edited.html';
      a.click();
      notice.textContent='✔ Downloaded!';
    }}
  }}

  setTimeout(()=>notice.textContent='',3000);
  btn.disabled=false;
}}

function escapeHtml(s){{
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}
</script>
</body>
</html>"""


def generate_report():
    if not steps:
        print("  Nothing to save. Exiting.")
        return

    html = build_html(steps, _title)
    safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in _title)[:48].strip()
    fname = f"steps_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe}.html".replace(" ", "_")
    out = os.path.join(OUTPUT_DIR, fname)

    with open(out, "w", encoding="utf-8") as f:
        f.write(html)

    abs_path = os.path.abspath(out)
    print(f"\n  ✔ Report saved: {abs_path}")

    try:
        import webbrowser
        webbrowser.open(f"file://{abs_path}")
        print("  Opened in your default browser.")
    except Exception:
        pass


_title = ""


def main():
    global _title, _kb_listener

    print("=" * 54)
    print("  Steps Recorder  —  MS Steps Recorder Replacement")
    print("=" * 54)
    print()
    print("  • F9          — start / stop recording")
    print("  • Screenshots taken on every mouse click")
    print("  • Debounce: double-clicks ignored (500 ms window)")
    print("  • Edit title, captions, reorder & delete in report")
    print()

    _title = input("  Session title (or ENTER for default): ").strip()
    if not _title:
        _title = f"Steps Recording - {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    print()
    print("  Press F9 to START recording.")
    print("  Press F9 again to STOP.")
    print("  Press ESC to finish and generate the report.")
    print()

    done = threading.Event()

    def on_key(key):
        if key == TOGGLE_KEY:
            on_key_press(key)
        elif key == keyboard.Key.esc:
            if recording:
                on_key_press(TOGGLE_KEY)  # stop recording first
            done.set()
            return False  # stop listener

    _kb_listener = keyboard.Listener(on_press=on_key)
    _kb_listener.start()

    done.wait()  # block until ESC pressed

    print(f"\n  Finalising — {len(steps)} step(s) captured.")
    generate_report()


if __name__ == "__main__":
    main()