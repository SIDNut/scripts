// SIDNut/scripts/servicenow/printAssetOffboardLabel.js
function printAssetOffboardLabel() {
  // --- Pull values directly from the form ---
  var lastAssigned = g_form.getDisplayValue('assigned_to') || '';
  var model        = g_form.getDisplayValue('model') || '';
  var serial       = g_form.getValue('serial_number') || '';

  // --- Compute quarantine expiry (+14 days) ---
  var d = new Date(); d.setDate(d.getDate() + 14);
  var qTo = d.toISOString().slice(0,10);

  // --- Escape for HTML safety ---
  var esc = s => String(s || '').replace(/[&<>"']/g,
      m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));

  // --- Build printable HTML (62 mm continuous tape) ---
  var html = `
<!doctype html><html><head><meta charset="utf-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap');
  @page { size:62mm auto; margin:0; }
  html,body{margin:0;padding:0;}
  body{
    width:62mm;
    font-family:'Montserrat',sans-serif;
    font-size:10px;
    color:#000;
  }
  .label{padding:6mm 4mm 4mm 4mm;}
  .h1{font-size:12pt;font-weight:700;margin-bottom:2mm;line-height:1.1;}
  .line{font-size:10pt;margin:1.4mm 0;white-space:pre-wrap;}
  .tiny{font-size:8pt;}
  hr{border:0;border-top:1px solid #000;margin:2mm 0;}
</style></head>
<body>
  <div class="label">
    <div class="tiny">Peel and stick to device here</div><br><hr>
    <div class="h1">Ex: ${esc(lastAssigned)}</div>
    <div class="line">Model: ${esc(model)}</div>
    <div class="line">Quarantine to: ${qTo}</div>
    <div class="line">Removed:SCCM☐/AD☐/IPAM☐</div>
    <div class="line">Wiped: 2025-___-___</div>
    <div class="line">Class: A  B  C  D</div>
    <div class="tiny" style="margin-top:2mm;">SN: ${esc(serial)}</div>
  </div>
  <script>
    window.onload = () => setTimeout(() => window.print(), 250);
  <\/script>
</body></html>`;

  // --- Open print window ---
  var w = window.open('', '_blank', 'width=600,height=800');
  w.document.open();
  w.document.write(html);
  w.document.close();
}

// Auto-run immediately when injected
printAssetOffboardLabel();
