// ====== Globals ======
const R = 6371000;
const el = id => document.getElementById(id);
const deg2rad = d => d*Math.PI/180;
const rad2deg = r => r*180/Math.PI;

// SheetJS workbook kept in memory so we can export it later
let WB = null;
let xlsxSheetNames = [];
let ballisticCSV = [];      // 2D array for the table preview
let fanGeoJSON = null;

// Leaflet (Step 1 FP picker map)
let fpMap, fpMarker, fpAccCircle;
let pickMode = false;

// ====== Geometry helpers ======
function brgRangeToXY(bearingDeg, rangeM){
  const th = deg2rad(bearingDeg);
  return { dx: rangeM * Math.sin(th), dy: rangeM * Math.cos(th) };
}
function xyToLL(lat0, lon0, dx, dy){
  const dLat = dy / R;
  const dLon = dx / (R * Math.cos(deg2rad(lat0)));
  return { lat: lat0 + rad2deg(dLat), lon: lon0 + rad2deg(dLon) };
}

// ====== UI utils ======
function download(name, content, mime){
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name; a.click();
  URL.revokeObjectURL(url);
}
function setStatus(msg, level=''){
  el('status').textContent = msg;
  el('status').className = `pill ${level||''}`;
}
const toast = el('toast');
function showToast(msg){
  toast.textContent = msg;
  toast.style.display = 'block';
  clearTimeout(showToast._t);
  showToast._t = setTimeout(()=>toast.style.display='none', 2000);
}

// ====== Table preview ======
function toTable(container, data){
  if(!data?.length){
    container.innerHTML = '<span class="hint">No data.</span>';
    return;
  }
  const [hdr, ...rows] = data;
  const f = (el('search')?.value || '').toLowerCase();
  const rowsF = rows.filter(r => !f || r.some(c => String(c).toLowerCase().includes(f)));

  container.innerHTML = [
    '<table><thead><tr>',
    ...hdr.map(h => `<th>${h}</th>`),
    '</tr></thead><tbody>',
    ...rowsF.map(r => `<tr>${r.map(c => `<td>${c ?? ''}</td>`).join('')}</tr>`),
    '</tbody></table>'
  ].join('');
}

function sheetTo2D(ws){
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:'' });
  const maxCols = Math.max(...rows.map(r => r.length), 0);
  return rows.map(r => {
    const row = Array.from({length:maxCols}, (_,i)=> r[i] ?? '');
    for(let i=row.length-1;i>=0;i--){ if(row[i]==='') row.pop(); else break; }
    return row;
  }).filter(r => r.length);
}

// ====== Auto-load workbook at startup ======
async function loadWorkbookFromAppFolder(){
  try{
    const res = await fetch('Safety Fan Calculator.xlsx');
    if(!res.ok) throw new Error('Workbook not found in app folder');
    const buf = await res.arrayBuffer();
    WB = XLSX.read(buf, { type:'array' });
    xlsxSheetNames = WB.SheetNames.slice();

    // Build weapon list from sheet names like "1 HE", "2 IM"
    hydrateWeaponListFromSheets();

    // Table preview uses SAFETY FAN DATA if present; else first sheet
    const previewName = WB.Sheets['SAFETY FAN DATA'] ? 'SAFETY FAN DATA' : WB.SheetNames[0];
    ballisticCSV = sheetTo2D(WB.Sheets[previewName]);
    toTable(el('tableWrap'), ballisticCSV);

    setStatus(`Loaded workbook: ${previewName}`, 'ok');
  }catch(err){
    setStatus(`Failed to auto-load workbook: ${err.message}`, 'bad');
  }
}

// ====== Weapon list from XLSX sheet names ======
function hydrateWeaponListFromSheets(){
  const sel = el('weapon');
  if(!sel) return;

  const hasAuto = !!sel.querySelector('option[value=""]');
  sel.innerHTML = hasAuto ? `<option value="">Auto (use Nature + Charge)</option>` : '';

  const profiles = [];
  for(let i=1;i<=7;i++) if(xlsxSheetNames.includes(`${i} HE`)) profiles.push(`HE C${i}`);
  for(let i=1;i<=7;i++) if(xlsxSheetNames.includes(`${i} IM`)) profiles.push(`IM C${i}`);
  sel.innerHTML += profiles.map(p => `<option>${p}</option>`).join('');

  pickDefaultWeapon();
}
function pickDefaultWeapon(){
  const sel = el('weapon'); if(!sel) return;
  const nat = (el('nature')?.value || 'HE').toUpperCase();
  const ch  = (el('charge')?.value || '1');
  const target = `${nat} C${ch}`;
  if(!sel.value){
    const match = [...sel.options].find(o => o.value === target || o.textContent === target);
    if (match) sel.value = match.value;
    else {
      const firstReal = [...sel.options].find(o => o.value);
      if (firstReal) sel.value = firstReal.value;
    }
  }
}
['nature','charge'].forEach(id=>{
  el(id)?.addEventListener('change', ()=>{
    const sel = el('weapon');
    if(sel && !sel.value) pickDefaultWeapon();
  });
});

// ====== MET (placeholder) ======
function metCorrections(az, windDir, windSpd, metScale){
  const rel = ((windDir - az + 540) % 360) - 180;
  const headTail = Math.cos(deg2rad(rel));
  const cross = Math.sin(deg2rad(rel));
  return {
    dRange:   metScale * windSpd * 10 * headTail, // meters
    dBearing: metScale * cross * 2                // degrees
  };
}

// ====== Fan builder (adaptive resolution) ======
const ARC_RES_DEG = 1.5;     // deg per segment on arcs
const RADIAL_SPACING_M = 75; // m per segment on radials

function buildFan(lat0, lon0, az, leftOff, rightOff, maxR, baseRange, dRange, dBearing){
  const left  = az - leftOff + dBearing;
  const right = az + rightOff + dBearing;

  const innerR = Math.max(0, (baseRange || 0) + dRange);
  const outerR = Math.max(innerR, maxR + dRange);

  const arcSpan = Math.max(1e-6, Math.abs(right - left));
  const stepsArc = Math.max(8, Math.ceil(arcSpan / ARC_RES_DEG));

  const radialSpan = Math.max(0, outerR - innerR);
  const stepsRadial = Math.max(2, Math.ceil(radialSpan / RADIAL_SPACING_M));

  const pts = [];

  // left radial (inner -> outer)
  for(let i=0;i<=stepsRadial;i++){
    const r = innerR + (radialSpan * i / stepsRadial);
    const {dx,dy}=brgRangeToXY(left, r); const {lat,lon}=xyToLL(lat0,lon0,dx,dy); pts.push([lon,lat]);
  }
  // outer arc (left -> right)
  for(let i=0;i<=stepsArc;i++){
    const b = left + (arcSpan * i / stepsArc);
    const {dx,dy}=brgRangeToXY(b, outerR); const {lat,lon}=xyToLL(lat0,lon0,dx,dy); pts.push([lon,lat]);
  }
  // right radial (outer -> inner)
  for(let i=0;i<=stepsRadial;i++){
    const r = outerR - (radialSpan * i / stepsRadial);
    const {dx,dy}=brgRangeToXY(right, r); const {lat,lon}=xyToLL(lat0,lon0,dx,dy); pts.push([lon,lat]);
  }
  // inner arc (right -> left)
  for(let i=0;i<=stepsArc;i++){
    const b = right - (arcSpan * i / stepsArc);
    const {dx,dy}=brgRangeToXY(b, innerR); const {lat,lon}=xyToLL(lat0,lon0,dx,dy); pts.push([lon,lat]);
  }

  if (pts.length) pts.push(pts[0]); // close polygon
  return pts;
}
function featureFromFan(coords){
  return { type:"Feature", properties:{ name:"Safety Fan" }, geometry:{ type:"Polygon", coordinates:[coords] } };
}

// ====== Wizard UI ======
let step = 0;
const steps = [...document.querySelectorAll('.step')];
const btnBack = el('btnBack');
const btnNext = el('btnNext');
const dots = el('progressDots');

function renderDots(){
  dots.innerHTML = steps.map((_,i)=>`<span class="dot ${i===step?'active':''}"></span>`).join('');
}
function showStep(i){
  step = Math.max(0, Math.min(steps.length-1, i));
  steps.forEach((s,idx)=>s.style.display = idx===step ? 'block' : 'none');
  btnBack.disabled = step===0;
  btnNext.style.display = step<steps.length-1 ? 'inline-block' : 'none';
  el('btnCompute').style.display = step===steps.length-1 ? 'inline-block' : 'none';
  renderDots();
  if (step === 0) initFpMap(); // show picker when on step 1
}
function buildReview(){
  const pairs = [
    ['Lat', lat.value],['Lon', lon.value],['Az', az.value],
    ['Left', leftOff.value],['Right', rightOff.value],['MaxR', fanDepth.value],
    ['Weapon', weapon.value || `(auto: ${nature.value.toUpperCase()} C${charge.value})`],
    ['Nature', nature.value],['Charge', charge.value],['Mode', mode.value],
    ['NomRange', range.value],['QE', qe.value],
    ['Temp', temp.value],['Press', press.value],
    ['WindFrom', windDir.value],['WindSpd', windSpd.value],['MetScale', metScale.value]
  ];
  el('review').innerHTML =
    `<table><tbody>${pairs.map(([k,v])=>`<tr><th style="text-align:left">${k}</th><td>${v}</td></tr>`).join('')}</tbody></table>`;
}
btnBack.onclick = ()=>showStep(step-1);
btnNext.onclick = ()=>{
  if(step===0 && (!lat.value || !lon.value)) return showToast('Set Lat/Lon first');
  if(step===1 && !weapon.value) pickDefaultWeapon();
  if(step===2) buildReview();
  showStep(step+1);
};

// ====== Step 1 Map Picker ======
const fpIcon = L.icon({
  iconUrl: 'data:image/svg+xml;utf8,' + encodeURIComponent(`
  <svg xmlns="http://www.w3.org/2000/svg" width="32" height="48" viewBox="0 0 32 48">
    <defs><filter id="s" x="-50%" y="-50%" width="200%" height="200%">
      <feDropShadow dx="0" dy="1.5" stdDeviation="1.5" flood-opacity=".4"/>
    </filter></defs>
    <path d="M16 47c0 0 12-16.4 12-26A12 12 0 0 0 16 9 12 12 0 0 0 4 21c0 9.6 12 26 12 26z" fill="#e74c3c" filter="url(#s)"/>
    <circle cx="16" cy="21" r="5.5" fill="#fff"/>
    <circle cx="16" cy="21" r="3.2" fill="#e74c3c"/>
  </svg>`),
  iconSize: [32,48],
  iconAnchor: [16,47],
  popupAnchor: [0,-40],
});
function initFpMap(){
  if (fpMap) return;
  const lat0 = parseFloat(el('lat').value) || 0;
  const lon0 = parseFloat(el('lon').value) || 0;

  fpMap = L.map('fpMap').setView([lat0, lon0], 12);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
    { attribution: '© OpenStreetMap' }).addTo(fpMap);

  fpMarker = L.marker([lat0, lon0], { draggable: true, icon: fpIcon }).addTo(fpMap)
    .bindPopup('Firing Point');

  fpMarker.on('dragend', () => {
    const p = fpMarker.getLatLng();
    el('lat').value = p.lat.toFixed(6);
    el('lon').value = p.lng.toFixed(6);
    fpMarker.setPopupContent(`Firing Point<br>${p.lat.toFixed(6)}, ${p.lng.toFixed(6)}`);
  });

  fpMap.on('click', (e) => {
    if (!pickMode) return;
    setFiringPoint(e.latlng.lat, e.latlng.lng, true);
    togglePickMode(false);
  });
}
function setFiringPoint(lat, lon, fly=false){
  el('lat').value = (+lat).toFixed(6);
  el('lon').value = (+lon).toFixed(6);
  if (fpMarker) {
    fpMarker.setLatLng([lat, lon]);
    fpMarker.setPopupContent(`Firing Point<br>${(+lat).toFixed(6)}, ${(+lon).toFixed(6)}`);
  }
  if (fpMap && fly) fpMap.flyTo([lat, lon], Math.max(12, fpMap.getZoom()));
}
function previewInputsOnMap(){
  const lat = parseFloat(el('lat').value);
  const lon = parseFloat(el('lon').value);
  if (Number.isNaN(lat) || Number.isNaN(lon)) return showToast('Enter a valid Lat/Lon first');
  setFiringPoint(lat, lon, true);
  if (fpMarker) fpMarker.openPopup();
}
function togglePickMode(on){
  pickMode = on;
  if (!fpMap) return;
  fpMap.getContainer().style.cursor = on ? 'crosshair' : '';
  showToast(on ? 'Tap the map to place the pin' : 'Pick mode off');
}
// Keep marker in sync when typing
['lat','lon'].forEach(id=>{
  el(id)?.addEventListener('change', ()=>{
    const lat = parseFloat(el('lat').value);
    const lon = parseFloat(el('lon').value);
    if (!Number.isNaN(lat) && !Number.isNaN(lon)) setFiringPoint(lat, lon);
  });
});
// Buttons
el('btnPickOnMap')?.addEventListener('click', ()=>{ initFpMap(); togglePickMode(!pickMode); });
el('btnPreviewOnMap')?.addEventListener('click', ()=>{ initFpMap(); previewInputsOnMap(); });
el('btnUseMyLocation')?.addEventListener('click', ()=>{
  initFpMap();
  if (!navigator.geolocation) return showToast('Geolocation not supported');
  navigator.geolocation.getCurrentPosition(
    (pos) => {
      const { latitude, longitude, accuracy } = pos.coords;
      setFiringPoint(latitude, longitude, true);
      if (fpAccCircle) fpMap.removeLayer(fpAccCircle);
      fpAccCircle = L.circle([latitude, longitude], { radius: accuracy || 30, color:'#4aa3', fillColor:'#4aa3', fillOpacity:0.2 });
      fpAccCircle.addTo(fpMap);
      showToast(`Set from device (±${Math.round(accuracy||0)} m)`);
    },
    (err) => showToast(`Location failed: ${err.code}`),
    { enableHighAccuracy: true, timeout: 8000, maximumAge: 0 }
  );
});

// ====== Compute & export ======
el('btnCompute').onclick = ()=>{
  const lat = parseFloat(el('lat').value);
  const lon = parseFloat(el('lon').value);
  const az = parseFloat(el('az').value);
  const leftOff = parseFloat(el('leftOff').value);
  const rightOff = parseFloat(el('rightOff').value);
  const maxR = parseFloat(el('fanDepth').value);
  const baseRange = parseFloat(el('range').value) || 0;

  const windDir = parseFloat(el('windDir').value);
  const windSpd = parseFloat(el('windSpd').value);
  const metScale = parseFloat(el('metScale').value);

  const { dRange, dBearing } = metCorrections(az, windDir, windSpd, metScale);
  el('metPreview').value = `ΔRange: ${dRange.toFixed(1)} m, ΔBearing: ${dBearing.toFixed(2)}°`;

  const poly = buildFan(lat, lon, az, leftOff, rightOff, maxR, baseRange, dRange, dBearing);
  fanGeoJSON = { type:"FeatureCollection", features:[ featureFromFan(poly) ] };
  el('debug').textContent = JSON.stringify(fanGeoJSON, null, 2);
  setStatus(`Computed fan: ${poly.length} pts`, 'ok');
};

el('btnExportGeoJSON').onclick = ()=>{
  if(!fanGeoJSON) return setStatus('Compute first','bad');
  download('safety_fan.geojson', JSON.stringify(fanGeoJSON), 'application/geo+json');
};
el('btnExportKML').onclick = ()=>{
  if(!fanGeoJSON) return setStatus('Compute first','bad');
  const coords = fanGeoJSON.features[0].geometry.coordinates[0]
    .map(([lon,lat]) => `${lon},${lat},0`).join(' ');
  const kml = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>Safety Fan</name><Placemark><name>Safety Fan</name>
<Style><LineStyle><width>2</width></LineStyle><PolyStyle><fill>1</fill><outline>1</outline></PolyStyle></Style>
<Polygon><outerBoundaryIs><LinearRing><coordinates>${coords}</coordinates></LinearRing></outerBoundaryIs></Polygon>
</Placemark></Document></kml>`;
  download('safety_fan.kml', kml, 'application/vnd.google-earth.kml+xml');
};

// Export the updated workbook with an APP_EXPORT sheet
el('btnSaveXLSX').onclick = ()=>{
  if(!WB){ return setStatus('No workbook loaded','bad'); }
  const rows = [
    ['OPTECH GUNNER APP export'],
    ['Datetime (UTC)', new Date().toISOString()],
    [],
    ['Lat', el('lat').value], ['Lon', el('lon').value], ['Az', el('az').value],
    ['Left', el('leftOff').value], ['Right', el('rightOff').value], ['MaxR', el('fanDepth').value],
    ['Weapon', el('weapon').value || `(auto: ${el('nature').value.toUpperCase()} C${el('charge').value})`],
    ['Nature', el('nature').value], ['Charge', el('charge').value], ['Mode', el('mode').value],
    ['NomRange', el('range').value], ['QE', el('qe').value],
    ['Temp', el('temp').value], ['Press', el('press').value],
    ['WindFrom', el('windDir').value], ['WindSpd', el('windSpd').value], ['MetScale', el('metScale').value],
    [],
    ['Computed?', fanGeoJSON ? 'yes' : 'no'],
    ['Fan points', fanGeoJSON ? fanGeoJSON.features[0].geometry.coordinates[0].length : 0]
  ];
  if (fanGeoJSON){
    rows.push([], ['lon','lat']);
    fanGeoJSON.features[0].geometry.coordinates[0].slice(0,20).forEach(([lo,la])=>rows.push([lo,la]));
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  if(WB.Sheets['APP_EXPORT']) delete WB.Sheets['APP_EXPORT'];
  WB.Sheets['APP_EXPORT'] = ws;
  if(!WB.SheetNames.includes('APP_EXPORT')) WB.SheetNames.push('APP_EXPORT');
  XLSX.writeFile(WB, 'Safety Fan Calculator.updated.xlsx');
  setStatus('Exported XLSX with APP_EXPORT sheet', 'ok');
};

// Manual reload button
el('btnReload').onclick = ()=> loadWorkbookFromAppFolder();

// Search filter
el('search')?.addEventListener('input', ()=>toTable(el('tableWrap'), ballisticCSV));

// Weather sync (Open-Meteo; uses current Lat/Lon from inputs / map pin)
el('btnWeather')?.addEventListener('click', async ()=>{
  try{
    const lat = parseFloat(el('lat').value), lon = parseFloat(el('lon').value);
    if(Number.isNaN(lat) || Number.isNaN(lon)) return showToast('Set a valid Lat/Lon first');
    // Snap marker to inputs (no fly)
    if (typeof setFiringPoint === 'function') setFiringPoint(lat, lon, false);

    const url = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}` +
                `&current=temperature_2m,wind_speed_10m,wind_direction_10m,pressure_msl`;
    setStatus('Fetching weather…');
    const res = await fetch(url);
    if(!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    const c = data.current || {};
    if(c.temperature_2m!=null) el('temp').value = c.temperature_2m;
    if(c.pressure_msl!=null)   el('press').value = c.pressure_msl;
    if(c.wind_speed_10m!=null) el('windSpd').value = c.wind_speed_10m;
    if(c.wind_direction_10m!=null) el('windDir').value = c.wind_direction_10m;
    setStatus(`Weather synced @ ${lat.toFixed(4)}, ${lon.toFixed(4)}`, 'ok');
    showToast('Weather synced');
  }catch(e){
    setStatus(`Weather sync failed: ${e.message}`, 'bad');
  }
});

// Wizard init + auto-load workbook
showStep(0);
renderDots();
window.addEventListener('beforeunload', e => { if(!fanGeoJSON) return; e.preventDefault(); e.returnValue = ''; });
loadWorkbookFromAppFolder();
