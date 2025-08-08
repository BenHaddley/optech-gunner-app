// ====== Globals ======
const R = 6371000;
const el = id => document.getElementById(id);
const deg2rad = d => d*Math.PI/180;
const rad2deg = r => r*180/Math.PI;

let WB = null;
let xlsxSheetNames = [];
let ballisticCSV = [];
let fanGeoJSON = null;

// Trajectory mode for results page
let mode = 'LA'; // or 'HA'

// Leaflet FP picker
let fpMap, fpMarker, fpAccCircle;
let pickMode = false;

// ====== Geometry ======
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

// ====== Table helpers ======
function render2DTable(container, data){
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
function pointsToTable(container, coords){
  if(!coords?.length){
    container.innerHTML = '<span class="hint">No points.</span>';
    return;
  }
  const head = ['lon','lat'];
  const body = coords.map(([lo,la]) => [lo,la]);
  container.innerHTML = [
    '<table><thead><tr>',
    ...head.map(h=>`<th>${h}</th>`),
    '</tr></thead><tbody>',
    ...body.map(r=>`<tr>${r.map(c=>`<td>${(+c).toFixed(6)}</td>`).join('')}</tr>`),
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

// ====== XLSX load & STAPD ======
async function loadWorkbookFromAppFolder(){
  try{
    const res = await fetch('Safety Fan Calculator.xlsx');
    if(!res.ok) throw new Error('Workbook not found in app folder');
    const buf = await res.arrayBuffer();
    WB = XLSX.read(buf, { type:'array' });
    xlsxSheetNames = WB.SheetNames.slice();

    hydrateWeaponList();

    const previewName = WB.Sheets['SAFETY FAN DATA'] ? 'SAFETY FAN DATA' : WB.SheetNames[0];
    ballisticCSV = sheetTo2D(WB.Sheets[previewName]);
    render2DTable(el('tableWrap'), ballisticCSV);

    setStatus(`Loaded workbook: ${previewName}`, 'ok');
  }catch(err){
    setStatus(`Failed to auto-load workbook: ${err.message}`, 'bad');
  }
}
function hydrateWeaponList(){
  const sel = el('weapon');
  if(!sel) return;

  const hasAuto = !!sel.querySelector('option[value=""]');
  sel.innerHTML = hasAuto ? `<option value="">Auto (use Nature + Charge)</option>` : '';

  let profiles = [];
  const st = WB?.Sheets?.['STAPD'];
  if (st){
    const data = sheetTo2D(st);
    const head = (data[0] || []).map(h=>String(h).toLowerCase());
    const idxName = head.indexOf('name');
    const idxNature = head.indexOf('nature');
    const idxCharge = head.indexOf('charge');
    for (let i=1;i<data.length;i++){
      const row = data[i];
      const n = row[idxName] || '';
      const nat = (row[idxNature]||'').toString().toUpperCase();
      const ch = (row[idxCharge]||'').toString();
      if (n || (nat && ch)) profiles.push(n || `${nat} C${ch}`);
    }
  }
  if (!profiles.length && xlsxSheetNames.length){
    for(let i=1;i<=7;i++) if(xlsxSheetNames.includes(`${i} HE`)) profiles.push(`HE C${i}`);
    for(let i=1;i<=7;i++) if(xlsxSheetNames.includes(`${i} IM`)) profiles.push(`IM C${i}`);
  }
  profiles = [...new Set(profiles)];
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

// ====== Met corrections (placeholder) ======
function metCorrections(az, windDir, windSpd, metScale){
  const rel = ((windDir - az + 540) % 360) - 180;
  const headTail = Math.cos(deg2rad(rel));
  const cross = Math.sin(deg2rad(rel));
  return {
    dRange:   metScale * windSpd * 10 * headTail,
    dBearing: metScale * cross * 2
  };
}

// ====== Fan builder with dynamic resolution ======
function dynamicResolution(maxR, leftOff, rightOff){
  const depth = Math.max(1, +maxR || 1000);
  const widthDeg = Math.max(0.1, (+leftOff||0) + (+rightOff||0));
  const arcResDeg = Math.max(0.5, Math.min(3, 60 / Math.sqrt(widthDeg + 1)));
  const radialSpacing = Math.max(10, Math.min(200, depth/60));
  return { arcResDeg, radialSpacing };
}
function buildFan(lat0, lon0, az, leftOff, rightOff, maxR, baseRange, dRange, dBearing){
  const { arcResDeg, radialSpacing } = dynamicResolution(maxR, leftOff, rightOff);

  const left  = az - leftOff + dBearing;
  const right = az + rightOff + dBearing;

  const innerR = Math.max(0, (baseRange || 0) + dRange);
  const outerR = Math.max(innerR, maxR + dRange);

  const arcSpan = Math.max(1e-6, Math.abs(right - left));
  const stepsArc = Math.max(8, Math.ceil(arcSpan / arcResDeg));

  const radialSpan = Math.max(0, outerR - innerR);
  const stepsRadial = Math.max(2, Math.ceil(radialSpan / radialSpacing));

  const pts = [];
  for(let i=0;i<=stepsRadial;i++){
    const r = innerR + (radialSpan * i / stepsRadial);
    const {dx,dy}=brgRangeToXY(left, r); const {lat,lon}=xyToLL(lat0,lon0,dx,dy); pts.push([lon,lat]);
  }
  for(let i=0;i<=stepsArc;i++){
    const b = left + (arcSpan * i / stepsArc);
    const {dx,dy}=brgRangeToXY(b, outerR); const {lat,lon}=xyToLL(lat0,lon0,dx,dy); pts.push([lon,lat]);
  }
  for(let i=0;i<=stepsRadial;i++){
    const r = outerR - (radialSpan * i / stepsRadial);
    const {dx,dy}=brgRangeToXY(right, r); const {lat,lon}=xyToLL(lat0,lon0,dx,dy); pts.push([lon,lat]);
  }
  for(let i=0;i<=stepsArc;i++){
    const b = right - (arcSpan * i / stepsArc);
    const {dx,dy}=brgRangeToXY(b, innerR); const {lat,lon}=xyToLL(lat0,lon0,dx,dy); pts.push([lon,lat]);
  }

  if (pts.length) pts.push(pts[0]);
  return pts;
}
function featureFromFan(coords){
  return { type:"Feature", properties:{ name:"Safety Fan", mode }, geometry:{ type:"Polygon", coordinates:[coords] } };
}

// ====== Page flow ======
const page1 = document.getElementById('page1');
const page2 = document.getElementById('page2');
function goPage(id){
  page1.style.display = id==='page1' ? 'block' : 'none';
  page2.style.display = id==='page2' ? 'block' : 'none';
}
el('goResults').addEventListener('click', ()=>{
  if(!el('lat').value || !el('lon').value) return showToast('Enter BC grid (Lat/Lon) first');
  goPage('page2');
  buildReview();
});
el('btnBackToInputs').addEventListener('click', ()=> goPage('page1'));
el('btnToggleTraj').addEventListener('click', ()=>{
  mode = (mode === 'LA' ? 'HA' : 'LA');
  showToast(`Trajectory: ${mode}`);
  buildReview();
});

// ====== Review (page 2) ======
function buildReview(){
  const pairs = [
    ['Lat', el('lat').value],['Lon', el('lon').value],['Az', el('az').value],
    ['Left', el('leftOff').value],['Right', el('rightOff').value],['MaxR', el('fanDepth').value],
    ['Weapon', el('weapon').value || `(auto: ${el('nature').value.toUpperCase()} C${el('charge').value})`],
    ['Nature', el('nature').value],['Charge', el('charge').value],['Mode', mode],
    ['NomRange', el('range').value],['QE', el('qe').value],
    ['Temp', el('temp').value],['Press', el('press').value],
    ['WindFrom', el('windDir').value],['WindSpd', el('windSpd').value],['MetScale', el('metScale').value]
  ];
  el('review').innerHTML =
    `<table><tbody>${pairs.map(([k,v])=>`<tr><th style="text-align:left">${k}</th><td>${v}</td></tr>`).join('')}</tbody></table>`;
}

// ====== Compute & export ======
el('btnCompute').addEventListener('click', ()=>{
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
  pointsToTable(el('pointsWrap'), poly);

  setStatus(`Computed fan: ${poly.length} pts (mode: ${mode})`, 'ok');
});

el('btnExportGeoJSON').addEventListener('click', ()=>{
  if(!fanGeoJSON) return setStatus('Compute first','bad');
  download('safety_fan.geojson', JSON.stringify(fanGeoJSON), 'application/geo+json');
});
el('btnExportKML').addEventListener('click', ()=>{
  if(!fanGeoJSON) return setStatus('Compute first','bad');
  const coords = fanGeoJSON.features[0].geometry.coordinates[0]
    .map(([lon,lat]) => `${lon},${lat},0`).join(' ');
  const kml = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.org/kml/2.2"><Document><name>Safety Fan</name><Placemark><name>Safety Fan</name>
<Style><LineStyle><width>2</width></LineStyle><PolyStyle><fill>1</fill><outline>1</outline></PolyStyle></Style>
<Polygon><outerBoundaryIs><LinearRing><coordinates>${coords}</coordinates></LinearRing></outerBoundaryIs></Polygon>
</Placemark></Document></kml>`;
  download('safety_fan.kml', kml, 'application/vnd.google-earth.kml+xml');
});

el('btnSaveXLSX').addEventListener('click', ()=>{
  if(!WB){ return setStatus('No workbook loaded','bad'); }
  const rows = [
    ['OPTECH GUNNER APP export'],
    ['Datetime (UTC)', new Date().toISOString()],
    [],
    ['Lat', el('lat').value], ['Lon', el('lon').value], ['Az', el('az').value],
    ['Left', el('leftOff').value], ['Right', el('rightOff').value], ['MaxR', el('fanDepth').value],
    ['Weapon', el('weapon').value || `(auto: ${el('nature').value.toUpperCase()} C${el('charge').value})`],
    ['Nature', el('nature').value], ['Charge', el('charge').value], ['Mode', mode],
    ['NomRange', el('range').value], ['QE', el('qe').value],
    ['Temp', el('temp').value], ['Press', el('press').value],
    ['WindFrom', el('windDir').value], ['WindSpd', el('windSpd').value], ['MetScale', el('metScale').value],
    [],
    ['Computed?', fanGeoJSON ? 'yes' : 'no'],
    ['Fan points', fanGeoJSON ? fanGeoJSON.features[0].geometry.coordinates[0].length : 0]
  ];
  if (fanGeoJSON){
    rows.push([], ['lon','lat']);
    fanGeoJSON.features[0].geometry.coordinates[0].slice(0,50).forEach(([lo,la])=>rows.push([lo,la]));
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  if(WB.Sheets['APP_EXPORT']) delete WB.Sheets['APP_EXPORT'];
  WB.Sheets['APP_EXPORT'] = ws;
  if(!WB.SheetNames.includes('APP_EXPORT')) WB.SheetNames.push('APP_EXPORT');
  XLSX.writeFile(WB, 'Safety Fan Calculator.updated.xlsx');
  setStatus('Exported XLSX with APP_EXPORT sheet', 'ok');
});

// ====== FP picker (Leaflet) ======
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
  if (Number.isNaN(lat) || Number.isNaN(lon)) {
    showToast('Enter a valid Lat/Lon first');
    return;
  }
  setFiringPoint(lat, lon, true);
  if (fpMarker) fpMarker.openPopup();
}
function togglePickMode(on){
  pickMode = on;
  if (!fpMap) return;
  fpMap.getContainer().style.cursor = on ? 'crosshair' : '';
  showToast(on ? 'Tap the map to place the pin' : 'Pick mode off');
}

// Buttons for map picker
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

// Keep inputs in sync with marker if typed manually
['lat','lon'].forEach(id=>{
  el(id)?.addEventListener('change', ()=>{
    const lat = parseFloat(el('lat').value);
    const lon = parseFloat(el('lon').value);
    if (!Number.isNaN(lat) && !Number.isNaN(lon)) { initFpMap(); setFiringPoint(lat, lon); }
  });
});

// ====== Weather sync (uses current Lat/Lon on Page 1) ======
el('btnWeather')?.addEventListener('click', async ()=>{
  try{
    const lat = parseFloat(el('lat').value), lon = parseFloat(el('lon').value);
    if(Number.isNaN(lat) || Number.isNaN(lon)) return showToast('Set valid Lat/Lon first');
    const url = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&current=temperature_2m,wind_speed_10m,wind_direction_10m,pressure_msl`;
    const res = await fetch(url);
    if(!res.ok) throw new Error('Weather fetch failed');
    const data = await res.json();
    const c = data.current || {};
    if(c.temperature_2m!=null) el('temp').value = c.temperature_2m;          // °C
    if(c.pressure_msl!=null)   el('press').value = c.pressure_msl;            // hPa
    if(c.wind_speed_10m!=null) el('windSpd').value = c.wind_speed_10m;        // m/s
    if(c.wind_direction_10m!=null) el('windDir').value = c.wind_direction_10m;// deg
    showToast('Weather synced from current Lat/Lon');
  }catch(e){ setStatus('Weather sync failed', 'bad'); }
});

// ====== Other wiring ======
el('btnReload').addEventListener('click', loadWorkbookFromAppFolder);
el('search')?.addEventListener('input', ()=>render2DTable(el('tableWrap'), ballisticCSV));

// Boot
window.addEventListener('load', initFpMap);
window.addEventListener('beforeunload', e => { if(!fanGeoJSON) return; e.preventDefault(); e.returnValue = ''; });
loadWorkbookFromAppFolder();
