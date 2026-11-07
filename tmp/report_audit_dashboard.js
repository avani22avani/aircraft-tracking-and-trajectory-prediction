
    const map=L.map("map",{zoomControl:false}).setView([22.2,79.2],5);
    const mapStatus=document.getElementById("map-status");
    const tileLayer=L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{maxZoom:18,attribution:"&copy; OpenStreetMap contributors"}).addTo(map);
    let tileErrors=0;
    const setOfflineMap=(offline)=>{mapStatus.style.display=offline?"block":"none"};
    tileLayer.on("tileload",()=>{tileErrors=0;setOfflineMap(false)});
    tileLayer.on("tileerror",()=>{tileErrors+=1;if(tileErrors>=2)setOfflineMap(true)});
    window.addEventListener("offline",()=>setOfflineMap(true));
    window.addEventListener("online",()=>{tileErrors=0;tileLayer.redraw()});
    if(!navigator.onLine)setOfflineMap(true);
    const state={layers:new Map(),rows:[],sel:null,route:null,routeEnds:[],hist:new Map(),info:new Map(),routeInfo:new Map()};
    const ICONS={jet:"/static/airplane.svg",prop:"/static/propplane.svg",helicopter:"/static/helicopter.svg",default:"/static/airplane.svg"};
    const iconFor=(category)=>L.icon({iconUrl:ICONS[category]||ICONS.default,iconSize:[30,30],iconAnchor:[15,15],className:"aircraft-icon"});
    function infoFor(icao){const e=state.info.get(icao);return e&&e.status==="done"?e.data:null}
    async function ensureInfo(icao){
      if(state.info.has(icao))return;
      state.info.set(icao,{status:"pending",data:null});
      try{
        const r=await fetch(`/api/aircraft/info?icao=${encodeURIComponent(icao)}`,{cache:"no-store"});
        const p=await r.json().catch(()=>null);
        const data=(p&&p.found)?p:null;
        state.info.set(icao,{status:"done",data});
        const l=state.layers.get(icao);
        if(l&&data&&data.category)l.m.setIcon(iconFor(data.category));
        if(state.sel===icao&&l)panels(l.data);
      }catch(e){
        state.info.set(icao,{status:"error",data:null});
      }
    }
    function routeInfoFor(callsign){const e=state.routeInfo.get(callsign);return e&&e.status==="done"?e.data:null}
    async function ensureRouteInfo(callsign,icao){
      if(state.routeInfo.has(callsign))return;
      state.routeInfo.set(callsign,{status:"pending",data:null});
      try{
        const r=await fetch(`/api/route?callsign=${encodeURIComponent(callsign)}`,{cache:"no-store"});
        const p=await r.json().catch(()=>null);
        const data=(p&&p.found)?p:null;
        state.routeInfo.set(callsign,{status:"done",data});
        if(state.sel===icao){const l=state.layers.get(icao);if(l)panels(l.data)}
      }catch(e){
        state.routeInfo.set(callsign,{status:"error",data:null});
      }
    }
    const els={tracked:document.getElementById("tracked"),visible:document.getElementById("visible"),status:document.getElementById("status"),list:document.getElementById("aircraftList"),search:document.getElementById("search")};
    const chart=new Chart(document.getElementById("altitudeChart"),{type:"line",data:{labels:[],datasets:[{data:[],borderColor:"#8ecae6",backgroundColor:"rgba(142,202,230,.18)",fill:true,borderWidth:2,pointRadius:1.4,tension:.24}]},options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{color:"#9ab7d1"},grid:{color:"rgba(154,183,209,.15)"}},y:{ticks:{color:"#9ab7d1"},grid:{color:"rgba(154,183,209,.15)"}}}}});
    const num=(v,f=0)=>{const n=Number(v);return Number.isFinite(n)?n:f};
    const M_TO_FT=3.28084,MS_TO_KT=1.943844,MS_TO_FPM=196.8504;
    const norm=(r)=>{const c=String(r.callsign??r.flight_no??"UNKNOWN").trim().toUpperCase()||"UNKNOWN";const i=String(r.icaoAddress??r.icao??r.hex??c).trim().toUpperCase()||"UNKNOWN";const lat=num(r.latDD??r.lat??r.latitude,NaN),lon=num(r.lonDD??r.lon??r.longitude,NaN);let a=num(r.altitudeMM??r.altitude??r.alt??0),h=num(r.headingDE2??r.heading??r.track??0),s=num(r.horVelocityCMS??r.speed??r.groundspeed??r.gs??0);if(a<100000)a*=1000;if(h<=360)h*=100;if(s<1000)s*=100;return{icaoAddress:i,callsign:c,latDD:lat,lonDD:lon,altitudeMM:a,headingDE2:h,horVelocityCMS:s,verVelocityCMS:num(r.verVelocityCMS??0,0),squawk:num(r.squawk??0,0),sourceIp:String(r.sourceIp??"").trim(),isPredicted:Boolean(r.isPredicted),predictionAgeSec:num(r.predictionAgeSec??0,0)};};
    const pop=(a)=>`<b>${a.callsign}</b>${a.isPredicted?' <span style="color:#ff7d7d">(Predicted)</span>':''}<br>ICAO: ${a.icaoAddress}<br>ALT: ${Math.round(a.altitudeMM/1000)} m<br>SPD: ${Math.round(a.horVelocityCMS/100)} m/s<br>HDG: ${Math.round(a.headingDE2/100)} deg${a.isPredicted?`<br>PRED AGE: ${Math.round(a.predictionAgeSec)} s`:''}`;
    const haversine=(a,b)=>{const R=6371000,rad=v=>v*Math.PI/180,p1=rad(a[0]),p2=rad(b[0]),dLat=rad(b[0]-a[0]),dLon=rad(b[1]-a[1]);const h=Math.sin(dLat/2)**2+Math.cos(p1)*Math.cos(p2)*Math.sin(dLon/2)**2;return 2*R*Math.atan2(Math.sqrt(h),Math.sqrt(1-h));};
    const PLACES=[
      {name:"Delhi",lat:28.6139,lon:77.2090},{name:"Mumbai",lat:19.0760,lon:72.8777},{name:"Bengaluru",lat:12.9716,lon:77.5946},
      {name:"Hyderabad",lat:17.3850,lon:78.4867},{name:"Chennai",lat:13.0827,lon:80.2707},{name:"Kolkata",lat:22.5726,lon:88.3639},
      {name:"Ahmedabad",lat:23.0225,lon:72.5714},{name:"Pune",lat:18.5204,lon:73.8567},{name:"Jaipur",lat:26.9124,lon:75.7873},
      {name:"Lucknow",lat:26.8467,lon:80.9462},{name:"Nagpur",lat:21.1458,lon:79.0882},{name:"Goa",lat:15.2993,lon:74.1240},
      {name:"Kochi",lat:9.9312,lon:76.2673},{name:"Thiruvananthapuram",lat:8.5241,lon:76.9366},{name:"Guwahati",lat:26.1445,lon:91.7362},
      {name:"Patna",lat:25.5941,lon:85.1376},{name:"Bhopal",lat:23.2599,lon:77.4126},{name:"Srinagar",lat:34.0837,lon:74.7973},
      {name:"Chandigarh",lat:30.7333,lon:76.7794},{name:"Indore",lat:22.7196,lon:75.8577},{name:"Doha",lat:25.2854,lon:51.5310},
      {name:"Dubai",lat:25.2048,lon:55.2708},{name:"Abu Dhabi",lat:24.4539,lon:54.3773},{name:"Karachi",lat:24.8607,lon:67.0011},
      {name:"Kathmandu",lat:27.7172,lon:85.3240},{name:"Bangkok",lat:13.6900,lon:100.7501},{name:"Singapore",lat:1.3644,lon:103.9915},
      {name:"Kuala Lumpur",lat:2.7456,lon:101.7099},{name:"Colombo",lat:7.1808,lon:79.8841},{name:"Dhaka",lat:23.8103,lon:90.4125},
      {name:"Male",lat:4.1755,lon:73.5093},{name:"Hong Kong",lat:22.3080,lon:113.9185},{name:"Shanghai",lat:31.1443,lon:121.8083},
      {name:"Tokyo",lat:35.5494,lon:139.7798},{name:"Seoul",lat:37.4602,lon:126.4407},{name:"Istanbul",lat:41.2753,lon:28.7519},
      {name:"Riyadh",lat:24.9576,lon:46.6988},{name:"Jeddah",lat:21.6796,lon:39.1565},{name:"Muscat",lat:23.5933,lon:58.2844},
      {name:"London",lat:51.4700,lon:-0.4543},{name:"Paris",lat:49.0097,lon:2.5479},{name:"Frankfurt",lat:50.0379,lon:8.5622},
      {name:"Amsterdam",lat:52.3105,lon:4.7683},{name:"Moscow",lat:55.9736,lon:37.4125},{name:"Sydney",lat:-33.9399,lon:151.1753},
      {name:"Melbourne",lat:-37.6690,lon:144.8410},{name:"New York",lat:40.6413,lon:-73.7781},{name:"Toronto",lat:43.6777,lon:-79.6248}
    ];
    const nearestPlace=(lat,lon)=>{let best=null,bestD=Infinity;for(const p of PLACES){const d=haversine([lat,lon],[p.lat,p.lon]);if(d<bestD){bestD=d;best=p;}}return best?{name:best.name,distanceKm:bestD/1000}:null;};
    const projectPoint=(lat,lon,headingDeg,speedMps,sec)=>{const dist=speedMps*sec,rad=headingDeg*Math.PI/180,dLat=(dist/111320)*Math.cos(rad),scale=Math.max(.2,Math.cos(lat*Math.PI/180)),dLon=(dist/(111320*scale))*Math.sin(rad);return [lat+dLat,lon+dLon];};
    function openTab(n){document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));document.querySelectorAll("#tabs button").forEach(x=>x.classList.remove("active"));document.getElementById(n).classList.add("active");document.querySelector(`#tabs button[data-tab="${n}"]`).classList.add("active")}
    function clearRoute(){if(state.route){map.removeLayer(state.route);state.route=null}state.routeEnds.forEach(m=>map.removeLayer(m));state.routeEnds=[]}
    function drawRoute(a){clearRoute();const l=state.layers.get(a.icaoAddress);if(!l)return;const live=l.tr.getLatLngs()||[];const pred=l.pr.getLatLngs()||[];const all=[...live,...pred].map(p=>[p.lat,p.lng]);if(all.length<2)return;state.route=L.polyline(all,{color:"#fdb813",weight:2,dashArray:"5 6",opacity:.9}).addTo(map);const f=L.circleMarker(all[0],{radius:4,color:"#fdb813",fillColor:"#fdb813",fillOpacity:1}).addTo(map);const t=L.circleMarker(all[all.length-1],{radius:4,color:"#ff7d7d",fillColor:"#ff7d7d",fillOpacity:1}).addTo(map);state.routeEnds=[f,t]}
    function flightStatus(vsFpm){if(vsFpm>150)return{label:"Climbing",cls:"fbadge-climb",arrow:"↑"};if(vsFpm<-150)return{label:"Descending",cls:"fbadge-descend",arrow:"↓"};return{label:"Level",cls:"fbadge-level",arrow:"→"};}
    function panels(a){
      const altM=Math.round(a.altitudeMM/1000),altFt=Math.round(altM*M_TO_FT);
      const spdMs=Math.round(a.horVelocityCMS/100),spdKt=Math.round(spdMs*MS_TO_KT);
      const vsMs=a.verVelocityCMS/100,vsFpm=Math.round(vsMs*MS_TO_FPM);
      const hdg=Math.round(a.headingDE2/100);
      const squawk=a.squawk?String(a.squawk).padStart(4,"0"):"————";
      const st=flightStatus(vsFpm);
      const infoEntry=state.info.get(a.icaoAddress);
      const info=infoEntry&&infoEntry.status==="done"?infoEntry.data:null;
      const infoPending=!infoEntry||infoEntry.status==="pending";
      const typeBits=info?[[info.manufacturer,info.type].filter(Boolean).join(" "),info.registration?`Reg ${info.registration}`:''].filter(Boolean):[];
      const subParts=[`ICAO24 ${a.icaoAddress}`,`Squawk ${squawk}`,...(typeBits.length?typeBits:(a.sourceIp?[`Feed ${a.sourceIp}`]:[]))];
      if(a.isPredicted)subParts.push(`Pred. age ${Math.round(a.predictionAgeSec)}s`);
      const subLine=subParts.join(" &nbsp;·&nbsp; ");
      let thumbHtml;
      if(info&&info.photoUrl){
        thumbHtml=`<a href="${info.photoLink||'#'}" target="_blank" rel="noopener"><img src="${info.photoUrl}" alt="${a.callsign} photo" loading="lazy"/></a><div class="fthumb-credit">${info.photoCredit||"planespotters.net"}</div>`;
      }else if(infoPending){
        thumbHtml=`<div class="fthumb-ph">Looking up photo…</div>`;
      }else{
        thumbHtml=`<div class="fthumb-ph">No photo available</div>`;
      }
      const l=state.layers.get(a.icaoAddress);const live=l?l.tr.getLatLngs():[];const pred=l?l.pr.getLatLngs():[];const all=[...live,...pred].map(p=>[p.lat,p.lng]);let dist=0;for(let i=1;i<all.length;i++)dist+=haversine(all[i-1],all[i]);const first=l&&l.origin?l.origin:(all.length?all[0]:[a.latDD,a.lonDD]),last=all.length?all[all.length-1]:[a.latDD,a.lonDD];const observedStart=nearestPlace(first[0],first[1]),observedCurrent=nearestPlace(last[0],last[1]);
      ensureRouteInfo(a.callsign,a.icaoAddress);
      const routeEntry=state.routeInfo.get(a.callsign);
      const routeInfo=routeEntry&&routeEntry.status==="done"?routeEntry.data:null;
      const routePending=!routeEntry||routeEntry.status==="pending";
      const realOrigin=routeInfo&&routeInfo.origin;
      const realDest=routeInfo&&routeInfo.destination;
      const fromName=realOrigin?(realOrigin.municipality||realOrigin.name)+(realOrigin.iataCode?` (${realOrigin.iataCode})`:""):(routePending?"Looking up...":"Route not available");
      const fromSub=realOrigin?realOrigin.name:(routePending?"Checking scheduled origin":"No verified origin");
      const toName=realDest?(realDest.municipality||realDest.name)+(realDest.iataCode?` (${realDest.iataCode})`:""):(routePending?"Looking up...":"Route not available");
      const toSub=realDest?realDest.name:(routePending?"Checking scheduled destination":"No verified destination");
      const toLabel="To";
      document.getElementById("flight").innerHTML=`
        <div class="fcard">
          <div class="ftoprow">
            <div class="fmeta">
              <div class="fhead">
                <div class="fcall">${a.callsign}</div>
                <div class="fbadge ${st.cls}">${st.arrow} ${st.label}</div>
                ${a.isPredicted?'<div class="fbadge fbadge-pred">Predicted</div>':''}
              </div>
              <div class="fsub">${subLine}</div>
            </div>
            <div class="fthumb">${thumbHtml}</div>
          </div>
          <div class="stat-grid">
            <div class="stat"><div class="sk">Altitude</div><div class="sv">${altFt.toLocaleString()} ft</div><div class="su">${altM.toLocaleString()} m</div></div>
            <div class="stat"><div class="sk">Ground speed</div><div class="sv">${spdKt} kt</div><div class="su">${spdMs} m/s</div></div>
            <div class="stat"><div class="sk">Vertical speed</div><div class="sv">${vsFpm>0?'+':''}${vsFpm} ft/min</div><div class="su">${vsMs.toFixed(1)} m/s</div></div>
            <div class="stat"><div class="sk">Heading</div><div class="sv">${hdg}°</div></div>
            <div class="stat"><div class="sk">Position</div><div class="sv">${a.latDD.toFixed(3)}, ${a.lonDD.toFixed(3)}</div></div>
            <div class="stat"><div class="sk">Distance flown</div><div class="sv">${(dist/1000).toFixed(0)} km</div></div>
          </div>
          <div class="routebar">
            <div class="pt from"><div class="pn">${fromName}</div><div class="pd">${fromSub}</div></div>
            <div class="arrow">&#9992;</div>
            <div class="pt to"><div class="pn">${toName}</div><div class="pd">${toSub}</div></div>
          </div>
        </div>`;
      const typeRows=info?`<dt>Aircraft type</dt><dd>${[info.manufacturer,info.type].filter(Boolean).join(" ")||"Unknown"}${info.icaoTypeCode?` (${info.icaoTypeCode})`:''}</dd><dt>Registration</dt><dd>${info.registration||"Unknown"}</dd><dt>Category</dt><dd>${info.category}</dd>`:'';
      document.getElementById("aircraft").innerHTML=`<dl class="kv"><dt>ICAO24</dt><dd>${a.icaoAddress}</dd><dt>Callsign</dt><dd>${a.callsign}</dd>${typeRows}<dt>Latitude</dt><dd>${a.latDD.toFixed(6)}</dd><dt>Longitude</dt><dd>${a.lonDD.toFixed(6)}</dd><dt>Altitude</dt><dd>${altFt.toLocaleString()} ft (${altM.toLocaleString()} m)</dd><dt>Ground speed</dt><dd>${spdKt} kt (${spdMs} m/s)</dd><dt>Vertical speed</dt><dd>${vsFpm} ft/min</dd><dt>Heading</dt><dd>${hdg}°</dd><dt>Squawk</dt><dd>${squawk}</dd><dt>Source feed</dt><dd>${a.sourceIp||"—"}</dd></dl>`;
      const airlineRow=routeInfo&&routeInfo.airlineName?`<dt>Airline</dt><dd>${routeInfo.airlineName}</dd>`:'';
      const routeStatus=realOrigin||realDest?"Confirmed by route database":(routePending?"Lookup in progress":"Not available - no verified schedule found");
      document.getElementById("route").innerHTML=`<dl class="kv">${airlineRow}<dt>Route status</dt><dd>${routeStatus}</dd><dt>From</dt><dd>${fromName}</dd><dt>${toLabel}</dt><dd>${toName}</dd><dt>Observed start</dt><dd>${observedStart?`${observedStart.name} area (${observedStart.distanceKm.toFixed(0)} km)`:"Unknown"}</dd><dt>Current area</dt><dd>${observedCurrent?`${observedCurrent.name} area (${observedCurrent.distanceKm.toFixed(0)} km)`:"Unknown"}</dd><dt>Start Coord</dt><dd>${first[0].toFixed(6)}, ${first[1].toFixed(6)}</dd><dt>Current Coord</dt><dd>${last[0].toFixed(6)}, ${last[1].toFixed(6)}</dd><dt>Track Distance</dt><dd>${(dist/1000).toFixed(2)} km</dd></dl>`;
    }
    function pushAlt(a){const h=state.hist.get(a.icaoAddress)||[];h.push({t:new Date(),v:Math.round(a.altitudeMM/1000)});if(h.length>60)h.shift();state.hist.set(a.icaoAddress,h);if(state.sel!==a.icaoAddress)return;chart.data.labels=h.map(p=>p.t.toLocaleTimeString("en-IN",{hour12:false,minute:"2-digit",second:"2-digit"}));chart.data.datasets[0].data=h.map(p=>p.v);chart.update("none")}
    function animateAircraft(l,target,heading){if(l.anim)cancelAnimationFrame(l.anim);const start=l.m.getLatLng(),jump=haversine([start.lat,start.lng],target);if(jump>100000){l.m.setLatLng(target);l.m.setRotationAngle(heading);l.heading=heading;l.anim=null;return}const startHeading=Number.isFinite(l.heading)?l.heading:heading,delta=((heading-startHeading+540)%360)-180,started=performance.now(),duration=1250;const step=(now)=>{const t=Math.min(1,(now-started)/duration);l.m.setLatLng([start.lat+(target[0]-start.lat)*t,start.lng+(target[1]-start.lng)*t]);l.m.setRotationAngle(startHeading+delta*t);if(t<1){l.anim=requestAnimationFrame(step)}else{l.anim=null;l.heading=heading}};l.anim=requestAnimationFrame(step)}
    function ensure(a){let l=state.layers.get(a.icaoAddress),hd=a.headingDE2/100;if(!l){const cachedInfo=infoFor(a.icaoAddress);const m=L.marker([a.latDD,a.lonDD],{icon:iconFor(cachedInfo&&cachedInfo.category),rotationAngle:hd}).addTo(map);m.bindTooltip(a.callsign,{permanent:true,direction:"right",offset:[10,0],className:"lbl"}).openTooltip();m.bindPopup(pop(a));m.on("click",()=>select(a.icaoAddress));ensureInfo(a.icaoAddress);const tr=L.polyline([],{color:"#8ecae6",weight:2,opacity:.5}).addTo(map),pr=L.polyline([],{color:"#ff5d5d",weight:2,dashArray:"7 7",opacity:.95}).addTo(map),ve=L.polyline([],{color:"#bde0fe",weight:1,dashArray:"6 6"}).addTo(map);l={m,tr,pr,ve,data:a,vis:true,origin:[a.latDD,a.lonDD],heading:hd,anim:null};state.layers.set(a.icaoAddress,l)}l.data=a;animateAircraft(l,[a.latDD,a.lonDD],hd);l.m.setPopupContent(pop(a));l.m.setTooltipContent(a.callsign+(a.isPredicted?" PRED":""));if(a.isPredicted){l.pr.addLatLng([a.latDD,a.lonDD]);const pPts=l.pr.getLatLngs();if(pPts.length>120)l.pr.setLatLngs(pPts.slice(-120));}else{l.tr.addLatLng([a.latDD,a.lonDD]);const pts=l.tr.getLatLngs();if(pts.length>120)l.tr.setLatLngs(pts.slice(-120));if(l.pr.getLatLngs().length>0)l.pr.setLatLngs([])}const dist=(a.horVelocityCMS/100)*20,rad=hd*Math.PI/180,dLat=(dist/111320)*Math.cos(rad),scale=Math.max(.2,Math.cos(a.latDD*Math.PI/180)),dLon=(dist/(111320*scale))*Math.sin(rad);const vColor=a.isPredicted?"#ff5d5d":"#bde0fe";const vDash=a.isPredicted?"4 6":"6 6";l.ve.setStyle({color:vColor,dashArray:vDash});l.ve.setLatLngs([[a.latDD,a.lonDD],[a.latDD+dLat,a.lonDD+dLon]]);pushAlt(a)}
    function showLayer(l,v){if(v&&!l.vis){l.m.addTo(map);l.tr.addTo(map);l.pr.addTo(map);l.ve.addTo(map);l.vis=true}else if(!v&&l.vis){map.removeLayer(l.m);map.removeLayer(l.tr);map.removeLayer(l.pr);map.removeLayer(l.ve);l.vis=false}}
    function cleanup(active){for(const [i,l] of state.layers){if(active.has(i))continue;if(l.anim)cancelAnimationFrame(l.anim);map.removeLayer(l.m);map.removeLayer(l.tr);map.removeLayer(l.pr);map.removeLayer(l.ve);state.layers.delete(i);state.hist.delete(i);if(state.sel===i){state.sel=null;clearRoute();document.getElementById("flight").innerHTML='<div class="empty">Selected aircraft left the feed.</div>';document.getElementById("aircraft").innerHTML='<div class="empty">No aircraft selected.</div>';document.getElementById("route").innerHTML='<div class="empty">Route details appear on selection.</div>';chart.data.labels=[];chart.data.datasets[0].data=[];chart.update("none")}}}
    function select(i){const l=state.layers.get(i);if(!l)return;state.sel=i;const a=l.data;map.setView([a.latDD,a.lonDD],6);l.m.setZIndexOffset(1000);l.m.openPopup();panels(a);drawRoute(a);openTab("flight");filterAll()}
    function renderList(){const q=els.search.value.trim().toUpperCase();const arr=state.rows.filter(a=>!q||a.callsign.includes(q)||a.icaoAddress.includes(q)).sort((a,b)=>a.callsign.localeCompare(b.callsign));if(!arr.length){els.list.innerHTML='<li class="empty">No flights match current search.</li>';return}els.list.innerHTML="";arr.forEach(a=>{const li=document.createElement("li");li.className=`row${state.sel===a.icaoAddress?" active":""}`;li.innerHTML=`<div class="call">${a.callsign}${a.isPredicted?' <span style="color:#ff7d7d">(PRED)</span>':''}</div><div class="meta">${a.icaoAddress} | ALT ${Math.round(a.altitudeMM/1000)}m | SPD ${Math.round(a.horVelocityCMS/100)}m/s${a.isPredicted?` | AGE ${Math.round(a.predictionAgeSec)}s`:''}</div>`;li.addEventListener("click",()=>select(a.icaoAddress));els.list.appendChild(li)})}
    function filterAll(){const q=els.search.value.trim().toUpperCase();let vis=0;for(const [i,l] of state.layers){const a=l.data,m=!q||a.callsign.includes(q)||a.icaoAddress.includes(q)||state.sel===i;showLayer(l,m);if(m)vis++}els.visible.textContent=String(vis);renderList()}
    // ---- mode/status badge ------------------------------------------------
    const modeBadge=document.getElementById("mode-badge");
    const modeToggle=document.getElementById("mode-toggle");
    const liveBanner=document.getElementById("live-banner");
    let _lastStatus=null;
    async function refreshStatus(){
      try{
        const r=await fetch("/api/status",{cache:"no-store"});
        if(!r.ok)return;
        const s=await r.json();
        _lastStatus=s;
        const simulationOn=s.mode==="simulation";
        modeToggle.textContent=`Simulation: ${simulationOn?"TRUE":"FALSE"}`;
        modeToggle.classList.toggle("sim-on",simulationOn);
        if(s.mode==="simulation"){
          modeBadge.textContent="⏺ SIM";
          modeBadge.className="sim";
          liveBanner.style.display="none";
        } else {
          // live mode
          if(s.connected){
            const ago=s.lastDataSecondsAgo!=null?` · data ${s.lastDataSecondsAgo}s ago`:"";
            modeBadge.textContent=`● LIVE${ago}`;
            modeBadge.className="live-ok";
            liveBanner.style.display="none";
          } else if(s.bindError){
            modeBadge.textContent="✖ LIVE ERR";
            modeBadge.className="live-err";
            liveBanner.style.display="block";
            liveBanner.innerHTML=`<b>⚠ PingStation socket error</b><br>${s.bindError}<br>Check that port ${s.listenPort} is not in use by another program.`;
          } else {
            modeBadge.textContent="◌ LIVE WAIT";
            modeBadge.className="live-wait";
            liveBanner.style.display="block";
            liveBanner.innerHTML=`<b>Waiting for PingStation connection…</b><br>
              Listening on port <b>${s.listenPort}</b>. &nbsp;
              Configure the PingStation device to send its RAW stream to
              <b>this PC's IP address&nbsp;:&nbsp;${s.listenPort}</b> (TCP).<br>
              Frames decoded so far: ${s.totalFramesDecoded} &nbsp;|&nbsp;
              Connections received: ${s.totalConnections}<br>
              <span style="font-size:10px;color:#9dbdde">Config file: ${s.configFile}</span>`;
          }
        }
      }catch(e){/* ignore status errors */}
    }

    // ---- aircraft data refresh -------------------------------------------
    async function refresh(){
      try{
        const r=await fetch("/api/aircraft",{cache:"no-store"});
        if(!r.ok)throw new Error(r.status);
        const p=await r.json(),src=Array.isArray(p)?p:(p.aircraft||[]),arr=src.map(norm).filter(a=>Number.isFinite(a.latDD)&&Number.isFinite(a.lonDD));
        state.rows=arr;
        els.tracked.textContent=String(arr.length);
        // Show "Live" / "Sim" in the Feed chip based on mode
        if(_lastStatus&&_lastStatus.mode==="simulation"){
          els.status.textContent="Sim";
          els.status.style.color="#8ecae6";
        } else {
          els.status.textContent=arr.length?"Live":"Live";
          els.status.style.color="var(--ok)";
        }
        const active=new Set();
        arr.forEach(a=>{active.add(a.icaoAddress);ensure(a);if(state.sel===a.icaoAddress){panels(a);drawRoute(a)}});
        cleanup(active);
        filterAll();
      }catch(e){
        els.status.textContent="Error";
        els.status.style.color="var(--bad)";
      }
    }

    document.getElementById("zin").addEventListener("click",()=>map.zoomIn());
    document.getElementById("zout").addEventListener("click",()=>map.zoomOut());
    document.getElementById("center").addEventListener("click",()=>map.setView([22.2,79.2],5));
    document.querySelectorAll("#tabs button").forEach(b=>b.addEventListener("click",()=>openTab(b.dataset.tab)));
    els.search.addEventListener("input",filterAll);
    modeToggle.addEventListener("click",async()=>{
      const simulation=!(_lastStatus&&_lastStatus.mode==="simulation");
      modeToggle.disabled=true;
      modeToggle.textContent="Switching...";
      try{
        const r=await fetch("/api/mode",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({simulation})});
        const response=await r.json().catch(()=>({}));
        if(!r.ok)throw new Error(response.error||`HTTP ${r.status}`);
        _lastStatus=response;
        await refreshStatus();
        await refresh();
      }catch(e){
        modeToggle.textContent=`Switch failed: ${e.message||"unknown error"}`;
        setTimeout(refreshStatus,2500);
      }finally{
        modeToggle.disabled=false;
      }
    });

    const bottomResizer=document.getElementById("bottom-resizer");
    const savedBottomHeight=Number(localStorage.getItem("adsbBottomHeight"));
    if(Number.isFinite(savedBottomHeight)&&savedBottomHeight>=150)document.documentElement.style.setProperty("--bottom-h",`${savedBottomHeight}px`);
    function beginBottomResize(startY,moveEventName,endEventName,getY){
      const startHeight=document.getElementById("bottom").getBoundingClientRect().height;
      const resize=moveEvent=>{
        if(moveEvent.cancelable)moveEvent.preventDefault();
        const height=Math.max(150,Math.min(window.innerHeight-150,startHeight+startY-getY(moveEvent)));
        document.documentElement.style.setProperty("--bottom-h",`${height}px`);
        localStorage.setItem("adsbBottomHeight",String(Math.round(height)));
        map.invalidateSize({pan:false});
        chart.resize();
      };
      const finish=()=>{
        document.removeEventListener(moveEventName,resize);
        document.removeEventListener(endEventName,finish);
      };
      document.addEventListener(moveEventName,resize,{passive:false});
      document.addEventListener(endEventName,finish);
    }
    bottomResizer.addEventListener("mousedown",event=>{
      event.preventDefault();
      beginBottomResize(event.clientY,"mousemove","mouseup",moveEvent=>moveEvent.clientY);
    });
    bottomResizer.addEventListener("touchstart",event=>{
      event.preventDefault();
      beginBottomResize(event.touches[0].clientY,"touchmove","touchend",moveEvent=>moveEvent.touches[0].clientY);
    },{passive:false});

    // Initial load and recurring timers
    refreshStatus();
    refresh();
    setInterval(refresh,1000);
    setInterval(refreshStatus,4000);
  