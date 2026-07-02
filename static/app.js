const body=document.querySelector('#forecast-body');
const head=document.querySelector('#head');
const statusEl=document.querySelector('#status');
const retrievedEl=document.querySelector('#retrieved');
const issuedEl=document.querySelector('#issued');
const nextRefreshEl=document.querySelector('#next-refresh');
const REFRESH_MS=30*60*1000;
let nextRefreshAt=Date.now()+REFRESH_MS;

function esc(value){return String(value??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
function regionSpan(rows,index){if(!rows[index].region)return 0;let count=1;for(let i=index+1;i<rows.length&&!rows[i].region;i++)count++;return count}
function render(data){
  const first=data.rows.find(r=>r.days.length)?.days||[];
  head.innerHTML='<th>Region</th><th>Site</th>'+first.slice(0,5).map(d=>`<th>${esc(d.date)}</th>`).join('');
  body.innerHTML=data.rows.map((row,index)=>{
    const region=row.region?`<td class="region" rowspan="${regionSpan(data.rows,index)}">${esc(row.region)}</td>`:'';
    const cells=row.days.slice(0,5).map(day=>`<td class="forecast"><div class="weather-top">${day.icon?`<img src="${esc(day.icon)}" alt="">`:''}<span class="condition">${esc(day.condition)}</span></div><div class="metrics"><span class="pill">Low ${esc(day.low)}</span><span class="pill">High ${esc(day.high)}</span><span class="pill">Rain ${day.rain_chance??'—'}%</span></div></td>`).join('');
    const missing=row.days.length?'':`<td class="error" colspan="5">Forecast unavailable</td>`;
    return `<tr>${region}<td class="site">${esc(row.site)}<span class="source">PAGASA source: ${esc(row.source_city)}</span></td>${cells}${missing}</tr>`;
  }).join('');
  issuedEl.textContent=data.issued;
  statusEl.textContent=`Source: PAGASA (${data.source_mode})`;
  retrievedEl.textContent=`Retrieved ${new Date(data.retrieved_at).toLocaleString()}`;
}
async function load(){
  statusEl.textContent='Refreshing PAGASA data…';
  try{const res=await fetch('/api/forecast',{cache:'no-store'});if(!res.ok)throw new Error((await res.json()).detail||'Request failed');render(await res.json());nextRefreshAt=Date.now()+REFRESH_MS}
  catch(err){body.innerHTML=`<tr><td class="error" colspan="7">${esc(err.message)}</td></tr>`;statusEl.textContent='Unable to load forecast'}
}
document.querySelector('#refresh').addEventListener('click',load);
setInterval(load,REFRESH_MS);
setInterval(()=>{
  const remaining=Math.max(0,nextRefreshAt-Date.now());
  const minutes=Math.floor(remaining/60000);
  const seconds=Math.floor((remaining%60000)/1000);
  nextRefreshEl.textContent=`Next refresh in ${minutes}:${String(seconds).padStart(2,'0')}`;
},1000);
load();
