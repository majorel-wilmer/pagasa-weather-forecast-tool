const body=document.querySelector('#forecast-body');
const head=document.querySelector('#head');
const statusEl=document.querySelector('#status');
const retrievedEl=document.querySelector('#retrieved');
const issuedEl=document.querySelector('#issued');
const nextRefreshEl=document.querySelector('#next-refresh');
const adminButton=document.querySelector('#admin-login');
const adminDialog=document.querySelector('#admin-dialog');
const adminForm=document.querySelector('#admin-form');
const adminPassword=document.querySelector('#admin-password');
const adminError=document.querySelector('#admin-error');
const REFRESH_MS=30*60*1000;
let nextRefreshAt=Date.now()+REFRESH_MS;
let adminMode=false;

function esc(value){return String(value??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
function regionSpan(rows,index){if(!rows[index].region)return 0;let count=1;for(let i=index+1;i<rows.length&&!rows[i].region;i++)count++;return count}
function severityLabel(level){return ({green:'LIGHT',yellow:'MODERATE',orange:'HEAVY',red:'RED OVERRIDE',none:'NO RAIN'})[level]||level.toUpperCase()}

function render(data){
  const first=data.rows.find(r=>r.days.length)?.days||[];
  head.innerHTML='<th>Region</th><th>Site</th>'+first.slice(0,5).map(d=>`<th>${esc(d.date)}</th>`).join('');
  body.innerHTML=data.rows.map((row,index)=>{
    const region=row.region?`<td class="region" rowspan="${regionSpan(data.rows,index)}">${esc(row.region)}</td>`:'';
    const cells=row.days.slice(0,5).map(day=>`<td class="forecast severity-${esc(day.severity)}" data-site="${esc(row.site)}" data-date="${esc(day.date)}"><button class="override-button${adminMode?' visible':''}${day.red_override?' active':''}" type="button" data-site="${esc(row.site)}" data-date="${esc(day.date)}" data-red="${day.red_override?'true':'false'}">${day.red_override?'Undo RED':'Set RED'}</button><span class="severity-badge">${esc(severityLabel(day.severity))}</span><div class="weather-top">${day.icon?`<img src="${esc(day.icon)}" alt="">`:''}<span class="condition">${esc(day.condition)}</span></div><div class="metrics"><span class="pill">Low ${esc(day.low)}</span><span class="pill">High ${esc(day.high)}</span><span class="pill">Rain ${day.rain_chance??'—'}%</span></div><div class="forecast-window">${esc(day.forecast_window)}</div></td>`).join('');
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

async function checkAdmin(){
  try{const res=await fetch('/api/admin/status',{cache:'no-store'});const data=await res.json();adminMode=Boolean(data.authenticated);adminButton.textContent=adminMode?'Admin: signed in':'Admin';adminButton.classList.toggle('admin-active',adminMode)}catch{adminMode=false}
}

async function setRed(site,date,red){
  const res=await fetch('/api/admin/override',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({site,date,red})});
  if(res.status===401){adminMode=false;await checkAdmin();adminDialog.showModal();throw new Error('Admin session expired. Sign in again.')}
  if(!res.ok)throw new Error((await res.json()).detail||'Unable to update warning');
  await load();
}

document.querySelector('#refresh').addEventListener('click',load);
adminButton.addEventListener('click',async()=>{if(adminMode){await fetch('/api/admin/logout',{method:'POST'});adminMode=false;adminButton.textContent='Admin';adminButton.classList.remove('admin-active');await load()}else{adminError.textContent='';adminPassword.value='';adminDialog.showModal();adminPassword.focus()}});
document.querySelector('.dialog-close').addEventListener('click',()=>adminDialog.close());
adminForm.addEventListener('submit',async event=>{event.preventDefault();adminError.textContent='';const res=await fetch('/api/admin/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:adminPassword.value})});if(!res.ok){adminError.textContent=(await res.json()).detail||'Login failed';return}adminMode=true;adminDialog.close();adminButton.textContent='Admin: signed in';adminButton.classList.add('admin-active');await load()});
body.addEventListener('click',async event=>{const button=event.target.closest('.override-button');if(!button||!adminMode)return;button.disabled=true;try{await setRed(button.dataset.site,button.dataset.date,button.dataset.red!=='true')}catch(error){alert(error.message)}finally{button.disabled=false}});
setInterval(load,REFRESH_MS);
setInterval(()=>{const remaining=Math.max(0,nextRefreshAt-Date.now());const minutes=Math.floor(remaining/60000);const seconds=Math.floor((remaining%60000)/1000);nextRefreshEl.textContent=`Next refresh in ${minutes}:${String(seconds).padStart(2,'0')}`},1000);
(async()=>{await checkAdmin();await load()})();
