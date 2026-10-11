// 주피터 웹판 (아이폰용, 2026-10-11: 아이폰은 웹, 안드로이드는 apk)
// 화면(index.html)은 안드로이드 앱과 같은 파일을 쓴다. 앱이 자바에 맡기던 일(window.Android)을 여기서 웹 방식으로 대신한다
//  - 문자·전화: sms: / tel: 주소 (전송·통화는 문자·전화 앱에서 직접)
//  - 명단 불러오기: 파일 고르기 → 엑셀(xlsx)·CSV를 이 안에서 읽음 (서버로 보내지 않음)
//  - 약속·부재중 알림: 캘린더 일정 (구글 캘린더 링크 또는 아이폰 캘린더 파일). 웹은 정한 시각에 알림을 직접 띄울 수 없어서 캘린더 알림으로 대신한다
//  - 연락처 저장: 연락처 파일(.vcf) → [새로운 연락처 생성]
// 자료(명단·기록·문구)는 이 휴대폰 브라우저 저장소에만 있다
(function(){
  const VERSION = '1.8.1';
  const isIOS = /iP(hone|ad|od)/.test(navigator.userAgent) || (navigator.platform==='MacIntel' && navigator.maxTouchPoints>1);
  const standalone = navigator.standalone === true || matchMedia('(display-mode: standalone)').matches;
  const get = (k, d) => { try{ const v = localStorage.getItem(k); return v==null ? d : JSON.parse(v); }catch(e){ return d; } };
  const put = (k, v) => { try{ localStorage.setItem(k, JSON.stringify(v)); }catch(e){} };
  const say = t => window.toast ? toast(t) : alert(t);

  // ---------- 캘린더 ----------
  // google = 구글 캘린더 새 일정 화면(링크) / ios = 아이폰 캘린더 파일(.ics, 알림 시각까지 들어감)
  const calMode = () => get('webCal', 'google');
  const utc = ms => new Date(ms).toISOString().replace(/[-:]/g,'').replace(/\.\d+Z$/,'Z');
  function googleCal(title, begin, end, place, desc){
    const q = new URLSearchParams({action:'TEMPLATE', text:title, dates:utc(begin)+'/'+utc(end), details:desc||'', location:place||''});
    return openOut('https://calendar.google.com/calendar/render?' + q);
  }
  const icsEsc = s => String(s||'').replace(/\\/g,'\\\\').replace(/;/g,'\\;').replace(/,/g,'\\,').replace(/\r?\n/g,'\\n');
  function icsText(title, begin, end, place, desc, alarmMin){
    return ['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//Hanho//Jupiter//KO','CALSCALE:GREGORIAN','METHOD:PUBLISH','BEGIN:VEVENT',
      'UID:'+begin+'-'+Math.random().toString(36).slice(2)+'@jupiter', 'DTSTAMP:'+utc(Date.now()), 'DTSTART:'+utc(begin), 'DTEND:'+utc(end),
      'SUMMARY:'+icsEsc(title), place?'LOCATION:'+icsEsc(place):'', desc?'DESCRIPTION:'+icsEsc(desc):'',
      'BEGIN:VALARM','ACTION:DISPLAY','DESCRIPTION:'+icsEsc(title),'TRIGGER:-PT'+(alarmMin||0)+'M','END:VALARM',
      'END:VEVENT','END:VCALENDAR'].filter(Boolean).join('\r\n');
  }
  function calendar(title, begin, end, place, desc, alarmMin){
    if(calMode()==='ios') return openFile('jupiter.ics', 'text/calendar', icsText(title, begin, end, place, desc, alarmMin));
    return googleCal(title, begin, end, place, desc);
  }

  // ---------- 파일 열기 (캘린더 .ics / 연락처 .vcf) ----------
  // 홈 화면에 추가한 웹앱 안에서는 이런 파일이 잘 열리지 않아, 앱 범위 밖의 도우미 쪽(../f/open.html)을 사파리 창으로 열어 거기서 연다
  function b64(s){ const u = new TextEncoder().encode(s); let b=''; u.forEach(x=>b+=String.fromCharCode(x)); return btoa(b).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,''); }
  function openFile(name, type, text){
    return openOut(new URL('../f/open.html', location.href).href + '#' + b64(JSON.stringify({n:name, t:type, d:text})));
  }
  // 새 창(사파리 창)으로 연다. 막히면 지금 창에서
  function openOut(url){
    const w = window.open(url, '_blank');
    if(!w) location.href = url;
    return true;
  }

  // ---------- 명단 파일 읽기 (앱의 MainActivity.readXlsx / readCsv 와 같은 규칙) ----------
  async function unzip(buf){
    const dv = new DataView(buf), u8 = new Uint8Array(buf), files = {};
    let eo = -1;
    for(let i = buf.byteLength-22; i >= Math.max(0, buf.byteLength-65557); i--) if(dv.getUint32(i, true)===0x06054b50){ eo = i; break; }
    if(eo<0) throw new Error('엑셀 파일이 아닙니다');
    const count = dv.getUint16(eo+10, true); let p = dv.getUint32(eo+16, true);
    const dec = new TextDecoder();
    for(let k=0; k<count; k++){
      if(dv.getUint32(p, true)!==0x02014b50) break;
      const method = dv.getUint16(p+10, true), csize = dv.getUint32(p+20, true);
      const nl = dv.getUint16(p+28, true), el = dv.getUint16(p+30, true), cl = dv.getUint16(p+32, true), lo = dv.getUint32(p+42, true);
      const name = dec.decode(u8.subarray(p+46, p+46+nl));
      p += 46+nl+el+cl;
      if(!(name==='xl/sharedStrings.xml' || name==='xl/workbook.xml' || name==='xl/_rels/workbook.xml.rels' || name.startsWith('xl/worksheets/sheet'))) continue;
      const start = lo + 30 + dv.getUint16(lo+26, true) + dv.getUint16(lo+28, true);
      const raw = u8.subarray(start, start+csize);
      if(method===0) files[name] = dec.decode(raw);
      else if(method===8){
        if(typeof DecompressionStream==='undefined') throw new Error('이 브라우저는 엑셀을 읽지 못합니다 (iOS 16.4 이상 필요). CSV로 저장해 고르세요');
        const out = await new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer();
        files[name] = dec.decode(out);
      }
    }
    return files;
  }
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const tags = (n, t) => Array.from(n.getElementsByTagNameNS('*', t));
  function colIndex(ref, fb){
    if(!ref) return fb; let c = 0;
    for(const ch of ref){ if(ch>='A' && ch<='Z') c = c*26 + (ch.charCodeAt(0)-64); else break; }
    return c>0 ? c-1 : fb;
  }
  async function readXlsx(buf){
    const f = await unzip(buf);
    const ss = f['xl/sharedStrings.xml'] ? tags(xml(f['xl/sharedStrings.xml']), 'si').map(si =>
      tags(si, 't').filter(t => !(t.parentNode && t.parentNode.localName==='rPh')).map(t=>t.textContent).join('')) : [];
    // 첫 번째 시트 (workbook 의 첫 sheet → rels 의 대상), 못 찾으면 sheet1.xml
    let sheet = 'xl/worksheets/sheet1.xml';
    try{
      const rid = /<sheet [^>]*r:id="([^"]+)"/.exec(f['xl/workbook.xml'])[1];
      const rels = f['xl/_rels/workbook.xml.rels'];
      const m = new RegExp('<Relationship [^>]*Id="'+rid+'"[^>]*Target="([^"]+)"|<Relationship [^>]*Target="([^"]+)"[^>]*Id="'+rid+'"').exec(rels);
      let t = m[1] || m[2]; t = t.startsWith('/') ? t.slice(1) : 'xl/' + t;
      if(f[t]) sheet = t;
    }catch(e){}
    if(!f[sheet]) throw new Error('엑셀 시트를 찾지 못했습니다');
    return tags(xml(f[sheet]), 'row').map(r => {
      const row = [];
      tags(r, 'c').forEach(c => {
        const type = c.getAttribute('t'), v = tags(c, 'v')[0];
        let s = v ? v.textContent : tags(c, 't').map(t=>t.textContent).join('');
        if(type==='s') s = ss[parseInt(s.trim(), 10)] ?? '';
        else if(/^-?\d+\.0+$/.test(s)) s = s.slice(0, s.indexOf('.'));   // 숫자로 저장된 번호
        const col = colIndex(c.getAttribute('r'), row.length);
        while(row.length < col) row.push('');
        row.push(s);
      });
      return row;
    });
  }
  function readCsv(buf){
    let t = new TextDecoder('utf-8').decode(buf);
    if(t.includes('\uFFFD')) t = new TextDecoder('euc-kr').decode(buf);
    if(t.charCodeAt(0)===0xFEFF) t = t.slice(1);
    const rows = []; let row = [], f = '', q = false;
    for(let i=0; i<t.length; i++){
      const c = t[i];
      if(q){ if(c==='"'){ if(t[i+1]==='"'){ f+='"'; i++; } else q = false; } else f += c; }
      else if(c==='"') q = true;
      else if(c===','){ row.push(f); f=''; }
      else if(c==='\n'){ row.push(f); f=''; rows.push(row); row=[]; }
      else if(c!=='\r') f += c;
    }
    if(f || row.length){ row.push(f); rows.push(row); }
    return rows;
  }
  function pickFile(){
    let inp = document.getElementById('webFile');
    if(!inp){
      inp = document.createElement('input'); inp.type = 'file'; inp.id = 'webFile'; inp.style.display = 'none';
      inp.accept = '.xlsx,.csv,.txt,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv';
      inp.onchange = async () => {
        const file = inp.files && inp.files[0]; inp.value = '';
        if(!file) return;
        try{
          const buf = await file.arrayBuffer(), u = new Uint8Array(buf);
          const rows = u.length>3 && u[0]===0x50 && u[1]===0x4B ? await readXlsx(buf) : readCsv(buf);
          window.onImported && window.onImported(rows);
        }catch(e){ window.onImportError && window.onImportError(e.message || String(e)); }
      };
      document.body.appendChild(inp);
    }
    inp.click();
  }

  // ---------- 앱의 window.Android 자리 ----------
  window.Android = {
    openSms(phone, text){
      // 아이폰은 sms:번호&body= , 안드로이드는 sms:번호?body=
      location.href = 'sms:' + phone + (text ? (isIOS ? '&' : '?') + 'body=' + encodeURIComponent(text) : '');
      return true;
    },
    openDial(phone){ location.href = 'tel:' + phone; return true; },
    openContact(name, phone, notes){
      const e = s => String(s||'').replace(/\\/g,'\\\\').replace(/;/g,'\\;').replace(/,/g,'\\,').replace(/\r?\n/g,'\\n');
      const v = ['BEGIN:VCARD','VERSION:3.0','N:;'+e(name)+';;;','FN:'+e(name),'TEL;TYPE=CELL:'+phone, notes?'NOTE:'+e(notes):'','END:VCARD'].filter(Boolean).join('\r\n');
      return openFile('contact.vcf', 'text/vcard', v);
    },
    // 부재중 다시 전화: 그 시각의 캘린더 일정(15분)으로. 아이폰 캘린더 파일이면 시각에 맞춰 알림이 들어간다
    setAlarm(id, when, title, text, phone){
      calendar(title, when, when + 900000, '', [text, '전화: tel:' + phone].filter(Boolean).join('\n'), 0);
      return 'cal';
    },
    // 캘린더에 넣은 일정은 웹에서 지울 수 없다
    cancelAlarm(id){ setTimeout(()=>say('캘린더에 넣은 다시 전화 일정은 캘린더에서 지워 주세요'), 2800); },
    // 약속: 캘린더 일정 (아이폰 캘린더 파일이면 30분 전 알림)
    openCalendar(title, begin, end, place, desc){ return calendar(title, begin, end, place, desc, 30); },
    isNightMode(){ return matchMedia('(prefers-color-scheme: dark)').matches; },
    setDark(dark){ const m = document.querySelector('meta[name=theme-color]'); if(m) m.content = dark ? '#1c1f27' : '#ffffff'; },
    pickFile,
    exit(){}
  };

  // 첫 화면 아래: 캘린더 고르기 + (사파리에서 열었으면) 홈 화면 추가 안내
  window.setWebCal = k => { put('webCal', k); show(home, false); };
  window.webExtra = function(){
    const m = calMode();
    return `<div class="theme"><span>캘린더</span>
      ${[['google','구글 캘린더'],['ios','아이폰 캘린더']].map(([k,t])=>`<button class="btn seg${m===k?' on':''}" onclick="setWebCal('${k}')">${t}</button>`).join('')}</div>
      <p class="note">약속과 부재중 '다시 전화'는 캘린더 일정으로 들어갑니다. 열린 캘린더에서 [저장]을 눌러야 그 시각에 알림이 옵니다.
      ${m==='google' ? '구글 캘린더 알림이 일정 시작 때 오게 하려면 구글 캘린더 설정 > 기본 알림을 \'이벤트 시작 시\'로 해 두세요.' : ''}</p>
      ${isIOS && !standalone ? `<p class="note" style="color:var(--miss)"><b>먼저 홈 화면에 추가하세요.</b> 사파리 아래 [공유] → [홈 화면에 추가] 후 홈 화면의 주피터로 여세요. 사파리에서 그냥 쓰면 오래 안 열 때 명단·기록이 지워질 수 있습니다.</p>` : ''}
      <p class="note">주피터 웹 v${VERSION} · 명단과 기록은 이 휴대폰에만 저장됩니다.</p>`;
  };

  // 다른 앱(캘린더·문자)에 다녀오면: 앱의 onResume 과 같게 (약속 문자 보낼지 묻기)
  document.addEventListener('visibilitychange', () => { if(document.visibilityState==='visible' && window.onResumeApp) window.onResumeApp(); });
  // 휴대폰 다크 모드가 바뀌면 바로 따라감
  matchMedia('(prefers-color-scheme: dark)').addEventListener?.('change', () => window.applyTheme && applyTheme());
  // 저장소를 지우지 말라고 브라우저에 요청 + 꺼진 상태에서도 열리게
  if(navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(()=>{});
  if('serviceWorker' in navigator && (location.protocol==='https:' || location.hostname==='localhost'))
    navigator.serviceWorker.register('sw.js').catch(()=>{});
})();
