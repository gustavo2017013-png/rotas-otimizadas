{
// ===== leitor/gravador de planilha sem internet =====
const CRC=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
function crc32(b){let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=CRC[(c^b[i])&255]^(c>>>8);return(c^0xFFFFFFFF)>>>0}
const unxml=s=>String(s).replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi,(m,e)=>e[0]==='#'?String.fromCodePoint(e[1]==='x'||e[1]==='X'?parseInt(e.slice(2),16):+e.slice(1)):({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"})[e.toLowerCase()]);
const toxml=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'');

async function unzip(buf){
  const u=new Uint8Array(buf),dv=new DataView(buf),files={};
  let e=-1;for(let i=u.length-22;i>=Math.max(0,u.length-70000);i--)if(dv.getUint32(i,true)===0x06054b50){e=i;break}
  if(e<0)throw new Error('zip');
  const n=dv.getUint16(e+10,true);let p=dv.getUint32(e+16,true);const dec=new TextDecoder();
  for(let k=0;k<n;k++){
    if(dv.getUint32(p,true)!==0x02014b50)break;
    const m=dv.getUint16(p+10,true),cs=dv.getUint32(p+20,true),nl=dv.getUint16(p+28,true),xl=dv.getUint16(p+30,true),cl=dv.getUint16(p+32,true),lo=dv.getUint32(p+42,true);
    const name=dec.decode(u.subarray(p+46,p+46+nl));
    const ds=lo+30+dv.getUint16(lo+26,true)+dv.getUint16(lo+28,true);
    files[name]={m,data:u.subarray(ds,ds+cs)};
    p+=46+nl+xl+cl;
  }
  return async name=>{
    const f=files[name]||files[Object.keys(files).find(k=>k.toLowerCase()===name.toLowerCase())];if(!f)return null;
    let out=f.data;
    if(f.m===8)out=new Uint8Array(await new Response(new Blob([f.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
    return dec.decode(out);
  };
}
const tagRe=(t,flags='g')=>new RegExp(`<(?:\\w+:)?${t}\\b[^>]*?(?:/>|>([\\s\\S]*?)</(?:\\w+:)?${t}>)`,flags);
const texts=x=>[...String(x||'').matchAll(tagRe('t'))].map(m=>unxml(m[1]||'')).join('');
const colNum=r=>{let n=0;for(const ch of r.replace(/\d+/g,''))n=n*26+ch.toUpperCase().charCodeAt(0)-64;return n-1};

async function readXlsx(buf){
  const get=await unzip(buf);
  let path='xl/worksheets/sheet1.xml';
  const wb=await get('xl/workbook.xml'),rels=await get('xl/_rels/workbook.xml.rels');
  if(wb&&rels){
    const sh=wb.match(/<(?:\w+:)?sheet\b[^>]*>/);const id=sh&&(sh[0].match(/r:id="([^"]+)"/)||[])[1];
    const rel=id&&[...rels.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)].map(m=>m[0]).find(r=>r.includes(`Id="${id}"`));
    const tg=rel&&(rel.match(/Target="([^"]+)"/)||[])[1];
    if(tg)path=tg.startsWith('/')?tg.slice(1):'xl/'+tg.replace(/^\.\//,'');
  }
  const ssx=await get('xl/sharedStrings.xml');
  const ss=ssx?[...ssx.matchAll(tagRe('si'))].map(m=>texts((m[1]||'').replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g,''))):[];
  const xml=await get(path);if(!xml)throw new Error('sheet');
  const aoa=[];
  for(const rm of xml.matchAll(tagRe('row'))){
    const head=rm[0].match(/^<[^>]*>/)[0];const rn=(head.match(/\br="(\d+)"/)||[])[1];
    const ri=rn?+rn-1:aoa.length;const row=[];let ci=0;
    for(const cm of String(rm[1]||'').matchAll(tagRe('c'))){
      const ch=cm[0].match(/^<[^>]*>/)[0];const r=(ch.match(/\br="([A-Z]+\d*)"/i)||[])[1];if(r)ci=colNum(r);
      const t=(ch.match(/\bt="(\w+)"/)||[])[1]||'n';const inner=cm[1]||'';
      const v=(inner.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/)||[])[1];
      let val='';
      if(t==='s')val=ss[+v]??'';else if(t==='inlineStr')val=texts(inner);else if(t==='b')val=v==='1'?'VERDADEIRO':'FALSO';else val=v==null?'':unxml(v);
      row[ci]=val;ci++;
    }
    for(let i=0;i<row.length;i++)if(row[i]==null)row[i]='';
    aoa[ri]=row;
  }
  for(let i=0;i<aoa.length;i++)if(!aoa[i])aoa[i]=[];
  return aoa;
}
function readCsv(text){
  text=text.replace(/^﻿/,'');
  const first=text.split(/\r?\n/)[0]||'';const d=(first.match(/;/g)||[]).length>(first.match(/,/g)||[]).length?';':(first.includes('\t')?'\t':',');
  const out=[];let row=[],f='',q=false;
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(q){if(c==='"'){if(text[i+1]==='"'){f+='"';i++}else q=false}else f+=c}
    else if(c==='"')q=true;else if(c===d){row.push(f);f=''}
    else if(c==='\n'||c==='\r'){if(c==='\r'&&text[i+1]==='\n')i++;row.push(f);out.push(row);row=[];f=''}
    else f+=c;
  }
  if(f||row.length){row.push(f);out.push(row)}
  return out;
}
function zip(entries){
  const enc=new TextEncoder(),parts=[],cd=[];let off=0;
  for(const [name,str] of entries){
    const nm=enc.encode(name),data=enc.encode(str),crc=crc32(data);
    const h=new DataView(new ArrayBuffer(30));
    h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,0x0800,true);h.setUint16(8,0,true);h.setUint16(10,0,true);h.setUint16(12,0x21,true);
    h.setUint32(14,crc,true);h.setUint32(18,data.length,true);h.setUint32(22,data.length,true);h.setUint16(26,nm.length,true);h.setUint16(28,0,true);
    parts.push(new Uint8Array(h.buffer),nm,data);
    const c=new DataView(new ArrayBuffer(46));
    c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x0800,true);c.setUint16(10,0,true);c.setUint16(12,0,true);c.setUint16(14,0x21,true);
    c.setUint32(16,crc,true);c.setUint32(20,data.length,true);c.setUint32(24,data.length,true);c.setUint16(28,nm.length,true);c.setUint32(42,off,true);
    cd.push(new Uint8Array(c.buffer),nm);
    off+=30+nm.length+data.length;
  }
  const cdSize=cd.reduce((t,a)=>t+a.length,0);
  const e=new DataView(new ArrayBuffer(22));
  e.setUint32(0,0x06054b50,true);e.setUint16(8,entries.length,true);e.setUint16(10,entries.length,true);e.setUint32(12,cdSize,true);e.setUint32(16,off,true);
  return new Blob([...parts,...cd,new Uint8Array(e.buffer)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function colName(i){let s='';i++;while(i){const m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26)}return s}
function writeXlsx(aoa,widths){
  const rows=aoa.map((r,ri)=>`<row r="${ri+1}">`+r.map((v,ci)=>{
    const ref=colName(ci)+(ri+1),s=String(v??'');
    if(ri>0&&typeof v==='number')return `<c r="${ref}"><v>${v}</v></c>`;
    if(ri>0&&/^-?\d{1,9}$/.test(s))return `<c r="${ref}"><v>${s}</v></c>`;
    return `<c r="${ref}" t="inlineStr"${ri===0?' s="1"':''}><is><t xml:space="preserve">${toxml(s)}</t></is></c>`;
  }).join('')+'</row>').join('');
  const cols=widths?`<cols>${widths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('')}</cols>`:'';
  return zip([
    ['[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
    ['_rels/.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Paradas" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
    ['xl/styles.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>'],
    ['xl/worksheets/sheet1.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>${cols}<sheetData>${rows}</sheetData></worksheet>`]
  ]);
}
// ===== fim =====

const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=s=>String(s??'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();

const FIELDS=[
  ['seq','Sequência',['sequence','sequencia','ordem','seq','posicao','order']],
  ['stop','Parada (Stop)',['stop','parada','paradas']],
  ['addr','Endereço',['destination address','endereco','address','logradouro','rua','destino']],
  ['num','Número',['numero','number','n']],
  ['comp','Complemento',['complemento','complement','address 2','apto']],
  ['bairro','Bairro',['bairro','neighborhood','district']],
  ['city','Cidade',['city','cidade','municipio']],
  ['cep','CEP',['zipcode','cep','postal','zip']],
  ['code','Código do pacote',['spx tn','tracking','rastreio','codigo','pedido','order id','tn','at id']],
  ['name','Destinatário',['recipient','destinatario','nome','cliente','name']],
  ['lat','Latitude',['latitude','lat']],
  ['lon','Longitude',['longitude','lng','lon']]
];

let rows=[],headers=[],map={},stops=[],fname='rota',srcName='';

function say(t){const m=$('msg');m.textContent=t;m.style.display=t?'block':'none'}

function guess(){
  map={};const used=new Set();
  for(const [k,,keys] of FIELDS){
    let best=-1,score=0;
    headers.forEach((h,i)=>{
      if(used.has(i))return;const n=norm(h);
      keys.forEach((kw,j)=>{const s=(n===kw?100:(kw.length>2&&n.includes(kw))?60:0)-j;if(s>score){score=s;best=i}});
    });
    if(best>=0&&score>0){map[k]=best;used.add(best)}
  }
}

function buildSel(){
  const opts=['<option value="">— nenhuma —</option>',...headers.map((h,i)=>`<option value="${i}">${esc(h||'Coluna '+(i+1))}</option>`)].join('');
  $('sel').innerHTML=FIELDS.filter(([k])=>SHOW.includes(k)).map(([k,l])=>`<label class="f">${l}<select data-k="${k}">${opts}</select></label>`).join('');
  $('sel').querySelectorAll('select').forEach(s=>{s.value=map[s.dataset.k]??'';s.addEventListener('change',()=>{s.value===''?delete map[s.dataset.k]:map[s.dataset.k]=+s.value;group()})});
}

const val=(r,k)=>map[k]==null?'':String(r[map[k]]??'').trim();

function seqOf(r){const v=val(r,'seq');const n=parseFloat(v.replace(',','.'));return v!==''&&!isNaN(n)?n:null}
function seqLabel(s){
  const nums=[...new Set(s.items.map(seqOf).filter(n=>n!=null))].sort((a,b)=>a-b);
  const plus=s.items.filter(r=>seqOf(r)==null).length;
  return [...nums.map(String),...Array(plus).fill('+')].join(', ');
}
function coord(r){const a=parseFloat(val(r,'lat').replace(',','.')),b=parseFloat(val(r,'lon').replace(',','.'));return isNaN(a)||isNaN(b)||(!a&&!b)?null:[a,b]}
function dist(p,q){if(!p||!q)return Infinity;const dy=(p[0]-q[0])*110540,dx=(p[1]-q[1])*111320*Math.cos(p[0]*Math.PI/180);return Math.hypot(dx,dy)}
// ===== regra de agrupamento: uma linha por endereço (mesma rua + mesmo número) =====
// A coluna Stop NÃO agrupa: a Shopee põe casas vizinhas com números diferentes na mesma Stop.
const STREET_TYPES=/\b(rua|r|avenida|av|avda|travessa|trav|tv|alameda|al|estrada|estr|est|rodovia|rod|praca|pca|pc|largo|lgo|viela|beco)\b/g;
const NUM_RE=/^(?:n[º°o]?\.?\s*)?(\d+)\s*([a-z])?(?![\p{L}\d])/iu;
// acha o número da casa: primeiro trecho (depois da rua) que começa com número; o que vem depois é ignorado
function findNum(a){
  const parts=String(a).split(/[,;]|\s[-–]\s/);
  for(let i=1;i<parts.length;i++){
    const m=parts[i].trim().match(NUM_RE);
    if(m)return {street:parts.slice(0,i).join(' '),num:m[1].replace(/^0+(?=\d)/,'')+(m[2]||'').toLowerCase()};
  }
  const m=String(a).match(/^(.*?\p{L}.*?)[\s,]+(?:n[º°o]?\.?\s*)?(\d+)\s*([a-z])?(?![\p{L}\d])/iu);
  return m?{street:m[1],num:m[2].replace(/^0+(?=\d)/,'')+(m[3]||'').toLowerCase()}:null;
}
function houseNum(r){
  const n=val(r,'num');
  if(n){const m=norm(n).match(/\d+[a-z]?/);if(m)return m[0].replace(/^0+(?=\d)/,'')}
  const f=findNum(val(r,'addr'));return f?f.num:'';
}
function addrKey(r,useComp){
  const a=val(r,'addr'),f=findNum(a),num=houseNum(r);
  const street=norm(f?f.street:a).replace(STREET_TYPES,' ').replace(/\s+/g,' ').trim();
  let k=num?street+'#'+num:norm(a).replace(STREET_TYPES,' ').replace(/\s+/g,' ').trim();
  if(!k)return '';
  if(useComp)k+='|'+norm(val(r,'comp'));
  return k;
}
function group(){
  if(map.addr==null){say('Não achei a coluna de endereço. Abra “Ajustar colunas” e escolha qual é.');$('cfg').open=true}else say('');
  if(map.seq==null)say('Não achei a coluna de sequência. Abra “Ajustar colunas” e escolha qual é.');
  const useComp=$('useComp').checked;
  // 1) junta linhas com o mesmo endereço, OU com latitude/longitude idênticas e o mesmo número
  //    (a Shopee às vezes dá a mesma coordenada para casas vizinhas, ex.: 513 e 529 — essas NÃO juntam)
  const par=rows.map((_,i)=>i);
  const find=i=>{while(par[i]!==i){par[i]=par[par[i]];i=par[i]}return i};
  const seen={};
  rows.forEach((r,i)=>{
    const k=addrKey(r,useComp),c=coord(r),keys=[];
    if(k)keys.push('A'+k);
    if(c)keys.push('C'+c[0]+','+c[1]+'#'+houseNum(r)+(useComp?'|'+norm(val(r,'comp')):''));
    for(const x of keys){if(x in seen){const a=find(i),b=find(seen[x]);if(a!==b)par[Math.max(a,b)]=Math.min(a,b)}else seen[x]=i}
  });
  const byRoot=new Map();let lastSeq=0;
  rows.forEach((r,i)=>{
    const sq=seqOf(r);if(sq!=null)lastSeq=sq;
    const g=find(i);let s=byRoot.get(g);
    if(!s){s={key:'G'+g,items:[],r,fallback:lastSeq+0.5+i*1e-6};byRoot.set(g,s)}
    s.items.push(r);
  });
  stops=[...byRoot.values()];
  // 2) ordem: menor sequência do endereço
  for(const s of stops){
    let best=null;for(const r of s.items){const q=seqOf(r);if(q!=null&&(best==null||q<seqOf(best)))best=r}
    if(best){s.r=best;s.pos=seqOf(best)}else s.extra=true;
  }
  // 3) endereço só com pacotes "+": linha própria logo depois da parada mais próxima
  stops.filter(s=>s.extra).forEach((s,j)=>{
    const c=coord(s.r);let near=null,bestD=Infinity;
    for(const t of stops){if(t.extra)continue;for(const x of t.items){const d=dist(coord(x),c);if(d<bestD){bestD=d;near=t}}}
    s.pos=near?near.pos+0.5+j*1e-6:s.fallback;
  });
  stops.sort((a,b)=>a.pos-b.pos);
  stops.forEach((s,i)=>{s.n=i+1;s.id=s.key;s.label=seqLabel(s)});
  render();
}
// colunas repetidas (ex.: Bairro): se for a mesma coisa escrita diferente, mostra uma só
function merged(s,c){
  if(c===map.seq)return s.label;
  if(c===map.lat||c===map.lon)return String(s.r[c]??'');
  const seen=new Set(),v=[];
  for(const r of s.items){const x=String(r[c]??'').trim(),k=norm(x)||x;if(x&&!seen.has(k)){seen.add(k);v.push(x)}}
  return v.join(c===map.code?', ':' / ');
}

// só os campos que mudam o resultado aparecem em "Ajustar colunas"
const SHOW=['seq','addr','num','comp','code','lat','lon','city','bairro','cep'];

// ===== partes do endereço: rua, número e complemento (o que vem depois do número) =====
const UFS=new Set('ac al ap am ba ce df es go ma mt ms mg pa pb pr pe pi rj rn rs ro rr sc sp se to'.split(' '));
const STATES=new Set(['acre','alagoas','amapa','amazonas','bahia','ceara','distrito federal','espirito santo','goias','maranhao','mato grosso','mato grosso do sul','minas gerais','para','paraiba','parana','pernambuco','piaui','rio de janeiro','rio grande do norte','rio grande do sul','rondonia','roraima','santa catarina','sao paulo','sergipe','tocantins','brasil','brazil']);
const isNA=n=>!n||/^(n a|na|null|undefined|sem complemento|nenhum)$/.test(n);
// pedaço do endereço que é só cidade / estado / CEP / "N/A" (não é complemento)
function placeTest(r){
  const city=norm(val(r,'city')),cep=val(r,'cep').replace(/\D/g,'');
  const placeTok=new Set([...city.split(' '),...[...STATES].join(' ').split(' '),...UFS].filter(Boolean));
  return p=>{
    const n=norm(p),d=String(p).replace(/\D/g,'');
    if(isNA(n))return true;
    if((cep&&d===cep&&!/[a-z]/.test(n))||/^\d{5}\s?-?\s?\d{3}$/.test(String(p).trim()))return true;
    if((city&&n===city)||STATES.has(n)||UFS.has(n))return true;
    const t=n.split(' ');
    return (city&&n.includes(city)||[...STATES].some(s=>s.length>3&&n.includes(s)))&&t.every(x=>placeTok.has(x)||x===cep);
  };
}
function addrParts(r){
  const a=String(val(r,'addr')),isPlace=placeTest(r);
  let street='',num='',rest=[];
  const parts=a.split(',');
  for(let i=1;i<parts.length;i++){
    const p=parts[i].trim(),m=p.match(NUM_RE);
    if(m){street=parts.slice(0,i).join(',').trim();num=m[1]+(m[2]?' '+m[2].toUpperCase():'');rest=[p.slice(m[0].length),...parts.slice(i+1)];break}
  }
  if(!num){
    const m=a.match(/^(.*?\p{L}.*?)[\s,]+(?:n[º°o]?\.?\s*)?(\d+)\s*([a-z])?(?![\p{L}\d])/iu);
    if(m){street=m[1].trim();num=m[2]+(m[3]?' '+m[3].toUpperCase():'');rest=a.slice(m[0].length).split(',')}
  }
  const n=val(r,'num');if(n&&!num)num=n;
  if(!street)street=parts[0].trim();
  rest=rest.map(x=>x.replace(/^[\s\-–:;.]+|[\s\-–:;.]+$/g,'').trim()).filter(x=>x&&!isPlace(x));
  const c=val(r,'comp');if(c&&!isNA(norm(c))&&!rest.some(x=>norm(x)===norm(c)))rest.push(c);
  return {street:street.replace(/\s+/g,' '),num,comp:rest.join(', ')};
}
// endereço limpo da linha: rua + número, da forma mais completa entre os pacotes, sem complemento
function cleanStreet(s){
  let best=null;
  for(const r of s.items){const p=addrParts(r);if(p.num&&(!best||p.street.length>best.street.length))best=p}
  if(!best){const p=addrParts(s.r);return p.street}
  return best.street+', '+best.num;
}
// "Apto 12 (1) · Apto 34 (2) · sem complemento (1)"; um pacote só: só o complemento dele
function complements(s){
  const list=s.items.map(r=>addrParts(r).comp);
  if(s.items.length===1)return list[0];
  const order=[],count={},disp={};let none=0;
  for(const c of list){const k=norm(c);if(!k){none++;continue}if(!(k in count)){count[k]=0;disp[k]=c;order.push(k)}count[k]++}
  if(!order.length)return '';
  return [...order.map(k=>`${disp[k]} (${count[k]})`),...(none?[`sem complemento (${none})`]:[])].join(' · ');
}
function originals(s){
  const seen=new Set(),v=[];
  for(const r of s.items){const a=String(val(r,'addr')).trim();if(a&&!seen.has(a)){seen.add(a);v.push(a)}}
  return v.join(' / ');
}
// endereço completo (só linhas sem GPS): "Rua Floraí, 123, Água Boa, Paiçandu - PR, 87140-000"
const HAS_TYPE=/^(rua|r|av|avenida|avda|travessa|trav|tv|alameda|al|estrada|estr|est|rodovia|rod)\b/;
function fullAddr(s){
  let st=cleanStreet(s);
  if(!HAS_TYPE.test(norm(st)))st='Rua '+st;
  const r=s.r,bairro=String(val(r,'bairro')).trim();
  let city=String(val(r,'city')).trim();
  let d=val(r,'cep').replace(/\D/g,'');
  if(!city||d.length!==8){ // sem as colunas: tenta achar no próprio endereço
    const isPlace=placeTest(r);
    for(const p of String(val(r,'addr')).split(',').map(x=>x.trim())){
      const n=norm(p),dd=p.replace(/\D/g,'');
      if(d.length!==8&&dd.length===8&&!/[a-z]/.test(n))d=dd;
      else if(!city&&isPlace(p)&&n&&!isNA(n)&&!STATES.has(n)&&!UFS.has(n)&&!/\d/.test(n))city=p;
    }
  }
  const cepTxt=d.length===8?d.slice(0,5)+'-'+d.slice(5):'';
  const b=bairro&&norm(bairro)!==norm(city)&&!isNA(norm(bairro))?bairro:'';
  return [st,b,(city?city+' - ':'')+'PR',cepTxt].filter(Boolean).join(', ');
}

// ===== Ajuste de GPS (cada rua/endereço ligado a uma cidade). Não muda o agrupamento. =====
// Ordem: 1) Endereços corrigidos → 2) Ponto de referência da rua → 3) Apagar GPS (rua da lista sem ponto)
//        → 4) todas as outras linhas: GPS da Shopee sem mexer.
// NÃO existe nenhuma regra automática por distância do centro da rota (ela apagava GPS certo, ex.: distrito Água Boa).
const RAIO_PADRAO=500;
const GPS_DEFAULT=['R. Alession Dal Pozo Boti','R. Carlos Comini','R. Darcy Zirondi','R. Antônio Dossi','Av. das Gloriosas','R. Idalina Kellen Françozo','R. Luzia Jociele de Godoy','R. Sd. Arlindo Boeno de Godoy','R. João Marques dos Reis','R. Paulo Afonso','R. Arcebispo Dom Jaime Luiz Coelho','R. José Rigobelli'].map(rua=>({rua,cidade:'Paiçandu',ref:null,raio:RAIO_PADRAO}));
const GPS_KEY='rotasRuasGpsErrado2',GPS_OLD='rotasRuasGpsErrado',CITY_KEY='rotasUltimaCidade',FIX_KEY='rotasEnderecosCorrigidos';
const clip=(s,n=120)=>String(s??'').replace(/[<>]/g,'').trim().replace(/\s+/g,' ').slice(0,n);

// coordenadas coladas do Google Maps: "-23.4521, -52.0533" (também aceita "-23,4521 -52,0533")
function parseCoord(v){
  if(Array.isArray(v))v=v.join(', ');
  const s=String(v??'').trim().replace(/[()°]/g,'');
  if(!s)return null;
  const m=s.match(/^(-?\d{1,3}(?:\.\d+)?)\s*[,;\s]\s*(-?\d{1,3}(?:\.\d+)?)$/)||s.match(/^(-?\d{1,3}(?:,\d+)?)\s*(?:;|\s|,\s)\s*(-?\d{1,3}(?:,\d+)?)$/);
  if(!m)return null;
  const a=+m[1].replace(',','.'),b=+m[2].replace(',','.');
  return isFinite(a)&&isFinite(b)&&Math.abs(a)<=90&&Math.abs(b)<=180?[a,b]:null;
}
const inBR=(a,b)=>a>=-34&&a<=6&&b>=-74&&b<=-32;
function readCoord(txt){
  const c=parseCoord(txt);
  if(!c)return {err:'Não entendi as coordenadas. Cole como o Google Maps mostra, por exemplo: -23.4521, -52.0533'};
  if(!inBR(c[0],c[1]))return {err:inBR(c[1],c[0])?'As coordenadas parecem invertidas: primeiro vem a latitude (ex.: -23.45), depois a longitude (ex.: -52.05).':'Essas coordenadas ficam fora do Brasil. Confira se copiou certo.'};
  return {c};
}
const fmtC=c=>c?`${+c[0].toFixed(7)}, ${+c[1].toFixed(7)}`:'';
function cleanStreetEntry(x){
  if(!x||typeof x!=='object'||!clip(x.rua))return null;
  const ref=x.ref?parseCoord(x.ref):null;
  const raio=Math.round(+x.raio)>0?Math.min(50000,Math.round(+x.raio)):RAIO_PADRAO;
  return {rua:clip(x.rua),cidade:clip(x.cidade||''),ref:ref&&inBR(ref[0],ref[1])?ref:null,raio};
}
function cleanFixEntry(x){
  if(!x||typeof x!=='object'||!clip(x.rua)||!clip(x.numero,20))return null;
  const c=parseCoord(x.coord);if(!c||!inBR(c[0],c[1]))return null;
  return {rua:clip(x.rua),numero:clip(x.numero,20),cidade:clip(x.cidade||''),coord:c};
}
let badStreets=(()=>{
  try{const v=JSON.parse(localStorage.getItem(GPS_KEY)||'null');if(Array.isArray(v))return v.map(cleanStreetEntry).filter(Boolean)}catch(e){}
  try{const v=JSON.parse(localStorage.getItem(GPS_OLD)||'null');if(Array.isArray(v))return v.map(rua=>cleanStreetEntry({rua,cidade:'Paiçandu'})).filter(Boolean)}catch(e){} // lista antiga: era tudo Paiçandu
  return GPS_DEFAULT.map(x=>({...x}));
})();
let fixes=(()=>{try{const v=JSON.parse(localStorage.getItem(FIX_KEY)||'null');if(Array.isArray(v))return v.map(cleanFixEntry).filter(Boolean)}catch(e){}return []})();
const saveBad=()=>{try{localStorage.setItem(GPS_KEY,JSON.stringify(badStreets))}catch(e){}};
const storeFix=()=>{try{localStorage.setItem(FIX_KEY,JSON.stringify(fixes))}catch(e){}};

// tipo/título da rua e palavras de ligação não contam na comparação
const TITLES=new Set('r rua av avenida avda al alameda tv trav travessa est estr estrada rod rodovia pc pca praca lgo largo viela beco sd soldado arcebispo dom d dr doutor prof professor pe padre cel coronel sgt sargento cap capitao ten tenente gen general ver vereador pres presidente eng engenheiro dep deputado gov governador mal marechal frei cb cabo'.split(' '));
const LINK=new Set(['de','da','do','das','dos','e']);
const streetTokens=s=>norm(s).split(' ').filter(t=>t&&!TITLES.has(t)&&!LINK.has(t));
const numKey=n=>norm(n).replace(/\s+/g,'').replace(/^0+(?=\d)/,'');
const entryKey=e=>norm(e.cidade)+'|'+streetTokens(e.rua).join(' ');
const fixKey=e=>entryKey(e)+'#'+numKey(e.numero);
// só o nome da rua (antes do número), para não comparar com bairro/cidade/referência
function streetOf(r){const a=val(r,'addr');const f=findNum(a);return f?f.street:String(a).split(',')[0]}
function lev(a,b){let p=[...Array(b.length+1).keys()];for(let i=1;i<=a.length;i++){const c=[i];for(let j=1;j<=b.length;j++)c[j]=Math.min(p[j]+1,c[j-1]+1,p[j-1]+(a[i-1]===b[j-1]?0:1));p=c}return p[b.length]}
// erro de digitação aceito: 1 letra em palavras de 4 a 6 letras, 2 letras em palavras maiores
function tokCost(a,b){if(a===b)return 0;const L=Math.max(a.length,b.length),lim=L>=7?2:L>=4?1:0;if(!lim||Math.abs(a.length-b.length)>lim)return -1;const d=lev(a,b);return d<=lim?d:-1}
// cada palavra de "small" precisa casar com uma palavra diferente de "big"; devolve o total de letras erradas e os pares casados
function subsetMatch(small,big){
  let cost=0;const used=new Set(),pairs=[];
  for(const t of small){let best=-1,bi=-1;big.forEach((u,j)=>{if(used.has(j))return;const d=tokCost(t,u);if(d>=0&&(best<0||d<best)){best=d;bi=j}});if(bi<0)return null;used.add(bi);cost+=best;pairs.push([t,big[bi]])}
  return {cost,pairs};
}
// mesma rua: o nome do endereço está todo dentro do cadastro (endereço incompleto) ou o cadastro (palavras-chave) está todo dentro do endereço.
// Precisa casar pelo menos 2 palavras PRINCIPAIS do cadastro. Palavra que aparece em mais de uma rua da lista da mesma cidade
// (ex.: "Godoy") não é principal: "Luzia de Godoy" não vira "Luzia Jociele de Godoy" nem "Arlindo Boeno de Godoy".
function sameStreet(addrT,entryT,common){
  if(!addrT.length||!entryT.length)return false;
  const main=entryT.filter(t=>!common.has(t));
  const need=main.length?Math.min(2,main.length):entryT.length;
  for(const [small,big,entryIsBig] of [[addrT,entryT,true],[entryT,addrT,false]]){
    if(small.length<Math.min(2,entryT.length))continue;
    const m=subsetMatch(small,big);if(!m||m.cost>2)continue;
    const hitEntry=m.pairs.map(p=>entryIsBig?p[1]:p[0]);
    const hits=main.length?hitEntry.filter(t=>!common.has(t)).length:hitEntry.length;
    if(hits>=need)return true;
  }
  return false;
}
// separa o cadastro por cidade (sem acento e sem maiúscula) e acha as palavras repetidas entre ruas da mesma cidade
function cityIndex(list){
  const by={},common={};
  for(const e of list){const t=streetTokens(e.rua),c=norm(e.cidade);if(t.length&&c)(by[c]=by[c]||[]).push({e,t})}
  for(const c in by){const cnt={};for(const {t} of by[c])for(const w of new Set(t))cnt[w]=(cnt[w]||0)+1;common[c]=new Set(Object.keys(cnt).filter(w=>cnt[w]>1))}
  return {by,common};
}
const GPS_LABEL={fix:'endereço corrigido',ref:'ponto da rua',off:'removido'};
function computeGps(){
  const S=cityIndex(badStreets),F=cityIndex(fixes);
  const tot={fix:0,ref:0,off:0};
  for(const s of stops){
    s.gps=null;
    if(map.lat==null||map.lon==null)continue;
    const shopee=coord(s.r);
    // 1) endereço corrigido: mesma cidade, mesma rua e mesmo número
    let hit=null;
    for(const r of s.items){
      const c=norm(val(r,'city'));if(!F.by[c])continue;
      const t=streetTokens(streetOf(r)),n=numKey(houseNum(r));if(!n)continue;
      hit=F.by[c].find(x=>numKey(x.e.numero)===n&&sameStreet(t,x.t,F.common[c]));if(hit)break;
    }
    if(hit){s.gps={tipo:'fix',coord:hit.e.coord};tot.fix++;continue}
    // 2 e 3) rua da lista de GPS errado
    let st=null;
    for(const r of s.items){
      const c=norm(val(r,'city'));if(!S.by[c])continue;
      const t=streetTokens(streetOf(r));st=S.by[c].find(x=>sameStreet(t,x.t,S.common[c]));if(st)break;
    }
    if(!st)continue; // 4) fora das listas: GPS da Shopee sem mexer
    if(st.e.ref){
      if(shopee&&dist(shopee,st.e.ref)<=st.e.raio)continue; // GPS da Shopee dentro do raio: fica como veio
      s.gps={tipo:'ref',coord:st.e.ref};tot.ref++;
    }else if(shopee){s.gps={tipo:'off'};tot.off++}
  }
  return tot;
}

// ----- tela: lista de ruas -----
let gEdit=-1,fEdit=-1;
const sayIn=id=>t=>{const m=$(id);m.textContent=t;m.style.display=t?'block':'none'};
const gSay=sayIn('gMsg'),fSay=sayIn('fMsg'),xSay=sayIn('xMsg');
function drawGps(){
  $('gN').textContent=badStreets.length;
  $('gList').innerHTML=badStreets.map((e,i)=>`<li><span>${esc(e.rua)}<small>${esc(e.cidade||'sem cidade')} · ${e.ref?`ponto ${esc(fmtC(e.ref))} (raio ${e.raio} m)`:'sem ponto: apaga o GPS'}</small></span><span class="lbtn"><button data-e="${i}" aria-label="Editar ${esc(e.rua)}">Editar</button><button data-i="${i}" aria-label="Remover ${esc(e.rua)}">Remover</button></span></li>`).join('')||'<li class="hint">Nenhuma rua cadastrada.</li>';
  $('fN').textContent=fixes.length;
  $('fList').innerHTML=fixes.map((e,i)=>`<li><span>${esc(e.rua)}, ${esc(e.numero)}<small>${esc(e.cidade||'sem cidade')} · ${esc(fmtC(e.coord))}</small></span><span class="lbtn"><button data-e="${i}" aria-label="Editar ${esc(e.rua)}">Editar</button><button data-i="${i}" aria-label="Remover ${esc(e.rua)}">Remover</button></span></li>`).join('')||'<li class="hint">Nenhum endereço cadastrado.</li>';
}
const lastCity=()=>{try{return localStorage.getItem(CITY_KEY)||'Paiçandu'}catch(e){return 'Paiçandu'}};
const rememberCity=c=>{try{localStorage.setItem(CITY_KEY,c)}catch(e){}};
function gFormReset(){gEdit=-1;$('gFormT').textContent='Adicionar rua';$('gSave').textContent='Adicionar';$('gCancel').hidden=true;$('gIn').value='';$('gRef').value='';$('gRaio').value=RAIO_PADRAO;$('gCity').value=lastCity()}
function saveStreet(){
  const rua=clip($('gIn').value),cidade=clip($('gCity').value),refTxt=$('gRef').value.trim(),raio=Math.round(+$('gRaio').value);
  const t=streetTokens(rua);
  if(!t.length){gSay('Escreva o nome da rua ou as palavras principais dela.');$('gIn').focus();return}
  if(!norm(cidade)){gSay('Escreva também a cidade da rua (igual à coluna City da planilha).');$('gCity').focus();return}
  let ref=null;if(refTxt){const r=readCoord(refTxt);if(r.err){gSay(r.err);$('gRef').focus();return}ref=r.c}
  if(!(raio>0)){gSay('O raio precisa ser um número de metros, por exemplo 500.');$('gRaio').focus();return}
  const e={rua,cidade,ref,raio:Math.min(50000,raio)};
  if(badStreets.some((x,i)=>i!==gEdit&&entryKey(x)===entryKey(e))){gSay('Essa rua já está na lista. Toque em Editar nela para mudar.');return}
  const editing=gEdit>=0;
  if(editing)badStreets[gEdit]=e;else badStreets.push(e);
  saveBad();rememberCity(cidade);drawGps();gFormReset();
  const word=rua.split(' ').find(w=>norm(w)===t[0])||rua;
  gSay((editing?'Salvo: ':'Adicionada: ')+`${rua} (${cidade})`+(ref?` com ponto de referência e raio de ${e.raio} m.`:', sem ponto: o GPS dessa rua será apagado.')+(t.length<2?` Atenção: com uma palavra só, vale para qualquer rua com “${word}” no nome em ${cidade}.`:''));
  if(stops.length)render();
}
$('gSave').addEventListener('click',saveStreet);
$('gCancel').addEventListener('click',()=>{gFormReset();gSay('')});
for(const id of ['gIn','gCity','gRef','gRaio'])$(id).addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();saveStreet()}});
$('gList').addEventListener('click',e=>{
  const ed=e.target.closest('button[data-e]'),rm=e.target.closest('button[data-i]');
  if(ed){const i=+ed.dataset.e,x=badStreets[i];gEdit=i;$('gIn').value=x.rua;$('gCity').value=x.cidade;$('gRef').value=fmtC(x.ref);$('gRaio').value=x.raio;
    $('gFormT').textContent='Editando: '+x.rua;$('gSave').textContent='Salvar';$('gCancel').hidden=false;gSay('');$('gRef').scrollIntoView({block:'center'});$('gRef').focus();return}
  if(rm){const i=+rm.dataset.i;const [gone]=badStreets.splice(i,1);if(gEdit===i)gFormReset();else if(gEdit>i)gEdit--;saveBad();drawGps();gSay(`Removida: ${gone.rua} (${gone.cidade})`);if(stops.length)render()}
});
$('gReset').addEventListener('click',()=>{
  const have=new Set(badStreets.map(entryKey));let n=0;
  for(const e of GPS_DEFAULT)if(!have.has(entryKey(e))){badStreets.push({...e});n++}
  saveBad();drawGps();gSay(n?`${n} rua${n>1?'s':''} da lista inicial recolocada${n>1?'s':''}.`:'Todas as ruas da lista inicial já estão aí.');if(stops.length)render();
});

// ----- tela: endereços corrigidos -----
function fFormReset(){fEdit=-1;$('fFormT').textContent='Adicionar endereço';$('fSave').textContent='Adicionar';$('fCancel').hidden=true;$('fRua').value='';$('fNum').value='';$('fCoord').value='';$('fCity').value=lastCity()}
function saveFix(){
  const rua=clip($('fRua').value),numero=clip($('fNum').value,20),cidade=clip($('fCity').value),cTxt=$('fCoord').value.trim();
  if(!streetTokens(rua).length){fSay('Escreva o nome da rua.');$('fRua').focus();return}
  if(!numKey(numero)){fSay('Escreva o número da casa.');$('fNum').focus();return}
  if(!norm(cidade)){fSay('Escreva a cidade (igual à coluna City da planilha).');$('fCity').focus();return}
  const r=readCoord(cTxt);if(r.err){fSay(r.err);$('fCoord').focus();return}
  const e={rua,numero,cidade,coord:r.c};
  if(fixes.some((x,i)=>i!==fEdit&&fixKey(x)===fixKey(e))){fSay('Esse endereço já está na lista. Toque em Editar nele para mudar.');return}
  const editing=fEdit>=0;
  if(editing)fixes[fEdit]=e;else fixes.push(e);
  storeFix();rememberCity(cidade);drawGps();fFormReset();
  fSay(`${editing?'Salvo':'Adicionado'}: ${rua}, ${numero} (${cidade}) → ${fmtC(r.c)}`);
  if(stops.length)render();
}
$('fSave').addEventListener('click',saveFix);
$('fCancel').addEventListener('click',()=>{fFormReset();fSay('')});
for(const id of ['fRua','fNum','fCity','fCoord'])$(id).addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();saveFix()}});
$('fList').addEventListener('click',e=>{
  const ed=e.target.closest('button[data-e]'),rm=e.target.closest('button[data-i]');
  if(ed){const i=+ed.dataset.e,x=fixes[i];fEdit=i;$('fRua').value=x.rua;$('fNum').value=x.numero;$('fCity').value=x.cidade;$('fCoord').value=fmtC(x.coord);
    $('fFormT').textContent=`Editando: ${x.rua}, ${x.numero}`;$('fSave').textContent='Salvar';$('fCancel').hidden=false;fSay('');$('fCoord').scrollIntoView({block:'center'});$('fCoord').focus();return}
  if(rm){const i=+rm.dataset.i;const [gone]=fixes.splice(i,1);if(fEdit===i)fFormReset();else if(fEdit>i)fEdit--;storeFix();drawGps();fSay(`Removido: ${gone.rua}, ${gone.numero} (${gone.cidade})`);if(stops.length)render()}
});

// ----- passar as listas para colegas (arquivo .json; nada vai para a internet) -----
const LIST_FILE='listas-gps-rotas-otimizadas.json';
function listsFile(){
  const data={app:'rotas-otimizadas',tipo:'listas-gps',versao:1,exportado:new Date().toISOString().slice(0,10),ruas:badStreets,enderecos:fixes};
  return new File([JSON.stringify(data,null,1)],LIST_FILE,{type:'application/json'});
}
$('xOut').addEventListener('click',()=>{
  const url=URL.createObjectURL(listsFile());
  const a=document.createElement('a');a.href=url;a.download=LIST_FILE;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),60000);
  xSay(`Arquivo baixado: ${LIST_FILE} (${badStreets.length} rua${badStreets.length===1?'':'s'}, ${fixes.length} endereço${fixes.length===1?'':'s'}). Procure em Downloads e mande para o colega.`);
});
$('xShare').addEventListener('click',async()=>{
  try{await navigator.share({files:[listsFile()],title:'Listas de GPS — Rotas Otimizadas'});xSay('Listas compartilhadas.')}
  catch(e){if(e&&e.name!=='AbortError')xSay('Não deu para compartilhar daqui. Use “Baixar arquivo das listas”.')}
});
try{const t=new File(['{}'],'a.json',{type:'application/json'});if(!(navigator.canShare&&navigator.canShare({files:[t]})))$('xShare').hidden=true}catch(e){$('xShare').hidden=true}
$('xIn').addEventListener('change',async e=>{
  const f=e.target.files[0];e.target.value='';if(!f)return;
  try{
    if(f.size>2e6)throw new Error('grande');
    const o=JSON.parse(await f.text());
    if(!o||o.app!=='rotas-otimizadas'||(!Array.isArray(o.ruas)&&!Array.isArray(o.enderecos)))throw new Error('formato');
    let rn=0,ru=0,fn=0,fu=0;
    for(const x of (o.ruas||[]).map(cleanStreetEntry).filter(Boolean)){const i=badStreets.findIndex(y=>entryKey(y)===entryKey(x));if(i>=0){badStreets[i]=x;ru++}else{badStreets.push(x);rn++}}
    for(const x of (o.enderecos||[]).map(cleanFixEntry).filter(Boolean)){const i=fixes.findIndex(y=>fixKey(y)===fixKey(x));if(i>=0){fixes[i]=x;fu++}else{fixes.push(x);fn++}}
    saveBad();storeFix();gFormReset();fFormReset();drawGps();
    xSay(`Listas importadas: ${rn} rua${rn===1?'':'s'} nova${rn===1?'':'s'}, ${ru} atualizada${ru===1?'':'s'}; ${fn} endereço${fn===1?'':'s'} novo${fn===1?'':'s'}, ${fu} atualizado${fu===1?'':'s'}.`);
    if(stops.length)render();
  }catch(err){xSay('Esse arquivo não é de listas do Rotas Otimizadas. Escolha o arquivo “'+LIST_FILE+'” que o colega mandou.')}
});
gFormReset();fFormReset();drawGps();

function render(){
  $('out').hidden=false;
  $('sPk').textContent=stops.reduce((t,s)=>t+s.items.length,0);
  $('sSt').textContent=stops.length;
  const g=computeGps(),total=g.fix+g.ref+g.off;
  $('sGps').textContent=total;
  $('sGpsInfo').textContent=total?'Ajuste de GPS: '+[g.fix?`${g.fix} endereço${g.fix>1?'s':''} corrigido${g.fix>1?'s':''}`:'',g.ref?`${g.ref} pelo ponto da rua`:'',g.off?`${g.off} removido${g.off>1?'s':''}`:''].filter(Boolean).join(' · ')+'. As outras linhas ficaram com o GPS da Shopee.':'Nenhum ajuste de GPS: todas as linhas ficaram com o GPS da Shopee.';
  $('sFile').textContent=srcName+' → '+fname+'-otimizada.xlsx';
}

function load(aoa,name){
  aoa=aoa.filter(r=>r&&r.some(c=>String(c??'').trim()!==''));
  if(aoa.length<2){say('A planilha parece vazia.');$('out').hidden=true;return}
  let h=0;for(let i=0;i<Math.min(10,aoa.length);i++){if(aoa[i].filter(c=>typeof c==='string'&&c.trim()).length>=3){h=i;break}}
  headers=aoa[h].map(c=>String(c??'').trim());rows=aoa.slice(h+1);
  srcName=name||'rota';fname=srcName.replace(/\.[^.]+$/,'');
  guess();buildSel();group();
}

async function readFile(f){
  if(!f)return;
  say('Lendo a planilha…');
  try{
    const buf=await f.arrayBuffer();const u=new Uint8Array(buf);
    const isZip=u[0]===0x50&&u[1]===0x4B;
    if(isZip&&typeof DecompressionStream==='undefined'){say('Este navegador é antigo para ler .xlsx. Atualize o Chrome ou abra no computador.');return}
    const aoa=isZip?await readXlsx(buf):readCsv(new TextDecoder().decode(u));
    say('');load(aoa,f.name);
  }catch(err){say('Não consegui ler esse arquivo. Ele é a planilha .xlsx ou .csv da rota?')}
}
$('file').addEventListener('change',e=>{readFile(e.target.files[0]);e.target.value=''});
const dz=$('drop');
['dragenter','dragover'].forEach(t=>dz.addEventListener(t,e=>{e.preventDefault();dz.classList.add('on')}));
['dragleave','drop'].forEach(t=>dz.addEventListener(t,e=>{e.preventDefault();dz.classList.remove('on')}));
dz.addEventListener('drop',e=>readFile(e.dataTransfer.files[0]));
$('useComp').addEventListener('change',group);

// colunas extras no fim: Qtd pacotes, Complementos, Ajuste de GPS, Endereço completo, Endereços originais
const EXTRA=['Qtd pacotes','Complementos','Ajuste de GPS','Endereço completo','Endereços originais'];
const OLD_EXTRA=['GPS removido']; // nome antigo da coluna: sai se a planilha já tinha sido otimizada
function lean(){
  computeGps();
  const cell=(s,i)=>{
    if(s.gps&&(i===map.lat||i===map.lon))return s.gps.coord?s.gps.coord[i===map.lat?0:1]:'';
    if(i===map.addr)return cleanStreet(s);
    return merged(s,i);
  };
  // se a planilha já foi otimizada antes, as colunas extras antigas saem e entram as novas (sem repetir)
  const extraN=new Set([...EXTRA,...OLD_EXTRA].map(norm)),keep=headers.map((h,i)=>i).filter(i=>!extraN.has(norm(headers[i])));
  return [[...keep.map(i=>headers[i]||'Coluna '+(i+1)),...EXTRA],
    ...stops.map(s=>[...keep.map(i=>cell(s,i)),s.items.length,complements(s),s.gps?GPS_LABEL[s.gps.tipo]:'',s.gps&&s.gps.tipo==='off'?fullAddr(s):'',originals(s)])];
}
function leanBlob(){
  const L=lean(),n=L[0].length-EXTRA.length;
  const w={'Qtd pacotes':12,'Complementos':40,'Ajuste de GPS':20,'Endereço completo':52,'Endereços originais':70};
  return writeXlsx(L,L[0].map((h,i)=>i>=n?w[h]:i===map.addr?40:i===map.code?40:i===map.seq?18:14));
}
$('dl').addEventListener('click',()=>{
  if(!stops.length)return;
  try{
    const url=URL.createObjectURL(leanBlob());
    const a=document.createElement('a');a.href=url;a.download=fname+'-otimizada.xlsx';document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),60000);
    say('Planilha otimizada baixada. Procure em Downloads: '+fname+'-otimizada.xlsx');
  }catch(e){say('Não consegui baixar a planilha.')}
});
$('sh').addEventListener('click',async()=>{
  if(!stops.length)return;
  const file=new File([leanBlob()],fname+'-otimizada.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  try{await navigator.share({files:[file],title:'Rota otimizada'})}catch(e){if(e&&e.name!=='AbortError')say('Não deu para compartilhar daqui. Use “Baixar planilha otimizada”.')}
});
try{const t=new File([''],'a.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});if(!(navigator.canShare&&navigator.canShare({files:[t]})))$('sh').hidden=true}catch(e){$('sh').hidden=true}

// app instalável: guarda tudo no celular para abrir sem internet (só quando aberto pelo link, não como arquivo)
if('serviceWorker' in navigator&&/^https?:$/.test(location.protocol))addEventListener('load',()=>navigator.serviceWorker.register('sw.js').catch(()=>{}));
}
