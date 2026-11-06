/* Test Bank Generator document import. Runs locally; never uploads documents. */
(function(global){
'use strict';
const cp1252=new TextDecoder('windows-1252'),utf8=new TextDecoder(),utf16=new TextDecoder('utf-16le');
const newId=()=>crypto.randomUUID?crypto.randomUUID():Date.now().toString(36)+Math.random().toString(36).slice(2);
const local=(node,name)=>Array.from(node.getElementsByTagName('*')).filter(n=>n.localName===name);
const attr=(n,k)=>n?.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main',k)||n?.getAttribute('w:'+k)||n?.getAttribute(k)||'';
const linesOf=text=>String(text).replace(/\r\n?/g,'\n').split('\n').map(text=>({text}));
function normalizeText(s){return String(s).replace(/\u00a0/g,' ').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').trim()}
function parseQuestions(blocks,name,options={}){
 const result=[],preamble=[],keys=new Map(),keyImages=new Map();let current=null,choice=null,keyMode=false,lastKey=null,section='';
 const push=()=>{if(!current)return;if(!current.stem.trim()){current=null;return}current.stem=current.stem.trim();current.choices=current.choices.filter(a=>a.text.trim());if(current.choices.length>=2){current.type='mc';current.choices.sort((a,b)=>a.label.localeCompare(b.label))}else{current.type='written';if(current.choices.length===1){current.stem+='\n'+current.choices[0].label+'. '+current.choices[0].text;current.choices=[];current.flags.push('A single answer choice was found. Check the question split.')}}result.push(current);current=null;choice=null};
 const answerFor=(q,value)=>{const v=value.trim();const letter=v.match(/^([a-h])(?:[.)]|\s|$)/i);if(q.type==='mc'&&letter&&q.choices.some(a=>a.label===letter[1].toLowerCase()))q.answer=letter[1].toLowerCase();else if(q.type==='mc'){const a=q.choices.find(a=>a.text.trim().toLowerCase()===v.toLowerCase());if(a)q.answer=a.label;else q.flags.push('An answer was found but could not be matched to a choice.')}else q.answer=v};
 for(const block of blocks){
  const images=block.images||[];let text=normalizeText(block.text||'');if(!text){if(keyMode&&lastKey&&images.length)keyImages.set(lastKey,[...(keyImages.get(lastKey)||[]),...images]);else if(current&&images.length)current.images.push(...images);continue}
  if(/^(?:answer\s*key|answers|teacher(?:'s)?\s+(?:key|answers))\s*:?(?:\s*[-–]\s*.*)?$/i.test(text)){push();keyMode=true;continue}
  if(keyMode){const matches=[...text.matchAll(/(?:^|\s)(\d{1,4})\s*[.):\-]\s*([\s\S]*?)(?=\s+\d{1,4}\s*[.):\-]\s*|$)/g)];for(const m of matches){lastKey=Number(m[1]);keys.set(lastKey,m[2].trim())}if(!matches.length&&lastKey)keys.set(lastKey,(keys.get(lastKey)||'')+' '+text);if(lastKey&&images.length)keyImages.set(lastKey,[...(keyImages.get(lastKey)||[]),...images]);continue}
  if(/^(?:multiple choice|short answer|essay|completion|fill[- ]in[- ]the[- ]blank|true\s*[/–-]\s*false|problems?)\s*:?(?:\s*\([^)]*\))?$/i.test(text)){push();section=text;continue}
  const q=text.match(/^\s*(?:_{2,}\s*)?(?:(?:question|q)\s*)?(\d{1,4})\s*[.):]\s+([\s\S]*)$/i);
  if(q){push();current={id:newId(),source:name,sourceNumber:Number(q[1]),sourceFormat:options.format||'document',sourcePage:block.page||null,category:name.replace(/\.(docx?|pdf)$/i,''),tags:'',type:'written',stem:'',choices:[],answer:'',images:[...images],referenceImages:[],flags:['Imported from '+(options.format||'document').toUpperCase()+'. Check question boundaries, choices, and the answer key.'],reviewed:false,points:1};text=q[2];choice=null;
   if(/true\s*[/–-]\s*false/i.test(section))current.choices=[{label:'a',text:'True'},{label:'b',text:'False'}];
  }else if(!current){preamble.push(text);continue}
  else if(images.length)current.images.push(...images);
  if(!current)continue;
  const answer=text.match(/^(?:correct\s+)?answer\s*:\s*(.+)$/i);if(answer){current.pendingAnswer=answer[1];continue}
  // Split lettered choices on one line, including choices laid out in columns.
  const chunks=[...text.matchAll(/(?:^|\s{2,}|\t|(?<=\S)\s+)(\*?)([a-h])\s*[.)]\s+(\*?)([\s\S]*?)(?=(?:\s{2,}|\t|\s+)[*]?[a-h]\s*[.)]\s+|$)/gi)];
  if(chunks.length&&(chunks[0].index===0||q)){
   const prefix=text.slice(0,chunks[0].index).trim();if(prefix)current.stem+=(current.stem?'\n':'')+prefix;
   for(const m of chunks){choice={label:m[2].toLowerCase(),text:m[4].trim()};current.choices.push(choice);if(m[1]||m[3])current.pendingAnswer=choice.label}
  }else if(choice)choice.text+=(choice.text?' ':'')+text;
  else current.stem+=(current.stem?'\n':'')+text;
 }
 push();
 const counts=new Map();for(const q of result){counts.set(q.sourceNumber,(counts.get(q.sourceNumber)||0)+1);const answer=q.pendingAnswer||keys.get(q.sourceNumber);if(answer)answerFor(q,answer);q.referenceImages=keyImages.get(q.sourceNumber)||[];delete q.pendingAnswer;if(q.type==='mc'&&!q.answer)q.flags.push('No correct answer was identified. Select one before grading exports.');if(options.warning)q.flags.push(options.warning);if(/\[Unrecovered/.test(q.stem))q.flags.push('This question contains an unrecovered embedded object. Repair it before use.')}
 for(const q of result)if(counts.get(q.sourceNumber)>1)q.flags.push('This question number appears more than once. Check its answer-key association.');
 return {questions:result,expected:result.length,preamble:preamble.join('\n'),blocks};
}
function parseXML(s){const doc=new DOMParser().parseFromString(s,'application/xml');if(doc.getElementsByTagName('parsererror').length)throw Error('A Word XML part is malformed.');return doc}
function toBase64(bytes){let s='';for(let i=0;i<bytes.length;i+=16384)s+=String.fromCharCode(...bytes.subarray(i,i+16384));return btoa(s)}
async function pngImage(bytes,mime){
 const image=new Image(),url='data:'+mime+';base64,'+toBase64(bytes);
 await new Promise((resolve,reject)=>{image.onload=resolve;image.onerror=()=>reject(Error('The browser could not decode an embedded image.'));image.src=url});
 if(!image.width||!image.height||image.width*image.height>20000000)throw Error('An embedded image is too large.');
 const canvas=document.createElement('canvas');canvas.width=image.width;canvas.height=image.height;canvas.getContext('2d').drawImage(image,0,0);
 return {data:canvas.toDataURL('image/png'),width:image.width,height:image.height};
}
async function readDocx(bytes){
 if(!global.JSZip)throw Error('The bundled Word importer is missing. Extract the entire ZIP, including vendor files.');
 const zip=await global.JSZip.loadAsync(bytes),part=zip.file('word/document.xml');if(!part)throw Error('This is not a valid Word .docx document.');
 if(Object.values(zip.files).reduce((sum,x)=>sum+(x._data?.uncompressedSize||0),0)>160*1024*1024)throw Error('The expanded Word document is too large.');
 const doc=parseXML(await part.async('string')),relations=new Map(),rels=zip.file('word/_rels/document.xml.rels');
 if(rels)for(const n of local(parseXML(await rels.async('string')),'Relationship'))if(n.getAttribute('TargetMode')!=='External'){
  let path=n.getAttribute('Target')||'';if(path.startsWith('/'))path=path.slice(1);else path='word/'+path;const pieces=[];for(const p of path.split('/'))if(p==='..')pieces.pop();else if(p!=='.')pieces.push(p);relations.set(n.getAttribute('Id'),pieces.join('/'));
 }
 const levels=new Map(),nums=new Map(),counters=new Map(),numberPart=zip.file('word/numbering.xml');
 if(numberPart){const nd=parseXML(await numberPart.async('string'));for(const a of local(nd,'abstractNum'))for(const l of local(a,'lvl'))levels.set(attr(a,'abstractNumId')+':'+attr(l,'ilvl'),{format:attr(local(l,'numFmt')[0],'val'),start:Number(attr(local(l,'start')[0],'val')||1)});for(const n of local(nd,'num'))nums.set(attr(n,'numId'),attr(local(n,'abstractNumId')[0],'val'))}
 const styles=new Map(),stylePart=zip.file('word/styles.xml');
 if(stylePart)for(const st of local(parseXML(await stylePart.async('string')),'style'))styles.set(attr(st,'styleId'),st);
 const styleNumbering=p=>{let id=attr(local(p,'pStyle')[0],'val'),seen=new Set();while(id&&!seen.has(id)){seen.add(id);const st=styles.get(id);if(!st)break;const n=local(st,'numPr')[0];if(n)return n;id=attr(local(st,'basedOn')[0],'val')}return null};
 const blocks=[],imageCache=new Map();let warning='';
 for(const p of local(doc,'p')){
  // XML math is not flattened into plausible-looking but incomplete equations.
  let text='';const images=[];
  for(const n of Array.from(p.getElementsByTagName('*'))){
   if(n.localName==='t'&&n.namespaceURI==='http://schemas.openxmlformats.org/wordprocessingml/2006/main'){
    let t=n.textContent;const run=n.parentElement,vert=local(run,'vertAlign')[0],val=attr(vert,'val');
    if(['superscript','subscript'].includes(val)){const digits=val==='superscript'?'⁰¹²³⁴⁵⁶⁷⁸⁹':'₀₁₂₃₄₅₆₇₈₉';t=t.replace(/\d/g,c=>digits[Number(c)])}text+=t;
   }else if(n.localName==='tab')text+='\t';else if(n.localName==='br')text+='\n';
   else if(n.localName==='oMath')text+=' [Unrecovered Word equation — review original] ';
   else if(n.localName==='blip'||n.localName==='imagedata'){
    const id=n.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships',n.localName==='blip'?'embed':'id'),path=relations.get(id),file=path&&zip.file(path);
    if(!file){text+=' [Unrecovered linked image] ';continue}
    if(!imageCache.has(path))imageCache.set(path,(async()=>{const mime=/\.png$/i.test(path)?'image/png':/\.jpe?g$/i.test(path)?'image/jpeg':/\.gif$/i.test(path)?'image/gif':/\.bmp$/i.test(path)?'image/bmp':null;if(!mime)throw Error('Unsupported Word image format');return pngImage(await file.async('uint8array'),mime)})());
    try{images.push(await imageCache.get(path))}catch(e){text+=' [Unrecovered Word image — review original] ';warning='Some Word images or equation objects could not be decoded.'}
   }
  }
  const numPr=local(p,'numPr')[0]||styleNumbering(p);if(numPr){const id=attr(local(numPr,'numId')[0],'val'),level=attr(local(numPr,'ilvl')[0],'val')||'0',spec=levels.get(nums.get(id)+':'+level)||{format:'decimal',start:1},key=id+':'+level;const value=counters.has(key)?counters.get(key)+1:spec.start;counters.set(key,value);for(const k of counters.keys())if(k.startsWith(id+':')&&Number(k.split(':')[1])>Number(level))counters.delete(k);if(!/^\s*\d+[.)]/.test(text)&&['decimal','lowerLetter','upperLetter'].includes(spec.format))text=(spec.format==='decimal'?value:String.fromCharCode((spec.format==='lowerLetter'?97:65)+(value-1)%26))+'. '+text}
  const lines=text.split('\n');lines.forEach((line,i)=>blocks.push({text:line,images:i===lines.length-1?images:[]}));
 }
 return {blocks,assets:blocks.flatMap(b=>b.images||[]).map((im,i)=>({...im,label:'Word image '+(i+1)})),warning};
}
function cfbStreams(bytes){
 const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),u32=(o)=>{if(o<0||o+4>bytes.length)throw Error('Corrupt Word file.');return v.getUint32(o,true)};
 if(bytes.length<512||u32(0)!==0xe011cfd0||u32(4)!==0xe11ab1a1)throw Error('Unsupported .doc format.');
 const shift=v.getUint16(30,true),miniShift=v.getUint16(32,true);if(![9,12].includes(shift)||miniShift!==6)throw Error('Unsupported Word compound-file version.');const sector=2**shift,mini=2**miniShift,maxSectors=Math.floor(bytes.length/sector),end=0xfffffffe;
 const sec=id=>{if(id>=maxSectors-1)throw Error('Word file contains an invalid sector.');return bytes.subarray((id+1)*sector,(id+2)*sector)};
 const difat=[];for(let i=0;i<109;i++){const id=u32(76+i*4);if(id<0xfffffffa)difat.push(id)}let next=u32(68),seen=new Set();for(let n=0;n<u32(72);n++){if(seen.has(next)||n>maxSectors)throw Error('Corrupt Word allocation table.');seen.add(next);const a=sec(next),w=new DataView(a.buffer,a.byteOffset,a.byteLength);for(let i=0;i<sector/4-1;i++){const id=w.getUint32(i*4,true);if(id<0xfffffffa)difat.push(id)}next=w.getUint32(sector-4,true)}
 const fat=[];for(const id of difat.slice(0,u32(44))){const a=sec(id),w=new DataView(a.buffer,a.byteOffset,a.byteLength);for(let i=0;i<sector;i+=4)fat.push(w.getUint32(i,true))}
 const chain=(start,table,get,sizeLimit=bytes.length)=>{let id=start,parts=[],size=0,seen=new Set();while(id!==end&&id!==0xffffffff){if(seen.has(id)||id>=table.length)throw Error('Corrupt Word stream chain.');seen.add(id);const a=get(id);parts.push(a);size+=a.length;if(size>Math.max(sizeLimit,sector)+sector)throw Error('Word stream is too large.');id=table[id]}const out=new Uint8Array(size);let at=0;for(const a of parts){out.set(a,at);at+=a.length}return out};
 const directory=chain(u32(48),fat,sec),entries=[];for(let i=0;i+128<=directory.length;i+=128){const d=new DataView(directory.buffer,i,128),len=d.getUint16(64,true),type=directory[i+66];if(!type||len<2||len>64)continue;entries.push({name:utf16.decode(directory.subarray(i,i+len-2)),type,start:d.getUint32(116,true),size:d.getUint32(120,true)})}
 const root=entries.find(e=>e.type===5);if(!root)throw Error('Word root stream is missing.');const miniStream=chain(root.start,fat,sec).subarray(0,root.size),miniFatBytes=u32(64)?chain(u32(60),fat,sec):new Uint8Array(),miniFat=[];
 const mv=new DataView(miniFatBytes.buffer);for(let i=0;i<miniFatBytes.length;i+=4)miniFat.push(mv.getUint32(i,true));
 const streams=new Map();for(const e of entries.filter(e=>e.type===2)){if(e.size>bytes.length)throw Error('Word stream size is invalid.');const a=e.size<u32(56)?chain(e.start,miniFat,id=>{if((id+1)*mini>miniStream.length)throw Error('Invalid Word mini stream.');return miniStream.subarray(id*mini,(id+1)*mini)}):chain(e.start,fat,sec);streams.set(e.name,a.subarray(0,e.size))}return streams;
}
function wordBinaryText(bytes){
 const streams=cfbStreams(bytes),word=streams.get('WordDocument');if(!word||word.length<432)throw Error('This .doc file has no readable Word document stream.');const v=new DataView(word.buffer,word.byteOffset,word.byteLength),flags=v.getUint16(10,true);
 if(flags&0x8100)throw Error('Password-protected Word files are not supported. Save an unprotected copy first.');const table=streams.get(flags&0x200?'1Table':'0Table');if(!table)throw Error('The Word piece table is missing.');
 const csw=v.getUint16(32,true),lw=34+csw*2,cslw=v.getUint16(lw,true),rgLw=lw+2,pairs=rgLw+cslw*4+2;if(pairs+272>word.length)throw Error('Unsupported legacy Word layout.');
 const ccpText=v.getUint32(rgLw+12,true),fc=v.getUint32(pairs+33*8,true),lcb=v.getUint32(pairs+33*8+4,true);if(!lcb||fc+lcb>table.length)throw Error('Unsupported or corrupt Word text table.');
 let at=fc;const tv=new DataView(table.buffer,table.byteOffset,table.byteLength);while(at<fc+lcb&&table[at]===1){at+=3+tv.getUint16(at+1,true)}if(table[at]!==2)throw Error('Word text pieces could not be located.');const len=tv.getUint32(at+1,true),base=at+5,count=(len-4)/12;if(!Number.isInteger(count)||count<1||base+len>table.length)throw Error('Invalid Word text pieces.');
 let text='';for(let i=0;i<count;i++){const begin=tv.getUint32(base+i*4,true),end=Math.min(tv.getUint32(base+(i+1)*4,true),ccpText);if(begin>=ccpText)break;const raw=tv.getUint32(base+(count+1)*4+i*8+2,true),compressed=!!(raw&0x40000000),off=(raw&0x3fffffff)/(compressed?2:1),size=(end-begin)*(compressed?1:2);if(off+size>word.length)throw Error('Word text points outside the file.');text+=(compressed?cp1252:utf16).decode(word.subarray(off,off+size))}
 // Keep displayed field results; discard instruction strings such as PAGE.
 text=text.replace(/\x13[^\x14\x15]*\x14/g,'').replace(/\x13[^\x15]*\x15/g,'').replace(/\x15/g,'').replace(/[\x01\x08]/g,' [Unrecovered Word object — review original] ').replace(/\x07/g,'\t').replace(/[\x0b\x0c]/g,'\r');
 return {blocks:linesOf(text),assets:[],warning:'Legacy .doc text was extracted. Verify numbering and formatting; embedded diagrams and equations may need attachment from the original.'};
}
function readRTF(bytes){
 const input=cp1252.decode(bytes),stack=[];let skip=false,uc=1,output='',fallback=0;
 const ignored=new Set(['fonttbl','colortbl','stylesheet','info','pict','object','header','footer','listtable','listoverridetable','fldinst','datastore','themedata']);
 for(let i=0;i<input.length;){const c=input[i++];if(c==='{'){stack.push({skip,uc});continue}if(c==='}'){const old=stack.pop();if(old){skip=old.skip;uc=old.uc}continue}if(c!=='\\'){if(fallback>0){fallback--;continue}if(!skip&&c!=='\r'&&c!=='\n')output+=c;continue}
  const rest=input.slice(i);let m;if((m=rest.match(/^'([a-f\d]{2})/i))){i+=3;if(fallback)fallback--;else if(!skip)output+=cp1252.decode(Uint8Array.of(parseInt(m[1],16)));continue}
  if(['\\','{','}'].includes(input[i])){if(!skip)output+=input[i];i++;continue}if(input[i]==='*'){skip=true;i++;continue}
  m=rest.match(/^([a-z]+)(-?\d+)? ?/i);if(!m){i++;continue}i+=m[0].length;const cmd=m[1],n=m[2]===undefined?null:Number(m[2]);if(ignored.has(cmd)){skip=true;continue}if(cmd==='uc'){uc=n;continue}if(cmd==='u'&&!skip){output+=String.fromCharCode(n<0?n+65536:n);fallback=uc;continue}if(cmd==='bin'&&n){i+=n;continue}if(skip)continue;if(cmd==='par'||cmd==='line')output+='\n';else if(cmd==='tab')output+='\t';else if(cmd==='emdash')output+='—';else if(cmd==='endash')output+='–';
 }
 return {blocks:linesOf(output),assets:[],warning:'RTF content in this .doc file was imported as text. Check diagrams and equation objects against the original.'};
}
async function readDoc(bytes){
 const head=cp1252.decode(bytes.subarray(0,400)).trimStart();if(head.startsWith('{\\rtf'))return readRTF(bytes);
 if(/^(?:<!doctype\s+html|<html|<\?xml)/i.test(head)){const d=new DOMParser().parseFromString(utf8.decode(bytes),'text/html');d.querySelectorAll('script,style').forEach(n=>n.remove());return {blocks:Array.from(d.querySelectorAll('p,tr,li,h1,h2,h3')).map(n=>({text:n.textContent})),assets:[],warning:'HTML content in this .doc file was imported as text. Check formatting and images.'}}
 return wordBinaryText(bytes);
}
async function readPDF(bytes){
 if(!global.TestBankPDF||!global.pdfjsWorker)throw Error('The bundled PDF importer is missing. Extract the entire ZIP, including vendor files.');
 const task=global.TestBankPDF.getDocument({data:bytes.slice(),isEvalSupported:false,useSystemFonts:true,disableFontFace:true,useWasm:false,verbosity:0});let pdf;
 try{pdf=await task.promise}catch(e){throw Error(/password/i.test(e.message)?'Password-protected PDFs are not supported. Save an unprotected copy.':'The PDF could not be read: '+e.message)}
 try{
 if(pdf.numPages>100)throw Error('Import up to 100 PDF pages at a time.');const blocks=[],assets=[];let scanned=0,previewFailures=0;
 for(let pageNumber=1;pageNumber<=pdf.numPages;pageNumber++){
  const page=await pdf.getPage(pageNumber),content=await page.getTextContent(),view=page.getViewport({scale:1}),rows=[];
  for(const item of content.items){if(!item.str?.trim())continue;const tr=view.transform,tx=item.transform[4],ty=item.transform[5],x=tr[0]*tx+tr[2]*ty+tr[4],y=tr[1]*tx+tr[3]*ty+tr[5];let row=rows.find(r=>Math.abs(r.y-y)<Math.max(2,Math.abs(item.transform[3])*.22));if(!row){row={y,items:[]};rows.push(row)}row.items.push({x,text:item.str,width:item.width||0})}
  rows.sort((a,b)=>a.y-b.y);for(const row of rows){row.items.sort((a,b)=>a.x-b.x);let text='',last=null;for(const item of row.items){if(last)text+=(item.x-last.x-last.width>14?'\t':' ');text+=item.text;last=item}blocks.push({text,page:pageNumber})}
  if(!rows.length)scanned++;
  const ops=await page.getOperatorList(),graphicOps=new Set([global.TestBankPDF.OPS.paintImageXObject,global.TestBankPDF.OPS.paintInlineImageXObject,global.TestBankPDF.OPS.paintImageMaskXObject,global.TestBankPDF.OPS.constructPath]);
  if(ops.fnArray.some(x=>graphicOps.has(x))){try{const scale=Math.min(1.4,1000/view.width),viewport=page.getViewport({scale}),c=document.createElement('canvas');c.width=Math.ceil(viewport.width);c.height=Math.ceil(viewport.height);await page.render({canvasContext:c.getContext('2d'),viewport}).promise;assets.push({label:'PDF page '+pageNumber+' — original layout',page:pageNumber,data:c.toDataURL('image/png'),width:c.width,height:c.height})}catch(e){previewFailures++}}
 }
 if(!blocks.length)throw Error('No selectable text was found. This appears to be a scanned PDF. Run OCR on it first, then import the searchable PDF.');
 return {blocks,assets,warning:'PDF text order and question splits need review. '+(scanned?scanned+' page(s) had no selectable text. ':'')+(previewFailures?previewFailures+' source page preview(s) could not be rendered. ':'')+(assets.length?'Original page previews are available for checking and cropping diagrams.':'Diagrams or vector equations may require comparison with the original.')};
 }finally{await pdf.destroy()}
}
async function read(bytes,name){const format=name.split('.').pop().toLowerCase();let data;if(format==='docx')data=await readDocx(bytes);else if(format==='doc')data=await readDoc(bytes);else if(format==='pdf')data=await readPDF(bytes);else throw Error('Choose a .tst, .doc, .docx, or .pdf test file.');const parsed=parseQuestions(data.blocks,name,{format,warning:data.warning});if(!parsed.questions.length)throw Error('No numbered questions were found. Use question numbers such as 1. or 1), with lettered choices such as A). Unnumbered or complex layouts need editing first.');return {...parsed,assets:data.assets||[],format};}
 global.TestBankDocuments={read,parseQuestions,readDocx,readDoc,wordBinaryText,readPDF,cfbStreams};
})(globalThis);
