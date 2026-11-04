import {ACTIONS,preview} from './engine.js';
import {reshapeText} from './reshape.js';
import {connectOffice,readSelection,applyPatches,undoPatches,createWorksheet} from './office.js';

const $=id=>document.getElementById(id);
const escape=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const clone=value=>structuredClone(value);
const displayValue=value=>typeof value==='string'?JSON.stringify(value):(value===null?'（空白）':String(value));
const officeMode=new URLSearchParams(location.search).get('office')==='1';
let connected=false,view='reshape',category='全部',actionId='trim',shaped=null,pending=null,undo=null,snapshot=null,busy=false;
let data=[],formulas=[],sourceHistory=[];
const titles={reshape:'粘贴排成表格',tools:'常用工具箱',formulas:'公式小抄',help:'3 分钟上手'};
function toast(message){$('toast').textContent=message;$('toast').classList.add('show');clearTimeout(toast.timer);toast.timer=setTimeout(()=>$('toast').classList.remove('show'),5500);}
async function safely(fn){if(busy)return;busy=true;document.body.setAttribute('aria-busy','true');try{await fn();}catch(error){toast(error.message||'操作未完成，请重新读取数据后再试。');}finally{busy=false;document.body.removeAttribute('aria-busy');}}
function columnName(index){let out='';for(let i=index+1;i;i=Math.floor((i-1)/26))out=String.fromCharCode(65+(i-1)%26)+out;return out;}
function tableHTML(matrix,{editable=false,limit=80}={}){
 if(!matrix?.length)return '<div class="empty-message">▦<br><br>表格会出现在这里</div>';
 const cols=matrix[0].length,shown=Math.min(cols,20);
 let html='<table><thead><tr><th></th>'+Array.from({length:shown},(_,i)=>`<th>${columnName(i)}</th>`).join('')+'</tr></thead><tbody>';
 matrix.slice(0,limit).forEach((row,r)=>{html+=`<tr><td>${r+1}</td>`+row.slice(0,shown).map((v,c)=>`<td ${editable?`contenteditable="plaintext-only" data-r="${r}" data-c="${c}" aria-label="单元格 ${columnName(c)}${r+1}"`:''} class="${v===''||v===null?'empty':''}" title="${escape(v)}">${escape(v)}</td>`).join('')+'</tr>';});
 html+='</tbody></table>';if(matrix.length>limit||cols>shown)html+=`<p class="helper">预览前 ${Math.min(limit,matrix.length)} 行、${shown} 列；生成和导出包含全部 ${matrix.length} 行、${cols} 列。</p>`;return html;
}
function setView(next){view=next;document.querySelectorAll('.view').forEach(el=>el.hidden=el.id!==`view-${next}`);document.querySelectorAll('[data-view]').forEach(el=>el.classList.toggle('active',el.dataset.view===next));$('view-name').textContent=titles[next];if(next==='tools')renderData();}
document.querySelectorAll('[data-view]').forEach(el=>el.addEventListener('click',()=>setView(el.dataset.view)));

function renderShape(){
 shaped=null;const text=$('raw-data').value,mode=$('dimension-mode').value;
 $('source-count').textContent=`${text.length.toLocaleString()} 个字符`;$('columns-field').hidden=mode==='rows';$('rows-field').hidden=mode==='columns';
 const explicitRows=!!$('row-delimiter').value;document.querySelectorAll('input[name=fill]').forEach(el=>el.disabled=explicitRows);
 document.querySelectorAll('[data-delimiter]').forEach(el=>el.classList.toggle('selected',el.dataset.delimiter===$('delimiter').value));
 const message=$('reshape-message');message.className='notice';$('reshape-warnings').replaceChildren();
 try{
  if(!text){message.textContent='从左边粘贴内容开始，每一步都会即时显示结果。';$('grid-size').textContent='等待数据';$('reshape-grid').innerHTML=tableHTML([]);return;}
  const opts={delimiter:$('delimiter').value,rowDelimiter:$('row-delimiter').value,boundary:$('boundary-direction').value,fill:document.querySelector('input[name=fill]:checked').value,keepEmpty:$('keep-empty').checked};
  if(mode!=='rows')opts.columns=Number($('columns').value);if(mode!=='columns')opts.rows=Number($('rows').value);
  shaped=reshapeText(text,opts);
  message.textContent=`已找到 ${shaped.count} 项数据，排成 ${shaped.columns} 列 × ${shaped.rows} 行${shaped.padded?`，末尾补 ${shaped.padded} 个空格子`:''}。先确认下面的顺序是否正确。`;
  $('grid-size').textContent=`${shaped.rows} 行 × ${shaped.columns} 列`;$('reshape-grid').innerHTML=tableHTML(shaped.matrix);
  $('reshape-warnings').innerHTML=shaped.warnings.map(w=>`<p>· ${escape(w)}</p>`).join('');
 }catch(error){message.classList.add('error');message.textContent=error.message;$('grid-size').textContent='请调整设置';$('reshape-grid').innerHTML=tableHTML([]);}
 finally{for(const id of ['import-grid','copy-grid','download-grid'])$(id).disabled=!shaped||(officeMode&&!connected&&id==='import-grid');}
}
for(const id of ['raw-data','delimiter','row-delimiter','boundary-direction','columns','rows','dimension-mode','keep-empty'])$(id).addEventListener('input',renderShape);
document.querySelectorAll('input[name=fill]').forEach(el=>el.addEventListener('change',renderShape));
document.querySelectorAll('[data-delimiter]').forEach(el=>el.addEventListener('click',()=>{$('delimiter').value=el.dataset.delimiter;renderShape();}));
const example='姓名,部门,城市,小林,设计部,杭州,小周,运营部,上海,小陈,研发部,深圳,小吴,市场部,成都';
$('load-example').onclick=()=>{$('raw-data').value=example;$('delimiter').value=',';$('row-delimiter').value='';$('columns').value=3;$('dimension-mode').value='columns';$('keep-empty').checked=true;document.querySelector('input[name=fill][value=row]').checked=true;renderShape();};
$('clear-input').onclick=()=>{$('raw-data').value='';renderShape();$('raw-data').focus();};

function literalTSV(matrix){return matrix.map(row=>row.map(v=>{let s=String(v??'');if(typeof v==='string'&&s!=='')s="'"+s;return /[\t\r\n"]/.test(s)?'"'+s.replaceAll('"','""')+'"':s;}).join('\t')).join('\r\n');}
async function copy(text){if(!navigator.clipboard?.writeText)throw Error('浏览器未开放剪贴板，请使用下载 TSV，或手动选择表格复制。');await navigator.clipboard.writeText(text);}
function download(text,name,type){const url=URL.createObjectURL(new Blob([text],{type})),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
$('copy-grid').onclick=()=>safely(async()=>{if(!shaped)return;await copy(literalTSV(shaped.matrix));toast('表格已复制。到 Excel 中选一个空白单元格，再粘贴即可。文本已加保护前缀。');});
$('download-grid').onclick=()=>{if(shaped){download('\uFEFF'+literalTSV(shaped.matrix),'格子助手-排好的表格.tsv','text/tab-separated-values;charset=utf-8');toast('已下载 TSV。文本前缀保护编号和公式样内容；不同软件的导入表现可能不同。');}};
async function outputMatrix(matrix){
 if(!matrix?.length||!matrix[0]?.length)throw Error('没有可输出的数据，原表保持不变。');
 if(officeMode){if(!connected)throw Error('请先在 Excel 加载项中打开本页。');const result=await createWorksheet(matrix);toast(`已生成新工作表：${result.name}。原始数据没有改动。`);}
 else{if(data.length)sourceHistory.push({data:clone(data),formulas:clone(formulas)});data=clone(matrix);formulas=data.map(row=>row.map(()=>''));undo=null;$('undo').disabled=true;snapshot=null;renderData();setView('tools');toast('表格已放入网页工作表；可返回上一张来源表。刷新前记得导出。');}
 invalidatePreview();
}
$('import-grid').onclick=()=>safely(async()=>{if(shaped)await outputMatrix(shaped.matrix);});

function resetDemo(){data=[['姓名','部门','工号','报销金额','完成率'],['  小林  ',' 设计部','00128','1250.50','85%'],['小周','运  营部','00129',-62.556,'90%'],['小陈','研\u200b发部','123456789012345678',330,'100%'],['小周','运  营部','00129',-62.556,'90%'],['小吴','','00131',null,' 75% '],['  alice  ','ＭＡＲＫＥＴ','00132',92.439,'60%']];formulas=data.map(row=>row.map(()=>''));formulas[3][3]='=SUM(300,30)';undo=null;snapshot=null;$('undo').disabled=true;invalidatePreview();renderData();}
function renderData(){
 $('return-source').hidden=officeMode||!sourceHistory.length;
 $('range-label').textContent=officeMode?(snapshot?.address||'请先读取 Excel 选区'):`网页工作表 · ${data.length} 行 × ${data[0]?.length||0} 列`;
 $('data-grid').innerHTML=tableHTML(data,{editable:!officeMode,limit:40});
 if(!officeMode)$('data-grid').querySelectorAll('[contenteditable]').forEach(cell=>cell.addEventListener('blur',()=>{const r=Number(cell.dataset.r),c=Number(cell.dataset.c),text=cell.textContent;if(String(data[r][c]??'')===text)return;if(text.length>32767){cell.textContent=String(data[r][c]??'');toast('单元格最多 32,767 个字符，请缩短内容。');return;}data[r][c]=text;formulas[r][c]='';invalidatePreview();}));
}
$('return-source').onclick=()=>{const previous=sourceHistory.pop();if(previous){data=previous.data;formulas=previous.formulas;undo=null;$('undo').disabled=true;invalidatePreview();renderData();toast('已返回上一张来源表。');}};
$('copy-current').onclick=()=>safely(async()=>{if(!data.length)throw Error('还没有数据。');await copy(literalTSV(data));toast('已复制当前工作表的全部数据（不止预览部分）。');});
$('download-current').onclick=()=>{if(data.length)download('\uFEFF'+literalTSV(data),'格子助手-当前工作表.tsv','text/tab-separated-values;charset=utf-8');else toast('还没有数据。');};
$('reset-demo').onclick=()=>{resetDemo();toast('已恢复内置示例。');};
function invalidatePreview(){pending=null;$('apply-action').hidden=true;$('action-result').replaceChildren();}
function renderTools(){
 const cats=['全部',...new Set(ACTIONS.map(a=>a.category))];$('categories').innerHTML=cats.map(c=>`<button class="${category===c?'active':''}" data-category="${escape(c)}">${escape(c)}</button>`).join('');
 $('categories').querySelectorAll('button').forEach(el=>el.onclick=()=>{category=el.dataset.category;renderTools();});
 const query=$('tool-search').value.toLowerCase();const visible=ACTIONS.filter(a=>(category==='全部'||a.category===category)&&(a.title+a.description).toLowerCase().includes(query));
 $('tool-list').innerHTML=visible.map(a=>`<button class="tool-button ${a.id===actionId?'selected':''}" data-action="${a.id}"><b>${escape(a.title)}</b><small>${escape(a.description)}</small></button>`).join('')||'<p class="helper">没有找到。试试「空白」「拆列」或「数字」。</p>';
 $('tool-list').querySelectorAll('[data-action]').forEach(el=>el.onclick=()=>{actionId=el.dataset.action;renderTools();renderAction();});
}
function renderAction(){
 const action=ACTIONS.find(a=>a.id===actionId);$('action-title').textContent=action.title;$('action-description').textContent=action.description;
 const controls={find:'<label class="field">查找什么<input id="option-find" placeholder="要查找的内容"></label>',replace:'<label class="field">替换成<input id="option-replace" placeholder="留空表示删除"></label>',delimiter:'<label class="field">拆列分隔符<input id="option-delimiter" value=","><small>先在 Excel 中选一列。网页可先用排表向导生成单列数据。</small></label>',digits:'<label class="field">保留几位小数<input id="option-digits" type="number" value="2" min="-10" max="10"></label>'};
 $('action-options').innerHTML=(action.needs||[]).map(k=>controls[k]||'').join('');$('action-options').querySelectorAll('input').forEach(el=>el.oninput=invalidatePreview);invalidatePreview();
}
$('tool-search').oninput=renderTools;$('has-header').onchange=invalidatePreview;
$('preview-action').onclick=()=>safely(async()=>{
 if(officeMode&&!snapshot)throw Error('先在 Excel 选中数据，然后点击「读取 Excel 选区」。');
 const action=ACTIONS.find(a=>a.id===actionId),options={formulas,hasHeader:$('has-header').checked};for(const key of action.needs||[])options[key]=key==='digits'?Number($(`option-${key}`).value):$(`option-${key}`).value;
 const result=preview(actionId,data,options);pending={result,values:clone(data),formulas:clone(formulas),actionId,snapshot};
 let html=`<div class="notice success">${action.kind==='transform'?`将修改 ${result.patches.length} 个单元格，跳过 ${result.summary.skippedFormulas} 个公式单元格。`:action.kind==='inspect'?`发现 ${result.findings.length} 个检查提示。只检查，不修改。`:`将生成 ${result.output?.length||0} 行 × ${result.output?.[0]?.length||0} 列的新表。只复制值，原始区域保留。`}</div>`;
 if(result.patches.length)html+='<div class="diff-list">'+result.patches.slice(0,60).map(p=>`<div class="diff-row"><span>${columnName(p.c)}${p.r+1}</span><span class="before" title="${escape(displayValue(p.before))}">${escape(displayValue(p.before))}</span><span>→</span><span class="after" title="${escape(displayValue(p.after))}">${escape(displayValue(p.after))}</span></div>`).join('')+'</div>';
 if(result.findings.length)html+='<div class="diff-list">'+result.findings.slice(0,60).map(f=>`<div>${columnName(f.c)}${f.r+1} · ${escape(f.reason)}</div>`).join('')+'</div>';
 if(result.output)html+='<div class="table-scroll compact">'+tableHTML(result.output,{limit:20})+'</div>';
 if(result.patches.length>60||result.findings.length>60)html+='<p class="helper">这里只展示前 60 项；数量包含完整区域。</p>';
 html+=result.warnings.map(w=>`<p class="helper">ⓘ ${escape(w)}</p>`).join('');
 if(result.patches.length>2000)html+='<div class="notice error">修改超过 2,000 个单元格，请缩小选区后分批操作。</div>';
 $('action-result').innerHTML=html;$('apply-action').hidden=action.kind==='inspect'||(!result.output?.length&&!result.patches.length)||result.patches.length>2000;
 $('apply-action').textContent=result.output?'确认生成新工作表 →':`确认修改 ${result.patches.length} 个单元格`;
});
$('apply-action').onclick=()=>safely(async()=>{
 if(!pending)return;const p=pending;if(JSON.stringify(p.values)!==JSON.stringify(data)||JSON.stringify(p.formulas)!==JSON.stringify(formulas))throw Error('数据已改变，请重新预览后再应用。');
 if(p.result.output){if(officeMode){const fresh=await readSelection();if(fresh.address!==p.snapshot.address||fresh.worksheetId!==p.snapshot.worksheetId||JSON.stringify(fresh.values)!==JSON.stringify(p.values)||JSON.stringify(fresh.formulas)!==JSON.stringify(p.formulas))throw Error('Excel 选区或内容已变化，请重新读取并预览，避免生成过期结果。');}await outputMatrix(p.result.output);return;}
 if(officeMode){undo=await applyPatches(p.snapshot,p.result.patches);snapshot=clone(p.snapshot);for(const patch of undo.patches){snapshot.values[patch.r][patch.c]=patch.after;snapshot.formulas[patch.r][patch.c]=patch.afterFormula;}data=snapshot.values;formulas=snapshot.formulas;}
 else{undo={patches:clone(p.result.patches),demo:true};for(const patch of p.result.patches)data[patch.r][patch.c]=patch.after;}
 $('undo').disabled=false;renderData();invalidatePreview();toast(`已修改 ${p.result.patches.length} 个单元格。需要时可撤销上次修改。`);
});
$('undo').onclick=()=>safely(async()=>{
 if(!undo)return;if(officeMode){await undoPatches(undo);try{snapshot=await readSelection();data=snapshot.values;formulas=snapshot.formulas;}catch{snapshot=null;data=[];formulas=[];}}
 else{for(const p of undo.patches)if(data[p.r]?.[p.c]!==p.after||String(formulas[p.r]?.[p.c]||'').startsWith('='))throw Error('这些格子后来又被修改了，无法安全撤销。请手动检查，避免覆盖新内容。');for(const p of undo.patches)data[p.r][p.c]=p.before;}
 undo=null;$('undo').disabled=true;invalidatePreview();renderData();toast('已撤销上一次单元格修改。');
});
$('refresh-data').onclick=()=>safely(async()=>{if(!connected)throw Error('请在 Excel 加载项内打开，才能读取选区。');snapshot=await readSelection();data=snapshot.values;formulas=snapshot.formulas;invalidatePreview();renderData();toast('已读取选区。勾选或取消「第一行是标题」，然后选择工具。');});

const formulaCards=[
 ['把一列数字加起来','=SUM(B2:B100)','统计 B2 到 B100 的合计。替换 B 和行号就能用。','文本格式的数字可能被忽略；先用「数字文本转数值」检查。'],
 ['计算平均值','=AVERAGE(B2:B100)','统计这一段数字的平均数。','空白不参与平均，但数字 0 会参与。'],
 ['统计符合条件的个数','=COUNTIF(C2:C100,"已完成")','数一数 C 列中有多少个「已完成」。','条件文字必须用英文双引号；可把条件换成别的内容。'],
 ['按条件求和','=SUMIF(A2:A100,"设计部",B2:B100)','只累加 A 列属于设计部的 B 列金额。','条件范围和求和范围要一一对应，长度相同。'],
 ['按编号找对应内容','=IFERROR(VLOOKUP(E2,A2:C100,3,FALSE),"未找到")','用 E2 的编号在 A 列找，返回第 3 列的内容。','编号要在查找区域的第一列；FALSE 表示精确匹配。'],
 ['根据条件显示文字','=IF(B2>=60,"通过","待改进")','B2 大于等于 60 时显示通过，否则显示待改进。','可以修改 60 和两段结果文字。'],
 ['把两格文字拼起来','=A2&" - "&B2','把 A2、分隔文字、B2 连接起来。','& 是连接符；中间的 " - " 可以改成空格或留空。'],
 ['计算两个日期相隔几天','=B2-A2','结束日期减去开始日期，得到相隔的天数。','A2、B2 必须是真正的日期；把结果单元格格式设为常规或数字。'],
];
$('formula-list').innerHTML=formulaCards.map((f,i)=>`<article class="card formula-card"><p class="eyebrow">FORMULA ${String(i+1).padStart(2,'0')}</p><h2>${escape(f[0])}</h2><p>${escape(f[2])}</p><pre>${escape(f[1])}</pre><p>小提醒：${escape(f[3])} 若 Excel 使用分号分隔参数，请把公式参数之间的逗号改成分号。</p><button class="secondary" data-copy-formula="${i}">复制公式</button></article>`).join('');
document.querySelectorAll('[data-copy-formula]').forEach(el=>el.onclick=()=>safely(async()=>{await copy(formulaCards[Number(el.dataset.copyFormula)][1]);toast('公式已复制。粘贴前记得把范围改成你的实际数据范围。');}));

$('tool-count').textContent=ACTIONS.length;$('refresh-data').hidden=!officeMode;$('reset-demo').hidden=officeMode;
if(officeMode){data=[];formulas=[];$('mode-badge').textContent='正在连接 Excel…';$('destination-title').textContent='生成到一张全新的工作表';$('destination-note').textContent='原始工作表保持不变；文字按文本写入。';$('selection-help').textContent='在 Excel 中选中要处理的区域，然后点击读取。这里只展示选区，不直接编辑。';
 connectOffice().then(()=>{connected=true;$('mode-badge').textContent='Excel 已连接';renderShape();toast('已连接 Excel。粘贴排表可直接生成新工作表。');}).catch(e=>{$('mode-badge').textContent='Excel 未连接';toast(e.message);});
}else resetDemo();
renderTools();renderAction();renderData();$('raw-data').value=example;renderShape();
