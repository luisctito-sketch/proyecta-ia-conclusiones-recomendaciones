
const $=id=>document.getElementById(id);
const clean=s=>(s??"").toString().replace(/[ \t]+/g," ").replace(/\s*\n\s*/g,"\n").trim();
const norm=s=>clean(s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();
const esc=s=>(s??"").toString().replace(/[&<>\"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const words=s=>clean(s).split(/\s+/).filter(Boolean);
const stop=new Set(`a al algo algunas algunos ante antes bajo bien cada como con contra cual cuando de del desde donde durante e el ella ellas ellos en entre era es esa ese eso esta estaba este esto fue ha hacia hasta la las le lo los mas mediante muy ni no o para pero por porque que se sin sobre son su sus tambien tan tanto un una uno unos unas y ya`.split(" "));
const domainStop=new Set(`objetivo objetivos especifico especificos general proyecto sistema prototipo desarrollar disenar realizar implementar efectuar elaborar analizar evaluar identificar determinar establecer crear mejorar`.split(" "));
function tokens(s){return [...new Set(norm(s).replace(/[^a-z0-9áéíóúñü%]+/gi," ").split(/\s+/).filter(x=>x.length>3&&!stop.has(x)&&!domainStop.has(x)))];}
function similarity(a,b){const A=tokens(a),B=new Set(tokens(b));if(!A.length)return 0;return A.filter(x=>B.has(x)).length/A.length;}
function splitObjectives(t){return clean(t).split(/\n+/).map(x=>x.replace(/^\s*(?:[-•*]|\d+[.)-]?|[a-z][.)])\s*/i,"").trim()).filter(x=>x.length>8);}
function splitParagraphs(t){let raw=(t??"").replace(/\r/g,"").trim();if(!raw)return[];let p=raw.split(/\n\s*\n+/).map(clean).filter(Boolean);if(p.length===1){const lines=raw.split(/\n+/).map(clean).filter(Boolean);if(lines.length>1&&lines.every(x=>x.length>35))p=lines;}return p;}
function firstPerson(t){return t.match(/\b(yo|nosotros|nosotras|nuestro|nuestra|nuestros|nuestras|hicimos|realizamos|desarrollamos|implementamos|diseñamos|disenamos|probamos|evaluamos|analizamos|obtuvimos|concluimos|recomendamos|utilizamos|empleamos|logramos)\b/gi)||[];}
function mechanical(t){return t.match(/objetivo\s+espec[ií]fico\s+(?:n[.º°o]*\s*)?\d+/gi)||[];}
function evidenceScore(t){let s=0;if(/\d/.test(t))s+=2;if(/%|°c|\bms\b|\bmbps\b|\bgbps\b|\bwh\b|\bkwh\b|\bdb\b|\bdbm\b|\bpa\b|\bv\b|\ba\b|\bhz\b|\bppm\b|\bbs\b/i.test(t))s+=1;if(/prueb|medici|compar|error|desviaci|precision|precisión|exactitud|confiab|disponib|eficien|latencia|tiempo|resultado|indicador|valid|calibr|alcanz|obtuvo|registr|mantuvo|redujo|increment/i.test(t))s+=1;return s;}
function genericClaims(t){const arr=[];const pats=[/funcion[oó]\s+correctamente/gi,/cumpli[oó]\s+(?:con\s+)?(?:todos\s+)?los\s+objetivos/gi,/datos\s+precisos\s+y\s+confiables/gi,/demuestra\s+ser\s+efectiv[oa]/gi,/resultados\s+esperados/gi,/fue\s+exitos[oa]/gi];pats.forEach(p=>{const m=t.match(p);if(m)arr.push(...m)});return arr;}
function recDimensions(t){const n=norm(t),out=[];const m={"Uso/operación":["operar","operacion","usuario","capacita","setpoint","interfaz"],"Mantenimiento":["mantenimiento","calibr","limpieza","inspeccion","reemplazo"],"Adaptación real":["pcb","gabinete","industrial","instalacion","certific","proteccion","puesta en marcha"],"Escalabilidad":["escalab","ampliar","cobertura","mas sensores","usuarios","capacidad"],"Manual/documentación":["manual","documentacion","protocolo","guia de usuario"],"Tecnología futura":["inteligencia artificial","aprendizaje automatico","pid","difusa","nueva plataforma","tecnologia emergente"],"Seguridad":["seguridad","ciberseg","aislamiento","fusible","sobretension","estado seguro","redundan"],"Pruebas futuras":["pruebas prolongadas","validacion","ensayo","instrumento patron","operacion prolongada"]};for(const [k,keys] of Object.entries(m))if(keys.some(x=>n.includes(x)))out.push(k);return out;}
function hasWhyBenefit(t){const n=norm(t);return /debido a|porque|a partir de|ante la|considerando|dado que/.test(n)&&/con el fin|de manera que|permitir|para mejorar|para reducir|para incrementar|para garantizar|contribuir/.test(n);}
function badge(score){return score>=.42?'<span class="score-pill score-good">COBERTURA ALTA</span>':score>=.22?'<span class="score-pill score-warn">COBERTURA PARCIAL</span>':'<span class="score-pill score-bad">COBERTURA BAJA</span>'}
function show(el,cls,html){el.className=`output-box ${cls}`;el.innerHTML=html;el.classList.remove('hidden')}

if(window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
async function extractFile(file){
  const name=file.name.toLowerCase();
  if(name.endsWith('.txt')) return await file.text();
  const ab=await file.arrayBuffer();
  if(name.endsWith('.docx')){
    if(!window.mammoth) throw new Error('No se pudo cargar el lector DOCX.');
    const r=await mammoth.extractRawText({arrayBuffer:ab}); return r.value;
  }
  if(name.endsWith('.pdf')){
    if(!window.pdfjsLib) throw new Error('No se pudo cargar el lector PDF.');
    const pdf=await pdfjsLib.getDocument({data:ab}).promise;let pages=[];
    for(let i=1;i<=pdf.numPages;i++){const p=await pdf.getPage(i),c=await p.getTextContent();pages.push(c.items.map(x=>x.str).join(' '));}
    return pages.join('\n\n');
  }
  throw new Error('Formato no admitido. Usa DOCX, PDF o TXT.');
}
document.querySelectorAll('input[type=file][data-target]').forEach(inp=>inp.addEventListener('change',async()=>{
  const file=inp.files?.[0],target=inp.dataset.target,status=document.querySelector(`[data-status="${target}"]`);if(!file)return;
  status.className='file-status';status.textContent='Leyendo archivo...';
  try{const text=await extractFile(file);$(target).value=text.trim();status.className='file-status ok';status.textContent=`Cargado: ${file.name} · ${words(text).length} palabras.`;}
  catch(e){status.className='file-status bad';status.textContent=e.message;}
}));

function reviewConclusions(){
 const t=clean($('conclusionsText').value),paras=splitParagraphs(t),objs=splitObjectives($('specificObjectives').value);if(!t){show($('conclusionsOutput'),'bad','Pega o carga tus conclusiones.');return;}
 let issues=[],strength=[];const fp=firstPerson(t),mech=mechanical(t),generic=genericClaims(t),ev=paras.filter(p=>evidenceScore(p)>=2).length,long=paras.filter(p=>words(p).length>120).length;
 if(fp.length)issues.push(`Se detectaron formas evidentes de primera persona: ${[...new Set(fp)].slice(0,8).join(', ')}.`);else strength.push('Predomina la redacción impersonal.');
 if(mech.length)issues.push('Se menciona mecánicamente el número de Objetivo Específico. La relación debe ser implícita.');else strength.push('No se detecta numeración mecánica de objetivos dentro de las conclusiones.');
 if(generic.length)issues.push(`Hay afirmaciones genéricas que necesitan evidencia: ${[...new Set(generic)].slice(0,6).join('; ')}.`);
 if(ev<Math.max(1,Math.ceil(paras.length/3)))issues.push('Pocos párrafos contienen señales claras de evidencia técnica (mediciones, pruebas, porcentajes, errores, tiempos o indicadores).');else strength.push('Varios párrafos incorporan evidencia o resultados técnicos.');
 if(long)issues.push(`${long} párrafo(s) superan aproximadamente 120 palabras. Comprueba en Word que ninguno exceda 10 líneas.`);
 if(objs.length&&!objs.every(o=>Math.max(0,...paras.map(p=>similarity(o,p)))>=.18))issues.push('Al menos un Objetivo Específico presenta cobertura textual débil en las conclusiones. Construye la matriz de cierre para identificar cuál.');
 if(!/[0-9]/.test(t))issues.push('No se detectan datos numéricos. Si el proyecto generó mediciones o indicadores relevantes, incorpora los valores principales ya demostrados en capítulos anteriores.');
 if(!/limit|dificult|sin embargo|no obstante|restric|pendiente|mejora/i.test(t))issues.push('No se reconocen con claridad dificultades, limitaciones o aspectos pendientes de optimización.');
 const cls=issues.length>=5?'bad':issues.length?'warn':'ok';
 show($('conclusionsOutput'),cls,`<h4>${issues.length>=5?'🟠 REQUIERE REFORMULACIÓN':issues.length?'🟡 CUMPLE PARCIALMENTE':'🟢 CUMPLE'}</h4>
 <div class="review-kpis"><div class="review-kpi"><span class="k">Párrafos</span><span class="v">${paras.length}</span><span class="d">Conclusiones detectadas</span></div><div class="review-kpi"><span class="k">Con evidencia</span><span class="v">${ev}</span><span class="d">Señales técnicas</span></div><div class="review-kpi"><span class="k">1.ª persona</span><span class="v">${fp.length}</span><span class="d">Formas detectadas</span></div><div class="review-kpi"><span class="k">Extensos</span><span class="v">${long}</span><span class="d">Estimación &gt;120 palabras</span></div></div>
 ${strength.length?`<h4>Fortalezas</h4><ul>${strength.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>`:''}
 ${issues.length?`<h4>Qué falta y cómo mejorarlo</h4><ul>${issues.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>`:''}
 <div class="rule-c6"><b>Control:</b> no agregues cifras nuevas aquí. Usa solamente resultados ya obtenidos y explicados en los capítulos previos.</div>`);
}
$('reviewConclusionsBtn').addEventListener('click',reviewConclusions);$('clearConclusionsBtn').addEventListener('click',()=>{$('conclusionsText').value='';$('conclusionsOutput').classList.add('hidden')});

function reviewRecommendations(){
 const t=clean($('recommendationsText').value),paras=splitParagraphs(t),concs=splitParagraphs($('conclusionsText').value);if(!t){show($('recommendationsOutput'),'bad','Pega o carga tus recomendaciones.');return;}
 const fp=firstPerson(t),long=paras.filter(p=>words(p).length>120).length;let orphan=0,incomplete=0,dim=new Set();
 const cards=paras.map((p,i)=>{const ds=recDimensions(p);ds.forEach(x=>dim.add(x));const link=concs.length?Math.max(...concs.map(c=>similarity(p,c))):0;if(concs.length&&link<.08)orphan++;const complete=hasWhyBenefit(p);if(!complete)incomplete++;return `<div class="paragraph-card"><h5>Recomendación ${i+1}</h5><div class="dimension-tags">${(ds.length?ds:['Sin categoría clara']).map(x=>`<span class="dimension-tag">${esc(x)}</span>`).join('')}</div><p>${esc(p.slice(0,360))}${p.length>360?'…':''}</p><p class="small"><b>Vínculo textual con conclusiones:</b> ${(link*100).toFixed(0)} % · <b>Estructura qué/por qué/beneficio:</b> ${complete?'probable':'incompleta'}</p></div>`}).join('');
 let issues=[];if(fp.length)issues.push('Se detecta primera persona; cambia a redacción impersonal.');if(long)issues.push(`${long} párrafo(s) parecen extensos; comprueba el límite de 10 líneas en Word.`);if(orphan)issues.push(`${orphan} recomendación(es) no muestran relación textual clara con resultados o limitaciones de las conclusiones. Justifica su origen.`);if(incomplete)issues.push(`${incomplete} recomendación(es) no expresan claramente qué mejorar, por qué y qué beneficio se espera.`);if(dim.size<3)issues.push('La cobertura de mejora futura es reducida. Revisa si corresponden uso, mantenimiento, adaptación real, escalabilidad, manual, seguridad, pruebas o tecnología futura.');
 show($('recommendationsOutput'),issues.length>=4?'bad':issues.length?'warn':'ok',`<h4>${issues.length>=4?'🟠 REQUIERE REFORMULACIÓN':issues.length?'🟡 CUMPLE PARCIALMENTE':'🟢 CUMPLE'}</h4><p><b>Dimensiones detectadas:</b> ${dim.size?[...dim].join(' · '):'ninguna clara'}.</p>${issues.length?`<h4>Qué falta y cómo mejorarlo</h4><ul>${issues.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>`:''}${cards}`);
}
$('reviewRecommendationsBtn').addEventListener('click',reviewRecommendations);$('clearRecommendationsBtn').addEventListener('click',()=>{$('recommendationsText').value='';$('recommendationsOutput').classList.add('hidden')});

let lastMatrix=[];
function buildMatrix(){
 const objs=splitObjectives($('specificObjectives').value),paras=splitParagraphs($('conclusionsText').value);if(!objs.length||!paras.length){$('matrixOutput').innerHTML='<div class="output-box bad">Se necesitan Objetivos Específicos y Conclusiones.</div>';return;}
 lastMatrix=objs.map((o,i)=>{let best={score:-1,p:'',idx:-1};paras.forEach((p,j)=>{const s=similarity(o,p);if(s>best.score)best={score:s,p,idx:j}});const ev=evidenceScore(best.p);let status=best.score>=.42&&ev>=2?'ALTA':best.score>=.22||ev>=2?'PARCIAL':'BAJA';return {i:i+1,o,p:best.p,pidx:best.idx+1,score:best.score,ev,status}});
 $('matrixOutput').innerHTML=`<table class="closure-table"><thead><tr><th>Objetivo Específico</th><th>Conclusión con mayor relación textual</th><th>Cobertura</th><th>Evidencia</th><th>Orientación</th></tr></thead><tbody>${lastMatrix.map(r=>`<tr><td><b>OE ${r.i}</b><br>${esc(r.o)}</td><td><b>Párrafo ${r.pidx}</b><br>${esc(r.p.slice(0,380))}${r.p.length>380?'…':''}</td><td>${badge(r.score)}<br>${(r.score*100).toFixed(0)} %</td><td>${r.ev>=2?'<span class="score-pill score-good">HAY SEÑALES</span>':'<span class="score-pill score-warn">DÉBIL</span>'}</td><td>${r.status==='ALTA'?'Verifica que las magnitudes citadas coincidan con resultados previos.':r.status==='PARCIAL'?'Haz más explícito el resultado alcanzado y su evidencia, sin nombrar el número del objetivo.':'No se identifica una conclusión que cierre claramente esta acción. Añade o reformula un párrafo sustentado en resultados reales.'}</td></tr>`).join('')}</tbody></table>`;
}
$('buildMatrixBtn').addEventListener('click',buildMatrix);
$('downloadMatrixBtn').addEventListener('click',()=>{if(!lastMatrix.length)buildMatrix();if(!lastMatrix.length)return;const rows=[['Objetivo','Conclusion relacionada','Cobertura %','Evidencia','Estado'],...lastMatrix.map(r=>[r.o,r.p,(r.score*100).toFixed(0),r.ev,r.status])];const csv=rows.map(r=>r.map(x=>'"'+String(x).replace(/"/g,'""')+'"').join(';')).join('\n');const blob=new Blob([csv],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='Matriz_Objetivos_Conclusiones_PROYECTA_IA.csv';a.click();URL.revokeObjectURL(a.href)});

function integratedReview(){
 const og=clean($('generalObjective').value),oe=splitObjectives($('specificObjectives').value),ps=clean($('problemStatement').value),pf=clean($('problemFormulation').value),c=clean($('conclusionsText').value),r=clean($('recommendationsText').value);
 let missing=[];if(!og)missing.push('Objetivo General');if(!oe.length)missing.push('Objetivos Específicos');if(!ps)missing.push('Planteamiento del Problema');if(!pf)missing.push('Formulación del Problema');if(!c)missing.push('Conclusiones');if(!r)missing.push('Recomendaciones');
 if(missing.length){show($('integralOutput'),'bad',`<h4>🔵 FALTA INFORMACIÓN</h4><p>Completa: ${missing.join(', ')}.</p>`);return;}
 const cp=splitParagraphs(c),rp=splitParagraphs(r),oeScores=oe.map(o=>Math.max(...cp.map(p=>similarity(o,p)))),covered=oeScores.filter(x=>x>=.22).length,ogScore=Math.max(similarity(og,c),similarity(pf,c),similarity(ps,c)),ev=cp.filter(p=>evidenceScore(p)>=2).length,fp=firstPerson(c+' '+r).length,generic=genericClaims(c).length,recComplete=rp.filter(hasWhyBenefit).length,recLinked=rp.filter(p=>Math.max(...cp.map(x=>similarity(p,x)))>=.08).length;
 let issues=[];if(covered<oe.length)issues.push(`${oe.length-covered} Objetivo(s) Específico(s) no presentan cierre textual suficiente.`);if(ogScore<.18)issues.push('El cierre global mantiene poca relación textual con el Objetivo General y/o con el problema.');if(ev<Math.ceil(cp.length/3))issues.push('La evidencia técnica dentro de las conclusiones es escasa.');if(generic)issues.push('Persisten afirmaciones genéricas que deberían acompañarse de evidencia.');if(fp)issues.push('Se detectan formas de primera persona.');if(recComplete<rp.length)issues.push('Varias recomendaciones no contienen simultáneamente acción, razón y beneficio esperado.');if(recLinked<Math.ceil(rp.length*.6))issues.push('Varias recomendaciones parecen desvinculadas de los resultados o limitaciones declaradas.');
 const status=issues.length>=5?['bad','🔴 NO CUMPLE']:issues.length>=2?['warn','🟠 REQUIERE REFORMULACIÓN']:issues.length?['warn','🟡 CUMPLE PARCIALMENTE']:['ok','🟢 CUMPLE'];
 show($('integralOutput'),status[0],`<h4>${status[1]} — CAPÍTULO VI</h4><div class="review-kpis"><div class="review-kpi"><span class="k">Objetivos cerrados</span><span class="v">${covered}/${oe.length}</span><span class="d">Cobertura textual</span></div><div class="review-kpi"><span class="k">Conclusiones</span><span class="v">${cp.length}</span><span class="d">Párrafos detectados</span></div><div class="review-kpi"><span class="k">Con evidencia</span><span class="v">${ev}</span><span class="d">Señales técnicas</span></div><div class="review-kpi"><span class="k">Recomendaciones completas</span><span class="v">${recComplete}/${rp.length}</span><span class="d">Qué + por qué + beneficio</span></div></div>
 <div class="report-section"><h4>A. Correspondencia con el problema y Objetivo General</h4><p>${ogScore>=.30?'La relación textual es clara.':'La relación necesita hacerse más explícita mediante una conclusión integradora basada en resultados.'}</p></div>
 <div class="report-section"><h4>B. Cumplimiento de Objetivos Específicos</h4><p>${covered} de ${oe.length} presentan una conclusión con cobertura textual al menos parcial. Utiliza la matriz para revisar los restantes.</p></div>
 <div class="report-section"><h4>C. Evidencia y grado de cumplimiento</h4><p>${ev} de ${cp.length} párrafos contienen señales claras de evidencia. El grado de cumplimiento debe justificarse con resultados previamente demostrados.</p></div>
 <div class="report-section"><h4>D. Problemas y dificultades</h4><p>${/limit|dificult|sin embargo|no obstante|restric|ruido|retard|fall|conect|calibr/i.test(c)?'Se detecta reconocimiento de dificultades o limitaciones.':'No se identifican dificultades técnicas con suficiente claridad; incorpora las que realmente se presentaron y cómo fueron tratadas.'}</p></div>
 <div class="report-section"><h4>E. Recomendaciones de continuidad</h4><p>${recLinked} de ${rp.length} muestran relación textual probable con las conclusiones. Prioriza acciones derivadas de limitaciones reales y de la transición prototipo → sistema definitivo.</p></div>
 ${issues.length?`<div class="report-section"><h4>F. Información que falta y cómo mejorarla</h4><ul>${issues.map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div>`:''}
 <div class="rule-c6"><b>Regla pedagógica:</b> PROYECTA-IA señala las brechas, pero no sustituye al estudiante redactando sus conclusiones o recomendaciones finales.</div>`);
}
$('integralReviewBtn').addEventListener('click',integratedReview);

function deepPrompt(){return `ACTÚA COMO REVISOR METODOLÓGICO Y ACADÉMICO DE PROYECTOS DE GRADO DE INGENIERÍA ELECTRÓNICA.

REVISA EXCLUSIVAMENTE:
CAPÍTULO VI. CONCLUSIONES Y RECOMENDACIONES
6.1. CONCLUSIONES
6.2. RECOMENDACIONES

TÍTULO DEL PROYECTO:
${$('projectTitle').value||'[NO INGRESADO]'}

OBJETIVO GENERAL:
${$('generalObjective').value||'[NO INGRESADO]'}

OBJETIVOS ESPECÍFICOS:
${$('specificObjectives').value||'[NO INGRESADOS]'}

PLANTEAMIENTO DEL PROBLEMA:
${$('problemStatement').value||'[NO INGRESADO]'}

FORMULACIÓN DEL PROBLEMA:
${$('problemFormulation').value||'[NO INGRESADA]'}

CONCLUSIONES DEL ESTUDIANTE:
${$('conclusionsText').value||'[NO INGRESADAS]'}

RECOMENDACIONES DEL ESTUDIANTE:
${$('recommendationsText').value||'[NO INGRESADAS]'}

REGLAS PARA LAS CONCLUSIONES:
- Son síntesis razonadas de los resultados; no repiten capítulos ni marco teórico.
- Deben responder a los Objetivos Específicos de forma implícita y natural.
- No escribir “objetivo específico N.º 1”, “en cumplimiento del objetivo 2”, etc.
- Deben indicar qué se logró, grado de cumplimiento, evidencia, principales resultados y dificultades/limitaciones.
- Las afirmaciones “funcionó correctamente”, “fue preciso/confiable/eficiente” requieren evidencia de pruebas, mediciones, comparaciones, cálculos o indicadores.
- No inventar resultados ni introducir datos nuevos que no hayan sido demostrados previamente.
- Conviene cerrar con una conclusión integradora respecto del Objetivo General y el problema formulado.

REGLAS PARA LAS RECOMENDACIONES:
- Deben derivarse de resultados, limitaciones o problemas realmente observados.
- Deben indicar qué se propone mejorar, por qué y qué beneficio se espera.
- Pueden abarcar uso, mantenimiento, adaptación a instalación real, escalabilidad, manual de usuario, documentación, seguridad, pruebas futuras, actualización tecnológica e investigación futura.
- Deben distinguir prototipo/sistema experimental de un producto o sistema final listo para operación.
- No recomendar tecnología nueva solamente porque es novedosa; justificar la necesidad.

REDACCIÓN:
- Científica, impersonal y en tercera persona.
- Evitar verbos en primera persona como hicimos, desarrollamos, implementamos, probamos, nuestro sistema.
- Párrafos, no viñetas, salvo necesidad estricta.
- Cada párrafo debe comprobarse en Word para no exceder aproximadamente 10 líneas en hoja tamaño carta, Times New Roman 12.
- Usar conectores gramaticales y evitar repetición excesiva.
- No redactes ni reescribas el Capítulo VI por el estudiante. Revisa, diagnostica y orienta.

RESPONDE CON:
1. DICTAMEN DE REVISIÓN — CAPÍTULO VI: 🟢 CUMPLE / 🟡 CUMPLE PARCIALMENTE / 🟠 REQUIERE REFORMULACIÓN / 🔴 NO CUMPLE / 🔵 FALTA INFORMACIÓN.
2. A. FORTALEZAS.
3. B. INFORMACIÓN QUE FALTA Y CÓMO MEJORARLA.
4. C. COHERENCIA CON EL PLANTEAMIENTO Y FORMULACIÓN DEL PROBLEMA.
5. D. COHERENCIA CON EL OBJETIVO GENERAL.
6. E. MATRIZ DE CORRESPONDENCIA: cada Objetivo Específico → conclusión que lo cierra → evidencia citada por el estudiante → grado de cobertura.
7. F. CALIDAD TÉCNICA DE LAS CONCLUSIONES.
8. G. PROBLEMAS, DIFICULTADES Y LIMITACIONES RECONOCIDAS.
9. H. RECOMENDACIONES: vínculo con resultados/limitaciones.
10. I. TRANSICIÓN DEL PROTOTIPO A SISTEMA FINAL: uso, mantenimiento, PCB/gabinete/protecciones, manuales, escalabilidad, seguridad y pruebas, solo cuando correspondan al proyecto.
11. J. REDACCIÓN CIENTÍFICA, TERCERA PERSONA, CONECTORES Y REPETICIONES.
12. K. PÁRRAFOS POSIBLEMENTE EXTENSOS PARA COMPROBAR EN WORD.
13. L. RECOMENDACIONES CONCRETAS AL ESTUDIANTE, SIN ENTREGAR PÁRRAFOS SUSTITUTOS.

REGLA PEDAGÓGICA FINAL:
NO redactes las conclusiones ni recomendaciones completas por el estudiante. Para cada falla indica qué falta, por qué importa y cómo debe corregirla por sí mismo.`;}
$('copyPromptBtn').addEventListener('click',async()=>{await navigator.clipboard.writeText(deepPrompt());$('promptStatus').textContent='Prompt copiado.';setTimeout(()=>$('promptStatus').textContent='',1600)});
