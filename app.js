let inventorCount = 0;
let figures = []; // {dataUrl, caption}
let priorArtKeywords = [];
let priorArtSelected = null; // Set, lazily initialized to all keywords on first extraction
let synonymSelected = {}; // { [term]: Set of selected synonym strings }
let queryMode = 'advanced'; // 'basic' | 'advanced' | 'patent'

// ---- Lazy-loaded export libraries ----
// jsPDF, docx.js and JSZip together are roughly 600KB of JS. They used to be
// loaded unconditionally in index.html's <head>, which meant every single
// page visit paid that cost even if the person never clicked an export
// button. Now each is fetched once, the first time it's actually needed,
// and cached here so a second export click in the same session doesn't
// re-fetch anything.
const _loadedScripts = {};
function loadScriptOnce(url){
  if(_loadedScripts[url]) return _loadedScripts[url];
  _loadedScripts[url] = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = url;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('Failed to load ' + url + ' — check your internet connection and try again.'));
    document.head.appendChild(s);
  });
  return _loadedScripts[url];
}
async function ensureJsPDFLoaded(){
  if(!window.jspdf){
    await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  }
}
async function ensureWordExportLibsLoaded(){
  if(!window.docx){
    await loadScriptOnce('https://unpkg.com/docx@8.5.0/build/index.umd.js');
  }
  if(!window.JSZip){
    await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
  }
}

// ---- Local autosave (survives refresh / accidental tab close) ----
// Everything here stays in localStorage on this device only — it is never
// sent anywhere, same as the rest of the tool. "Clear all fields" below
// wipes this saved copy too, not just the on-screen fields.
const AUTOSAVE_KEY = 'idf_builder_draft_v1';
const SIMPLE_FIELD_IDS = ['discNo','title','briefSummary','description','discDetail',
  'priorArt','problemSolved','novelty','advantages','otherDesc','priorLit','addlNotes'];
let autosaveTimer = null;

function collectDraftState(includeFigures){
  const fields = {};
  SIMPLE_FIELD_IDS.forEach(id => {
    const el = document.getElementById(id);
    if(el) fields[id] = el.value;
  });
  const discYNEl = document.querySelector('input[name=discYN]:checked');
  const state = {
    v: 1,
    fields,
    discYN: discYNEl ? discYNEl.value : 'No',
    inventors: getInventorRecords(),
  };
  if(includeFigures) state.figures = figures;
  return state;
}

function saveDraftNow(){
  try{
    localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(collectDraftState(true)));
  }catch(e){
    // Most likely a quota error caused by embedded figure images. Retry
    // without figures so the text fields — the part that's actually painful
    // to retype — still get saved.
    try{
      localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(collectDraftState(false)));
      console.warn('Autosave: figures omitted (browser storage limit reached).', e);
    }catch(e2){
      console.warn('Autosave failed:', e2);
    }
  }
}

function saveDraftDebounced(){
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(saveDraftNow, 500);
}

function clearSavedDraft(){
  try{ localStorage.removeItem(AUTOSAVE_KEY); }catch(e){ /* ignore */ }
}

function restoreDraft(){
  let raw;
  try{ raw = localStorage.getItem(AUTOSAVE_KEY); }catch(e){ return; }
  if(!raw) return;
  let saved;
  try{ saved = JSON.parse(raw); }catch(e){ return; }
  if(!saved || typeof saved !== 'object') return;

  if(saved.fields){
    SIMPLE_FIELD_IDS.forEach(id => {
      const el = document.getElementById(id);
      if(el && typeof saved.fields[id] === 'string') el.value = saved.fields[id];
    });
  }
  if(saved.discYN === 'Yes'){
    const el = document.querySelector('input[name=discYN][value="Yes"]');
    if(el) el.checked = true;
  }
  if(Array.isArray(saved.inventors) && saved.inventors.length){
    document.getElementById('inventorList').innerHTML = '';
    inventorCount = 0;
    saved.inventors.forEach(rec => {
      addInventor();
      const row = document.getElementById('inventorList').lastElementChild;
      if(row){
        const nameEl = row.querySelector('.inv-name');
        const addrEl = row.querySelector('.inv-address');
        const emailEl = row.querySelector('.inv-email');
        const phoneEl = row.querySelector('.inv-phone');
        if(nameEl) nameEl.value = rec.name || '';
        if(addrEl) addrEl.value = rec.address || '';
        if(emailEl) emailEl.value = rec.email || '';
        if(phoneEl) phoneEl.value = rec.phone || '';
      }
    });
  }
  if(Array.isArray(saved.figures) && saved.figures.length){
    figures = saved.figures.map(f => ({dataUrl: f.dataUrl, caption: f.caption || '', source: f.source || 'image'}));
    renderFigList();
  }
  const titleEl = document.getElementById('title');
  if(titleEl) enforceTitleWordLimit(titleEl);
}

const STOPWORDS = new Set([
  "a","an","the","of","to","in","on","for","and","or","with","by","at","from","as","is","are","was","were",
  "be","been","being","this","that","these","those","it","its","which","who","whom","whose","what","when",
  "where","why","how","not","no","nor","but","if","then","than","so","such","can","will","may","might","must",
  "shall","should","would","could","do","does","did","has","have","had","into","onto","over","under","between",
  "within","without","through","during","before","after","above","below","up","down","out","off","again",
  "further","once","here","there","all","each","few","more","most","other","some","any","both","own","same",
  "very","also","thereof","thereby","wherein","whereby","said","one","two","three","first","second","third",
  // patent-boilerplate terms that would otherwise dominate every extraction regardless of subject matter
  "invention","present","disclosure","embodiment","embodiments","comprising","comprise","comprises","including",
  "include","includes","object","objects","aspect","aspects","provide","provides","providing","configured",
  "system","method","apparatus","device","relates","relating","related","art","prior","field","background",
  "summary","claim","claims","figure","figures","description","detailed","accordance","plurality","means",
]);

// ---- Domain-specific dictionary: typed classification + synonym expansion ----
// Static, hand-curated, offline lookup table — deliberately NOT AI-generated at
// runtime. Coverage is intentionally scoped to the domains most SIU/SIDTM
// disclosures fall into (AI, Telecom, Blockchain, IoT, Healthcare, FinTech)
// plus a handful of common cross-domain terms (sustainability/agriculture)
// used in worked examples. A keyword that doesn't match any entry here is
// left unclassified rather than guessed — an honest "not covered" beats a
// fabricated category or invented synonym list. Extend this object directly
// to add domains/terms; no other code needs to change.
const DOMAIN_DICTIONARY = {
  // ---- AI ----
  "artificial intelligence": { category: "Technology", domain: "AI", synonyms: ["machine learning","cognitive computing","predictive analytics"] },
  "machine learning": { category: "Technology", domain: "AI", synonyms: ["statistical learning","predictive modeling","ml"] },
  "neural network": { category: "Component", domain: "AI", synonyms: ["artificial neural network","deep neural network","neural net"] },
  "deep learning": { category: "Technology", domain: "AI", synonyms: ["deep neural network learning","representation learning"] },
  "reinforcement learning": { category: "Technology", domain: "AI", synonyms: ["reward-based learning","policy optimization"] },
  "large language model": { category: "Technology", domain: "AI", synonyms: ["llm","transformer model","generative language model"] },
  "natural language processing": { category: "Technology", domain: "AI", synonyms: ["nlp","text analytics","language understanding"] },
  "computer vision": { category: "Technology", domain: "AI", synonyms: ["image recognition","visual perception system"] },
  "predictive model": { category: "Component", domain: "AI", synonyms: ["forecasting model","predictive algorithm"] },

  // ---- Telecom ----
  "5g": { category: "Technology", domain: "Telecom", synonyms: ["fifth generation network","5g nr","new radio"] },
  "6g": { category: "Technology", domain: "Telecom", synonyms: ["sixth generation network"] },
  "massive mimo": { category: "Technology", domain: "Telecom", synonyms: ["multiple input multiple output","large-scale antenna system"] },
  "beamforming": { category: "Technology", domain: "Telecom", synonyms: ["beam steering","directional signal transmission"] },
  "network slicing": { category: "Technology", domain: "Telecom", synonyms: ["virtual network partitioning","logical network segmentation"] },
  "base station": { category: "Component", domain: "Telecom", synonyms: ["cell site","radio access node"] },
  "spectrum allocation": { category: "Application", domain: "Telecom", synonyms: ["frequency allocation","spectrum management"] },

  // ---- Blockchain ----
  "blockchain": { category: "Technology", domain: "Blockchain", synonyms: ["distributed ledger","dlt","immutable ledger","decentralized ledger","shared ledger"] },
  "distributed ledger": { category: "Technology", domain: "Blockchain", synonyms: ["blockchain","dlt","decentralized ledger"] },
  "smart contract": { category: "Component", domain: "Blockchain", synonyms: ["self-executing contract","chaincode","programmable contract"] },
  "consensus mechanism": { category: "Technology", domain: "Blockchain", synonyms: ["consensus protocol","proof of work","proof of stake"] },
  "tokenization": { category: "Application", domain: "Blockchain", synonyms: ["asset tokenization","digital token issuance"] },
  "cryptographic hash": { category: "Component", domain: "Blockchain", synonyms: ["hash function","hashing algorithm"] },

  // ---- IoT ----
  "iot": { category: "Technology", domain: "IoT", synonyms: ["internet of things","connected device network","machine to machine"] },
  "sensor": { category: "Component", domain: "IoT", synonyms: ["wireless sensor","smart sensor","remote monitoring device","embedded sensor","environmental sensor"] },
  "soil sensor": { category: "Component", domain: "Agriculture", synonyms: ["soil moisture sensor","soil probe","in-ground sensor","agricultural sensor"] },
  "sensor network": { category: "Technology", domain: "IoT", synonyms: ["wireless sensor network","distributed sensor array"] },
  "edge computing": { category: "Technology", domain: "IoT", synonyms: ["edge processing","fog computing"] },
  "remote monitoring": { category: "Application", domain: "IoT", synonyms: ["telemetry","condition monitoring","remote sensing"] },

  // ---- Healthcare ----
  "wearable device": { category: "Component", domain: "Healthcare", synonyms: ["wearable sensor","body-worn device"] },
  "electronic health record": { category: "Component", domain: "Healthcare", synonyms: ["ehr","electronic medical record","emr"] },
  "telemedicine": { category: "Application", domain: "Healthcare", synonyms: ["telehealth","remote patient monitoring","virtual consultation"] },
  "diagnostic model": { category: "Component", domain: "Healthcare", synonyms: ["diagnostic algorithm","clinical decision support"] },
  "patient monitoring": { category: "Application", domain: "Healthcare", synonyms: ["vital sign monitoring","physiological monitoring"] },

  // ---- FinTech ----
  "digital payment": { category: "Application", domain: "FinTech", synonyms: ["electronic payment","mobile payment","cashless transaction"] },
  "fraud detection": { category: "Application", domain: "FinTech", synonyms: ["anomaly detection","transaction risk scoring"] },
  "credit scoring": { category: "Application", domain: "FinTech", synonyms: ["credit risk assessment","creditworthiness model"] },
  "digital wallet": { category: "Component", domain: "FinTech", synonyms: ["e-wallet","mobile wallet"] },
  "know your customer": { category: "Application", domain: "FinTech", synonyms: ["kyc verification","identity verification"] },

  // ---- Cross-domain (sustainability / agriculture — common in worked examples) ----
  "carbon credit": { category: "Application", domain: "Sustainability", synonyms: ["emission credit","carbon offset","carbon allowance"] },
  "carbon sequestration": { category: "Outcome", domain: "Sustainability", synonyms: ["carbon capture","carbon storage","carbon tracking"] },
  "precision agriculture": { category: "Application", domain: "Agriculture", synonyms: ["smart farming","precision farming"] },
};

// Exact match first; otherwise falls back to the longest dictionary key that
// is a substring of the term or vice versa (so a bigram like "language
// model" partially matches "large language model", and "sensor network"
// matches "sensor network" or falls back to "sensor"). No fuzzy/AI matching
// — a term with no textual overlap with any key is left unclassified.
function classifyKeyword(term){
  const t = term.toLowerCase();
  if(DOMAIN_DICTIONARY[t]) return Object.assign({matchedKey: t}, DOMAIN_DICTIONARY[t]);
  let best = null, bestLen = 0;
  for(const key in DOMAIN_DICTIONARY){
    if(t.includes(key) || key.includes(t)){
      if(key.length > bestLen){ bestLen = key.length; best = key; }
    }
  }
  if(best) return Object.assign({matchedKey: best}, DOMAIN_DICTIONARY[best]);
  return null;
}

function tokenizeForKeywords(text){
  if(!text) return [];
  return text.toLowerCase().replace(/[^a-z0-9\s-]/g, " ").split(/\s+/).filter(w => w.length > 2 && !STOPWORDS.has(w) && !/^\d+$/.test(w));
}

function extractKeywords(sourceText, maxKeywords = 10){
  const words = tokenizeForKeywords(sourceText);
  if(words.length === 0) return [];

  const freq = {};
  words.forEach(w => { freq[w] = (freq[w] || 0) + 1; });

  // bigram frequency (adjacent non-stopword pairs) — compound technical terms
  // like "proximity sensor" or "waste management" carry more search signal as a unit
  const bigramFreq = {};
  for(let i = 0; i < words.length - 1; i++){
    const bg = words[i] + " " + words[i + 1];
    bigramFreq[bg] = (bigramFreq[bg] || 0) + 1;
  }

  // score: bigrams that occur 2+ times outrank unigrams of similar frequency,
  // since a recurring two-word technical term is more specific/useful for a
  // patent search than either word alone.
  const candidates = [];
  Object.entries(bigramFreq).forEach(([bg, count]) => {
    if(count >= 2) candidates.push({term: bg, score: count * 2.2});
  });
  Object.entries(freq).forEach(([w, count]) => {
    candidates.push({term: w, score: count});
  });

  candidates.sort((a,b) => b.score - a.score);

  // de-duplicate: drop a unigram if it's already fully contained in a higher-ranked bigram
  const chosen = [];
  const usedWords = new Set();
  for(const c of candidates){
    const parts = c.term.split(" ");
    if(parts.length === 1 && usedWords.has(parts[0])) continue;
    chosen.push(c.term);
    parts.forEach(p => usedWords.add(p));
    if(chosen.length >= maxKeywords) break;
  }
  return chosen;
}

function getPriorArtSourceText(){
  // IDF tool has no separate Field-of-Invention section, so this draws from
  // Title + Description only (the closest equivalents to the richer
  // Title+Field+Background+Detailed-Description source used in the full
  // Complete Specification builder).
  const title = document.getElementById('title').value || '';
  const description = document.getElementById('description').value || '';
  return [title, description].filter(Boolean).join(' ');
}

function runKeywordExtraction(){
  const sourceText = getPriorArtSourceText();
  priorArtKeywords = extractKeywords(sourceText, 10);
  priorArtSelected = new Set(priorArtKeywords);
  synonymSelected = {}; // fresh extraction may match different dictionary terms — don't carry stale selections
  renderPriorArtTool();
}

function toggleKeyword(term){
  if(!priorArtSelected) priorArtSelected = new Set(priorArtKeywords);
  if(priorArtSelected.has(term)) priorArtSelected.delete(term);
  else priorArtSelected.add(term);
  renderPriorArtTool();
}

function toggleSynonym(term, syn){
  if(!synonymSelected[term]) synonymSelected[term] = new Set();
  if(synonymSelected[term].has(syn)) synonymSelected[term].delete(syn);
  else synonymSelected[term].add(syn);
  renderPriorArtTool();
}

function setQueryMode(mode){
  queryMode = mode;
  renderPriorArtTool();
}

// Google Patents and Lens.org both document full support for AND/OR/NOT,
// quoted exact phrases, and parenthesised grouping in their query field
// (default operator is AND, left-associative — exactly the format
// buildBooleanQuery's "advanced"/"patent" modes already produce). So those
// two links use the current Boolean/synonym query when one exists, falling
// back to the plain space-joined keywords only if no query has been built
// yet. Espacenet, WIPO Patentscope, and InPASS use different field-coded
// query languages, so they keep the plain keyword join — passing them the
// same Boolean string is not confirmed to work and could look like a typo
// to their parsers.
function buildSearchLinks(selectedTerms, booleanQuery){
  const q = selectedTerms.join(' ');
  const encoded = encodeURIComponent(q);
  const boolEncoded = encodeURIComponent(booleanQuery && booleanQuery.trim() ? booleanQuery : q);
  return [
    { name: 'Google Patents', url: `https://patents.google.com/?q=${boolEncoded}`, status: 'verified' },
    { name: 'Lens.org', url: `https://www.lens.org/lens/search/patent/list?q=${boolEncoded}`, status: 'verified' },
    { name: 'Espacenet (EPO)', url: `https://worldwide.espacenet.com/patent/search?q=${encoded}`, status: 'guess',
      fallbackUrl: 'https://worldwide.espacenet.com/patent/search' },
    { name: 'WIPO Patentscope', url: `https://patentscope.wipo.int/search/en/result.jsf?query=${encoded}`, status: 'guess',
      fallbackUrl: 'https://patentscope.wipo.int/search/en/search.jsf' },
    { name: 'IP India / InPASS', url: 'https://www.ipindia.gov.in/public-search', status: 'guess-weak',
      fallbackUrl: 'https://www.ipindia.gov.in/public-search' },
  ];
}

// Derwent Innovation (Clarivate) has no public, unauthenticated search URL —
// it is a subscription product behind product login, unlike the five engines
// above. So this builds a ready-to-paste Boolean search string in Derwent's
// own field-code syntax (TI= title field, AB= abstract field) rather than a
// clickable deep-link, for users who have institutional Derwent access.
function buildDerwentSearchString(selectedTerms){
  if(!selectedTerms.length) return '';
  const quoted = selectedTerms.map(t => t.includes(' ') ? `"${t}"` : t);
  const titleClause = quoted.map(t => `TI=(${t})`).join(' OR ');
  const abstractClause = quoted.map(t => `AB=(${t})`).join(' OR ');
  return `(${titleClause}) OR (${abstractClause})`;
}

// Builds the concept groups (a selected keyword plus whichever of its
// dictionary synonyms the user has opted into) that feed the query builder.
function getQueryConceptGroups(selectedTerms){
  return selectedTerms.map(term => {
    const dict = classifyKeyword(term);
    const chosenSyns = dict ? Array.from(synonymSelected[term] || []) : [];
    return { term, variants: [term, ...chosenSyns] };
  });
}

function buildBooleanQuery(mode, selectedTerms){
  if(!selectedTerms.length) return '';
  const groups = getQueryConceptGroups(selectedTerms);
  const quote = t => t.includes(' ') ? `"${t}"` : t;

  if(mode === 'basic'){
    return groups.map(g => quote(g.term)).join(' AND ');
  }
  if(mode === 'patent'){
    return groups.map(g => {
      const phrases = g.variants.map(v => `"${v}"`);
      return phrases.length > 1 ? '(' + phrases.join(' OR ') + ')' : phrases[0];
    }).join(' AND ');
  }
  // advanced (default) — kept on a single line (no embedded newlines) so this
  // same string can be used verbatim as a URL query for Google Patents/Lens.org,
  // not just displayed in the copy box.
  return groups.map(g => {
    if(g.variants.length > 1) return '(' + g.variants.map(quote).join(' OR ') + ')';
    return quote(g.term);
  }).join(' AND ');
}

function escapeAttr(str){
  return String(str).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function renderPriorArtTool(){
  const body = document.getElementById('priorArtToolBody');
  if(!body) return;

  if(priorArtKeywords.length === 0){
    body.innerHTML = `<div class="prior-art-empty">No keywords extracted yet. Draft the Title and Description first — the more technical content available, the more specific the extracted terms will be. Click extract above when ready.</div>`;
    return;
  }

  if(!priorArtSelected) priorArtSelected = new Set(priorArtKeywords);

  let html = '<div class="keyword-chips">';
  priorArtKeywords.forEach(term => {
    const selected = priorArtSelected.has(term);
    html += `<button type="button" class="kw-chip${selected ? ' kw-selected' : ''}" onclick="toggleKeyword('${escapeAttr(term).replace(/'/g, "\\'")}')">${escapeAttr(term)}</button>`;
  });
  html += '</div>';
  html += `<div class="kw-hint">Click a term to include/exclude it from the search links, classification, and query builder below. Two-word terms are compound technical phrases found repeated in your text.</div>`;

  const selectedTerms = priorArtKeywords.filter(k => priorArtSelected.has(k));

  // ---- Typed classification table ----
  html += '<div class="classify-block"><div class="classify-block-title">Typed classification</div>';
  html += '<table class="classify-table"><tr><th>Term</th><th>Type</th><th>Domain</th></tr>';
  priorArtKeywords.forEach(term => {
    const dict = classifyKeyword(term);
    if(dict){
      html += `<tr><td>${escapeAttr(term)}</td><td><span class="type-badge">${escapeAttr(dict.category)}</span></td><td>${escapeAttr(dict.domain)}</td></tr>`;
    } else {
      html += `<tr><td>${escapeAttr(term)}</td><td><span class="type-badge unclassified">unclassified</span></td><td>—</td></tr>`;
    }
  });
  html += '</table>';
  html += `<div class="classify-note">Classification comes from a static, offline dictionary covering AI, Telecom, Blockchain, IoT, Healthcare, and FinTech (plus a few cross-domain terms) — it is not AI-guessed. Terms outside this coverage are honestly marked "unclassified" rather than assigned a made-up category; extend the dictionary in the tool's source to widen coverage.</div>`;
  html += '</div>';

  // ---- Synonym expansion (only for selected + classified terms) ----
  const classifiedSelected = selectedTerms.map(t => ({term: t, dict: classifyKeyword(t)})).filter(x => x.dict);
  html += '<div class="classify-block"><div class="classify-block-title">Synonym expansion</div>';
  if(classifiedSelected.length === 0){
    html += `<div class="synonym-empty">No selected term matches the domain dictionary yet, so there are no known synonyms to show. Synonym data only appears for terms recognised in the static dictionary above.</div>`;
  } else {
    classifiedSelected.forEach(({term, dict}) => {
      html += `<div class="synonym-row"><div class="synonym-term-label">${escapeAttr(term)}</div><div class="synonym-chips">`;
      const chosen = synonymSelected[term] || new Set();
      dict.synonyms.forEach(syn => {
        const isSel = chosen.has(syn);
        html += `<button type="button" class="syn-chip${isSel ? ' syn-selected' : ''}" onclick="toggleSynonym('${escapeAttr(term).replace(/'/g,"\\'")}','${escapeAttr(syn).replace(/'/g,"\\'")}')">${escapeAttr(syn)}</button>`;
      });
      html += '</div></div>';
    });
    html += `<div class="classify-note">Click a synonym to include it in the query builder below (OR-grouped with its parent term). Unselected synonyms are shown for reference only and are not searched.</div>`;
  }
  html += '</div>';

  // ---- Boolean query builder ----
  html += '<div class="classify-block"><div class="classify-block-title">Search query builder</div>';
  html += '<div class="query-mode-tabs">';
  [['basic','Basic'],['advanced','Advanced'],['patent','Patent-style']].forEach(([mode,label]) => {
    html += `<button type="button" class="query-mode-tab${queryMode===mode ? ' qm-active' : ''}" onclick="setQueryMode('${mode}')">${label}</button>`;
  });
  html += '</div>';
  const queryString = buildBooleanQuery(queryMode, selectedTerms);
  html += `<div class="query-string" id="builtQueryString">${escapeAttr(queryString || '(select at least one keyword above)')}</div>`;
  html += `<button type="button" class="copy-btn" onclick="copyQueryString(this)">Copy query string</button>`;
  html += `<div class="query-caveat">This query now drives the Google Patents and Lens.org links below directly — both engines document full support for AND/OR/NOT, quoted phrases, and parenthesised grouping in their search field (default operator is AND, so OR groups must stay inside parentheses, which is exactly how Basic/Advanced/Patent-style build them here). Espacenet and WIPO Patentscope use their own field-coded query languages instead, so their links below still use the plain keyword chips, not this string — paste this string into Espacenet's or WIPO's advanced/CQL search box yourself if you want to try it there.</div>`;
  html += '</div>';

  const links = buildSearchLinks(selectedTerms, queryString);
  const statusCopy = {
    verified: { tag: 'opens pre-filled with the Boolean/synonym query above', cls: 'tag-verified' },
    guess: { tag: 'best-guess pre-fill — unverified, may not apply your keywords correctly', cls: 'tag-guess' },
    'guess-weak': { tag: 'best-guess pre-fill — unconfirmed and likely unsupported by this site; fallback link provided', cls: 'tag-guess-weak' },
  };

  html += '<div class="search-links">';
  links.forEach(link => {
    const info = statusCopy[link.status];
    html += `<div class="search-link-row">
      <a href="${link.url}" target="_blank" rel="noopener noreferrer" class="search-link">${link.name}</a>
      <span class="search-link-tag ${info.cls}">${info.tag}</span>`;
    if(link.fallbackUrl){
      html += `<a href="${link.fallbackUrl}" target="_blank" rel="noopener noreferrer" class="search-link-fallback">[plain search page]</a>`;
    }
    html += '</div>';
  });
  html += '</div>';

  const derwentString = buildDerwentSearchString(selectedTerms);
  html += `<div class="derwent-block">
    <div class="derwent-head">
      <span class="search-link-tag tag-guess-weak">requires institutional Derwent Innovation login — no public search link exists</span>
    </div>
    <div class="derwent-string" id="derwentSearchString">${escapeAttr(derwentString)}</div>
    <button type="button" class="copy-btn" onclick="copyDerwentString(this)">Copy Derwent search string</button>
  </div>`;

  html += `<div class="prior-art-footnote">These links open public search engines in a new tab. Nothing is searched automatically and no results are fetched into this tool — you review and judge relevance yourself. Google Patents and Lens.org links use the Boolean/synonym query from the query builder above (both engines document AND/OR/NOT, quoted phrases, and parenthesised grouping as supported query syntax); if you haven't selected any keywords, both fall back to a plain keyword search instead. Espacenet and WIPO Patentscope pre-fill links use the plain keyword list, are educated guesses based on historically documented query parameters, and have not been independently confirmed to still work — if the keywords don't appear pre-filled, use the [plain search page] link and enter them manually. IP India's InPASS has no documented query-string support at all (it is form- and captcha-based); there is no working pre-fill for it here, only a direct link to the official public-search page, and multiple independent patent-search guides recommend using Google Patents or Lens.org for keyword searching and reserving InPASS for confirming the legal status of a specific Indian application you've already identified by number. Derwent Innovation (Clarivate) is a subscription product with no public unauthenticated search URL, so a ready-to-paste Boolean search string in Derwent's own field-code syntax is generated above instead of a link — paste it into Derwent's search bar once logged in with an institutional account.</div>`;

  body.innerHTML = html;
}

function copyDerwentString(btn){
  const text = document.getElementById('derwentSearchString').textContent;
  navigator.clipboard.writeText(text).then(()=>{
    const original = btn.textContent;
    btn.textContent = 'Copied ✓';
    btn.classList.add('copied');
    setTimeout(()=>{ btn.textContent = original; btn.classList.remove('copied'); }, 1500);
  }).catch(()=>{
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    btn.textContent = 'Copied ✓';
    setTimeout(()=>{ btn.textContent = 'Copy Derwent search string'; }, 1500);
  });
}

function copyQueryString(btn){
  const text = document.getElementById('builtQueryString').textContent;
  navigator.clipboard.writeText(text).then(()=>{
    const original = btn.textContent;
    btn.textContent = 'Copied ✓';
    btn.classList.add('copied');
    setTimeout(()=>{ btn.textContent = original; btn.classList.remove('copied'); }, 1500);
  }).catch(()=>{
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    btn.textContent = 'Copied ✓';
    setTimeout(()=>{ btn.textContent = 'Copy query string'; }, 1500);
  });
}

function toggleFullscreen(){
  if(!document.fullscreenElement){ document.documentElement.requestFullscreen(); }
  else{ document.exitFullscreen(); }
}

// Hard-blocks the title field at 15 words. Splits on whitespace, and if the
// 16th word would be added, truncates back to the first 15 words rather than
// letting the field exceed the limit — applied on every keystroke/paste so
// pasted text is capped too, not just typed text.
function enforceTitleWordLimit(el){
  const maxWords = 15;
  const raw = el.value;
  const words = raw.split(/\s+/).filter(Boolean);
  if(words.length > maxWords){
    const trimmed = words.slice(0, maxWords).join(' ');
    // preserve a trailing space if the user is mid-word-16 (still typing it)
    // so the cursor doesn't appear to jump backwards unexpectedly
    el.value = trimmed;
  }
  const count = el.value.split(/\s+/).filter(Boolean).length;
  const counter = document.getElementById('titleWordCount');
  if(counter){
    counter.textContent = count + ' / ' + maxWords + ' words';
    counter.style.color = count >= maxWords ? '#a01217' : '#888';
  }
}

// ---- Dynamic prompt generation ----
// Prompt text is not static: each template is a function of the current form
// state, so a prompt for a later section pulls in whatever has already been
// written in earlier sections (title, description, prior art, etc.) instead
// of asking the user to re-paste it. Bracketed [placeholders] remain for the
// few inputs a template genuinely cannot infer from other fields.

function getPromptContext(){
  const discYN = (document.querySelector('input[name=discYN]:checked') || {}).value || 'No';
  return {
    briefSummary: (document.getElementById('briefSummary').value || '').trim(),
    title: (document.getElementById('title').value || '').trim(),
    description: (document.getElementById('description').value || '').trim(),
    discYN: discYN,
    discDetail: (document.getElementById('discDetail').value || '').trim(),
    priorArt: (document.getElementById('priorArt').value || '').trim(),
    problemSolved: (document.getElementById('problemSolved').value || '').trim(),
    novelty: (document.getElementById('novelty').value || '').trim(),
    advantages: (document.getElementById('advantages').value || '').trim(),
    otherDesc: (document.getElementById('otherDesc').value || '').trim(),
    priorLit: (document.getElementById('priorLit').value || '').trim(),
    addlNotes: (document.getElementById('addlNotes').value || '').trim(),
  };
}

const PROMPT_TEMPLATES = {

  pb1: (ctx) => `Give me a precise title for the following invention, suitable for filing under the Indian Patents Act, 1970, in the style the Indian Patent Office (IPO) expects. The title should name the device/system, its core technical mechanism, and its primary application domain in one line — no marketing language, no "novel" or "innovative," no abbreviations unless standard in the art, and avoid trade names per IPO drafting norms. Keep it strictly under 15 words — this is a hard limit, not a guideline.

Invention summary (from Section 1 above):
${ctx.briefSummary || "[not yet filled — go back and complete the Brief invention summary field above]"}`,

  pb2: () => `No drafting needed for this section — it's factual contact information (names, institute addresses, institutional emails, contact numbers). Do not generate or guess inventor details; enter them directly from institutional records.`,

  pb3: (ctx) => `Write a formal description of the following invention in the style of the "detailed description" section of a complete specification filed under the Indian Patents Act, 1970 and Patent Rules, 2003 (Form 2 format used by the Indian Patent Office). Requirements:
- Assign a reference numeral to each major component (e.g., Component (201), Component (202)) and use those numerals consistently throughout, matching the convention used in IPO-filed specifications.
- Describe the operational sequence in order: what triggers the system, what each component does, how components interact, and the end result.
- Use precise technical language, third person, no marketing adjectives ("innovative," "revolutionary," "smart").
- Write as continuous prose paragraphs, not bullet points — this field is read as a single block on the disclosure form and should map cleanly onto the IPO's expected specification structure.
- Length: 400-700 words, dense with technical detail rather than restating the same point.

Invention title (from Section 1):
${ctx.title || "[not yet filled — go back and complete Section 1, Title of the invention]"}

Brief invention summary (from Section 1):
${ctx.briefSummary || "[not yet filled — go back and complete the Brief invention summary field]"}

Components and how they work: [list each component/module and a few lines on what it does]
Overall purpose/use case: [who uses this and why]`,

  pb4: (ctx) => `This is a factual disclosure question, not a drafting task — answer it from your own records rather than generating content. If you did present, publish, or exhibit this invention publicly, write one or two factual sentences stating: the venue/publication, the date, and exactly what was disclosed (e.g., title shown, abstract published, demo given). Do not embellish or imply disclosure that did not happen.

Under the Indian Patents Act, 1970, Section 31, only narrow categories of pre-filing disclosure are excused from destroying novelty: display at a government-notified exhibition, presentation before a "learned society" or publication in its transactions, or disclosure by the inventor with their consent — and even then, only if the complete specification is filed within 12 months of that disclosure. Since the 2024 Patent Rules amendment (Rule 29A), claiming this grace period also requires filing Form 31 with supporting evidence of the disclosure. A conference paper, seminar talk, or exhibition outside these categories, or beyond the 12-month window, can still destroy novelty — so state the venue and date precisely; do not assume a disclosure is automatically protected.

Invention title (from Section 1, for reference only):
${ctx.title || "[not yet filled]"}

Venue/publication: [name]
Date: [date]
What was shown/published: [brief description]`,

  pb5: (ctx) => `Use the prior art search assistant above to find candidate patents first — extract keywords from your Title and Description, then check the Google Patents and Lens.org links (most reliable), with Espacenet/WIPO Patentscope/InPASS as secondary checks, and the generated Derwent search string if your institution has Derwent Innovation access. Once you've identified relevant patent numbers and read their abstracts, use this prompt: for each prior-art patent or publication found, write a short, factual entry with two parts: (1) the patent/publication number, (2) a 1-2 sentence summary of what it discloses based on its abstract — main components/method and stated purpose only. This will be used to assess anticipation under Sections 13 and 29 of the Indian Patents Act, 1970 (absolute novelty standard — prior use or publication anywhere in the world counts, not just India), so be precise about what the document actually claims rather than rounding up to something broader. Keep each entry to 2-3 sentences total — this is a reference list, not a background essay. Do not invent patent numbers or details not present in the source; if the abstract doesn't specify something, say only what is supported. Critically: only summarize a document you have actually opened and read — if you have not opened the link and read the real abstract, do not ask an AI assistant to summarize it from the patent number alone, as it will likely fabricate a plausible-sounding but unfounded summary.

Invention title (from Section 1):
${ctx.title || "[not yet filled — go back and complete Section 1]"}

Invention description, for identifying what to search against (from Section 3):
${ctx.description || "[not yet filled — go back and complete Section 3, Description of the Invention]"}

Patent numbers / publications to summarize: [paste patent numbers, or paste abstracts/text you already have]`,

  pb6: (ctx) => `In 3-4 sentences, explain the unmet need this invention addresses and how it solves that need, suitable for a disclosure intended for filing under the Indian Patents Act, 1970. Name the specific problem/gap in current practice in one sentence, then state how the invention's components close that gap, then end with the outcome (efficiency gain, risk reduction, cost reduction, etc.) as a direct consequence — this will support the inventive-step argument under Section 2(1)(ja), so keep it specific rather than promotional. Avoid generic phrases like "revolutionizes" or "transforms." This is a short field, not an essay — do not pad it.

Invention description, so the components you reference actually exist in the disclosure (from Section 3):
${ctx.description || "[not yet filled — go back and complete Section 3, Description of the Invention]"}

Problem/unmet need: [describe the gap in current practice]
How the invention addresses it: [1-2 lines mapping components to the problem they solve]`,

  pb7: (ctx) => `List the novel/non-obvious elements of this invention against the prior art below, framed against the Indian Patents Act, 1970 test for inventive step under Section 2(1)(ja): a feature qualifies if it shows technical advance, or has economic significance, or both, and is not obvious to a "person skilled in the art." For each element, give one short line: name it, say which pathway it satisfies, and what distinguishes it from the closest prior art. List 3-5 elements as a short list, not a paragraph — keep each line under 20 words where possible. Do not invent novelty for a standard technique merely applied in a new context.

Components/features of the invention (from Section 3 description):
${ctx.description || "[not yet filled — go back and complete Section 3, Description of the Invention]"}

Closest prior art already identified (from Section 5):
${ctx.priorArt || "[not yet filled — go back and complete Section 5, Prior Filings / Closest Prior Art]"}`,

  pb8: (ctx) => `Write a short "usefulness and advantages" section for a patent disclosure intended for filing under the Indian Patents Act, 1970. One opening line on the overall value, then 3-4 short bullet-style lines, each tied to a specific component or mechanism — not vague benefits. Add SDG alignment only if genuinely applicable, as one short line naming the specific goal(s), not a padded list. Where an advantage is a cost/efficiency gain rather than a technical improvement, say so plainly — under Section 2(1)(ja), economic significance is an independent basis for inventive step in India, so don't undersell it. No superlatives ("groundbreaking," "best-in-class"). Keep the whole thing short — this is a summary field, not an essay.

Novel/non-obvious elements already identified (from Section 7 — ground the advantages in these, don't invent new ones):
${ctx.novelty || "[not yet filled — go back and complete Section 7, Novel / Non-Obvious Elements]"}

Key advantages to cover: [list them, each with the mechanism behind it]
Relevant SDGs (if any, leave blank if not applicable): [list]`,

  pbfig: (ctx) => `Create a system block diagram as a single SVG, based on the invention description below, matching this exact style (this is the established format used for these disclosures, not a generic patent drawing):

- Layout: rectangular boxes for each component, connected by arrows showing data/signal flow between them. One component may be an ellipse if it represents a network/cloud/external medium rather than a physical unit.
- Each box: white fill, black 2px stroke, no shading or gradients, no color anywhere in the diagram.
- Inside each box: the component name (wrapped across 2 lines if needed) plus its reference numeral in parentheses, e.g. "Central Control Unit (202)" — numeral goes inside the box with the name, not outside with a pointer line.
- Reference numerals must exactly match the numerals already used in the invention description (Section 3) — do not renumber or invent new ones.
- Arrows: black, 2px stroke, solid arrowhead at the destination end, drawn point-to-point between boxes in the direction data/control actually flows per the description.
- Label each arrow with a short 2-4 word caption describing what passes along it (e.g., "Occupancy Data," "Availability Data"), placed near the arrow, plain text, no background box.
- Title: one line at the top in bold, format "FIG. [number] – [invention title]," sized larger than body text.
- Font: a single plain sans-serif (e.g., Arial) throughout, one size for body labels, one larger size for the title — no decorative fonts, no italics.
- Canvas: roughly 1200×700 viewBox, landscape orientation, white background, components arranged left-to-right or top-to-bottom following the actual data flow sequence in the description (sensors/inputs on one side, processing in the middle, output/display/user-facing components on the other side).
- Output only the SVG code — no explanation before or after.

Invention description to base the figure on (from Section 3):
${ctx.description || "[not yet filled — go back and complete Section 3, Description of the Invention]"}

Figure number and title: [e.g., "FIG. 1 – ${ctx.title || '[invention title]'}"]`,

  pb9: (ctx) => `These four fields are administrative, not narrative — most disclosures will genuinely be "Nil" for some of them, and an LLM should not pad them with invented content. Use this prompt only if you have real material to summarize (e.g., a manuscript, a sketch description, or specific references), otherwise leave the field as "-Nil-".

If you have a manuscript or attachment to describe: write one or two sentences identifying what it is (e.g., journal manuscript title and status) — do not summarize its content here, just identify it.

If you have figures to caption: for each uploaded sketch, write a short caption (under 12 words) naming what the figure shows and which reference numerals from the description it depicts — keep numerals consistent with the IPO drawing-sheet convention used in the description section.

If you have literature to distinguish from prior art: list each reference with a one-line note on how it differs from this invention — not a restatement of the prior-art section above. This list will be used to support the novelty (Section 13/29) and inventive-step (Section 2(1)(ja)) arguments before the Indian Patent Office, so be precise rather than exhaustive.

Prior art already on file (from Section 5, for reference so you don't repeat it):
${ctx.priorArt || "[none entered yet in Section 5]"}

Material to describe: [paste manuscript title/status, figure contents, or reference list]`,
};

function renderPromptText(escapedTemplate){
  return escapedTemplate.replace(/\[([^\]]+)\]/g, (m, inner) => `<span class="placeholder">[${inner}]</span>`);
}

function refreshAllPrompts(){
  const ctx = getPromptContext();
  Object.keys(PROMPT_TEMPLATES).forEach(id => {
    const node = document.getElementById(id + '-text');
    if(!node) return;
    const raw = PROMPT_TEMPLATES[id](ctx);
    node.innerHTML = renderPromptText(escapeHtml(raw));
  });
}

function togglePrompt(id){
  const body = document.getElementById(id);
  const arrow = document.getElementById(id + '-arrow');
  const isOpen = body.classList.contains('open');
  body.classList.toggle('open');
  arrow.textContent = isOpen ? 'show ▾' : 'hide ▴';
}

function copyPrompt(textId, btn){
  const text = document.getElementById(textId).textContent;
  navigator.clipboard.writeText(text).then(()=>{
    const original = btn.textContent;
    btn.textContent = 'Copied ✓';
    btn.classList.add('copied');
    setTimeout(()=>{ btn.textContent = original; btn.classList.remove('copied'); }, 1500);
  }).catch(()=>{
    // fallback for environments without clipboard API
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    btn.textContent = 'Copied ✓';
    setTimeout(()=>{ btn.textContent = 'Copy prompt'; }, 1500);
  });
}

function addInventor(prefillName, prefillAddress, prefillEmail, prefillPhone){
  inventorCount++;
  const id = inventorCount;
  const row = document.createElement('tr');
  row.className = 'inventor-row';
  row.id = 'inventor-' + id;
  row.innerHTML = `
    <td><input type="text" class="inv-name" placeholder="Full name" oninput="renderPreview()"></td>
    <td><input type="text" class="inv-address" placeholder="Department, institute, campus" oninput="renderPreview()"></td>
    <td><input type="text" class="inv-email" placeholder="name@institute.edu" oninput="renderPreview()"></td>
    <td><input type="text" class="inv-phone" placeholder="+91 ..." oninput="renderPreview()"></td>
    <td><button class="remove-btn" onclick="removeInventor(${id})">remove</button></td>`;
  document.getElementById('inventorList').appendChild(row);
  renderPreview();
}

function removeInventor(id){
  const el = document.getElementById('inventor-' + id);
  if(el) el.remove();
  renderPreview();
}

// seed with one inventor row
addInventor();


function escapeHtml(str){
  const div = document.createElement('div');
  div.innerText = str;
  return div.innerHTML;
}

function fillField(id, value, fallback){
  const el = document.getElementById(id);
  if(value && value.trim() !== ''){
    el.innerHTML = escapeHtml(value).replace(/\n/g, '<br>');
    el.classList.remove('empty');
  } else {
    el.textContent = fallback || '—';
    el.classList.add('empty');
  }
}

// Collects one record per inventor row, skipping rows where every field is
// blank (so an unused trailing row added by mistake doesn't get exported).
function getInventorRecords(){
  const names = Array.from(document.querySelectorAll('.inv-name')).map(i=>i.value.trim());
  const addresses = Array.from(document.querySelectorAll('.inv-address')).map(i=>i.value.trim());
  const emails = Array.from(document.querySelectorAll('.inv-email')).map(i=>i.value.trim());
  const phones = Array.from(document.querySelectorAll('.inv-phone')).map(i=>i.value.trim());
  const records = [];
  for(let i = 0; i < names.length; i++){
    const r = {name: names[i] || '', address: addresses[i] || '', email: emails[i] || '', phone: phones[i] || ''};
    if(r.name || r.address || r.email || r.phone) records.push(r);
  }
  return records;
}

function renderInventorPreviewTable(){
  const target = document.getElementById('p-inventorTable');
  const records = getInventorRecords();
  if(records.length === 0){
    target.innerHTML = '<div style="padding:8px 10px; color:#aaa; font-style:italic; font-family:var(--sans); font-size:12.5px;">—</div>';
    return;
  }
  let html = '<table class="doc-table" style="margin:0;"><tr><th>Name</th><th>Institute address</th><th>Email ID</th><th>Contact number</th></tr>';
  records.forEach(r => {
    html += `<tr><td${r.name?'':' class="empty"'}>${escapeHtml(r.name) || '—'}</td><td${r.address?'':' class="empty"'}>${escapeHtml(r.address) || '—'}</td><td${r.email?'':' class="empty"'}>${escapeHtml(r.email) || '—'}</td><td${r.phone?'':' class="empty"'}>${escapeHtml(r.phone) || '—'}</td></tr>`;
  });
  html += '</table>';
  target.innerHTML = html;
}

function countFilled(){
  let filled = 0;
  const checks = [
    document.getElementById('title').value.trim() || document.getElementById('discNo').value.trim() || document.getElementById('briefSummary').value.trim(),
    document.querySelectorAll('.inv-name')[0] && document.querySelectorAll('.inv-name')[0].value.trim(),
    document.getElementById('description').value.trim(),
    document.getElementById('discDetail').value.trim() || true, // section 4 always has a Yes/No default
    document.getElementById('priorArt').value.trim(),
    document.getElementById('problemSolved').value.trim(),
    document.getElementById('novelty').value.trim(),
    document.getElementById('advantages').value.trim(),
    document.getElementById('otherDesc').value.trim() || document.getElementById('priorLit').value.trim() || document.getElementById('addlNotes').value.trim() || figures.length>0
  ];
  checks.forEach((c,i)=>{
    const statusEl = document.getElementById('status-'+(i+1));
    if(c){ filled++; statusEl.textContent='filled'; statusEl.classList.add('filled'); }
    else{ statusEl.textContent='empty'; statusEl.classList.remove('filled'); }
  });
  document.getElementById('progress-label').textContent = filled + ' / 9 sections';
  document.getElementById('progress-fill').style.width = (filled/9*100) + '%';
}

function renderPreview(){
  fillField('p-briefSummary', document.getElementById('briefSummary').value);
  fillField('p-discNo', document.getElementById('discNo').value);
  fillField('p-title', document.getElementById('title').value);

  renderInventorPreviewTable();

  fillField('p-description', document.getElementById('description').value);


  const discYN = document.querySelector('input[name=discYN]:checked').value;
  const discDetail = document.getElementById('discDetail').value.trim();
  fillField('p-disclosure', discYN + (discDetail ? ('\n' + discDetail) : ''));

  fillField('p-priorArt', document.getElementById('priorArt').value);
  fillField('p-problemSolved', document.getElementById('problemSolved').value);
  fillField('p-novelty', document.getElementById('novelty').value);
  fillField('p-advantages', document.getElementById('advantages').value);
  fillField('p-otherDesc', document.getElementById('otherDesc').value);

  const figsCell = document.getElementById('p-figs');
  if(figures.length){
    figsCell.classList.remove('empty');
    figsCell.innerHTML = '<div class="doc-figs">' + figures.map(f =>
      `<div><img src="${f.dataUrl}"><div style="font-size:10px;text-align:center;margin-top:2px;">${escapeHtml(f.caption||'')}</div></div>`
    ).join('') + '</div>';
  } else {
    figsCell.textContent = '—';
    figsCell.classList.add('empty');
  }

  fillField('p-priorLit', document.getElementById('priorLit').value);
  fillField('p-addlNotes', document.getElementById('addlNotes').value);

  countFilled();
  refreshAllPrompts();
  saveDraftDebounced();
}

document.querySelectorAll('textarea, input[type=text]').forEach(el=>{
  el.addEventListener('input', renderPreview);
});
document.querySelectorAll('input[name=discYN]').forEach(el=>{
  el.addEventListener('change', renderPreview);
});

// ---- figure artwork: SVG paste-and-render + image upload ----
const uploadZone = document.getElementById('uploadZone');
const figInput = document.getElementById('figInput');

figInput.addEventListener('change', e => handleFiles(e.target.files));
uploadZone.addEventListener('dragover', e=>{ e.preventDefault(); uploadZone.classList.add('dragover'); });
uploadZone.addEventListener('dragleave', ()=> uploadZone.classList.remove('dragover'));
uploadZone.addEventListener('drop', e=>{
  e.preventDefault();
  uploadZone.classList.remove('dragover');
  handleFiles(e.dataTransfer.files);
});

function setArtworkStatus(msg){
  const el = document.getElementById('artworkStatus');
  if(el) el.textContent = msg || '';
}

// Renders an SVG string onto a white-backed canvas and returns a PNG data URL.
// Rasterizing here (rather than embedding raw SVG) keeps figure handling
// uniform with uploaded images for both the live preview and jsPDF export,
// which only consistently accepts raster formats via addImage.
function svgToPngDataUrl(svgString, width = 800, height = 560){
  return new Promise((resolve, reject) => {
    try{
      const svgBlob = new Blob([svgString], {type: 'image/svg+xml;charset=utf-8'});
      const url = URL.createObjectURL(svgBlob);
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0, width, height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL('image/png'));
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Could not parse this as a valid SVG image. Check the markup starts with <svg> and is well-formed.')); };
      img.src = url;
    }catch(err){ reject(err); }
  });
}

// Normalizes any uploaded image format to a PNG data URL via canvas, so
// every figure in the `figures` array (SVG-rendered or uploaded) shares one
// type, and caps the longest side so large phone-camera photos stay light.
function imageDataUrlToPng(dataUrl, maxDim = 1400){
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      let w = img.naturalWidth, h = img.naturalHeight;
      if(Math.max(w, h) > maxDim){
        const scale = maxDim / Math.max(w, h);
        w = Math.round(w * scale);
        h = Math.round(h * scale);
      }
      const canvas = document.createElement('canvas');
      canvas.width = w; canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => reject(new Error('Could not decode this image file. Try a different format (PNG/JPEG).'));
    img.src = dataUrl;
  });
}

function fileToDataUrl(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('Could not read this file.'));
    reader.readAsDataURL(file);
  });
}

async function renderPastedSvg(){
  const textarea = document.getElementById('svgPasteArea');
  const svgSource = textarea.value.trim();
  if(!svgSource){ setArtworkStatus('Paste SVG markup first.'); return; }
  if(!/^<svg[\s>]/.test(svgSource)){ setArtworkStatus("Doesn't look like it starts with <svg> — check you copied the whole tag."); return; }
  setArtworkStatus('Rendering…');
  try{
    const pngDataUrl = await svgToPngDataUrl(svgSource, 800, 560);
    figures.push({dataUrl: pngDataUrl, caption: '', source: 'svg'});
    setArtworkStatus('Rendered ✓');
    textarea.value = '';
    renderFigList();
    renderPreview();
  }catch(err){
    setArtworkStatus('Render failed: ' + err.message);
  }
}

async function handleFiles(fileList){
  for(const file of Array.from(fileList)){
    if(!file.type.startsWith('image/')) continue;
    setArtworkStatus('Processing image…');
    try{
      const rawDataUrl = await fileToDataUrl(file);
      const normalizedPng = await imageDataUrlToPng(rawDataUrl);
      figures.push({dataUrl: normalizedPng, caption: '', source: 'image'});
      setArtworkStatus('Image added ✓');
      renderFigList();
      renderPreview();
    }catch(err){
      setArtworkStatus('Upload failed: ' + err.message);
    }
  }
}

function renderFigList(){
  const list = document.getElementById('figList');
  list.innerHTML = '';
  figures.forEach((f, idx)=>{
    const thumb = document.createElement('div');
    thumb.className = 'figure-thumb';
    const sourceLabel = f.source === 'svg' ? 'rendered from SVG' : 'uploaded image';
    thumb.innerHTML = `
      <button class="remove-x" onclick="removeFigure(${idx})">×</button>
      <img src="${f.dataUrl}">
      <input type="text" placeholder="Caption (e.g. Fig 1)" value="${escapeHtml(f.caption)}" oninput="updateCaption(${idx}, this.value)">
      <div class="artwork-source-tag">${sourceLabel}</div>
    `;
    list.appendChild(thumb);
  });
}

function updateCaption(idx, val){
  figures[idx].caption = val;
  renderPreview();
}

function removeFigure(idx){
  figures.splice(idx,1);
  renderFigList();
  renderPreview();
}

function clearAll(){
  if(!confirm('Clear every field? This cannot be undone.')) return;
  document.querySelectorAll('textarea, input[type=text]').forEach(el=>el.value='');
  document.getElementById('inventorList').innerHTML='';
  inventorCount = 0;
  addInventor();
  figures = [];
  renderFigList();
  document.querySelector('input[name=discYN][value="No"]').checked = true;
  const counter = document.getElementById('titleWordCount');
  if(counter){ counter.textContent = '0 / 15 words'; counter.style.color = '#888'; }
  priorArtKeywords = [];
  priorArtSelected = null;
  synonymSelected = {};
  queryMode = 'advanced';
  renderPriorArtTool();
  renderPreview();
  clearSavedDraft();
}

function exportForSpecBuilder(){
  // Bridge format consumed by the Complete Specification Builder's "Import
  // from Disclosure" feature. Only fields that transfer honestly are
  // included: title and figures carry over as-is (no rewriting needed).
  // briefSummary/description/priorArt/problemSolved/novelty/advantages carry
  // over as unrefined SEED TEXT only — the receiving tool must visibly mark
  // these as not-yet-drafted-to-specification-standard, never auto-finalize
  // them into Field / Detailed Description / Background / Objects / Summary
  // prose. Inventor contact details, disclosure No., and disclosure-history
  // fields have no corresponding slot in a Form 2 specification and are
  // intentionally NOT included here — only inventor names are carried, for
  // reference.
  const names = getInventorRecords().map(r => r.name).filter(Boolean);

  const bridge = {
    schema: "complete-spec-builder-bridge",
    schemaVersion: 1,
    sourceTool: "Invention Disclosure Form Builder",
    exportedAt: new Date().toISOString(),
    title: document.getElementById('title').value.trim(),
    inventorNamesForReference: names, // informational only; spec builder has no field for this
    seeds: {
      fieldOfInvention: document.getElementById('briefSummary').value.trim(),
      detailedDescription: document.getElementById('description').value.trim(),
      background: document.getElementById('priorArt').value.trim(),
      objects: document.getElementById('problemSolved').value.trim(),
      novelty: document.getElementById('novelty').value.trim(),
      advantages: document.getElementById('advantages').value.trim(),
    },
    figures: figures.map(f => ({ dataUrl: f.dataUrl, caption: f.caption || '' })),
  };

  const blob = new Blob([JSON.stringify(bridge, null, 2)], {type: 'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const safeTitle = (bridge.title || 'disclosure').replace(/[^a-z0-9]+/gi, '_').slice(0, 60);
  a.href = url;
  a.download = safeTitle + '_bridge.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

function getDisclosureFields(){
  const discYN = document.querySelector('input[name=discYN]:checked').value;
  const discDetail = document.getElementById('discDetail').value.trim();
  const disclosureText = discYN + (discDetail ? ('\n' + discDetail) : '');

  function getVal(label, text){
    return {label, text: (text && text.trim()) ? text.trim() : '—', kind: 'text'};
  }

  return [
    getVal('Brief invention summary', document.getElementById('briefSummary').value),
    getVal('Invention disclosure No.', document.getElementById('discNo').value),
    getVal('Title of the invention', document.getElementById('title').value),
    { label: 'Inventor(s)', kind: 'table', rows: getInventorRecords(),
      columns: ['Name', 'Institute address', 'Email ID', 'Contact number'] },
    getVal('Description of the invention', document.getElementById('description').value),
    getVal('Public disclosure in seminars/exhibitions or plans to disclose', disclosureText),
    getVal('Prior filings: disclosure document or provisional application filed', document.getElementById('priorArt').value),
    getVal('How are the problems solved / outstanding need met?', document.getElementById('problemSolved').value),
    getVal('Elements considered novel / non-obvious', document.getElementById('novelty').value),
    getVal('Usefulness / advantages over currently available technology', document.getElementById('advantages').value),
    getVal('Other additional description / attachments / manuscript', document.getElementById('otherDesc').value),
    getVal('Relevant literature distinguishing prior art', document.getElementById('priorLit').value),
    getVal('Additional notes or comments', document.getElementById('addlNotes').value),
  ];
}

async function exportPDF(){
  const exportBtn = document.querySelector('button[onclick="exportPDF()"]');
  const originalLabel = exportBtn ? exportBtn.textContent : null;
  if(exportBtn){ exportBtn.textContent = 'Loading PDF engine…'; exportBtn.disabled = true; }
  try{
    await ensureJsPDFLoaded();
  }catch(err){
    alert('Could not load the PDF export library: ' + err.message);
    if(exportBtn){ exportBtn.textContent = originalLabel; exportBtn.disabled = false; }
    return;
  }
  if(exportBtn){ exportBtn.textContent = originalLabel; exportBtn.disabled = false; }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({unit:'pt', format:'a4'});
  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 40;
  const usableWidth = pageWidth - margin*2;
  let y = 50;

  const logoEl = document.querySelector('.topbar-left img');
  if(logoEl && logoEl.src && logoEl.src.startsWith('data:image')){
    const logoW = 130, logoH = 38;
    try{ doc.addImage(logoEl.src, margin, 28, logoW, logoH); }catch(e){}
  }

  doc.setFont('times','bold');
  doc.setFontSize(16);
  doc.text('INVENTION DISCLOSURE FORM', pageWidth/2, y, {align:'center'});
  y += 26;
  doc.setDrawColor(196,22,28);
  doc.setLineWidth(1.5);
  doc.line(margin, y, pageWidth-margin, y);
  y += 20;

  const fields = getDisclosureFields();

  doc.setFontSize(10);

  function checkPageBreak(neededHeight){
    if(y + neededHeight > doc.internal.pageSize.getHeight() - 40){
      doc.addPage();
      y = 50;
    }
  }

  fields.forEach(f=>{
    checkPageBreak(30);
    doc.setFont('helvetica','bold');
    doc.setFontSize(10);
    doc.text(f.label, margin, y);
    y += 14;

    if(f.kind === 'table'){
      const rows = f.rows && f.rows.length ? f.rows : [{name:'—', address:'—', email:'—', phone:'—'}];
      const colWidths = [usableWidth*0.22, usableWidth*0.34, usableWidth*0.24, usableWidth*0.20];
      const colKeys = ['name','address','email','phone'];
      const rowH = 16;
      checkPageBreak(rowH * (rows.length + 1) + 10);

      // header row
      doc.setFont('helvetica','bold');
      doc.setFontSize(9);
      let cx = margin;
      f.columns.forEach((colLabel, ci)=>{
        doc.rect(cx, y, colWidths[ci], rowH);
        doc.text(colLabel, cx + 4, y + rowH - 5);
        cx += colWidths[ci];
      });
      y += rowH;

      // data rows
      doc.setFont('helvetica','normal');
      rows.forEach(r=>{
        checkPageBreak(rowH);
        cx = margin;
        colKeys.forEach((key, ci)=>{
          doc.rect(cx, y, colWidths[ci], rowH);
          const cellText = doc.splitTextToSize(r[key] || '—', colWidths[ci] - 8)[0] || '—';
          doc.text(cellText, cx + 4, y + rowH - 5);
          cx += colWidths[ci];
        });
        y += rowH;
      });
      y += 10;
    } else {
      doc.setFont('helvetica','normal');
      doc.setFontSize(10);
      const lines = doc.splitTextToSize(f.text, usableWidth);
      lines.forEach(line=>{
        checkPageBreak(14);
        doc.text(line, margin, y);
        y += 13;
      });
      y += 10;
    }

    doc.setDrawColor(210,210,210);
    doc.line(margin, y-4, pageWidth-margin, y-4);
    y += 6;
  });

  // figures section
  if(figures.length){
    checkPageBreak(30);
    doc.setFont('helvetica','bold');
    doc.text('Labelled sketches / figures', margin, y);
    y += 16;

    figures.forEach(f=>{
      const imgW = 200, imgH = 150;
      checkPageBreak(imgH + 30);
      try{
        doc.addImage(f.dataUrl, margin, y, imgW, imgH, undefined, 'MEDIUM');
      }catch(e){}
      doc.setFont('helvetica','normal');
      doc.setFontSize(9);
      doc.text(f.caption || '', margin, y + imgH + 12);
      y += imgH + 26;
    });
  }

  doc.save('Invention_Disclosure_Form.pdf');
}

// Reads a data URL's natural pixel dimensions, used to size embedded
// figures in the Word export proportionally rather than with a fixed
// guessed aspect ratio.
function getImageNaturalSize(dataUrl){
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({width: img.naturalWidth || 200, height: img.naturalHeight || 150});
    img.onerror = () => resolve({width: 200, height: 150});
    img.src = dataUrl;
  });
}

function scaleToMaxWidthPx(naturalW, naturalH, maxWidthPx = 420){
  if(!naturalW || !naturalH) return {width: maxWidthPx, height: Math.round(maxWidthPx * 0.7)};
  if(naturalW <= maxWidthPx) return {width: naturalW, height: naturalH};
  const ratio = maxWidthPx / naturalW;
  return {width: maxWidthPx, height: Math.round(naturalH * ratio)};
}

// jsPDF embeds a font-table content type, but omits the corresponding
// relationship entry in word/_rels/document.xml.rels — this leaves the
// .docx technically invalid (an unreferenced part) even though
// Word/LibreOffice generally open it without complaint. Patched here at
// generation time so every export is a clean, fully-valid OOXML package
// rather than relying on viewer tolerance.
async function fixMissingFontTableRelationship(blob){
  try{
    const zip = await JSZip.loadAsync(blob);
    const relsPath = 'word/_rels/document.xml.rels';
    const relsFile = zip.file(relsPath);
    if(!relsFile) return blob; // unexpected package shape; leave untouched rather than guess

    let relsXml = await relsFile.async('string');
    if(relsXml.includes('relationships/fontTable')) return blob; // already referenced, nothing to fix

    const existingIds = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g)).map(m => parseInt(m[1], 10));
    const nextId = (existingIds.length ? Math.max(...existingIds) : 0) + 1;
    const fontTableRel = `<Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`;

    if(!relsXml.includes('</Relationships>')) return blob; // not the shape we expect; don't risk corrupting it
    relsXml = relsXml.replace('</Relationships>', fontTableRel + '</Relationships>');

    zip.file(relsPath, relsXml);
    return await zip.generateAsync({type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  }catch(e){
    console.warn('fontTable relationship patch skipped:', e.message);
    return blob; // fall back to the unpatched (still functional in practice) file rather than fail the export
  }
}

async function exportWord(){
  const exportBtn = document.querySelector('button[onclick="exportWord()"]');
  const originalLabel = exportBtn ? exportBtn.textContent : null;
  if(exportBtn){ exportBtn.textContent = 'Generating…'; exportBtn.disabled = true; }

  try{
    await ensureWordExportLibsLoaded();
    const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType, BorderStyle,
            Table, TableRow, TableCell, WidthType, ShadingType } = window.docx;

    const fields = getDisclosureFields();
    const children = [];

    // Header: logo (if available) + title, matching the PDF export's header treatment.
    const logoEl = document.querySelector('.topbar-left img');
    const headerRuns = [];
    if(logoEl && logoEl.src && logoEl.src.startsWith('data:image')){
      try{
        headerRuns.push(new ImageRun({
          data: logoEl.src,
          transformation: {width: 140, height: 41},
          type: 'png',
        }));
        headerRuns.push(new TextRun({text: '   ', size: 1}));
      }catch(e){ /* logo embed is non-critical; continue without it */ }
    }
    headerRuns.push(new TextRun({text: 'INVENTION DISCLOSURE FORM', bold: true, size: 32, font: 'Arial'}));
    children.push(new Paragraph({
      children: headerRuns,
      alignment: AlignmentType.CENTER,
      spacing: {after: 200},
      border: {bottom: {style: BorderStyle.SINGLE, size: 12, color: 'C4161C', space: 4}},
    }));

    // Field rows: bold label paragraph, then either a normal-weight value
    // paragraph or (for the inventor field) a real table, matching the PDF's layout.
    const tableBorder = {style: BorderStyle.SINGLE, size: 2, color: 'AAAAAA'};
    const tableBorders = {top: tableBorder, bottom: tableBorder, left: tableBorder, right: tableBorder};
    fields.forEach(f => {
      children.push(new Paragraph({
        children: [new TextRun({text: f.label, bold: true, size: 22, font: 'Arial'})],
        spacing: {before: 160, after: 40},
      }));

      if(f.kind === 'table'){
        const rows = (f.rows && f.rows.length) ? f.rows : [{name:'—', address:'—', email:'—', phone:'—'}];
        const colKeys = ['name','address','email','phone'];
        const colWidths = [1800, 3000, 2400, 2160]; // sums to 9360 DXA (US Letter content width)
        function headerCell(text, width){
          return new TableCell({
            borders: tableBorders, width: {size: width, type: WidthType.DXA},
            shading: {fill: 'F0EEE5', type: ShadingType.CLEAR},
            margins: {top: 60, bottom: 60, left: 90, right: 90},
            children: [new Paragraph({children: [new TextRun({text, bold: true, size: 18, font: 'Arial'})]})],
          });
        }
        function dataCell(text, width){
          return new TableCell({
            borders: tableBorders, width: {size: width, type: WidthType.DXA},
            margins: {top: 60, bottom: 60, left: 90, right: 90},
            children: [new Paragraph({children: [new TextRun({text: text || '—', size: 18, font: 'Arial'})]})],
          });
        }
        const headerRow = new TableRow({children: f.columns.map((c,i) => headerCell(c, colWidths[i]))});
        const dataRows = rows.map(r => new TableRow({children: colKeys.map((k,i) => dataCell(r[k], colWidths[i]))}));
        children.push(new Table({
          width: {size: 9360, type: WidthType.DXA},
          columnWidths: colWidths,
          rows: [headerRow, ...dataRows],
        }));
        children.push(new Paragraph({children: [], spacing: {after: 80}}));
      } else {
        const valueLines = f.text.split('\n');
        valueLines.forEach(line => {
          children.push(new Paragraph({
            children: [new TextRun({text: line, size: 22, font: 'Arial'})],
            spacing: {after: 20},
          }));
        });
        children.push(new Paragraph({
          children: [],
          border: {bottom: {style: BorderStyle.SINGLE, size: 4, color: 'D2D2D2', space: 1}},
          spacing: {after: 80},
        }));
      }
    });

    // Figures: same content as the PDF export, sized proportionally per image.
    if(figures.length){
      children.push(new Paragraph({
        children: [new TextRun({text: 'Labelled sketches / figures', bold: true, size: 22, font: 'Arial'})],
        spacing: {before: 160, after: 100},
      }));

      for(const f of figures){
        const {width: naturalW, height: naturalH} = await getImageNaturalSize(f.dataUrl);
        const {width: w, height: h} = scaleToMaxWidthPx(naturalW, naturalH, 420);
        try{
          children.push(new Paragraph({
            children: [new ImageRun({data: f.dataUrl, transformation: {width: w, height: h}, type: 'png'})],
            alignment: AlignmentType.CENTER,
            spacing: {after: 60},
          }));
        }catch(e){
          children.push(new Paragraph({
            children: [new TextRun({text: '[ Figure could not be embedded: ' + e.message + ' ]', italics: true, color: 'AA3333', size: 20})],
          }));
        }
        children.push(new Paragraph({
          children: [new TextRun({text: f.caption || '', size: 18, italics: true, font: 'Arial'})],
          alignment: AlignmentType.CENTER,
          spacing: {after: 200},
        }));
      }
    }

    const doc = new Document({
      styles: {default: {document: {run: {font: 'Arial', size: 22}}}},
      sections: [{
        properties: {
          page: {
            size: {width: 11906, height: 16838}, // A4, matching the PDF export's page format
            margin: {top: 1080, right: 1080, bottom: 1080, left: 1080},
          },
        },
        children,
      }],
    });

    const rawBlob = await Packer.toBlob(doc);
    const fixedBlob = await fixMissingFontTableRelationship(rawBlob);
    const url = URL.createObjectURL(fixedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Invention_Disclosure_Form.docx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }catch(err){
    alert('Word export failed: ' + err.message + '. Try the PDF export instead, or check the browser console for details.');
    console.error('exportWord error:', err);
  }finally{
    if(exportBtn){ exportBtn.textContent = originalLabel; exportBtn.disabled = false; }
  }
}

renderPreview();
renderPriorArtTool();
restoreDraft();
