  // Section toggle
  document.querySelectorAll('.section-header').forEach(header => {
    header.addEventListener('click', () => {
      header.classList.toggle('collapsed');
      const row = header.nextElementSibling;
      if (row && row.classList.contains('cards-row')) {
        row.classList.toggle('hidden');
      }
    });
  });


  // Tooltip system — document-level for reliability
  const ftip = document.getElementById('floating-tooltip');

  function showTip(icon) {
    const tplId = icon.getAttribute('data-tooltip-id');
    if (!tplId) return;
    const tpl = document.getElementById(tplId);
    if (!tpl) return;
    ftip.innerHTML = '';
    ftip.appendChild(tpl.content.cloneNode(true));
    ftip.classList.add('show');
    const r = icon.getBoundingClientRect();
    let left = r.right - 230;
    let top  = r.bottom + 8;
    if (left < 8) left = 8;
    if (top + 300 > window.innerHeight) top = r.top - 300;
    ftip.style.left = left + 'px';
    ftip.style.top  = top  + 'px';
  }

  document.addEventListener('mouseover', (e) => {
    const icon = e.target.closest('.info-icon[data-tooltip-id]');
    if (icon) {
      showTip(icon);
    } else if (!e.target.closest('#floating-tooltip')) {
      ftip.classList.remove('show');
    }
  });

  // Card expand/collapse with FLIP animation
  (function(){
    const backdrop = document.createElement('div');
    backdrop.className = 'card-backdrop';
    document.body.appendChild(backdrop);

    let expandedCard = null;
    let placeholder = null;

    function collapseCard(){
      if (!expandedCard) return;
      const card = expandedCard;
      // If team-view is open, restore original innerHTML first so collapse animates correctly
      if (typeof teamState !== 'undefined' && teamState && teamState.card === card) {
        card.classList.remove('team-view');
        card.innerHTML = teamState.originalHTML;
        document.removeEventListener('click', closeTeamDD);
        teamState = null;
      }
      const actions = card.querySelector('.card-actions');
      if (actions) actions.remove();

      const first = card.getBoundingClientRect();
      card.classList.remove('expanded');
      card.style.cssText = '';
      if (placeholder) { placeholder.replaceWith(card); placeholder = null; }
      const last = card.getBoundingClientRect();

      const dx = first.left - last.left;
      const dy = first.top  - last.top;
      const sx = first.width  / last.width;
      const sy = first.height / last.height;

      card.style.transition = 'none';
      card.style.transformOrigin = 'top left';
      card.style.transform = `translate(${dx}px, ${dy}px) scale(${sx}, ${sy})`;
      requestAnimationFrame(() => {
        card.style.transition = 'transform 0.3s cubic-bezier(0.4,0,0.2,1)';
        card.style.transform = '';
      });
      setTimeout(() => { card.style.cssText = ''; }, 320);

      backdrop.classList.remove('show');
      expandedCard = null;
    }

    function expandCard(card){
      if (expandedCard) collapseCard();

      const first = card.getBoundingClientRect();
      placeholder = document.createElement('div');
      placeholder.style.width  = first.width  + 'px';
      placeholder.style.height = first.height + 'px';
      card.after(placeholder);

      const targetW = Math.min(first.width * 2, 470, window.innerWidth - 60);
      card.classList.add('expanded');
      card.style.width = targetW + 'px';
      card.style.left  = ((window.innerWidth - targetW) / 2) + 'px';
      card.style.top   = '50%';
      card.style.transform = 'translateY(-50%)';

      const actions = document.createElement('div');
      actions.className = 'card-actions';
      actions.innerHTML =
        '<button type="button" class="btn-action">' +
          '<span class="list-icon"><span></span><span></span><span></span></span>' +
          'Equipe' +
        '</button>' +
        '<button type="button" class="btn-close" aria-label="Fechar">&times;</button>';
      card.appendChild(actions);
      actions.querySelector('.btn-close').addEventListener('click', (e)=>{ e.stopPropagation(); collapseCard(); });
      actions.querySelector('.btn-action').addEventListener('click', (e)=>{
        e.stopPropagation();
        if (card.id === 'card-fat-total') openTeamView(card, actions);
      });

      const last = card.getBoundingClientRect();
      const dx = first.left - last.left;
      const dy = first.top  - last.top;
      const sx = first.width  / last.width;
      const sy = first.height / last.height;

      card.style.transition = 'none';
      card.style.transformOrigin = 'top left';
      card.style.transform = `translateY(-50%) translate(${dx}px, ${dy}px) scale(${sx}, ${sy})`;
      requestAnimationFrame(() => {
        card.style.transition = 'transform 0.32s cubic-bezier(0.4,0,0.2,1), box-shadow 0.32s ease';
        card.style.transform = 'translateY(-50%)';
      });

      backdrop.classList.add('show');
      expandedCard = card;
    }

    document.querySelectorAll('.kpi-card').forEach(card => {
      card.style.cursor = 'pointer';
      card.addEventListener('click', (e) => {
        if (e.target.closest('.info-icon')) return;
        if (e.target.closest('.card-actions')) return;
        if (card.classList.contains('expanded')) return;
        expandCard(card);
      });
    });
    backdrop.addEventListener('click', () => { collapseCard(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') collapseCard(); });

    // ── Team-view feature (somente card Total de Faturamento) ──
    const FIELDS = [
      { key: 'cv', label: 'Gerente' },
      { key: 'sv', label: 'Supervisor' },
      { key: 'rv', label: 'Vendedor' },
    ];
    let teamState = null; // { card, originalHTML, originalActions, field }

    function buildTeamRows(field){
      const api = window.dashboardAPI;
      if (!api) return [];
      // name -> Set of ck (rv|uf) keys for that person
      const groupMap = new Map();
      api.rvEntries.forEach(r => {
        if (!api.passes(r)) return;
        const name = field === 'rv' ? r.rvName : r[field];
        if (!name) return;
        if (!groupMap.has(name)) groupMap.set(name, new Set());
        groupMap.get(name).add(api.ck(r.rv, r.uf));
      });
      const realMap = new Map();
      api.vendas.forEach(x => {
        const k = api.ck(x.rv, x.uf);
        groupMap.forEach((keys, name) => {
          if (keys.has(k)) realMap.set(name, (realMap.get(name) || 0) + (x.v || 0));
        });
      });
      const metaMap = new Map();
      api.metas.forEach(m => {
        const k = api.ck(m.rv, m.uf);
        groupMap.forEach((keys, name) => {
          if (keys.has(k)) metaMap.set(name, (metaMap.get(name) || 0) + (m.total || 0));
        });
      });
      const rows = [];
      groupMap.forEach((_, name) => {
        const real = realMap.get(name) || 0;
        const meta = metaMap.get(name) || 0;
        const gap = meta - real;
        const pct = meta > 0 ? real / meta : 0;
        rows.push({ name, real, meta, gap, pct });
      });
      return rows.sort((a, b) => b.pct - a.pct);
    }

    function renderTeamTable(wrap, field){
      const rows = buildTeamRows(field);
      const api = window.dashboardAPI;
      const barColor = wrap.dataset.barColor || '#003DA5';
      if (!rows.length) {
        wrap.innerHTML = '<div class="team-bars-empty">Sem dados</div>';
        return;
      }
      const fmt = api ? api.fmtMoney : (v)=>v;
      wrap.innerHTML = '<div class="team-bars">' + rows.map(r => {
        const pctTxt = (r.pct * 100).toFixed(1).replace('.', ',') + '%';
        const w = Math.min(100, Math.max(0, r.pct * 100)).toFixed(1);
        return (
          '<div class="team-bar-item">' +
            '<div class="team-bar-head">' +
              '<div class="team-bar-name">' + r.name + '</div>' +
              '<div class="team-bar-pct">' + pctTxt + '</div>' +
            '</div>' +
            '<div class="team-bar-meta">' +
              '<span>Meta: <b>' + fmt(r.meta) + '</b></span>' +
              '<span>Gap: <b class="gap-val">' + fmt(Math.abs(r.gap)) + '</b></span>' +
              '<span>Realizado: <b>' + fmt(r.real) + '</b></span>' +
            '</div>' +
            '<div class="team-bar-track"><div class="team-bar-fill" style="width:' + w + '%;background:' + barColor + ';"></div></div>' +
          '</div>'
        );
      }).join('') + '</div>';
    }


    function openTeamView(card, actions){
      if (!window.dashboardAPI) return;
      const startH = card.getBoundingClientRect().height;
      const originalHTML = card.innerHTML;
      // capture bar color from existing progress-fill so bars match the card theme
      let barColor = '#003DA5';
      const pf = card.querySelector('.progress-fill');
      if (pf) {
        const c = getComputedStyle(pf).backgroundColor;
        if (c && c !== 'rgba(0, 0, 0, 0)') barColor = c;
      }
      teamState = { card, originalHTML, field: 'cv' };

      // Build team view
      card.innerHTML = '';
      const newActions = document.createElement('div');
      newActions.className = 'card-actions';
      newActions.innerHTML =
        '<div class="team-dropdown" id="team-dd">' +
          '<span class="td-label">Gerente</span><span class="td-arrow">▾</span>' +
          '<div class="team-dropdown-pop">' +
            FIELDS.map(f => `<div class="td-item" data-field="${f.key}">${f.label}</div>`).join('') +
          '</div>' +
        '</div>' +
        '<button type="button" class="btn-close" id="team-back" aria-label="Voltar">&times;</button>';
      card.appendChild(newActions);

      const wrap = document.createElement('div');
      wrap.className = 'team-view-wrap';
      wrap.innerHTML =
        '<div class="team-view-header"><div class="team-view-title">Faturamento Total por Equipe</div></div>' +
        '<div class="team-table-wrap" id="team-table-wrap"></div>';
      card.appendChild(wrap);
      card.classList.add('team-view');

      const tableWrap = wrap.querySelector('#team-table-wrap');
      tableWrap.dataset.barColor = barColor;
      renderTeamTable(tableWrap, teamState.field);

      // Dropdown logic
      const dd = newActions.querySelector('#team-dd');
      const ddLabel = dd.querySelector('.td-label');
      dd.addEventListener('click', (e) => {
        e.stopPropagation();
        dd.classList.toggle('open');
        // mark selected
        dd.querySelectorAll('.td-item').forEach(it => {
          it.classList.toggle('selected', it.dataset.field === teamState.field);
        });
      });
      dd.querySelectorAll('.td-item').forEach(it => {
        it.addEventListener('click', (e) => {
          e.stopPropagation();
          teamState.field = it.dataset.field;
          ddLabel.textContent = FIELDS.find(f => f.key === teamState.field).label;
          dd.classList.remove('open');
          renderTeamTable(tableWrap, teamState.field);
        });
      });
      document.addEventListener('click', closeTeamDD);

      // Back button restores original card
      newActions.querySelector('#team-back').addEventListener('click', (e) => {
        e.stopPropagation();
        closeTeamView();
      });

      // Smooth height expand
      card.style.height = startH + 'px';
      card.style.overflow = 'hidden';
      // measure target height
      card.style.transition = 'none';
      const prevH = card.style.height;
      card.style.height = 'auto';
      const targetH = card.getBoundingClientRect().height;
      card.style.height = prevH;
      // animate
      requestAnimationFrame(() => {
        card.style.transition = 'height 0.32s cubic-bezier(0.4,0,0.2,1), transform 0.32s cubic-bezier(0.4,0,0.2,1)';
        card.style.height = targetH + 'px';
        // keep vertically centered
        card.style.top = '50%';
        card.style.transform = 'translateY(-50%)';
      });
      setTimeout(() => {
        if (teamState && teamState.card === card) {
          card.style.height = '';
          card.style.overflow = '';
        }
      }, 360);
    }

    function closeTeamDD(e){
      if (!teamState) return;
      const dd = teamState.card.querySelector('#team-dd');
      if (dd && !dd.contains(e.target)) dd.classList.remove('open');
    }

    function closeTeamView(){
      if (!teamState) return;
      const { card, originalHTML } = teamState;
      const startH = card.getBoundingClientRect().height;
      document.removeEventListener('click', closeTeamDD);
      card.classList.remove('team-view');
      card.innerHTML = originalHTML;

      // rebind actions on restored content
      const actions = card.querySelector('.card-actions');
      if (actions) {
        actions.querySelector('.btn-close').addEventListener('click', (e)=>{ e.stopPropagation(); collapseCard(); });
        const btn = actions.querySelector('.btn-action');
        if (btn) btn.addEventListener('click', (e)=>{
          e.stopPropagation();
          if (card.id === 'card-fat-total') openTeamView(card, actions);
        });
      }

      // measure natural height and animate back
      card.style.transition = 'none';
      card.style.height = 'auto';
      card.style.overflow = 'hidden';
      const targetH = card.getBoundingClientRect().height;
      card.style.height = startH + 'px';
      requestAnimationFrame(() => {
        card.style.transition = 'height 0.3s cubic-bezier(0.4,0,0.2,1)';
        card.style.height = targetH + 'px';
      });
      setTimeout(() => {
        card.style.height = '';
        card.style.overflow = '';
        card.style.transition = '';
      }, 340);

      teamState = null;
    }

  })();

/* ====== Data loading & filtering (KPI Ranking P&G) ====== */
(async function(){
  let data;
  try {
    const res = await fetch('./data.json', {cache:'no-store'});
    data = await res.json();
  } catch(e) { console.error('Falha ao carregar data.json', e); return; }

  // Atualiza data/hora no header
  if (data.generated_at) {
    try {
      const d = new Date(data.generated_at);
      const meses = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
      const txt = `Atualizado em ${String(d.getDate()).padStart(2,'0')} ${meses[d.getMonth()]} ${d.getFullYear()} às ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
      const el = document.getElementById('last-update');
      if (el) el.textContent = txt;
    } catch(e) {}
  }

  // comercial: [{rv,uf,rvName,sv,cv}] - composite key rv+uf
  const rvEntries = data.comercial.map(r=>({uf:r.uf, cv:r.cv, sv:r.sv, rvName:r.rvName, rv:String(r.rv)}));
  const metas = data.metas;     // [{rv,uf,total,ec,sp,ali,far}]
  const vendas = data.vendas;   // [{rv,uf,v,ec,sp,ali,far}]
  const ck = (rv,uf) => String(rv)+'|'+String(uf);

  // ── Filtros: dropdown customizado (clique = único; Ctrl/Cmd = múltiplo) ──
  const uniq = (arr)=>[...new Set(arr.filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'pt-BR'));
  const filterValues = {
    uf: uniq(rvEntries.map(r=>r.uf)),
    cv: uniq(rvEntries.map(r=>r.cv)),
    sv: uniq(rvEntries.map(r=>r.sv)),
    rv: uniq(rvEntries.map(r=>r.rvName)),
  };
  // Cada filtro guarda um Set de valores selecionados (vazio = todos)
  const filters = { uf:new Set(), cv:new Set(), sv:new Set(), rv:new Set() };

  function updateChipLabel(chip){
    const key = chip.dataset.filter;
    const def = chip.dataset.defaultLabel;
    const sel = filters[key];
    const lbl = chip.querySelector('.chip-label');
    if (sel.size === 0) lbl.textContent = def;
    else if (sel.size === 1) lbl.textContent = [...sel][0];
    else lbl.textContent = `${sel.size} selecionados`;
  }

  const FIELD_ACCESSOR = { uf:r=>r.uf, cv:r=>r.cv, sv:r=>r.sv, rv:r=>r.rvName };
  function availableValues(key){
    const acc = FIELD_ACCESSOR[key];
    const set = new Set();
    rvEntries.forEach(r=>{
      for (const k of ['uf','cv','sv','rv']) {
        if (k === key) continue;
        if (filters[k].size && !filters[k].has(FIELD_ACCESSOR[k](r))) return;
      }
      const v = acc(r);
      if (v) set.add(v);
    });
    return [...set].sort((a,b)=>String(a).localeCompare(String(b),'pt-BR'));
  }

  function renderPop(chip, search=''){
    const key = chip.dataset.filter;
    const pop = chip.querySelector('.filter-pop');
    const list = pop.querySelector('.fp-list');
    const q = search.trim().toLowerCase();
    const all = availableValues(key);
    filters[key].forEach(v=>{ if (!all.includes(v)) all.unshift(v); });
    const filtered = q ? all.filter(v=>String(v).toLowerCase().includes(q)) : all;

    list.innerHTML = '';
    // "Todos"
    const allItem = document.createElement('div');
    allItem.className = 'fp-item all' + (filters[key].size===0 ? ' selected' : '');
    allItem.innerHTML = '<span class="fp-radio"></span><span>Todos</span>';
    allItem.addEventListener('click', ()=>{
      filters[key].clear();
      updateChipLabel(chip);
      renderPop(chip, pop.querySelector('.fp-search').value);
      recompute();
    });
    list.appendChild(allItem);
    if (filtered.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'fp-empty';
      empty.textContent = 'Nenhum resultado';
      list.appendChild(empty);
      return;
    }
    filtered.forEach(v=>{
      const it = document.createElement('div');
      const isSel = filters[key].has(v);
      it.className = 'fp-item' + (isSel ? ' selected' : '');
      it.innerHTML = `<span class="fp-radio"></span><span>${v}</span>`;
      it.addEventListener('click', (e)=>{
        const multi = e.ctrlKey || e.metaKey;
        if (multi) {
          if (filters[key].has(v)) filters[key].delete(v);
          else filters[key].add(v);
        } else {
          // Seleção única: toggle (desseleciona se já era o único selecionado)
          if (filters[key].has(v) && filters[key].size === 1) {
            filters[key].clear();
          } else {
            filters[key].clear();
            filters[key].add(v);
          }
        }
        updateChipLabel(chip);
        renderPop(chip, pop.querySelector('.fp-search').value);
        recompute();
      });
      list.appendChild(it);
    });
  }

  function closeAllPops(except){
    document.querySelectorAll('.filter-chip.open').forEach(c=>{
      if (c !== except) c.classList.remove('open');
    });
  }

  document.querySelectorAll('.filter-chip[data-filter]').forEach(chip=>{
    // monta o popover uma única vez
    const pop = document.createElement('div');
    pop.className = 'filter-pop';
    pop.innerHTML = '<input type="text" class="fp-search" placeholder="Buscar..." /><div class="fp-list"></div>';
    chip.appendChild(pop);
    pop.addEventListener('click', e=>e.stopPropagation());
    const search = pop.querySelector('.fp-search');
    search.addEventListener('input', ()=>renderPop(chip, search.value));

    chip.addEventListener('click', (e)=>{
      if (chip.classList.contains('open')) return;
      closeAllPops(chip);
      chip.classList.add('open');
      const rect = chip.getBoundingClientRect();
      pop.style.left = rect.left + 'px';
      pop.style.top = (rect.bottom + 6) + 'px';
      search.value = '';
      renderPop(chip, '');
      setTimeout(()=>search.focus(), 0);
    });
    updateChipLabel(chip);
  });

  function lockChipWidths(){
    document.querySelectorAll('.filter-chip[data-filter]').forEach(chip=>{
      chip.style.width = '';
      chip.style.width = Math.ceil(chip.getBoundingClientRect().width) + 'px';
    });
  }
  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(lockChipWidths);
  } else {
    window.addEventListener('load', lockChipWidths);
  }

  document.getElementById('clear-filters').addEventListener('click', ()=>{
    Object.values(filters).forEach(s => s.clear());
    document.querySelectorAll('.filter-chip[data-filter]').forEach(chip=>updateChipLabel(chip));
    recompute();
  });

  document.addEventListener('click', (e)=>{
    if (!e.target.closest('.filter-chip')) closeAllPops(null);
  });

  

  function passes(r){
    if (filters.uf.size && !filters.uf.has(r.uf)) return false;
    if (filters.cv.size && !filters.cv.has(r.cv)) return false;
    if (filters.sv.size && !filters.sv.has(r.sv)) return false;
    if (filters.rv.size && !filters.rv.has(r.rvName)) return false;
    return true;
  }

  function fmtMoney(v){
    const abs = Math.abs(v);
    let s;
    if (abs >= 1e9) s = (v/1e9).toFixed(2).replace('.',',')+'B';
    else if (abs >= 1e6) s = (v/1e6).toFixed(2).replace('.',',')+'M';
    else if (abs >= 1e3) s = (v/1e3).toFixed(1).replace('.',',')+'K';
    else s = v.toFixed(0);
    return 'R$ '+s;
  }
  function fmtPct(v){ return (v*100).toFixed(1).replace('.',',')+'%'; }

  function recompute(){
    const eligibleKeys = new Set();
    rvEntries.forEach(r=>{
      if (!passes(r)) return;
      eligibleKeys.add(ck(r.rv, r.uf));
    });


    const sums = {v:0, ec:0, sp:0, ali:0, far:0};
    vendas.forEach(x=>{
      if (!eligibleKeys.has(ck(x.rv, x.uf))) return;
      sums.v += x.v||0; sums.ec += x.ec||0; sums.sp += x.sp||0;
      sums.ali += x.ali||0; sums.far += x.far||0;
    });

    const metaSums = {total:0, ec:0, sp:0, ali:0, far:0};
    metas.forEach(m=>{
      if (!eligibleKeys.has(ck(m.rv, m.uf))) return;
      metaSums.total += m.total||0;
      metaSums.ec    += m.ec||0;
      metaSums.sp    += m.sp||0;
      metaSums.ali   += m.ali||0;
      metaSums.far   += m.far||0;
    });


    function paint(cardId, total, metaVal){
      const card = document.getElementById(cardId);
      if (!card) return;
      const gap = metaVal - total;
      const pct = metaVal > 0 ? total/metaVal : 0;
      card.querySelector('[data-bind=valor]').textContent = fmtMoney(total);
      card.querySelector('[data-bind=meta]').textContent  = fmtMoney(metaVal);
      const gapEl = card.querySelector('[data-bind=gap]');
      gapEl.textContent = fmtMoney(Math.abs(gap));
      gapEl.style.color = gap > 0 ? '#DC2626' : '#059669';
      card.querySelector('[data-bind=pct]').textContent   = fmtPct(pct);
      card.querySelector('[data-bind=bar]').style.width   = Math.min(100, Math.max(0, pct*100)).toFixed(1)+'%';
    }
    paint('card-fat-total', sums.v,   metaSums.total);
    paint('card-fat-ec',    sums.ec,  metaSums.ec);
    paint('card-fat-sp',    sums.sp,  metaSums.sp);
    paint('card-fat-ali',   sums.ali, metaSums.ali);
    paint('card-fat-far',   sums.far, metaSums.far);
  }

  window.dashboardAPI = { rvEntries, vendas, metas, passes, ck, fmtMoney };
  recompute();
})();

