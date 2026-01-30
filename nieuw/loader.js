// loader.js
(function () {
  const EXCEL_URL = '/nieuw/data/links.xlsx';
  const SHEET_NAME = 'Links';

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    try {
      const rows = await loadExcel(EXCEL_URL, SHEET_NAME);
      buildPageFromData(rows);
      setupSearchAndWideCard();
    } catch (err) {
      console.error('Fout bij laden Excel:', err);
    }
  }

  async function loadExcel(url, sheetName) {
    const res = await fetch(url);
    if (!res.ok) throw new Error('Kan Excel niet laden: ' + res.status);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheet =
      wb.Sheets[sheetName] || wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    return rows;
  }

  function buildPageFromData(rows) {
    const blocks = document.getElementById('blocks');
    if (!blocks) return;

    const cardsMap = new Map();

    rows.forEach((row, idx) => {
      const section = (row.section || '').toString().trim().toLowerCase() || 'grid';
      const title = (row.card_title || '').toString().trim();
      const color = (row.card_color || '').toString().trim();
      if (!title) return;

      const key = section + '||' + title + '||' + color;
      if (!cardsMap.has(key)) {
        cardsMap.set(key, {
          section,
          title,
          color,
          order: Number(row.card_order) || 999,
          items: []
        });
      }
      const card = cardsMap.get(key);

      card.items.push({
        wide_col: row.wide_col ? Number(row.wide_col) : null,
        item_group: row.item_group
          ? row.item_group.toString()
          : String(idx + 1),
        item_kind: (row.item_kind || 'link').toString().toLowerCase(),
        text: (row.text || '').toString(),
        url: (row.url || '').toString(),
        bold: String(row.bold || '').toLowerCase() === 'true',
        underline: String(row.underline || '').toLowerCase() === 'true'
      });
    });

    const cards = Array.from(cardsMap.values());
    const bySection = (sec) =>
      cards
        .filter((c) => c.section === sec)
        .sort((a, b) => a.order - b.order);

    // Bovenste grid (4 kolommen)
    const gridCards = bySection('grid');
    if (gridCards.length) {
      const gridDiv = document.createElement('div');
      gridDiv.className = 'grid';
      gridCards.forEach((c) => gridDiv.appendChild(createNormalCard(c)));
      blocks.appendChild(gridDiv);
    }

    // Rij met 3 kaarten (Warenwetten / EU regelgeving / EU Guides)
    const rowCards = bySection('row');
    if (rowCards.length) {
      const rowDiv = document.createElement('div');
      rowDiv.className = 'row flow';
      rowCards.forEach((c) => rowDiv.appendChild(createNormalCard(c)));
      blocks.appendChild(rowDiv);
    }

    // Brede uitklapkaart(en)
    const wideCards = bySection('wide');
    wideCards.forEach((c) => {
      const wrap = document.createElement('div');
      wrap.className = 'card_wide';
      wrap.appendChild(createWideCard(c));
      blocks.appendChild(wrap);
    });
  }

  function groupItems(items) {
    const groups = {};
    items.forEach((row, idx) => {
      const key = row.item_group || String(idx + 1);
      if (!groups[key]) groups[key] = [];
      groups[key].push(row);
    });
    return Object.values(groups);
  }

  function createNormalCard(card) {
    const section = document.createElement('section');
    section.className = 'card' + (card.color ? ' ' + card.color : '');

    const h2 = document.createElement('h2');
    h2.textContent = card.title;
    section.appendChild(h2);

    const ul = document.createElement('ul');
    const groups = groupItems(card.items);

    groups.forEach((groupRows) => {
      const li = document.createElement('li');
      const span = document.createElement('span');
      span.className = 'ph';
      li.appendChild(span);

      groupRows.forEach((row, index) => {
        if (row.item_kind === 'heading' || (!row.url && row.text)) {
          let target = li;
          if (row.bold) {
            const b = document.createElement('b');
            target.appendChild(b);
            target = b;
          }
          if (row.underline) {
            const u = document.createElement('u');
            target.appendChild(u);
            target = u;
          }
          target.appendChild(document.createTextNode(row.text));
        } else if (row.url) {
          if (index > 0) {
            li.appendChild(document.createTextNode(' | '));
          }
          const a = document.createElement('a');
          a.href = row.url;
          a.textContent = row.text;
          a.target = '_blank';
          a.rel = 'noopener noreferrer';
          li.appendChild(a);
        }
      });

      ul.appendChild(li);
    });

    section.appendChild(ul);
    return section;
  }

  function createWideCard(card) {
    const div = document.createElement('div');
    div.className = 'card wide' + (card.color ? ' ' + card.color : '');

    const h2 = document.createElement('h2');
    h2.textContent = card.title;
    div.appendChild(h2);

    const colsDiv = document.createElement('div');
    colsDiv.className = 'cols';

    const groupedByCol = {};
    card.items.forEach((row) => {
      const col = row.wide_col || 1;
      if (!groupedByCol[col]) groupedByCol[col] = [];
      groupedByCol[col].push(row);
    });

    const colNums = Object.keys(groupedByCol)
      .map(Number)
      .sort((a, b) => a - b);

    colNums.forEach((colNum) => {
      const ul = document.createElement('ul');
      const groups = groupItems(groupedByCol[colNum]);

      groups.forEach((groupRows) => {
        const li = document.createElement('li');

        groupRows.forEach((row, index) => {
          if (row.item_kind === 'heading' || (!row.url && row.text)) {
            let target = li;
            if (row.bold) {
              const b = document.createElement('b');
              target.appendChild(b);
              target = b;
            }
            if (row.underline) {
              const u = document.createElement('u');
              target.appendChild(u);
              target = u;
            }
            target.appendChild(document.createTextNode(row.text));
          } else if (row.url) {
            if (index > 0) {
              li.appendChild(document.createTextNode(' | '));
            }
            const a = document.createElement('a');
            a.href = row.url;
            a.textContent = row.text;
            a.target = '_blank';
            a.rel = 'noopener noreferrer';
            li.appendChild(a);
          }
        });

        ul.appendChild(li);
      });

      colsDiv.appendChild(ul);
    });

    div.appendChild(colsDiv);
    return div;
  }

  function setupSearchAndWideCard() {
    const blocks = document.getElementById('blocks');
    const input = document.getElementById('searchInput');
    const clearBtn = document.getElementById('clearBtn');
    const form = document.getElementById('searchForm');
    const wideCard = document.querySelector('.card.wide');

    function filterLinks(q) {
      const term = q.trim().toLowerCase();
      const cards = blocks.querySelectorAll('.card');
      let anyVisibleInPage = false;

      cards.forEach((card) => {
        const items = card.querySelectorAll('li');
        let visibleInCard = 0;

        items.forEach((li) => {
          const text = li.textContent.toLowerCase();
          const hrefs = Array.from(li.querySelectorAll('a'))
            .map((a) => a.href.toLowerCase())
            .join(' ');
          const match =
            !term || text.includes(term) || hrefs.includes(term);
          li.style.display = match ? '' : 'none';
          if (match) visibleInCard++;
        });

        card.style.display = visibleInCard > 0 ? '' : 'none';
        if (visibleInCard > 0) anyVisibleInPage = true;
      });

      blocks.dataset.hasresults = anyVisibleInPage ? '1' : '0';
    }

    if (input && clearBtn && form) {
      input.addEventListener('input', (e) => filterLinks(e.target.value));
      clearBtn.addEventListener('click', () => {
        input.value = '';
        filterLinks('');
        input.focus();
      });

      form.addEventListener('submit', (e) => {
        e.preventDefault();
        const term = input.value.trim();
        if (!term) return;
        const url =
          'https://www.google.com/search?q=' + encodeURIComponent(term);
        window.open(url, '_blank', 'noopener');
      });

      // initiale state
      filterLinks('');
    }

    // Hover open/dicht voor brede kaart
    if (wideCard) {
      let hoverTimeout;

      wideCard.addEventListener('mouseenter', () => {
        clearTimeout(hoverTimeout);
        wideCard.classList.add('open');
      });

      wideCard.addEventListener('mouseleave', () => {
        hoverTimeout = setTimeout(() => {
          wideCard.classList.remove('open');
        }, 200);
      });

      // Klik op titel togglet ook (lekker UX)
      const header = wideCard.querySelector('h2');
      if (header) {
        header.addEventListener('click', () => {
          wideCard.classList.toggle('open');
        });
      }
    }
  }
})();
