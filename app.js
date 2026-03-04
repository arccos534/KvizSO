const teamsBody = document.getElementById('teamsBody');
const rowTemplate = document.getElementById('rowTemplate');
const statusEl = document.getElementById('status');

function setStatus(message) {
  statusEl.textContent = message;
}

function clampScore(value) {
  const num = Number(value);
  if (!Number.isFinite(num) || num < 0) return 0;
  return Math.floor(num);
}

function recalcRow(row) {
  const r1 = clampScore(row.querySelector('.round1').value);
  const r2 = clampScore(row.querySelector('.round2').value);
  row.querySelector('.sum-cell').textContent = String(r1 + r2);
}

function renumberRows() {
  [...teamsBody.children].forEach((row, index) => {
    row.querySelector('.row-number').textContent = String(index + 1);
  });
}

function attachRowHandlers(row) {
  row.querySelectorAll('.score-input').forEach((input) => {
    input.addEventListener('input', () => recalcRow(row));
    input.addEventListener('blur', () => {
      const cleaned = clampScore(input.value);
      input.value = cleaned > 0 ? String(cleaned) : '';
      recalcRow(row);
    });
  });
}

function addRow(name = '', r1 = 0, r2 = 0) {
  const row = rowTemplate.content.firstElementChild.cloneNode(true);
  row.querySelector('.team-input').value = name;
  row.querySelector('.round1').value = r1 > 0 ? String(clampScore(r1)) : '';
  row.querySelector('.round2').value = r2 > 0 ? String(clampScore(r2)) : '';
  attachRowHandlers(row);
  teamsBody.appendChild(row);
  recalcRow(row);
  renumberRows();
}

function removeRow() {
  if (teamsBody.lastElementChild) {
    teamsBody.removeChild(teamsBody.lastElementChild);
  }
  renumberRows();
}

function clearRows() {
  teamsBody.innerHTML = '';
  addRow('', 0, 0);
}

function collectRows(sorted = true) {
  const rows = [...teamsBody.children].map((row, i) => {
    const name = row.querySelector('.team-input').value.trim() || `Команда ${i + 1}`;
    const round1 = clampScore(row.querySelector('.round1').value);
    const round2 = clampScore(row.querySelector('.round2').value);
    const total = round1 + round2;
    return {
      place: i + 1,
      team_name: name,
      round1,
      round2,
      total,
    };
  });

  if (sorted) {
    rows.sort((a, b) => b.total - a.total || a.place - b.place);
    rows.forEach((r, i) => {
      r.place = i + 1;
    });
  }

  return rows;
}

async function exportPresentation() {
  const rows = collectRows(true);
  if (!rows.length) {
    setStatus('Добавьте хотя бы одну команду.');
    return;
  }

  if (typeof PptxGenJS === 'undefined') {
    setStatus('Не удалось загрузить библиотеку презентаций.');
    return;
  }

  setStatus('Генерирую презентацию...');

  try {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'Quiz Scoreboard';
    pptx.subject = 'Результаты квиза';
    pptx.title = 'Квиз - таблица результатов';

    const slide = pptx.addSlide();

    const headerRow = [
      { text: 'Место', options: { bold: false, align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
      { text: 'Название команды', options: { bold: false, align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
      { text: 'Раунд 1', options: { bold: false, align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
      { text: 'Раунд 2', options: { bold: false, align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
      { text: 'Результат', options: { bold: false, align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
    ];

    const bodyRows = rows.map((r) => ([
      { text: String(r.place), options: { align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
      { text: String(r.team_name), options: { align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
      { text: String(r.round1), options: { align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
      { text: String(r.round2), options: { align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
      { text: String(r.total), options: { align: 'center', valign: 'middle', fill: { color: 'E59C85' } } },
    ]));

    const tableRows = [headerRow, ...bodyRows];
    const rowHeights = new Array(tableRows.length).fill(0.42);

    slide.addTable(tableRows, {
      x: 0,
      y: 0,
      w: 13.333,
      fontFace: 'Calibri',
      fontSize: 17,
      color: '000000',
      border: { pt: 1, color: '000000' },
      align: 'center',
      valign: 'middle',
      colW: [1.2, 4.8, 2.0, 2.0, 3.333],
      rowH: rowHeights,
      margin: 0.03,
      autoPage: false,
    });

    const now = new Date();
    const stamp = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}`;
    await pptx.writeFile({ fileName: `quiz_results_${stamp}.pptx` });
    setStatus('Презентация сохранена.');
  } catch (err) {
    console.error(err);
    setStatus('Не удалось создать презентацию.');
  }
}

document.getElementById('addRowBtn').addEventListener('click', () => addRow());
document.getElementById('removeRowBtn').addEventListener('click', removeRow);
document.getElementById('clearBtn').addEventListener('click', clearRows);
document.getElementById('exportBtn').addEventListener('click', exportPresentation);

addRow('', 0, 0);
