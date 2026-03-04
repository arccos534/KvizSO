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
  row.querySelector('.round1').value = r1;
  row.querySelector('.round2').value = r2;
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
  });
}

function addRow(name = '', r1 = 0, r2 = 0) {
  const row = rowTemplate.content.firstElementChild.cloneNode(true);
  row.querySelector('.team-input').value = name;
  row.querySelector('.round1').value = clampScore(r1);
  row.querySelector('.round2').value = clampScore(r2);
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
  setStatus('Генерирую презентацию...');

  try {
    const response = await fetch('/export-pptx', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ teams: rows }),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error || 'Ошибка генерации файла');
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const now = new Date();
    const stamp = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}`;
    a.href = url;
    a.download = `quiz_results_${stamp}.pptx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setStatus('Презентация сохранена.');
  } catch (err) {
    setStatus(err.message || 'Не удалось создать презентацию.');
  }
}

document.getElementById('addRowBtn').addEventListener('click', () => addRow());
document.getElementById('removeRowBtn').addEventListener('click', removeRow);
document.getElementById('clearBtn').addEventListener('click', clearRows);
document.getElementById('exportBtn').addEventListener('click', exportPresentation);

addRow('', 0, 0);
