(function () {
  const SUBJECTS = [
    { id: 'turkce', name: 'Türkçe' },
    { id: 'matematik', name: 'Matematik' },
    { id: 'sosyal', name: 'Sosyal Bilgiler' },
    { id: 'ingilizce', name: 'İngilizce' },
    { id: 'din', name: 'Din Kültürü ve Ahlak Bilgisi' }
  ];

  const STORAGE_KEY = 'lgs_tracker_entries_v1';

  const el = {
    subject: document.getElementById('subject'),
    date: document.getElementById('date'),
    solved: document.getElementById('solved'),
    correct: document.getElementById('correct'),
    wrong: document.getElementById('wrong'),
    blank: document.getElementById('blank'),
    form: document.getElementById('entry-form'),
    formHint: document.getElementById('form-hint'),
    totalsGrid: document.getElementById('totals-grid'),
    overallTotals: document.getElementById('overall-totals'),
    entriesTbody: document.getElementById('entries-tbody'),
    exportBtn: document.getElementById('export-btn'),
    resetBtn: document.getElementById('reset-btn')
  };

  function readEntries() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed;
    } catch (e) {
      console.error('readEntries error', e);
      return [];
    }
  }

  function writeEntries(entries) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
  }

  function addEntry(entry) {
    const entries = readEntries();
    entries.push(entry);
    writeEntries(entries);
  }

  function deleteEntry(id) {
    const entries = readEntries().filter(e => e.id !== id);
    writeEntries(entries);
  }

  function generateId() {
    return 'e_' + Math.random().toString(36).slice(2, 10) + '_' + Date.now().toString(36);
  }

  function formatDate(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    return d.toLocaleDateString('tr-TR', { year: 'numeric', month: '2-digit', day: '2-digit' });
  }

  function nonNegativeInt(value) {
    const num = Number(value);
    if (!Number.isFinite(num) || num < 0) return null;
    return Math.floor(num);
  }

  function populateSubjects() {
    el.subject.innerHTML = SUBJECTS.map(s => `<option value="${s.id}">${s.name}</option>`).join('');
  }

  function subjectNameById(id) {
    const f = SUBJECTS.find(s => s.id === id);
    return f ? f.name : id;
  }

  function validateForm() {
    const date = el.date.value;
    const subject = el.subject.value;
    const solved = nonNegativeInt(el.solved.value);
    const correct = nonNegativeInt(el.correct.value);
    const wrong = nonNegativeInt(el.wrong.value);
    const blank = nonNegativeInt(el.blank.value);

    let error = '';
    if (!date) error = 'Tarih gerekli.';
    else if (!subject) error = 'Ders seçiniz.';
    else if (solved === null || correct === null || wrong === null || blank === null) error = 'Sayılar 0 veya pozitif tam sayı olmalı.';
    else if (correct + wrong + blank > solved) error = 'Doğru + Yanlış + Boş, çözüleni aşamaz.';

    return {
      ok: !error,
      error,
      value: error ? null : { date, subject, solved, correct, wrong, blank }
    };
  }

  function clearFormAfterSave() {
    el.solved.value = '';
    el.correct.value = '';
    el.wrong.value = '';
    el.blank.value = '';
    el.formHint.textContent = 'Kaydedildi.';
    setTimeout(() => { el.formHint.textContent = ''; }, 1500);
  }

  function computeTotals(entries) {
    const perSubject = {};
    const overall = { solved: 0, correct: 0, wrong: 0, blank: 0 };

    for (const s of SUBJECTS) {
      perSubject[s.id] = { solved: 0, correct: 0, wrong: 0, blank: 0 };
    }

    for (const e of entries) {
      const t = perSubject[e.subject] || (perSubject[e.subject] = { solved: 0, correct: 0, wrong: 0, blank: 0 });
      t.solved += e.solved;
      t.correct += e.correct;
      t.wrong += e.wrong;
      t.blank += e.blank;

      overall.solved += e.solved;
      overall.correct += e.correct;
      overall.wrong += e.wrong;
      overall.blank += e.blank;
    }

    return { perSubject, overall };
  }

  function renderTotals(entries) {
    const { perSubject, overall } = computeTotals(entries);
    el.totalsGrid.innerHTML = SUBJECTS.map(s => {
      const t = perSubject[s.id];
      const answered = t.correct + t.wrong;
      const left = t.solved - answered - t.blank;
      const acc = t.solved > 0 ? Math.round((t.correct / Math.max(1, answered)) * 100) : 0;
      return `
        <div class="total-card">
          <h4>${s.name}</h4>
          <div class="totals-line">
            <span class="pill">Çözülen: ${t.solved}</span>
            <span class="pill">Doğru: ${t.correct}</span>
            <span class="pill warn">Yanlış: ${t.wrong}</span>
            <span class="pill">Boş: ${t.blank}</span>
            <span>Net: ${t.correct - t.wrong}</span>
            <span>Doğruluk: ${acc}%</span>
          </div>
        </div>`;
    }).join('');

    const answeredOverall = overall.correct + overall.wrong;
    const overallAcc = overall.solved > 0 ? Math.round((overall.correct / Math.max(1, answeredOverall)) * 100) : 0;
    el.overallTotals.innerHTML = `
      <span class="pill">Çözülen: ${overall.solved}</span>
      <span class="pill">Doğru: ${overall.correct}</span>
      <span class="pill warn">Yanlış: ${overall.wrong}</span>
      <span class="pill">Boş: ${overall.blank}</span>
      <span>Net: ${overall.correct - overall.wrong}</span>
      <span>Doğruluk: ${overallAcc}%</span>
    `;
  }

  function renderEntries(entries) {
    const sorted = [...entries].sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
    el.entriesTbody.innerHTML = sorted.map(e => `
      <tr>
        <td>${formatDate(e.date)}</td>
        <td>${subjectNameById(e.subject)}</td>
        <td>${e.solved}</td>
        <td>${e.correct}</td>
        <td>${e.wrong}</td>
        <td>${e.blank}</td>
        <td><button class="btn danger" data-del="${e.id}">Sil</button></td>
      </tr>
    `).join('');

    // Wire delete buttons
    el.entriesTbody.querySelectorAll('button[data-del]').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-del');
        if (!id) return;
        if (confirm('Bu girişi silmek istiyor musunuz?')) {
          deleteEntry(id);
          hydrate();
        }
      });
    });
  }

  function exportCSV(entries) {
    const header = ['id','tarih','ders','cozulen','dogru','yanlis','bos'];
    const rows = entries.map(e => [
      e.id,
      e.date,
      subjectNameById(e.subject),
      e.solved,
      e.correct,
      e.wrong,
      e.blank
    ]);
    const csv = [header, ...rows].map(r => r.map(v => `"${String(v).replaceAll('"', '""')}"`).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'lgs_takip.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function hydrate() {
    const entries = readEntries();
    renderTotals(entries);
    renderEntries(entries);
  }

  function initDefaults() {
    const today = new Date();
    const iso = today.toISOString().slice(0,10);
    el.date.value = iso;
  }

  function bindEvents() {
    el.form.addEventListener('submit', (ev) => {
      ev.preventDefault();
      const v = validateForm();
      if (!v.ok) {
        el.formHint.textContent = v.error;
        return;
      }
      const entry = { id: generateId(), ...v.value };
      addEntry(entry);
      hydrate();
      clearFormAfterSave();
    });

    el.exportBtn.addEventListener('click', () => {
      const entries = readEntries();
      if (entries.length === 0) {
        alert('Dışa aktarılacak kayıt yok.');
        return;
      }
      exportCSV(entries);
    });

    el.resetBtn.addEventListener('click', () => {
      if (confirm('Tüm verileri silmek istediğinize emin misiniz?')) {
        writeEntries([]);
        hydrate();
      }
    });
  }

  function main() {
    populateSubjects();
    initDefaults();
    bindEvents();
    hydrate();
  }

  document.addEventListener('DOMContentLoaded', main);
})();

