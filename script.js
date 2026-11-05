(() => {
  'use strict';

  const DATA_URL = 'data.xlsx';
  const LOW_BATTERY = 20; // dưới mức này (%) thì cảnh báo pin thấp

  const nf1 = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 1 });
  const $ = (id) => document.getElementById(id);

  const el = {
    brand: $('brand'),
    search: $('search'),
    status: $('data-status'),
    fallback: $('load-fallback'),
    fileInput: $('file-input'),
    results: $('results'),
    picker: $('picker'),
    picked: $('picked'),
    pickedName: $('picked-name'),
    pickedSource: $('picked-source'),
    changeCar: $('change-car'),
    range: $('range'),
    capacity: $('capacity'),
    rangeHint: $('range-hint'),
    rangeHintText: $('range-hint-text'),
    rangeReset: $('range-reset'),
    distance: $('distance'),
    round: $('round-trip'),
    roundNote: $('round-note'),
    slider: $('battery-slider'),
    battery: $('battery-number'),
    used: $('res-used'),
    usedKwh: $('res-used-kwh'),
    left: $('res-left'),
    leftKwh: $('res-left-kwh'),
    barLeft: $('bar-left'),
    barUsed: $('bar-used'),
    visual: $('battery-visual'),
    msg: $('res-msg'),
    total: $('res-total'),
    totalNote: $('res-total-note'),
    max: $('res-max'),
  };

  let cars = [];
  let selected = null;
  let sourceName = DATA_URL;

  /* ---------- Tiện ích ---------- */

  // Bỏ dấu tiếng Việt, chữ thường: "Mẫu xe" -> "mau xe"
  const fold = (s) =>
    String(s ?? '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd')
      .replace(/Đ/g, 'D')
      .toLowerCase()
      .trim();

  const clamp = (n, min, max) => Math.min(max, Math.max(min, n));

  // Đọc số từ ô Excel: 326, "326,4", "326 km"
  function toNumber(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : NaN;
    const s = String(v ?? '').trim().replace(/\s/g, '');
    if (!s) return NaN;
    const normalized = s.includes(',') ? s.replace(/\./g, '').replace(',', '.') : s;
    const n = parseFloat(normalized.replace(/[^0-9.\-]/g, ''));
    return Number.isFinite(n) ? n : NaN;
  }

  const fmtKm = (n) => `${nf1.format(n)} km`;
  const fmtPct = (n) => `${nf1.format(n)}%`;
  const fmtKwh = (n) => `${nf1.format(n)} kWh`;

  /* ---------- Đọc data.xlsx ---------- */

  function findCol(headers, keywords, skip = []) {
    return headers.findIndex(
      (h, i) => !skip.includes(i) && keywords.some((k) => h.includes(k))
    );
  }

  function parseWorkbook(buffer) {
    if (typeof XLSX === 'undefined') {
      throw new Error('Không tải được thư viện đọc Excel (SheetJS).');
    }
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws || !ws['!ref']) throw new Error('Sheet đầu tiên trống.');

    const startRow = XLSX.utils.decode_range(ws['!ref']).s.r;
    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: '',
      blankrows: true,
      raw: true,
    });

    const headerIdx = rows.findIndex(
      (r) => r.filter((c) => String(c).trim() !== '').length >= 2
    );
    if (headerIdx < 0) throw new Error('Không tìm thấy dòng tiêu đề.');

    const headers = rows[headerIdx].map(fold);
    const cBrand = Math.max(0, findCol(headers, ['hang', 'brand', 'thuong hieu', 'make']));
    const cRange = findCol(headers, ['km', 'quang duong', 'range'], [cBrand]);
    const cModel = findCol(headers, ['mau', 'model', 'ten'], [cBrand, cRange]);
    const cCap = findCol(headers, ['kwh', 'dung luong', 'pin', 'battery'], [cBrand, cRange, cModel]);
    const cStd = findCol(headers, ['chuan', 'standard'], [cBrand, cRange, cModel, cCap]);

    const col = {
      brand: cBrand,
      model: cModel >= 0 ? cModel : 1,
      range: cRange >= 0 ? cRange : 2,
      cap: cCap >= 0 ? cCap : 3,
      std: cStd >= 0 ? cStd : 4,
    };

    const list = [];
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const brand = String(r[col.brand] ?? '').trim();
      const model = String(r[col.model] ?? '').trim();
      const range = toNumber(r[col.range]);
      if ((!brand && !model) || !(range > 0)) continue;

      const capacity = toNumber(r[col.cap]);
      const standard = String(r[col.std] ?? '').trim();
      const name =
        model && brand && !fold(model).startsWith(fold(brand)) ? `${brand} ${model}` : model || brand;

      list.push({
        brand: brand || model,
        model,
        name,
        range,
        capacity: capacity > 0 ? capacity : null,
        standard,
        row: startRow + i + 1, // số dòng thật trong Excel
        haystack: fold(`${brand} ${model} ${standard}`),
      });
    }
    if (!list.length) throw new Error('Không có dòng xe hợp lệ (cần cột km/lần sạc là số).');
    return list;
  }

  async function loadFromUrl() {
    const res = await fetch(DATA_URL, { cache: 'no-cache' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.arrayBuffer();
  }

  function applyCars(list, name) {
    cars = list;
    sourceName = name;
    el.fallback.hidden = true;
    el.status.classList.remove('error');
    el.status.textContent = `Đã tải ${cars.length} xe từ ${sourceName}`;

    const brands = [...new Set(cars.map((c) => c.brand))].sort((a, b) => a.localeCompare(b, 'vi'));
    el.brand.length = 1; // giữ "Tất cả hãng"
    brands.forEach((b) => {
      const opt = document.createElement('option');
      opt.value = b;
      opt.textContent = `${b} (${cars.filter((c) => c.brand === b).length})`;
      el.brand.appendChild(opt);
    });
    el.brand.disabled = false;
    el.search.disabled = false;
    selected = null;
    showPicker(false);
    renderList();
    calculate();
  }

  function showLoadError(err) {
    console.error(err);
    el.status.classList.add('error');
    el.status.textContent = `Lỗi tải ${DATA_URL}: ${err.message || err}`;
    el.fallback.hidden = false;
  }

  /* ---------- Danh sách kết quả ---------- */

  function renderList() {
    const tokens = fold(el.search.value).split(/\s+/).filter(Boolean);
    const brand = el.brand.value;
    const hits = cars.filter(
      (c) => (!brand || c.brand === brand) && tokens.every((t) => c.haystack.includes(t))
    );

    el.results.textContent = '';
    el.results.hidden = false;

    if (!hits.length) {
      const li = document.createElement('li');
      li.className = 'empty-row';
      li.textContent = 'Không có xe phù hợp trong data.xlsx. Thử từ khóa khác hoặc chọn "Tất cả hãng".';
      el.results.appendChild(li);
    }

    hits.forEach((car) => {
      const li = document.createElement('li');
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'result-btn';
      btn.dataset.row = String(car.row);

      const left = document.createElement('span');
      const name = document.createElement('span');
      name.className = 'result-name';
      name.textContent = car.name;
      left.appendChild(name);
      left.appendChild(document.createElement('br'));
      const brand = document.createElement('span');
      brand.className = 'result-brand';
      brand.textContent = car.brand;
      left.appendChild(brand);

      const right = document.createElement('span');
      const km = document.createElement('span');
      km.className = 'result-km';
      km.textContent = fmtKm(car.range);
      right.appendChild(km);
      const metaParts = [];
      if (car.capacity) metaParts.push(fmtKwh(car.capacity));
      if (car.standard) metaParts.push(car.standard);
      if (metaParts.length) {
        right.appendChild(document.createElement('br'));
        const meta = document.createElement('span');
        meta.className = 'result-meta';
        meta.textContent = metaParts.join(', ');
        right.appendChild(meta);
      }
      right.style.textAlign = 'right';

      btn.appendChild(left);
      btn.appendChild(right);
      btn.addEventListener('click', () => selectCar(car));
      li.appendChild(btn);
      el.results.appendChild(li);
    });
  }

  function showPicker(focus) {
    el.picker.hidden = false;
    el.picked.hidden = true;
    if (focus) el.search.focus();
  }

  function selectCar(car) {
    selected = car;
    el.range.value = car.range;
    el.capacity.value = car.capacity ?? '';
    el.pickedName.textContent = car.name;
    el.pickedSource.textContent =
      `Từ ${sourceName}, dòng ${car.row}` + (car.standard ? `, chuẩn ${car.standard}` : '');
    el.picker.hidden = true;
    el.picked.hidden = false;
    calculate();
  }

  /* ---------- Tính toán (tự động) ---------- */

  function setEmpty(message) {
    el.used.textContent = '–';
    el.left.textContent = '–';
    el.usedKwh.textContent = '';
    el.leftKwh.textContent = '';
    el.total.textContent = '–';
    el.totalNote.textContent = '';
    el.max.textContent = '–';
    el.barLeft.style.width = '0%';
    el.barUsed.style.width = '0%';
    el.visual.setAttribute('aria-label', 'Biểu đồ pin: chưa có dữ liệu');
    el.msg.className = 'msg';
    el.msg.textContent = message;
  }

  function calculate() {
    updateRangeHint();

    const range = parseFloat(el.range.value);
    const capacity = parseFloat(el.capacity.value);
    const dist = Math.max(0, parseFloat(el.distance.value) || 0);
    const battery = clamp(parseFloat(el.battery.value) || 0, 0, 100);
    const roundTrip = el.round.checked;

    if (!(range > 0)) {
      setEmpty('Chọn xe hoặc nhập km mỗi lần sạc để bắt đầu tính.');
      return;
    }

    const total = roundTrip ? dist * 2 : dist;
    const usedPct = (total / range) * 100;        // pin cần cho chuyến đi
    const leftPct = Math.max(0, battery - usedPct);
    const maxKm = (range * battery) / 100;
    const enough = usedPct <= battery;

    // Hiển thị số
    el.used.textContent = fmtPct(usedPct);
    el.left.textContent = fmtPct(leftPct);
    el.total.textContent = fmtKm(total);
    el.totalNote.textContent = roundTrip
      ? `2 × ${fmtKm(dist)} (khứ hồi)`
      : 'Một chiều';
    el.max.textContent = fmtKm(maxKm);

    if (capacity > 0) {
      el.usedKwh.textContent = fmtKwh((usedPct / 100) * capacity);
      el.leftKwh.textContent = fmtKwh((leftPct / 100) * capacity);
    } else {
      el.usedKwh.textContent = '';
      el.leftKwh.textContent = '';
    }

    // Thanh pin: phần còn lại (xanh) + phần tiêu thụ (vằn vàng), tối đa bằng pin hiện tại
    const barUsed = Math.min(usedPct, battery);
    el.barLeft.style.width = `${leftPct}%`;
    el.barUsed.style.width = `${barUsed}%`;
    el.visual.setAttribute(
      'aria-label',
      `Biểu đồ pin: còn lại ${fmtPct(leftPct)}, tiêu thụ ${fmtPct(barUsed)}`
    );

    // Thông báo
    if (dist === 0) {
      el.msg.className = 'msg';
      el.msg.textContent = 'Nhập quãng đường để tính pin tiêu thụ.';
    } else if (!enough) {
      const shortKm = total - maxKm;
      el.msg.className = 'msg bad';
      el.msg.textContent =
        `Không đủ pin: thiếu khoảng ${fmtKm(shortKm)} (${fmtPct(usedPct - battery)} pin). ` +
        'Cần sạc thêm trước khi đi hoặc dọc đường.';
    } else if (leftPct < LOW_BATTERY) {
      el.msg.className = 'msg low';
      el.msg.textContent = `Đủ pin nhưng chỉ còn ${fmtPct(leftPct)}. Nên tìm trạm sạc trước khi đi.`;
    } else {
      el.msg.className = 'msg ok';
      el.msg.textContent = roundTrip
        ? `Đủ pin cho cả đi và về, còn dư ${fmtPct(leftPct)}.`
        : `Đủ pin cho chuyến đi, còn dư ${fmtPct(leftPct)}.`;
    }
  }

  function updateRangeHint() {
    if (!selected) {
      el.rangeHint.hidden = true;
      return;
    }
    const now = parseFloat(el.range.value);
    const changed = !(Math.abs(now - selected.range) < 1e-9);
    el.rangeHint.hidden = !changed;
    if (changed) {
      el.rangeHintText.textContent = `Đã sửa. Giá trị trong ${sourceName}: ${fmtKm(selected.range)}.`;
    }
  }

  /* ---------- Sự kiện ---------- */

  el.brand.addEventListener('change', renderList);
  el.search.addEventListener('input', renderList);
  el.search.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const first = el.results.querySelector('.result-btn');
      if (first) first.click();
    }
  });
  el.changeCar.addEventListener('click', () => showPicker(true));

  el.range.addEventListener('input', calculate);
  el.capacity.addEventListener('input', calculate);
  el.distance.addEventListener('input', calculate);
  el.rangeReset.addEventListener('click', () => {
    if (!selected) return;
    el.range.value = selected.range;
    calculate();
  });

  el.round.addEventListener('change', () => {
    el.roundNote.textContent = el.round.checked
      ? 'Bật: tính cả đi và về (×2)'
      : 'Tắt: chỉ tính một chiều';
    calculate();
  });

  el.slider.addEventListener('input', () => {
    el.battery.value = el.slider.value;
    calculate();
  });
  el.battery.addEventListener('input', () => {
    const n = parseFloat(el.battery.value);
    if (Number.isFinite(n)) el.slider.value = clamp(n, 0, 100);
    calculate();
  });
  el.battery.addEventListener('blur', () => {
    const n = clamp(parseFloat(el.battery.value) || 0, 0, 100);
    el.battery.value = n;
    el.slider.value = n;
    calculate();
  });

  // Chọn file thủ công (dùng khi mở bằng file:// hoặc tải lỗi)
  el.fileInput.addEventListener('change', async () => {
    const file = el.fileInput.files && el.fileInput.files[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      applyCars(parseWorkbook(buf), file.name);
    } catch (err) {
      showLoadError(err);
    }
  });

  /* ---------- Khởi động ---------- */

  setEmpty('Chọn xe hoặc nhập km mỗi lần sạc để bắt đầu tính.');
  calculate();

  loadFromUrl()
    .then((buf) => applyCars(parseWorkbook(buf), DATA_URL))
    .catch(showLoadError);
})();
