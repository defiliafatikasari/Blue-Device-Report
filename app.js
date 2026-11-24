/**
 * TELKOMSEL BLUE DEVICE — CONTROLLER LAPORAN KINERJA
 * PIC Lapangan: TIKA (Area Pamekasan)
 * Supervisor: Bpk. Choirul Umam
 * Data: 100% Real dari BLUE DEVICE TIKA.xlsx
 */

(function () {
  'use strict';

  // --- APPLICATION STATE ---
  const state = {
    allData: [],
    evidenceStore: {},
    currentPic: 'tika', // 'tika' (default) | 'NURUL' | 'VINDA' | 'ROS' | '...' | 'all'
    currentSubFilter: 'all', // 'all' | 'caring' | 'visit' | 'setuju'
    searchQuery: '',
    currentPage: 1,
    pageSize: 25,
    selectedRecord: null,
    bulkPendingFiles: [],
    charts: {
      caring: null,
      interest: null
    }
  };

  // --- INITIALIZATION ---
  document.addEventListener('DOMContentLoaded', () => {
    initEvidenceStore();
    initDataWithAutoFetch();
    setupEventListeners();
  });

  // --- ROBUST INDEXEDDB STORAGE (MENDUKUNG DATA BESAR TANPA BATAS KUOTA) ---
  const DB_NAME = 'TELKOMSEL_BLUE_DEVICE_DB';
  const DB_VERSION = 1;
  const STORE_NAME = 'app_data';

  function getDB() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) {
        return reject(new Error('IndexedDB not supported'));
      }
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME);
        }
      };
      request.onsuccess = (e) => resolve(e.target.result);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  async function idbGet(key) {
    try {
      const db = await getDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const req = store.get(key);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    } catch (err) {
      console.warn('idbGet error:', err);
      return null;
    }
  }

  async function idbSet(key, value) {
    try {
      const db = await getDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        const req = store.put(value, key);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
      });
    } catch (err) {
      console.warn('idbSet error:', err);
      return false;
    }
  }

  async function idbClear() {
    try {
      const db = await getDB();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, 'readwrite');
        const store = tx.objectStore(STORE_NAME);
        const req = store.clear();
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
      });
    } catch (err) {
      console.warn('idbClear error:', err);
      return false;
    }
  }

  // --- EVIDENCE STORE ---
  async function initEvidenceStore() {
    try {
      const saved = await idbGet('TELKOMSEL_BD_REAL_EVIDENCE');
      if (saved && typeof saved === 'object') {
        state.evidenceStore = saved;
        return;
      }
    } catch (e) {}

    const saved = localStorage.getItem('TELKOMSEL_BD_REAL_EVIDENCE');
    if (saved) {
      try {
        state.evidenceStore = JSON.parse(saved);
      } catch (e) {
        state.evidenceStore = {};
      }
    } else {
      state.evidenceStore = {};
    }
  }

  async function saveEvidenceStore() {
    await idbSet('TELKOMSEL_BD_REAL_EVIDENCE', state.evidenceStore);
    try {
      localStorage.setItem('TELKOMSEL_BD_REAL_EVIDENCE', JSON.stringify(state.evidenceStore));
    } catch (e) {
      // Data aman tersimpan di IndexedDB
    }
  }

  // --- DATA LOADING & PERSISTENCE DENGAN INDEXEDDB ---
  async function initDataWithAutoFetch() {
    // 1. Cek apakah ada data Excel yang pernah diupload di IndexedDB
    let savedData = null;
    let savedMeta = null;

    try {
      savedData = await idbGet('ACTIVE_EXCEL_DATA');
      savedMeta = await idbGet('ACTIVE_EXCEL_META');
    } catch (e) {
      console.warn('IndexedDB read notice:', e);
    }

    if (savedData && Array.isArray(savedData) && savedData.length > 0) {
      // DATA TERSIMPAN DI INDEXEDDB DITEMUKAN!
      // INI MENJAMIN DATA TETAP SAMA DAN TIDAK AKAN BERUBAH/RESET SAAT REFRESH F5!
      state.allData = savedData;
      if (savedMeta) updateSyncStatusUI(savedMeta);
      renderApp();
      return;
    }

    // 2. Jika belum pernah upload, gunakan data master dari data.js
    if (window.REAL_EXCEL_DATA && window.REAL_EXCEL_DATA.length) {
      state.allData = JSON.parse(JSON.stringify(window.REAL_EXCEL_DATA));
      const defaultMeta = {
        sourceName: 'BLUE DEVICE TIKA.xlsx',
        timeStr: 'Data Master Bawaan',
        rowCount: state.allData.length,
        isUserUpload: false
      };
      updateSyncStatusUI(defaultMeta);
      renderApp();
    }
  }

  function updateSyncStatusUI(meta) {
    const statusEl = document.getElementById('excelSyncStatus');
    const badgeEl = document.getElementById('excelRowsCountBadge');
    const btnReset = document.getElementById('btnResetMasterExcel');

    if (statusEl) {
      if (meta && meta.isUserUpload) {
        statusEl.innerHTML = `<span class="text-emerald-400 font-bold">● File Upload: ${escapeHtml(meta.sourceName)}</span> <span class="text-slate-400">(${escapeHtml(meta.timeStr)})</span>`;
      } else {
        statusEl.innerHTML = `<span class="text-slate-300">File Master: ${escapeHtml(meta?.sourceName || 'BLUE DEVICE TIKA.xlsx')}</span>`;
      }
    }

    if (badgeEl && state.allData) {
      badgeEl.textContent = state.allData.length.toLocaleString('id-ID') + ' Baris';
    }

    if (btnReset) {
      if (meta && meta.isUserUpload) {
        btnReset.classList.remove('hidden');
      } else {
        btnReset.classList.add('hidden');
      }
    }
  }

  async function processAndStoreRows(rawRows, sourceName, isUserUpload = false) {
    // Pertahankan seluruh baris pelanggan valid dari Excel yang diupload
    const validRows = rawRows.filter(r => r && (r['Idn Pelanggan'] || r['No Telepon'] || r['NAMA PELANGGAN']));
    state.allData = validRows.length > 0 ? validRows : rawRows;

    const timeStr = new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) + ' WIB';
    const meta = {
      sourceName: sourceName || 'BLUE DEVICE TIKA.xlsx',
      timeStr: timeStr,
      rowCount: state.allData.length,
      isUserUpload: Boolean(isUserUpload)
    };

    // SIMPAN KE INDEXEDDB: Data tersimpan permanen di browser, tidak akan hilang/kembali saat di-refresh!
    await idbSet('ACTIVE_EXCEL_DATA', state.allData);
    await idbSet('ACTIVE_EXCEL_META', meta);

    // Simpan metadata ringan ke localStorage
    localStorage.setItem('TELKOMSEL_BD_LAST_SYNC', timeStr);
    localStorage.setItem('TELKOMSEL_BD_SOURCE_NAME', meta.sourceName);
    localStorage.setItem('TELKOMSEL_BD_IS_USER_UPLOAD', isUserUpload ? '1' : '0');

    updateSyncStatusUI(meta);
    state.currentPage = 1;
    renderApp();
  }

  // --- EVENT LISTENERS ---
  function setupEventListeners() {
    // PIC Filter Buttons (5 Caring Categories + All)
    const picButtons = document.querySelectorAll('.pic-filter-btn');
    picButtons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        picButtons.forEach(b => {
          b.classList.remove('active', 'bg-white', 'text-[#E60000]', 'border-red-200', 'shadow-xs', 'font-bold');
          b.classList.add('text-slate-600', 'font-medium');
        });
        btn.classList.add('active', 'bg-white', 'text-[#E60000]', 'border-red-200', 'shadow-xs', 'font-bold');
        btn.classList.remove('text-slate-600', 'font-medium');

        state.currentPic = btn.getAttribute('data-pic');
        state.currentPage = 1;
        renderApp();
      });
    });

    // Daily Excel Update Button (Tersimpan Permanen di IndexedDB)
    const btnReload = document.getElementById('btnReloadExcel');
    const dailyInput = document.getElementById('dailyExcelFileInput');
    if (btnReload && dailyInput) {
      btnReload.addEventListener('click', () => dailyInput.click());
      dailyInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        if (typeof XLSX === 'undefined') {
          alert('Library Excel sedang dimuat. Coba 1 detik lagi.');
          return;
        }

        const reader = new FileReader();
        reader.onload = async function (evt) {
          try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

            if (!rawRows || rawRows.length === 0) {
              alert('File Excel kosong atau tidak terbaca.');
              return;
            }

            await processAndStoreRows(rawRows, file.name, true);
            alert('Sukses! Data Excel berhasil diperbarui: ' + state.allData.length.toLocaleString('id-ID') + ' baris dimuat.\n\nData telah disimpan permanen di IndexedDB browser Anda dan TIDAK AKAN HILANG atau kembali ke data awal saat halaman di-refresh (F5).');
          } catch (err) {
            alert('Gagal membaca file Excel: ' + err.message);
          }
        };
        reader.readAsArrayBuffer(file);
      });
    }

    // Reset ke Data Master Awal Button
    const btnResetMaster = document.getElementById('btnResetMasterExcel');
    if (btnResetMaster) {
      btnResetMaster.addEventListener('click', async () => {
        if (confirm('Apakah Anda yakin ingin mereset tampilan dan mengembalikan ke data master awal (BLUE DEVICE TIKA.xlsx)?')) {
          await idbClear();
          localStorage.removeItem('TELKOMSEL_BD_LAST_SYNC');
          localStorage.removeItem('TELKOMSEL_BD_SOURCE_NAME');
          localStorage.removeItem('TELKOMSEL_BD_IS_USER_UPLOAD');
          state.allData = JSON.parse(JSON.stringify(window.REAL_EXCEL_DATA));
          const defaultMeta = {
            sourceName: 'BLUE DEVICE TIKA.xlsx',
            timeStr: 'Data Master Bawaan',
            rowCount: state.allData.length,
            isUserUpload: false
          };
          updateSyncStatusUI(defaultMeta);
          state.currentPage = 1;
          renderApp();
          alert('Berhasil! Tampilan laporan telah dikembalikan ke data master bawaan.');
        }
      });
    }

    // Main Tab Navigation
    document.getElementById('navTabTable').addEventListener('click', () => switchMainTab('tabTable'));
    document.getElementById('navTabGallery').addEventListener('click', () => switchMainTab('tabGallery'));
    document.getElementById('navTabLaporan').addEventListener('click', () => switchMainTab('tabRekapLaporan'));

    // Sub-Filter Buttons in Table View
    const subFilterBtns = document.querySelectorAll('.sub-filter-btn');
    subFilterBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        subFilterBtns.forEach(b => {
          b.classList.remove('active', 'font-bold', 'bg-white', 'text-slate-900', 'shadow-xs');
          b.classList.add('font-medium', 'text-slate-600');
        });
        btn.classList.add('active', 'font-bold', 'bg-white', 'text-slate-900', 'shadow-xs');
        btn.classList.remove('font-medium', 'text-slate-600');

        state.currentSubFilter = btn.getAttribute('data-subfilter');
        state.currentPage = 1;
        renderTableAndPagination();
      });
    });

    // Search Input
    const searchInput = document.getElementById('searchInput');
    const btnClearSearch = document.getElementById('btnClearSearch');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        state.searchQuery = e.target.value.trim().toLowerCase();
        if (state.searchQuery) {
          btnClearSearch.classList.remove('hidden');
        } else {
          btnClearSearch.classList.add('hidden');
        }
        state.currentPage = 1;
        renderTableAndPagination();
      });

      btnClearSearch.addEventListener('click', () => {
        searchInput.value = '';
        state.searchQuery = '';
        btnClearSearch.classList.add('hidden');
        state.currentPage = 1;
        renderTableAndPagination();
      });
    }

    // Page Size Selector
    const pageSizeSelect = document.getElementById('pageSizeSelect');
    if (pageSizeSelect) {
      pageSizeSelect.addEventListener('change', (e) => {
        if (e.target.value === 'all') {
          state.pageSize = 999999;
        } else {
          state.pageSize = parseInt(e.target.value, 10);
        }
        state.currentPage = 1;
        renderTableAndPagination();
      });
    }

    // Tombol Tampilkan Semua Baris (All)
    const btnShowAllRows = document.getElementById('btnShowAllRows');
    if (btnShowAllRows) {
      btnShowAllRows.addEventListener('click', () => {
        state.pageSize = 999999;
        if (pageSizeSelect) pageSizeSelect.value = 'all';
        state.currentPage = 1;
        renderTableAndPagination();
      });
    }

    // Excel Export
    const btnExport = document.getElementById('btnExportExcel');
    if (btnExport) {
      btnExport.addEventListener('click', handleExcelExport);
    }

    // Print Button
    const btnPrint = document.getElementById('btnPrintReport');
    if (btnPrint) {
      btnPrint.addEventListener('click', () => window.print());
    }

    // Modal Close Handlers
    document.getElementById('btnCloseCustomerModal').addEventListener('click', closeCustomerModal);
    document.getElementById('btnCloseModalBottom').addEventListener('click', closeCustomerModal);
    document.getElementById('customerDetailModal').addEventListener('click', (e) => {
      if (e.target.id === 'customerDetailModal') closeCustomerModal();
    });

    // Bulk Upload Handlers
    document.getElementById('btnBulkUploadModal').addEventListener('click', openBulkUploadModal);
    document.getElementById('btnCloseBulkUpload').addEventListener('click', closeBulkUploadModal);
    document.getElementById('btnCancelBulkUpload').addEventListener('click', closeBulkUploadModal);
    document.getElementById('bulkUploadModal').addEventListener('click', (e) => {
      if (e.target.id === 'bulkUploadModal') closeBulkUploadModal();
    });

    const bulkDropZone = document.getElementById('bulkDropZone');
    const bulkFileInput = document.getElementById('bulkFileInput');
    if (bulkDropZone && bulkFileInput) {
      bulkDropZone.addEventListener('click', () => bulkFileInput.click());
      bulkFileInput.addEventListener('change', handleBulkFilesSelect);

      bulkDropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        bulkDropZone.classList.add('border-red-500', 'bg-red-50/40');
      });
      bulkDropZone.addEventListener('dragleave', () => {
        bulkDropZone.classList.remove('border-red-500', 'bg-red-50/40');
      });
      bulkDropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        bulkDropZone.classList.remove('border-red-500', 'bg-red-50/40');
        if (e.dataTransfer.files && e.dataTransfer.files.length) {
          handleBulkFiles(Array.from(e.dataTransfer.files));
        }
      });
    }

    document.getElementById('btnSaveBulkUpload').addEventListener('click', saveBulkUploadToGallery);

    // Lightbox Handlers
    document.getElementById('btnCloseLightbox').addEventListener('click', closeLightbox);
    document.getElementById('lightboxModal').addEventListener('click', (e) => {
      if (e.target.id === 'lightboxModal') closeLightbox();
    });

    // Modal Single Photo Upload
    const modalPhotoInput = document.getElementById('modalSinglePhotoUpload');
    if (modalPhotoInput) {
      modalPhotoInput.addEventListener('change', handleSingleModalPhotoUpload);
    }

    // Modal Photo container click -> Lightbox
    const modalPhotoContainer = document.getElementById('modalPhotoContainer');
    if (modalPhotoContainer) {
      modalPhotoContainer.addEventListener('click', () => {
        const img = document.getElementById('modalEvidenceImage');
        const caption = document.getElementById('modalCustomerName').textContent + ' — ' + document.getElementById('modalCustomerIdnPhone').textContent;
        if (img && img.src) {
          openLightbox(img.src, caption);
        }
      });
    }

    // Download Photo inside Modal
    const btnDownloadEvidence = document.getElementById('btnDownloadEvidence');
    if (btnDownloadEvidence) {
      btnDownloadEvidence.addEventListener('click', () => {
        const img = document.getElementById('modalEvidenceImage');
        if (img && img.src) {
          const a = document.createElement('a');
          a.href = img.src;
          a.download = (document.getElementById('col_IdnPelanggan').textContent || 'evidence_tika') + '.jpg';
          a.click();
        }
      });
    }

    // Escape Key
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        closeCustomerModal();
        closeBulkUploadModal();
        closeLightbox();
      }
    });
  }

  // --- TAB SWITCHING ---
  function switchMainTab(tabName) {
    const btnTable = document.getElementById('navTabTable');
    const btnGallery = document.getElementById('navTabGallery');
    const btnLaporan = document.getElementById('navTabLaporan');

    const paneTable = document.getElementById('contentTabTable');
    const paneGallery = document.getElementById('contentTabGallery');
    const paneLaporan = document.getElementById('contentTabLaporan');
    const tableTopControls = document.getElementById('tableTopControls');

    [btnTable, btnGallery, btnLaporan].forEach(b => {
      b.classList.remove('active', 'bg-slate-900', 'text-white');
      b.classList.add('bg-white', 'text-slate-700');
    });

    paneTable.classList.add('hidden');
    paneGallery.classList.add('hidden');
    paneLaporan.classList.add('hidden');

    if (tabName === 'tabTable') {
      btnTable.classList.add('active', 'bg-slate-900', 'text-white');
      btnTable.classList.remove('bg-white', 'text-slate-700');
      paneTable.classList.remove('hidden');
      if (tableTopControls) tableTopControls.classList.remove('hidden');
      renderTableAndPagination();
    } else if (tabName === 'tabGallery') {
      btnGallery.classList.add('active', 'bg-slate-900', 'text-white');
      btnGallery.classList.remove('bg-white', 'text-slate-700');
      paneGallery.classList.remove('hidden');
      if (tableTopControls) tableTopControls.classList.add('hidden');
      renderEvidenceGallery();
    } else if (tabName === 'tabRekapLaporan') {
      btnLaporan.classList.add('active', 'bg-slate-900', 'text-white');
      btnLaporan.classList.remove('bg-white', 'text-slate-700');
      paneLaporan.classList.remove('hidden');
      if (tableTopControls) tableTopControls.classList.add('hidden');
      renderRekapLaporanTable();
    }
  }

  window.app = {
    switchMainTab: switchMainTab
  };

  // --- MATCHING LOGIC FOR PIC CARING & VISIT ---
  function isRowMatchedPIC(row, pic) {
    if (!pic || pic === 'all') return true;
    const called = String(row['CALLED BY'] || '').trim();
    const calledLower = called.toLowerCase();

    if (pic === 'kosong') {
      // Pelanggan yang belum pernah di-caring (CALLED BY kosong/strip/null)
      return !called || called === '-' || called === '#N/A';
    } else if (pic === '...') {
      // Tanpa caring tapi sudah aktif
      return called.includes('...') || called.includes('....');
    } else if (pic.toLowerCase() === 'tika') {
      // Caring murni oleh TIKA (282 baris di Excel)
      return calledLower.includes('tika');
    }
    return calledLower.includes(pic.toLowerCase());
  }

  function isTandemVisit(noteAntar) {
    if (!noteAntar) return false;
    const note = String(noteAntar).toLowerCase();
    return note.includes('&') || note.includes('dan') || note.includes('/') || note.includes(',');
  }

  function getFilteredPICData() {
    return state.allData.filter(row => isRowMatchedPIC(row, state.currentPic));
  }

  function getTableDisplayData() {
    let list = [];

    if (state.currentSubFilter === 'visit') {
      if (state.currentPic === 'all') {
        list = state.allData.filter(r => r['NOTE PIC ANTAR'] && String(r['NOTE PIC ANTAR']).trim());
      } else if (state.currentPic.toLowerCase() === 'tika') {
        list = state.allData.filter(r => String(r['NOTE PIC ANTAR'] || '').toLowerCase().includes('tika'));
      } else if (state.currentPic === 'kosong') {
        list = [];
      } else {
        list = state.allData.filter(r => String(r['NOTE PIC ANTAR'] || '').toLowerCase().includes(state.currentPic.toLowerCase()));
      }
    } else {
      list = getFilteredPICData();

      if (state.currentSubFilter === 'caring') {
        list = list.filter(r => Boolean((r['DATE CALL'] && String(r['DATE CALL']).trim()) || (r['STATUS CALL'] && String(r['STATUS CALL']).trim())));
      } else if (state.currentSubFilter === 'setuju') {
        list = list.filter(r => String(r['BERMINAT?'] || '').trim().toUpperCase() === 'TERTARIK');
      } else if (state.currentSubFilter === 'needfu') {
        list = list.filter(r => String(r['STATUS ORDER'] || '').trim().toUpperCase() === 'NEED FU');
      } else if (state.currentSubFilter === 'doneact') {
        list = list.filter(r => String(r['STATUS ORDER'] || '').trim().toUpperCase() === 'DONE ACT');
      }
    }

    if (state.searchQuery) {
      const q = state.searchQuery;
      list = list.filter(r => {
        const idn = String(r['Idn Pelanggan'] || '').toLowerCase();
        const nama = String(r['NAMA PELANGGAN'] || '').toLowerCase();
        const telp = String(r['No Telepon'] || '').toLowerCase();
        const odp = String(r['ODP'] || '').toLowerCase();
        const alamat = String(r['ALAMAT'] || '').toLowerCase();
        const pic = String(r['CALLED BY'] || '').toLowerCase() + ' ' + String(r['NOTE PIC ANTAR'] || '').toLowerCase();
        return idn.includes(q) || nama.includes(q) || telp.includes(q) || odp.includes(q) || alamat.includes(q) || pic.includes(q);
      });
    }

    return list;
  }

  // --- KPI CALCULATIONS (SESUAI PERINTAH ATASAN) ---
  function calculateKPIs() {
    const isKosongFilter = state.currentPic === 'kosong';
    const picData = getFilteredPICData();

    // 1. HASIL CARING (Panggilan)
    let caringTotal = 0;
    let diangkat = 0;
    let tdkDiangkat = 0;
    let tdkSambung = 0;

    if (!isKosongFilter) {
      picData.forEach(r => {
        const hasCallData = Boolean((r['DATE CALL'] && String(r['DATE CALL']).trim()) || (r['STATUS CALL'] && String(r['STATUS CALL']).trim()));
        if (hasCallData) {
          caringTotal++;
          const s = String(r['STATUS CALL'] || '').toUpperCase().trim();
          if (s.includes('DIANGKAT') && !s.includes('TIDAK')) {
            diangkat++;
          } else if (s.includes('TIDAK DIANGKAT')) {
            tdkDiangkat++;
          } else if (s.includes('MENYAMBUNG')) {
            tdkSambung++;
          } else {
            tdkDiangkat++;
          }
        }
      });
    }

    // 2. HASIL VISIT LAPANGAN BLUE DEVICE
    let visitTotal = 0;
    let visitMandiri = 0;
    let visitTandem = 0;

    if (!isKosongFilter) {
      let visitRows = [];
      if (state.currentPic === 'all') {
        visitRows = state.allData.filter(r => r['NOTE PIC ANTAR'] && String(r['NOTE PIC ANTAR']).trim());
      } else if (state.currentPic.toLowerCase() === 'tika') {
        visitRows = state.allData.filter(r => String(r['NOTE PIC ANTAR'] || '').toLowerCase().includes('tika'));
      } else {
        visitRows = state.allData.filter(r => String(r['NOTE PIC ANTAR'] || '').toLowerCase().includes(state.currentPic.toLowerCase()));
      }

      visitRows.forEach(r => {
        visitTotal++;
        const antar = String(r['NOTE PIC ANTAR'] || '').trim();
        if (isTandemVisit(antar)) {
          visitTandem++;
        } else {
          visitMandiri++;
        }
      });
    }

    // 3. JUMLAH BLUE DEVICE YANG PELANGGAN SETUJU & STATUS ORDER
    let setujuTotal = 0;
    let tertarikCount = 0;
    let doneActCount = 0;
    let needFuCount = 0;
    let raguCount = 0;
    let tdkTertarikCount = 0;

    if (!isKosongFilter) {
      picData.forEach(r => {
        const berm = String(r['BERMINAT?'] || '').toUpperCase().trim();
        const order = String(r['STATUS ORDER'] || '').toUpperCase().trim();

        // Di Excel: DONE ACT murni jika STATUS ORDER === 'DONE ACT'
        const isDoneAct = order === 'DONE ACT';
        const isNeedFu = order === 'NEED FU';
        const isTertarik = berm === 'TERTARIK';

        if (isTertarik) {
          tertarikCount++;
          setujuTotal++;
        }
        if (isDoneAct) {
          doneActCount++;
        }
        if (isNeedFu) {
          needFuCount++;
        }

        if (berm === 'MASIH RAGU') raguCount++;
        if (berm === 'TIDAK TERTARIK') tdkTertarikCount++;
      });
    }

    const contactRate = (caringTotal > 0) ? ((diangkat / caringTotal) * 100).toFixed(1) : '0';
    const convRate = (caringTotal > 0) ? ((setujuTotal / caringTotal) * 100).toFixed(1) : '0';

    return {
      isKosongFilter,
      queueTotal: picData.length,
      caringTotal,
      diangkat,
      tdkDiangkat,
      tdkSambung,
      contactRate,
      visitTotal,
      visitMandiri,
      visitTandem,
      setujuTotal,
      tertarikCount,
      doneActCount,
      needFuCount,
      convRate,
      raguCount,
      tdkTertarikCount
    };
  }

  // --- RENDER APP ---
  function renderApp() {
    // 1. Dynamic Badge Counts across all 10,809 rows (Direct Excel Sum)
    const countTika = state.allData.filter(r => isRowMatchedPIC(r, 'tika')).length;
    const countNurul = state.allData.filter(r => isRowMatchedPIC(r, 'NURUL')).length;
    const countVinda = state.allData.filter(r => isRowMatchedPIC(r, 'VINDA')).length;
    const countRos = state.allData.filter(r => isRowMatchedPIC(r, 'ROS')).length;
    const countDots = state.allData.filter(r => isRowMatchedPIC(r, '...')).length;
    const countKosong = state.allData.filter(r => isRowMatchedPIC(r, 'kosong')).length;
    const countAll = state.allData.length;

    const elTika = document.getElementById('badgeCount_tika');
    const elNurul = document.getElementById('badgeCount_NURUL');
    const elVinda = document.getElementById('badgeCount_VINDA');
    const elRos = document.getElementById('badgeCount_ROS');
    const elDots = document.getElementById('badgeCount_dots');
    const elKosong = document.getElementById('badgeCount_kosong');
    const elAll = document.getElementById('badgeCount_all');
    const excelBadge = document.getElementById('excelRowsCountBadge');

    if (elTika) elTika.textContent = countTika.toLocaleString('id-ID');
    if (elNurul) elNurul.textContent = countNurul.toLocaleString('id-ID');
    if (elVinda) elVinda.textContent = countVinda.toLocaleString('id-ID');
    if (elRos) elRos.textContent = countRos.toLocaleString('id-ID');
    if (elDots) elDots.textContent = countDots.toLocaleString('id-ID');
    if (elKosong) elKosong.textContent = countKosong.toLocaleString('id-ID');
    if (elAll) elAll.textContent = countAll.toLocaleString('id-ID');
    if (excelBadge) excelBadge.textContent = countAll.toLocaleString('id-ID') + ' Baris';

    const picLabel = document.getElementById('currentPicLabel');
    if (picLabel) {
      if (state.currentPic === 'tika') picLabel.textContent = 'TIKA (Uplift Agent)';
      else if (state.currentPic === 'all') picLabel.textContent = 'Semua Data Tim';
      else if (state.currentPic === '...') picLabel.textContent = 'Tanpa Caring / Sudah Aktif';
      else if (state.currentPic === 'kosong') picLabel.textContent = 'Kosong / Belum di-Caring (Antrian Database Pelanggan)';
      else picLabel.textContent = state.currentPic;
    }

    const kpi = calculateKPIs();

    // Tampilkan / Sembunyikan Alert Banner Khusus Data Antrian Kosong
    const uncalledBanner = document.getElementById('uncalledAlertBanner');
    if (uncalledBanner) {
      if (kpi.isKosongFilter) {
        uncalledBanner.classList.remove('hidden');
      } else {
        uncalledBanner.classList.add('hidden');
      }
    }

    // 1. KPI Card 1 (Caring)
    if (kpi.isKosongFilter) {
      document.getElementById('kpiCaringTotal').textContent = kpi.queueTotal.toLocaleString('id-ID');
      document.getElementById('kpiCaringDiangkat').textContent = '0';
      document.getElementById('kpiCaringTdkDiangkat').textContent = '0';
      document.getElementById('kpiCaringTdkSambung').textContent = '0';
      document.getElementById('kpiContactRate').textContent = 'Antrian Call';
    } else {
      document.getElementById('kpiCaringTotal').textContent = kpi.caringTotal.toLocaleString('id-ID');
      document.getElementById('kpiCaringDiangkat').textContent = kpi.diangkat.toLocaleString('id-ID');
      document.getElementById('kpiCaringTdkDiangkat').textContent = kpi.tdkDiangkat.toLocaleString('id-ID');
      document.getElementById('kpiCaringTdkSambung').textContent = kpi.tdkSambung.toLocaleString('id-ID');
      document.getElementById('kpiContactRate').textContent = kpi.contactRate + '% Tersambung';
    }

    // 2. KPI Card 2 (Visit)
    document.getElementById('kpiVisitTotal').textContent = kpi.visitTotal.toLocaleString('id-ID');
    document.getElementById('kpiVisitMandiri').textContent = kpi.visitMandiri.toLocaleString('id-ID');
    document.getElementById('kpiVisitTandem').textContent = kpi.visitTandem.toLocaleString('id-ID');

    // 3. KPI Card 3 (Pelanggan Setuju & Need FU)
    document.getElementById('kpiSetujuTotal').textContent = kpi.setujuTotal.toLocaleString('id-ID');
    document.getElementById('kpiTertarikCount').textContent = kpi.tertarikCount.toLocaleString('id-ID');
    const needFuEl = document.getElementById('kpiNeedFuCount');
    if (needFuEl) needFuEl.textContent = kpi.needFuCount.toLocaleString('id-ID') + ' Need FU';
    document.getElementById('kpiDoneActCount').textContent = kpi.doneActCount.toLocaleString('id-ID') + ' Done Act';
    document.getElementById('kpiConversionRate').textContent = kpi.convRate + '% Konversi';

    // 4. KPI Card 4 (Aktivasi Tuntas Done Act)
    document.getElementById('kpiAktivasiTotal').textContent = kpi.doneActCount.toLocaleString('id-ID');
    const card4NeedFu = document.getElementById('kpiCard4NeedFu');
    if (card4NeedFu) card4NeedFu.textContent = kpi.needFuCount.toLocaleString('id-ID') + ' Need FU';

    // Sub-Filter Badges (Need FU & Done Act)
    const badgeNeedFu = document.getElementById('subfilterBadge_needfu');
    const badgeDoneAct = document.getElementById('subfilterBadge_doneact');
    if (badgeNeedFu) badgeNeedFu.textContent = kpi.needFuCount.toLocaleString('id-ID');
    if (badgeDoneAct) badgeDoneAct.textContent = kpi.doneActCount.toLocaleString('id-ID');

    // Tab badges
    const tabTableBadge = document.getElementById('tabBadgeRowCount');
    if (tabTableBadge) tabTableBadge.textContent = getFilteredPICData().length.toLocaleString('id-ID');

    const tabEvidenceBadge = document.getElementById('tabBadgeEvidenceCount');
    if (tabEvidenceBadge) tabEvidenceBadge.textContent = Object.keys(state.evidenceStore).length;

    // Charts
    renderCharts(kpi);

    // Table
    renderTableAndPagination();

    // Laporan Table
    renderRekapLaporanTable();
  }

  // --- CHARTS (CHART.JS) ---
  function renderCharts(kpi) {
    if (typeof Chart === 'undefined') return;

    const ctx1 = document.getElementById('chartCaringStatus');
    if (ctx1) {
      if (state.charts.caring) state.charts.caring.destroy();

      state.charts.caring = new Chart(ctx1, {
        type: 'doughnut',
        data: {
          labels: ['Diangkat', 'Tidak Diangkat', 'Tidak Menyambung'],
          datasets: [{
            data: [kpi.diangkat, kpi.tdkDiangkat, kpi.tdkSambung],
            backgroundColor: ['#16A34A', '#E11D48', '#D97706'],
            borderWidth: 2,
            borderColor: '#FFFFFF',
            hoverOffset: 4
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'bottom',
              labels: {
                boxWidth: 10,
                font: { size: 10, family: 'Inter' }
              }
            }
          },
          cutout: '68%'
        }
      });

      const total = (kpi.diangkat + kpi.tdkDiangkat + kpi.tdkSambung) || 1;
      document.getElementById('statDiangkatPct').textContent = Math.round((kpi.diangkat / total) * 100) + '%';
      document.getElementById('statTdkDiangkatPct').textContent = Math.round((kpi.tdkDiangkat / total) * 100) + '%';
      document.getElementById('statTdkSambungPct').textContent = Math.round((kpi.tdkSambung / total) * 100) + '%';
    }

    const ctx2 = document.getElementById('chartInterestStatus');
    if (ctx2) {
      if (state.charts.interest) state.charts.interest.destroy();

      state.charts.interest = new Chart(ctx2, {
        type: 'bar',
        data: {
          labels: ['Setuju / Tertarik', 'Masih Ragu', 'Tdk Tertarik', 'Done Act'],
          datasets: [{
            label: 'Pelanggan',
            data: [kpi.tertarikCount, kpi.raguCount, kpi.tdkTertarikCount, kpi.doneActCount],
            backgroundColor: ['#16A34A', '#D97706', '#94A3B8', '#7C3AED'],
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false }
          },
          scales: {
            y: {
              beginAtZero: true,
              ticks: { precision: 0, font: { size: 10 } },
              grid: { color: '#F1F5F9' }
            },
            x: {
              ticks: { font: { size: 10 } },
              grid: { display: false }
            }
          }
        }
      });

      document.getElementById('statTertarikCount').textContent = kpi.tertarikCount;
      document.getElementById('statRaguCount').textContent = kpi.raguCount;
      document.getElementById('statTdkTertarikCount').textContent = kpi.tdkTertarikCount;
    }
  }

  // --- TABLE RENDERING & PAGINATION ---
  function renderTableAndPagination() {
    const list = getTableDisplayData();
    const tbody = document.getElementById('tableBody');
    const emptyState = document.getElementById('tableEmptyState');
    const infoEl = document.getElementById('filteredRowsInfo');
    const summaryEl = document.getElementById('paginationSummary');
    const buttonsEl = document.getElementById('paginationButtons');

    if (infoEl) infoEl.textContent = 'Menampilkan ' + list.length.toLocaleString('id-ID') + ' data tersaring';

    if (list.length === 0) {
      tbody.innerHTML = '';
      emptyState.classList.remove('hidden');
      summaryEl.textContent = '0 data';
      buttonsEl.innerHTML = '';
      return;
    }

    emptyState.classList.add('hidden');

    const isShowingAll = state.pageSize >= 999999 || state.pageSize >= list.length;
    const totalPages = isShowingAll ? 1 : Math.ceil(list.length / state.pageSize);
    if (state.currentPage > totalPages) state.currentPage = totalPages || 1;

    const startIdx = isShowingAll ? 0 : (state.currentPage - 1) * state.pageSize;
    const endIdx = isShowingAll ? list.length : Math.min(startIdx + state.pageSize, list.length);
    const pageRows = list.slice(startIdx, endIdx);

    summaryEl.textContent = isShowingAll
      ? 'Menampilkan 1 - ' + list.length.toLocaleString('id-ID') + ' dari ' + list.length.toLocaleString('id-ID') + ' data (Semua Baris)'
      : 'Menampilkan ' + (startIdx + 1) + ' - ' + endIdx + ' dari ' + list.length.toLocaleString('id-ID') + ' data';

    let rowsHtml = '';
    pageRows.forEach((row, idx) => {
      const globalNo = startIdx + idx + 1;
      const idn = String(row['Idn Pelanggan'] || '-');
      const nama = String(row['NAMA PELANGGAN'] || 'Pelanggan #' + idn.slice(-4));
      const telp = String(row['No Telepon'] || '-');
      const called = String(row['CALLED BY'] || '-');
      const dateCall = String(row['DATE CALL'] || '-');
      const statusCall = String(row['STATUS CALL'] || '-');
      const berminat = String(row['BERMINAT?'] || '-');
      const picAntar = String(row['NOTE PIC ANTAR'] || '-');
      const dateCompleted = String(row['DATE COMPLETED'] || '-');
      const dateAct = String(row['DATE ACT'] || '-');
      const statusOrder = String(row['STATUS ORDER'] || '').trim().toUpperCase();

      const isUncalled = !row['CALLED BY'] || String(row['CALLED BY']).trim() === '' || String(row['CALLED BY']).trim() === '-';

      let scallBadgeClass = 'badge-neutral';
      let scallText = escapeHtml(statusCall);

      if (isUncalled) {
        scallBadgeClass = 'bg-amber-50 text-amber-700 border border-amber-200 font-semibold';
        scallText = 'Antrian Call';
      } else if (statusCall.toUpperCase().includes('DIANGKAT') && !statusCall.toUpperCase().includes('TIDAK')) {
        scallBadgeClass = 'badge-success';
      } else if (statusCall.toUpperCase().includes('TIDAK DIANGKAT')) {
        scallBadgeClass = 'badge-danger';
      } else if (statusCall.toUpperCase().includes('MENYAMBUNG')) {
        scallBadgeClass = 'badge-warning';
      }

      let bermBadgeClass = 'badge-neutral';
      if (berminat.toUpperCase().includes('TERTARIK') && !berminat.toUpperCase().includes('TIDAK')) {
        bermBadgeClass = 'badge-success';
      } else if (berminat.toUpperCase().includes('RAGU')) {
        bermBadgeClass = 'badge-warning';
      } else if (berminat.toUpperCase().includes('TIDAK TERTARIK')) {
        bermBadgeClass = 'badge-neutral';
      }

      // Order status badge (DONE ACT vs NEED FU)
      let orderBadgeHtml = '';
      if (statusOrder === 'DONE ACT') {
        orderBadgeHtml = '<span class="badge-status bg-purple-100 text-purple-900 border border-purple-300 font-black ml-1 text-[10px]">DONE ACT</span>';
      } else if (statusOrder === 'NEED FU') {
        orderBadgeHtml = '<span class="badge-status bg-amber-100 text-amber-900 border border-amber-300 font-black ml-1 text-[10px]">NEED FU</span>';
      }

      const isTandem = isTandemVisit(picAntar);
      const tandemBadgeHtml = isTandem ? '<span class="badge-tandem ml-1.5" title="Kunjungan bersama tim">Visit Tandem / Tim</span>' : '';

      const hasPhoto = Boolean(state.evidenceStore[idn] || state.evidenceStore[telp]);
      const photoIndicator = hasPhoto ? '<span class="w-2 h-2 rounded-full bg-emerald-500 inline-block mr-1" title="Foto bukti terlampir"></span>' : '';

      const calledHtml = isUncalled
        ? '<span class="text-slate-400 italic font-mono text-[11px]">- (Belum Call)</span>'
        : `<span class="px-2 py-0.5 rounded bg-slate-100 font-medium">${escapeHtml(called)}</span>`;

      rowsHtml += `
        <tr class="hover:bg-slate-50 border-b border-slate-100">
          <td class="text-center font-medium text-slate-400 text-xs">${globalNo}</td>
          <td class="font-mono text-slate-800 font-semibold text-xs whitespace-nowrap">${photoIndicator}${idn}</td>
          <td class="font-bold text-slate-900 whitespace-nowrap">${escapeHtml(nama)}</td>
          <td class="font-mono text-slate-600 text-xs whitespace-nowrap">${escapeHtml(telp)}</td>
          <td class="text-slate-700 whitespace-nowrap">${calledHtml}</td>
          <td class="text-slate-500 text-xs whitespace-nowrap">${isUncalled ? '<span class="text-slate-400">-</span>' : escapeHtml(dateCall)}</td>
          <td class="whitespace-nowrap"><span class="badge-status ${scallBadgeClass}">${scallText}</span></td>
          <td class="whitespace-nowrap">
            ${isUncalled ? '<span class="text-slate-400 text-xs">-</span>' : `<span class="badge-status ${bermBadgeClass}">${escapeHtml(berminat)}</span>`}
            ${orderBadgeHtml}
          </td>
          <td class="whitespace-nowrap text-slate-800">
            ${isUncalled ? '<span class="text-slate-400 text-xs">-</span>' : `<span class="font-medium">${escapeHtml(picAntar)}</span>` + tandemBadgeHtml}
          </td>
          <td class="text-slate-500 text-xs whitespace-nowrap">${escapeHtml(dateCompleted)}</td>
          <td class="text-slate-800 font-medium text-xs whitespace-nowrap">
            ${dateAct && dateAct !== '#N/A' ? '<span class="text-purple-700 font-bold bg-purple-50 px-2 py-0.5 rounded">' + escapeHtml(dateAct) + '</span>' : '<span class="text-slate-400">-</span>'}
          </td>
          <td class="text-center whitespace-nowrap">
            <button onclick="window.app.openDetail('${idn}')" class="inline-flex items-center gap-1 bg-slate-900 hover:bg-[#E60000] text-white text-[11px] font-semibold px-2.5 py-1.5 rounded-lg shadow-xs transition-all">
              <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"></path></svg>
              Detail &amp; Foto
            </button>
          </td>
        </tr>
      `;
    });

    tbody.innerHTML = rowsHtml;

    if (isShowingAll || totalPages <= 1) {
      buttonsEl.innerHTML = '<span class="text-xs text-slate-500 font-medium px-2.5 py-1 bg-white rounded border border-slate-200">Semua Baris Ditampilkan</span>';
    } else {
      let btnHtml = '';
      btnHtml += `<button ${state.currentPage === 1 ? 'disabled class="opacity-40 cursor-not-allowed"' : 'onclick="window.app.setPage(' + (state.currentPage - 1) + ')"'} class="px-2.5 py-1 rounded border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold">Prev</button>`;

      const startP = Math.max(1, state.currentPage - 2);
      const endP = Math.min(totalPages, startP + 4);

      for (let p = startP; p <= endP; p++) {
        const activeClass = p === state.currentPage ? 'bg-slate-900 text-white font-bold border-slate-900' : 'bg-white text-slate-700 hover:bg-slate-50 border-slate-200';
        btnHtml += `<button onclick="window.app.setPage(${p})" class="px-3 py-1 rounded border ${activeClass} text-xs">${p}</button>`;
      }

      btnHtml += `<button ${state.currentPage === totalPages ? 'disabled class="opacity-40 cursor-not-allowed"' : 'onclick="window.app.setPage(' + (state.currentPage + 1) + ')"'} class="px-2.5 py-1 rounded border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold">Next</button>`;
      buttonsEl.innerHTML = btnHtml;
    }
  }

  window.app.setPage = function (p) {
    state.currentPage = p;
    renderTableAndPagination();
  };

  // --- EVIDENCE GALLERY RENDERING ---
  function renderEvidenceGallery() {
    const grid = document.getElementById('evidenceGrid');
    const totalCountEl = document.getElementById('galleryTotalCount');
    const picData = getFilteredPICData();

    const items = [];

    // 1. Uploaded photos
    Object.keys(state.evidenceStore).forEach(key => {
      const ev = state.evidenceStore[key];
      const customer = state.allData.find(r => String(r['Idn Pelanggan']) === String(key) || String(r['No Telepon']) === String(key)) || {};
      
      if (state.currentPic !== 'all') {
        const pic = (ev.pic || customer['NOTE PIC ANTAR'] || customer['CALLED BY'] || '').toLowerCase();
        if (!pic.includes(state.currentPic.toLowerCase())) return;
      }

      items.push({
        idn: key,
        title: ev.title || 'Bukti Kunjungan / Aktivasi',
        cust: ev.cust || customer['NAMA PELANGGAN'] || ('Pelanggan #' + key),
        phone: ev.phone || customer['No Telepon'] || '-',
        loc: ev.loc || customer['ALAMAT'] || 'Area Pamekasan',
        pic: ev.pic || customer['NOTE PIC ANTAR'] || 'TIKA',
        date: ev.date || customer['DATE ACT'] || 'November 2026',
        url: ev.url,
        fileName: ev.fileName || (key + '.jpg'),
        hasRealPhoto: true
      });
    });

    // 2. Visits by selected PIC (so user can easily upload photos)
    const visitedRows = picData.filter(r => r['NOTE PIC ANTAR'] && String(r['NOTE PIC ANTAR']).toLowerCase().includes(state.currentPic.toLowerCase()));
    visitedRows.forEach(row => {
      const idn = String(row['Idn Pelanggan']);
      const alreadyIn = items.some(it => String(it.idn) === idn);
      if (!alreadyIn) {
        items.push({
          idn: idn,
          title: 'Kunjungan Lapangan Blue Device',
          cust: row['NAMA PELANGGAN'] || ('Pelanggan #' + idn),
          phone: row['No Telepon'] || '-',
          loc: row['ALAMAT'] || 'Area Pamekasan',
          pic: row['NOTE PIC ANTAR'] || 'TIKA',
          date: row['DATE ACT'] || row['DATE COMPLETED'] || 'November 2026',
          url: null,
          fileName: idn + '.jpg',
          hasRealPhoto: false
        });
      }
    });

    if (totalCountEl) totalCountEl.textContent = Object.keys(state.evidenceStore).length + ' Foto Diupload';

    if (items.length === 0) {
      grid.innerHTML = `
        <div class="col-span-full py-16 text-center text-slate-400">
          <svg class="w-12 h-12 mx-auto mb-3 text-slate-300" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"></path></svg>
          <p class="font-bold text-slate-600">Belum ada foto evidence</p>
          <p class="text-xs text-slate-400 mt-1">Gunakan tombol "Upload Foto Evidence Baru" di atas untuk memasukkan foto bukti lapangan</p>
        </div>
      `;
      return;
    }

    let cardsHtml = '';
    items.forEach(item => {
      const isTandem = isTandemVisit(item.pic);
      const thumbnailSrc = item.hasRealPhoto ? item.url : createEvidencePlaceholder(item.idn, item.cust, item.pic);

      cardsHtml += `
        <div class="evidence-card flex flex-col justify-between">
          <div class="relative group cursor-pointer" onclick="window.app.handleGalleryClick('${item.idn}', ${item.hasRealPhoto ? `'${item.url}'` : 'null'})">
            <img src="${thumbnailSrc}" alt="${escapeHtml(item.title)}" class="evidence-thumbnail">
            <div class="absolute inset-0 bg-slate-900/40 opacity-0 group-hover:opacity-100 flex items-center justify-center transition-all">
              <span class="bg-white/90 text-slate-900 font-bold text-xs px-2.5 py-1.5 rounded-lg shadow-sm flex items-center gap-1">
                ${item.hasRealPhoto ? 'Perbesar Foto' : 'Upload / Buka Detail'}
              </span>
            </div>
            <div class="absolute bottom-2 left-2 bg-slate-900/80 backdrop-blur-xs text-white text-[10px] font-mono px-2 py-0.5 rounded">
              ID: ${escapeHtml(item.idn)}
            </div>
          </div>

          <div class="p-3.5 flex flex-col justify-between flex-1">
            <div>
              <div class="flex items-center justify-between text-[11px] text-slate-400 mb-1">
                <span class="font-mono text-slate-600">${escapeHtml(item.phone)}</span>
                <span class="font-bold ${item.hasRealPhoto ? 'text-emerald-700 bg-emerald-50' : 'text-slate-600 bg-slate-100'} px-1.5 py-0.5 rounded text-[10px]">
                  ${item.hasRealPhoto ? 'TERPASANG' : 'BELUM ADA FOTO'}
                </span>
              </div>
              <h4 class="font-bold text-slate-900 text-xs line-clamp-1">${escapeHtml(item.cust)}</h4>
              <p class="text-[11px] text-slate-500 line-clamp-1 mt-0.5">📍 ${escapeHtml(item.loc)}</p>
            </div>

            <div class="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-xs">
              <div class="text-[11px]">
                <span class="text-slate-500">PIC:</span>
                <span class="font-bold text-slate-800">${escapeHtml(item.pic.split('&')[0].trim())}</span>
                ${isTandem ? '<span class="badge-tandem scale-90 ml-1">Tandem</span>' : ''}
              </div>

              <button onclick="window.app.openDetail('${item.idn}')" class="text-red-600 hover:text-red-700 font-bold text-xs">
                ${item.hasRealPhoto ? 'Detail' : '+ Upload'}
              </button>
            </div>
          </div>
        </div>
      `;
    });

    grid.innerHTML = cardsHtml;
  }

  window.app.handleGalleryClick = function (idn, url) {
    if (url) {
      const row = state.allData.find(r => String(r['Idn Pelanggan']) === String(idn));
      const caption = (row ? row['NAMA PELANGGAN'] : 'Pelanggan') + ' (ID: ' + idn + ')';
      openLightbox(url, caption);
    } else {
      window.app.openDetail(idn);
    }
  };

  // --- LEMBAR REKAPITULASI LAPORAN KINERJA (SESUAI ARAHAN ATASAN) ---
  function renderRekapLaporanTable() {
    const kpi = calculateKPIs();
    const tableBody = document.getElementById('laporanTableSummary');
    if (!tableBody) return;

    tableBody.innerHTML = `
      <tr class="hover:bg-slate-50">
        <td class="p-3">
          <div class="font-bold text-slate-800">1. Hasil Caring (Panggilan Telepon)</div>
          <div class="text-[11px] text-slate-500">Panggilan keluar ke database nomor kontak pelanggan Pamekasan</div>
        </td>
        <td class="p-3 text-center font-bold text-slate-900 text-sm">${kpi.caringTotal} Call</td>
        <td class="p-3 text-slate-700">
          <span class="badge-status badge-success">${kpi.diangkat} Diangkat</span>
          <span class="badge-status badge-neutral ml-1">${kpi.tdkDiangkat} Tidak Diangkat</span>
          <span class="badge-status badge-warning ml-1">${kpi.tdkSambung} Tidak Sambung</span>
        </td>
        <td class="p-3 text-right font-bold text-emerald-700 text-xs">${kpi.contactRate}% Tersambung</td>
      </tr>

      <tr class="hover:bg-slate-50">
        <td class="p-3">
          <div class="font-bold text-slate-800">2. Hasil Visit Lapangan (Kunjungan Rumah)</div>
          <div class="text-[11px] text-slate-500">Kunjungan fisik langsung ke alamat pelanggan oleh TIKA</div>
        </td>
        <td class="p-3 text-center font-bold text-slate-900 text-sm">${kpi.visitTotal} Kunjungan</td>
        <td class="p-3 text-slate-700">
          <span class="badge-mandiri">${kpi.visitMandiri} Visit Mandiri</span>
          <span class="badge-tandem ml-1">${kpi.visitTandem} Visit Tandem</span>
        </td>
        <td class="p-3 text-right font-bold text-sky-700 text-xs">Tercatat Lapangan</td>
      </tr>

      <tr class="hover:bg-slate-50">
        <td class="p-3">
          <div class="font-bold text-slate-800">3. Jumlah Blue Device yang Pelanggan Setuju</div>
          <div class="text-[11px] text-slate-500">Pelanggan berminat/setuju pasang perangkat kuota keluarga Blue Device</div>
        </td>
        <td class="p-3 text-center font-bold text-emerald-700 text-sm">${kpi.setujuTotal} Pelanggan</td>
        <td class="p-3 text-slate-700">
          <span class="text-xs font-semibold text-emerald-700">${kpi.tertarikCount} Respon Tertarik</span>
        </td>
        <td class="p-3 text-right font-bold text-emerald-700 text-xs">${kpi.convRate}% Rasio Persetujuan</td>
      </tr>

      <tr class="hover:bg-slate-50 bg-amber-50/40">
        <td class="p-3">
          <div class="font-bold text-amber-900">4. Tindak Lanjut: Pelanggan Perlu Follow-Up (Need FU)</div>
          <div class="text-[11px] text-amber-700">Pelanggan dalam status konfirmasi &amp; penjadwalan ulang kunjungan</div>
        </td>
        <td class="p-3 text-center font-bold text-amber-800 text-sm">${kpi.needFuCount} Pelanggan</td>
        <td class="p-3 text-amber-800 text-xs font-medium">Order Status: NEED FU (Proses Follow-Up)</td>
        <td class="p-3 text-right font-bold text-amber-800 text-xs">Antrian FU</td>
      </tr>

      <tr class="hover:bg-slate-50 bg-purple-50/50">
        <td class="p-3">
          <div class="font-bold text-purple-900">5. Realisasi Aktivasi Lapangan (Done Act)</div>
          <div class="text-[11px] text-purple-700">Paket kuota keluarga Blue Device yang telah aktif di aplikasi MyTelkomsel &amp; DigiPOS</div>
        </td>
        <td class="p-3 text-center font-bold text-purple-800 text-sm">${kpi.doneActCount} Perangkat</td>
        <td class="p-3 text-purple-800 text-xs font-medium">Order Status: DONE ACT (100% Selesai)</td>
        <td class="p-3 text-right font-bold text-purple-800 text-xs">Aktif di MyTelkomsel &amp; DigiPOS</td>
      </tr>
    `;
  }

  // --- MODAL 1: 23-COLUMN CUSTOMER DETAIL ---
  window.app.openDetail = function (idn) {
    const row = state.allData.find(r => String(r['Idn Pelanggan']) === String(idn));
    if (!row) return;

    state.selectedRecord = row;

    const modal = document.getElementById('customerDetailModal');
    const nama = row['NAMA PELANGGAN'] || 'Pelanggan #' + String(idn).slice(-4);
    const telp = String(row['No Telepon'] || '-');

    document.getElementById('modalCustomerName').textContent = nama;
    document.getElementById('modalCustomerIdnPhone').textContent = 'IDN: ' + idn + '  |  No Telepon: ' + telp;
    document.getElementById('modalCustomerPriority').textContent = row['Prioritas'] || 'P1 - High';

    // Group 1
    document.getElementById('col_IdnPelanggan').textContent = idn;
    document.getElementById('col_PaketInternet').textContent = row['Paket Internet'] || '-';
    document.getElementById('col_ODP').textContent = row['ODP'] || '-';
    document.getElementById('col_Prioritas').textContent = row['Prioritas'] || '-';

    // Group 2
    document.getElementById('col_NoTelepon').textContent = telp;
    document.getElementById('col_Email').textContent = row['E-mail'] || '-';
    document.getElementById('col_CalledBy').textContent = row['CALLED BY'] || '-';
    document.getElementById('col_DateCall').textContent = row['DATE CALL'] || '-';

    const scall = String(row['STATUS CALL'] || '-');
    const scallEl = document.getElementById('col_StatusCall');
    scallEl.textContent = scall;
    scallEl.className = 'badge-status ' + (scall.toUpperCase().includes('DIANGKAT') && !scall.toUpperCase().includes('TIDAK') ? 'badge-success' : (scall.toUpperCase().includes('TIDAK DIANGKAT') ? 'badge-danger' : 'badge-warning'));

    const berm = String(row['BERMINAT?'] || '-');
    const bermEl = document.getElementById('col_Berminat');
    bermEl.textContent = berm;
    bermEl.className = 'badge-status ' + (berm.toUpperCase().includes('TERTARIK') && !berm.toUpperCase().includes('TIDAK') ? 'badge-success' : (berm.toUpperCase().includes('RAGU') ? 'badge-warning' : 'badge-neutral'));

    // Group 3
    document.getElementById('col_PackageKuota').textContent = row['Package Kuota Keluarga BTL'] || '-';
    document.getElementById('col_NomorPerdana').textContent = row['Nomor Perdana'] || '-';
    document.getElementById('col_StatusOrder').textContent = row['STATUS ORDER'] || '-';
    document.getElementById('col_StatusMO').textContent = row['STATUS MO'] || '-';
    document.getElementById('col_REV').textContent = row['REV'] || '-';
    document.getElementById('col_DateAct').textContent = row['DATE ACT'] || '-';

    // Group 4
    const picAntar = String(row['NOTE PIC ANTAR'] || '-');
    document.getElementById('col_NotePicAntar').textContent = picAntar;
    document.getElementById('col_DateCompleted').textContent = row['DATE COMPLETED'] || '-';
    document.getElementById('col_Alamat').textContent = row['ALAMAT'] || 'Area Telkomsel Pamekasan';

    const tandemBadge = document.getElementById('modalTandemBadge');
    if (isTandemVisit(picAntar)) {
      tandemBadge.classList.remove('hidden');
    } else {
      tandemBadge.classList.add('hidden');
    }

    const lat = row['Latitude'];
    const lng = row['Longitude'];
    const shareloc = row['SHARELOC'];
    const coordsEl = document.getElementById('col_Coords');
    coordsEl.textContent = 'Lat: ' + (lat || '-') + ', Long: ' + (lng || '-');

    const mapsBtn = document.getElementById('btnOpenGoogleMaps');
    if (lat && lng && lat !== '#N/A' && lng !== '#N/A') {
      mapsBtn.href = `https://www.google.com/maps/search/?api=1&query=${lat},${lng}`;
    } else if (shareloc && shareloc.startsWith('http')) {
      mapsBtn.href = shareloc;
    } else {
      mapsBtn.href = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent((row['ALAMAT'] || '') + ' Pamekasan')}`;
    }

    const callBtn = document.getElementById('btnCallPhone');
    if (telp && telp !== '-') {
      let cleanPhone = telp.replace(/[^0-9]/g, '');
      if (cleanPhone.startsWith('0')) cleanPhone = '62' + cleanPhone.slice(1);
      callBtn.href = 'https://wa.me/' + cleanPhone;
      callBtn.classList.remove('hidden');
    } else {
      callBtn.classList.add('hidden');
    }

    // Photo Preview
    const evidence = state.evidenceStore[idn] || state.evidenceStore[telp] || null;
    const photoImg = document.getElementById('modalEvidenceImage');
    const photoFileName = document.getElementById('modalEvidenceFileName');
    const photoTimestamp = document.getElementById('modalEvidenceTimestamp');
    const driveLink = document.getElementById('modalDriveLink');
    const statusBadge = document.getElementById('evidenceVerificationStatus');

    if (evidence) {
      photoImg.src = evidence.url;
      photoFileName.textContent = evidence.fileName || (idn + '.jpg');
      photoTimestamp.textContent = evidence.date || 'Tersimpan';
      statusBadge.textContent = 'FOTO TERPASANG';
      statusBadge.className = 'badge-status badge-success text-[10px]';

      if (evidence.driveUrl) {
        driveLink.href = evidence.driveUrl;
        driveLink.classList.remove('hidden');
      } else {
        driveLink.classList.add('hidden');
      }
    } else {
      photoImg.src = createEvidencePlaceholder(idn, nama, picAntar);
      photoFileName.textContent = 'Belum ada foto diupload';
      photoTimestamp.textContent = '-';
      statusBadge.textContent = 'BELUM ADA BUKTI FOTO';
      statusBadge.className = 'badge-status badge-warning text-[10px]';
      driveLink.classList.add('hidden');
    }

    modal.classList.remove('hidden');
  };

  function closeCustomerModal() {
    document.getElementById('customerDetailModal').classList.add('hidden');
    state.selectedRecord = null;
  }

  function handleSingleModalPhotoUpload(e) {
    if (!state.selectedRecord || !e.target.files || !e.target.files[0]) return;
    const file = e.target.files[0];
    const idn = String(state.selectedRecord['Idn Pelanggan']);

    const reader = new FileReader();
    reader.onload = function (evt) {
      const dataUrl = evt.target.result;
      state.evidenceStore[idn] = {
        url: dataUrl,
        title: 'Bukti Kunjungan Lapangan',
        cust: state.selectedRecord['NAMA PELANGGAN'] || ('Pelanggan #' + idn),
        phone: state.selectedRecord['No Telepon'] || '-',
        loc: state.selectedRecord['ALAMAT'] || 'Area Pamekasan',
        pic: state.selectedRecord['NOTE PIC ANTAR'] || 'TIKA',
        date: new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) + ' WIB',
        fileName: file.name
      };

      saveEvidenceStore();

      document.getElementById('modalEvidenceImage').src = dataUrl;
      document.getElementById('modalEvidenceFileName').textContent = file.name;
      document.getElementById('modalEvidenceTimestamp').textContent = 'Baru diupload';
      document.getElementById('evidenceVerificationStatus').textContent = 'FOTO TERPASANG';
      document.getElementById('evidenceVerificationStatus').className = 'badge-status badge-success text-[10px]';

      renderApp();
    };
    reader.readAsDataURL(file);
  }

  // --- BULK UPLOAD FOTO EVIDENCE ---
  function openBulkUploadModal() {
    state.bulkPendingFiles = [];
    document.getElementById('bulkFilesContainer').classList.add('hidden');
    document.getElementById('bulkFileInput').value = '';
    document.getElementById('bulkUploadModal').classList.remove('hidden');
  }

  function closeBulkUploadModal() {
    document.getElementById('bulkUploadModal').classList.add('hidden');
    state.bulkPendingFiles = [];
  }

  function handleBulkFilesSelect(e) {
    if (e.target.files && e.target.files.length) {
      handleBulkFiles(Array.from(e.target.files));
    }
  }

  function handleBulkFiles(files) {
    state.bulkPendingFiles = [];
    const listEl = document.getElementById('bulkFilesList');
    const container = document.getElementById('bulkFilesContainer');
    const countEl = document.getElementById('bulkFileCount');
    const matchedCountEl = document.getElementById('bulkMatchedCount');

    let matchedCount = 0;

    files.forEach(file => {
      const nameWithoutExt = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
      const cleanDigits = nameWithoutExt.replace(/[^0-9]/g, '');

      let matchedRow = null;
      if (cleanDigits) {
        matchedRow = state.allData.find(r => {
          const idn = String(r['Idn Pelanggan'] || '');
          const telp = String(r['No Telepon'] || '').replace(/[^0-9]/g, '');
          return idn.includes(cleanDigits) || cleanDigits.includes(idn) || (telp && (telp.includes(cleanDigits) || cleanDigits.includes(telp)));
        });
      }

      if (!matchedRow && nameWithoutExt.length > 3) {
        matchedRow = state.allData.find(r => {
          const cName = String(r['NAMA PELANGGAN'] || '').toLowerCase();
          return cName && (cName.includes(nameWithoutExt.toLowerCase()) || nameWithoutExt.toLowerCase().includes(cName));
        });
      }

      if (matchedRow) matchedCount++;

      state.bulkPendingFiles.push({
        file: file,
        matchedRow: matchedRow,
        manualIdn: matchedRow ? String(matchedRow['Idn Pelanggan']) : ''
      });
    });

    countEl.textContent = files.length;
    matchedCountEl.textContent = matchedCount + ' Cocok Otomatis';

    let html = '';
    state.bulkPendingFiles.forEach((item, idx) => {
      const isMatched = Boolean(item.matchedRow);
      const custName = item.matchedRow ? (item.matchedRow['NAMA PELANGGAN'] || 'Pelanggan #' + item.matchedRow['Idn Pelanggan']) : 'Belum Cocok';
      const idnVal = item.matchedRow ? item.matchedRow['Idn Pelanggan'] : '';

      html += `
        <div class="py-2 flex items-center justify-between gap-3 text-xs">
          <div class="flex items-center gap-2 truncate">
            <span class="w-5 h-5 rounded-full flex items-center justify-center shrink-0 ${isMatched ? 'bg-emerald-100 text-emerald-700 font-bold' : 'bg-amber-100 text-amber-700'}">
              ${isMatched ? '✓' : '?'}
            </span>
            <div class="truncate">
              <div class="font-semibold text-slate-800 truncate">${escapeHtml(item.file.name)}</div>
              <div class="text-[10px] text-slate-500">${(item.file.size / 1024).toFixed(0)} KB</div>
            </div>
          </div>

          <div class="flex items-center gap-2 shrink-0">
            ${isMatched ? `
              <span class="text-emerald-700 font-bold bg-emerald-50 px-2 py-0.5 rounded text-[11px] truncate max-w-[180px]">
                ${escapeHtml(custName)} (${idnVal})
              </span>
            ` : `
              <select onchange="window.app.setManualMatch(${idx}, this.value)" class="bg-white border border-slate-300 rounded px-2 py-1 text-xs focus:ring-red-500 max-w-[200px]">
                <option value="">Pasangkan ke Pelanggan...</option>
                ${state.allData.slice(0, 50).map(r => `<option value="${r['Idn Pelanggan']}">${escapeHtml(r['NAMA PELANGGAN'] || r['Idn Pelanggan'])} (${r['Idn Pelanggan']})</option>`).join('')}
              </select>
            `}
          </div>
        </div>
      `;
    });

    listEl.innerHTML = html;
    container.classList.remove('hidden');
  }

  window.app.setManualMatch = function (idx, idn) {
    if (state.bulkPendingFiles[idx]) {
      state.bulkPendingFiles[idx].manualIdn = idn;
      state.bulkPendingFiles[idx].matchedRow = state.allData.find(r => String(r['Idn Pelanggan']) === String(idn));
    }
  };

  function saveBulkUploadToGallery() {
    if (!state.bulkPendingFiles.length) {
      alert('Silakan pilih foto terlebih dahulu.');
      return;
    }

    let processed = 0;
    const total = state.bulkPendingFiles.length;

    state.bulkPendingFiles.forEach(item => {
      const idn = item.manualIdn || (item.matchedRow ? String(item.matchedRow['Idn Pelanggan']) : null);
      if (!idn) {
        processed++;
        if (processed === total) finalizeBulkUpload();
        return;
      }

      const reader = new FileReader();
      reader.onload = function (e) {
        const dataUrl = e.target.result;
        const cust = item.matchedRow ? item.matchedRow['NAMA PELANGGAN'] : ('Pelanggan #' + idn);
        const phone = item.matchedRow ? item.matchedRow['No Telepon'] : '-';
        const loc = item.matchedRow ? item.matchedRow['ALAMAT'] : 'Area Pamekasan';
        const pic = item.matchedRow ? (item.matchedRow['NOTE PIC ANTAR'] || item.matchedRow['CALLED BY']) : 'TIKA';

        state.evidenceStore[idn] = {
          url: dataUrl,
          title: 'Bukti Kunjungan / Aktivasi',
          cust: cust || ('Pelanggan #' + idn),
          phone: phone || '-',
          loc: loc || 'Area Pamekasan',
          pic: pic || 'TIKA',
          date: new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) + ' WIB',
          fileName: item.file.name
        };

        processed++;
        if (processed === total) finalizeBulkUpload();
      };
      reader.readAsDataURL(item.file);
    });
  }

  function finalizeBulkUpload() {
    saveEvidenceStore();
    closeBulkUploadModal();
    renderApp();
    switchMainTab('tabGallery');
    alert('Foto evidence berhasil disimpan!');
  }

  // --- LIGHTBOX PREVIEW ---
  function openLightbox(url, caption) {
    document.getElementById('lightboxImage').src = url;
    document.getElementById('lightboxCaption').textContent = caption || 'Bukti Foto Kunjungan';
    document.getElementById('lightboxModal').classList.remove('hidden');
  }

  function closeLightbox() {
    document.getElementById('lightboxModal').classList.add('hidden');
  }

  // --- EXCEL EXPORT ---
  function handleExcelExport() {
    if (typeof XLSX === 'undefined') {
      alert('Library SheetJS tidak tersedia.');
      return;
    }

    const list = getTableDisplayData();
    if (!list.length) {
      alert('Tidak ada data untuk diekspor.');
      return;
    }

    const ws = XLSX.utils.json_to_sheet(list);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Rekap Blue Device');

    const fileName = 'Laporan_Rekap_Blue_Device_' + (state.currentPic.toUpperCase()) + '_Closing_30Nov2026.xlsx';
    XLSX.writeFile(wb, fileName);
  }

  // --- UTILITIES ---
  function escapeHtml(str) {
    if (!str && str !== 0) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function createEvidencePlaceholder(idn, name, pic) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 450" width="600" height="450">
      <rect width="600" height="450" fill="#0F172A"/>
      <rect x="20" y="20" width="560" height="410" rx="12" fill="#1E293B" stroke="#334155" stroke-width="2"/>
      <rect x="20" y="20" width="560" height="45" rx="12" fill="#E60000"/>
      <rect x="20" y="55" width="560" height="10" fill="#E60000"/>
      <text x="40" y="48" font-family="sans-serif" font-size="14" font-weight="bold" fill="#FFFFFF">TELKOMSEL — BLUE DEVICE EVIDENCE</text>
      <circle cx="300" cy="180" r="45" fill="#334155"/>
      <path d="M285 180 L295 190 L315 170" stroke="#4ADE80" stroke-width="4" fill="none" stroke-linecap="round"/>
      <text x="300" y="255" font-family="sans-serif" font-size="18" font-weight="bold" fill="#F8FAFC" text-anchor="middle">${idn}</text>
      <text x="300" y="280" font-family="sans-serif" font-size="13" fill="#94A3B8" text-anchor="middle">${name}</text>
      <rect x="80" y="320" width="440" height="70" rx="8" fill="#0F172A" stroke="#334155"/>
      <text x="100" y="345" font-family="sans-serif" font-size="11" fill="#94A3B8">PIC KUNJUNGAN: <tspan fill="#FFFFFF" font-weight="bold">${pic || 'TIKA'}</tspan></text>
      <text x="100" y="370" font-family="monospace" font-size="11" fill="#22C55E">● TITIK LOKASI TELKOMSEL AREA PAMEKASAN</text>
    </svg>`;
    return 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg)));
  }

})();
