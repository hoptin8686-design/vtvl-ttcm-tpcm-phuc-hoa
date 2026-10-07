/* ==============================================================================
   app.js – Đề Án Vị Trí Việc Làm Viên Chức – THPT Phục Hòa
   Căn cứ Nghị định số 232/2026/NĐ-CP & Phụ lục I Danh mục VTVL của Trường
   Hỗ trợ tương tác 35 Biên chế: 07 Quản lý · 22 GV 13 Môn học · 06 Hỗ trợ
   ============================================================================== */

document.addEventListener('DOMContentLoaded', () => {
  // ---- Scroll-top button ----
  const scrollBtn = document.getElementById('scrollTopBtn');
  window.addEventListener('scroll', () => {
    if (window.scrollY > 400) {
      scrollBtn.classList.add('visible');
    } else {
      scrollBtn.classList.remove('visible');
    }
  });

  // ---- Intersection Observer – animate-in elements ----
  const observerOptions = {
    threshold: 0.08,
    rootMargin: '0px 0px -40px 0px'
  };

  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, observerOptions);

  const animTargets = document.querySelectorAll(
    '.block-box, .subject-card, .info-card, .task-item, .product-card, .kpi-card, .stat-card, .comp-table-wrap, .category-header, .master-table-card'
  );
  animTargets.forEach((el, i) => {
    el.classList.add('animate-in');
    el.style.transitionDelay = `${Math.min(i * 0.02, 0.2)}s`;
    observer.observe(el);
  });

  // ---- Toast Message Helper ----
  const toastMsg = document.getElementById('toastMsg');
  const toastText = document.getElementById('toastText');
  function showToast(text) {
    if (!toastMsg) return;
    toastText.textContent = text;
    toastMsg.classList.add('show');
    setTimeout(() => {
      toastMsg.classList.remove('show');
    }, 3500);
  }

  // Hook download buttons
  const downloadBtns = document.querySelectorAll('#btn-excel-hero, #btn-download-excel-banner, #floatingExcelBtn');
  downloadBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      showToast('✓ Đang tải xuống trọn bộ 07 Sheet Đề án 35 Biên chế (.xlsx)...');
    });
  });

  // ---- Category Tabs Switcher / Filter ----
  const roleTabBtns = document.querySelectorAll('.role-tab-btn[data-filter]');
  const catBlocks = document.querySelectorAll('.category-block');

  roleTabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      roleTabBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const filter = btn.getAttribute('data-filter');

      if (filter === 'all') {
        catBlocks.forEach(blk => {
          blk.style.display = 'block';
        });
        showToast('Hiển thị toàn cảnh Đề án: Đủ 35 Biên chế (03 Danh mục VTVL)');
      } else {
        catBlocks.forEach(blk => {
          if (blk.getAttribute('data-cat') === filter) {
            blk.style.display = 'block';
            blk.scrollIntoView({ behavior: 'smooth', block: 'start' });
          } else {
            blk.style.display = 'none';
          }
        });
        const catNames = {
          'master': 'Bảng tổng hợp Phụ lục I: 35 Biên chế người làm việc',
          'dm1': 'Danh mục 1: Vị trí Quản lý (07 người: Hiệu trưởng, 2 Phó HT, 2 TTCM, 2 TPCM)',
          'dm2': 'Danh mục 2: Vị trí Chuyên môn (22 GV: Đầy đủ 13 môn học CT 2018)',
          'dm3': 'Danh mục 3: Vị trí Hỗ trợ (06 người: Kế toán, Văn thư, TB-TN, Giáo vụ, Y tế, Thủ quỹ)'
        };
        showToast(`Đang lọc: ${catNames[filter] || filter}`);
      }
    });
  });

  // ---- 13 Subjects Sub-Filter ----
  const subjFilterBtns = document.querySelectorAll('.subj-filter-btn[data-sub]');
  const subjectCards = document.querySelectorAll('.subject-card[data-sub-key]');

  subjFilterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      subjFilterBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const subKey = btn.getAttribute('data-sub');

      if (subKey === 'all') {
        subjectCards.forEach(card => card.style.display = 'flex');
        showToast('Hiển thị đầy đủ 13 môn học giảng dạy (22 GV)');
      } else {
        subjectCards.forEach(card => {
          if (card.getAttribute('data-sub-key') === subKey) {
            card.style.display = 'flex';
            card.scrollIntoView({ behavior: 'smooth', block: 'center' });
          } else {
            card.style.display = 'none';
          }
        });
        showToast(`Đang hiển thị môn: ${btn.textContent.trim()}`);
      }
    });
  });

  // ---- Live Search Filter ----
  const searchInput = document.getElementById('liveSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      const query = e.target.value.trim().toLowerCase();
      const taskItems = document.querySelectorAll('.task-item');
      const kpiCards = document.querySelectorAll('.kpi-card');
      const blockBoxes = document.querySelectorAll('.block-box');
      const subCards = document.querySelectorAll('.subject-card');
      const tableRows = document.querySelectorAll('.master-table tbody tr:not(.group-header-row):not(.group-total-row)');

      if (!query) {
        catBlocks.forEach(blk => blk.style.display = 'block');
        taskItems.forEach(item => {
          item.style.display = 'flex';
          item.style.backgroundColor = '';
        });
        kpiCards.forEach(card => card.style.display = 'block');
        blockBoxes.forEach(box => box.style.display = 'block');
        subCards.forEach(card => card.style.display = 'flex');
        tableRows.forEach(row => row.style.display = '');
        return;
      }

      // Show all category blocks while searching
      catBlocks.forEach(blk => blk.style.display = 'block');
      roleTabBtns.forEach(b => b.classList.remove('active'));
      document.querySelector('.role-tab-btn[data-filter="all"]')?.classList.add('active');

      // Filter subject cards
      subCards.forEach(card => {
        const text = card.textContent.toLowerCase();
        if (text.includes(query)) {
          card.style.display = 'flex';
        } else {
          card.style.display = 'none';
        }
      });

      // Filter block boxes
      blockBoxes.forEach(box => {
        const text = box.textContent.toLowerCase();
        if (text.includes(query)) {
          box.style.display = 'block';
        } else {
          box.style.display = 'none';
        }
      });

      // Filter table rows
      tableRows.forEach(row => {
        const text = row.textContent.toLowerCase();
        if (text.includes(query)) {
          row.style.display = '';
          row.style.backgroundColor = 'rgba(254, 240, 138, 0.35)';
        } else {
          row.style.display = 'none';
          row.style.backgroundColor = '';
        }
      });

      // Filter KPI cards
      kpiCards.forEach(card => {
        const text = card.textContent.toLowerCase();
        if (text.includes(query)) {
          card.style.display = 'block';
        } else {
          card.style.display = 'none';
        }
      });
    });
  }

  // ---- Table row hover effects ----
  document.querySelectorAll('.comp-table tbody tr, .master-table tbody tr').forEach(row => {
    row.addEventListener('mouseenter', () => {
      row.style.background = 'rgba(245, 158, 11, 0.08)';
    });
    row.addEventListener('mouseleave', () => {
      row.style.background = '';
    });
  });

  console.log(
    '%c🏛️ TRƯỜNG THPT PHỤC HÒA – CAO BẰNG\n' +
    '%cĐỀ ÁN VỊ TRÍ VIỆC LÀM VIÊN CHỨC (35 BIÊN CHẾ - 03 DANH MỤC)\n' +
    '%cCăn cứ Nghị định số 232/2026/NĐ-CP & Phụ lục I C3 Phục Hòa\n' +
    '%cTự động đẩy lên GitHub & Vercel | Năm học 2026–2027',
    'color: #f59e0b; font-size: 16px; font-weight: bold;',
    'color: #3b82f6; font-size: 14px; font-weight: bold;',
    'color: #10b981; font-size: 12px;',
    'color: #94a3b8; font-size: 11px;'
  );
});
