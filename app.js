/* ==============================================================================
   app.js – Đề Án Vị Trí Việc Làm Viên Chức – THPT Phục Hòa
   Căn cứ Nghị định số 232/2026/NĐ-CP (Phụ lục I, II, III & IV)
   Hỗ trợ tương tác trọn bộ 03 Danh mục: Quản lý, Chuyên môn, Hỗ trợ
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
    '.block-box, .info-card, .task-item, .product-card, .kpi-card, .stat-card, .comp-table-wrap, .category-header'
  );
  animTargets.forEach((el, i) => {
    el.classList.add('animate-in');
    el.style.transitionDelay = `${Math.min(i * 0.03, 0.25)}s`;
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
      showToast('✓ Đang tải xuống trọn bộ 07 Sheet Đề án 03 Danh mục VTVL (.xlsx)...');
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
        showToast('Hiển thị toàn cảnh Đề án: Đủ 03 Danh mục VTVL');
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
          'dm1': 'Danh mục 1: Vị trí Quản lý (Hiệu trưởng, Phó HT, TTCM, TPCM)',
          'dm2': 'Danh mục 2: Vị trí Chuyên môn (Giáo viên, Thư viện, CNTT)',
          'dm3': 'Danh mục 3: Vị trí Hỗ trợ (Kế toán, Văn thư, Thiết bị TN, Giáo vụ, Y tế)'
        };
        showToast(`Đang lọc: ${catNames[filter] || filter}`);
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

      if (!query) {
        catBlocks.forEach(blk => blk.style.display = 'block');
        taskItems.forEach(item => {
          item.style.display = 'flex';
          item.style.backgroundColor = '';
        });
        kpiCards.forEach(card => card.style.display = 'block');
        blockBoxes.forEach(box => box.style.display = 'block');
        return;
      }

      // Show all category blocks while searching
      catBlocks.forEach(blk => blk.style.display = 'block');
      roleTabBtns.forEach(b => b.classList.remove('active'));
      document.querySelector('.role-tab-btn[data-filter="all"]')?.classList.add('active');

      taskItems.forEach(item => {
        const text = item.textContent.toLowerCase();
        if (text.includes(query)) {
          item.style.display = 'flex';
          item.style.backgroundColor = 'rgba(254, 240, 138, 0.25)';
        } else {
          item.style.display = 'none';
          item.style.backgroundColor = '';
        }
      });

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

  // ---- Competency table row hover glow ----
  document.querySelectorAll('.comp-table tbody tr').forEach(row => {
    row.addEventListener('mouseenter', () => {
      row.style.background = 'rgba(245, 158, 11, 0.08)';
    });
    row.addEventListener('mouseleave', () => {
      row.style.background = '';
    });
  });

  console.log(
    '%c🏛️ TRƯỜNG THPT PHỤC HÒA – CAO BẰNG\n' +
    '%cĐỀ ÁN VỊ TRÍ VIỆC LÀM VIÊN CHỨC (03 DANH MỤC TOÀN DIỆN)\n' +
    '%cCăn cứ Nghị định số 232/2026/NĐ-CP & Thông tư số 15/2026/TT-BGDĐT\n' +
    '%cTự động đẩy lên GitHub & Vercel | Năm học 2026–2027',
    'color: #f59e0b; font-size: 16px; font-weight: bold;',
    'color: #3b82f6; font-size: 14px; font-weight: bold;',
    'color: #10b981; font-size: 12px;',
    'color: #94a3b8; font-size: 11px;'
  );
});
