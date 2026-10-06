/* ==============================================================================
   app.js – Đề Án Vị Trí Việc Làm Viên Chức Quản Lý – THPT Phục Hòa
   Căn cứ Nghị định số 232/2026/NĐ-CP & Thông tư 15/2026/TT-BGDĐT
   Hỗ trợ tương tác 04 Vị trí: Hiệu trưởng, Phó Hiệu trưởng, TTCM, TPCM
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
    '.block-box, .info-card, .task-item, .product-card, .kpi-card, .stat-card, .comp-table-wrap'
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
      showToast('✓ Đang tải xuống trọn bộ 07 Sheet Đề án VTVL (.xlsx)...');
    });
  });

  // ---- Role Tabs Switcher / Filter ----
  const roleTabBtns = document.querySelectorAll('.role-tab-btn[data-filter]');
  const roleSections = document.querySelectorAll('.role-section');

  roleTabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      roleTabBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const filter = btn.getAttribute('data-filter');

      if (filter === 'all') {
        roleSections.forEach(sec => {
          sec.style.display = 'block';
        });
        showToast('Hiển thị toàn cảnh Đề án 04 Vị trí Quản lý');
      } else {
        roleSections.forEach(sec => {
          if (sec.getAttribute('data-role') === filter) {
            sec.style.display = 'block';
            sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
          } else {
            sec.style.display = 'none';
          }
        });
        const roleNames = {
          'ht': '1. Hiệu trưởng (HT-THPT-01)',
          'pht': '2. Phó Hiệu trưởng (PHT-THPT-01)',
          'ttcm': '3. Tổ trưởng chuyên môn (TTCM-THPT-01)',
          'tpcm': '4. Tổ phó chuyên môn (TPCM-THPT-01)'
        };
        showToast(`Đang lọc chi tiết vị trí: ${roleNames[filter] || filter}`);
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

      if (!query) {
        taskItems.forEach(item => item.style.display = 'flex');
        kpiCards.forEach(card => card.style.display = 'block');
        return;
      }

      // If user is searching, show all sections so matches are visible
      roleSections.forEach(sec => sec.style.display = 'block');
      roleTabBtns.forEach(b => b.classList.remove('active'));
      document.querySelector('.role-tab-btn[data-filter="all"]')?.classList.add('active');

      taskItems.forEach(item => {
        const text = item.textContent.toLowerCase();
        if (text.includes(query)) {
          item.style.display = 'flex';
          item.style.backgroundColor = 'rgba(254, 240, 138, 0.2)';
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
    '%cĐỀ ÁN VỊ TRÍ VIỆC LÀM VIÊN CHỨC QUẢN LÝ (04 VỊ TRÍ)\n' +
    '%cCăn cứ Nghị định số 232/2026/NĐ-CP & Thông tư số 15/2026/TT-BGDĐT\n' +
    '%cTự động đẩy lên GitHub & Vercel | Năm học 2026–2027',
    'color: #f59e0b; font-size: 16px; font-weight: bold;',
    'color: #3b82f6; font-size: 14px; font-weight: bold;',
    'color: #10b981; font-size: 12px;',
    'color: #94a3b8; font-size: 11px;'
  );
});
