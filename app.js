/* ============================================
   app.js – VTVL TTCM & TPCM – THPT Phục Hòa
   ============================================ */

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

// Add animation class to sections and cards
document.addEventListener('DOMContentLoaded', () => {
  const animTargets = document.querySelectorAll(
    '.block-box, .info-card, .task-item, .product-card, .standard-row, .compare-col, .comp-table-wrap'
  );

  animTargets.forEach((el, i) => {
    el.classList.add('animate-in');
    el.style.transitionDelay = `${Math.min(i * 0.04, 0.3)}s`;
    observer.observe(el);
  });
});

// ---- Smooth active nav highlight ----
const sections = ['ttcm-section', 'tpcm-section', 'khung-nang-luc', 'so-sanh'];
const navButtons = {
  'ttcm-section':    document.getElementById('btn-ttcm'),
  'tpcm-section':    document.getElementById('btn-tpcm'),
  'khung-nang-luc':  document.getElementById('btn-khnl'),
};

const sectionObserver = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    const id = entry.target.id;
    const btn = navButtons[id];
    if (!btn) return;
    if (entry.isIntersecting) {
      btn.style.boxShadow = '0 0 0 3px rgba(255,255,255,0.4)';
    } else {
      btn.style.boxShadow = '';
    }
  });
}, { threshold: 0.3 });

sections.forEach(id => {
  const el = document.getElementById(id);
  if (el) sectionObserver.observe(el);
});

// ---- Competency table row hover glow ----
document.querySelectorAll('.comp-table tbody tr').forEach(row => {
  row.addEventListener('mouseenter', () => {
    row.style.background = 'rgba(245,158,11,0.06)';
  });
  row.addEventListener('mouseleave', () => {
    row.style.background = '';
  });
});

// ---- Task items - subtle interaction ----
document.querySelectorAll('.task-item').forEach(item => {
  item.addEventListener('mouseenter', () => {
    item.style.borderLeftColor = '#3b82f6';
    item.style.borderLeftWidth = '3px';
  });
  item.addEventListener('mouseleave', () => {
    item.style.borderLeftColor = '';
    item.style.borderLeftWidth = '';
  });
});

// ---- Print-friendly tip ----
console.log(
  '%c📋 THPT Phục Hòa – VTVL TTCM & TPCM\n' +
  '%cNghị định 232/2026/NĐ-CP | Năm học 2026–2027\n' +
  '%cĐể in tài liệu: Ctrl+P hoặc Cmd+P',
  'color: #f59e0b; font-size: 16px; font-weight: bold;',
  'color: #3b82f6; font-size: 12px;',
  'color: #94a3b8; font-size: 11px;'
);
