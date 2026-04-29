// ================================================================
// tools 共通スクリプト（inspection / payment / property / schedule）
// ハンバーガー → サイドドロワーの開閉
// ================================================================
// ===== プルトゥリフレッシュ（全ツール共通） =====
document.addEventListener('DOMContentLoaded', function() {
    const ind = document.createElement('div');
    ind.id = 'tool-ptr';
    ind.innerHTML =
        '<svg class="tool-ptr-arrow" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M5 12l7 7 7-7"/></svg>' +
        '<div class="tool-ptr-ring"></div>';
    document.body.appendChild(ind);

    let startY = 0, distance = 0, pulling = false, busy = false;
    const THRESHOLD = 68, MAX_PULL = 100;

    function setY(d) {
        const ty = -52 + (Math.min(d, MAX_PULL) / MAX_PULL) * 72;
        ind.style.transform = 'translateX(-50%) translateY(' + ty + 'px)';
    }
    function reset() {
        ind.style.transform = 'translateX(-50%) translateY(-52px)';
        ind.classList.remove('ptr-visible', 'ptr-ready', 'ptr-loading');
        distance = 0; pulling = false;
    }

    document.addEventListener('touchstart', function(e) {
        if (busy || window.scrollY > 0 || typeof window._ptrRefresh !== 'function') return;
        startY = e.touches[0].clientY; pulling = true;
    }, { passive: true });

    document.addEventListener('touchmove', function(e) {
        if (!pulling) return;
        if (window.scrollY > 0) { pulling = false; return; }
        distance = e.touches[0].clientY - startY;
        if (distance <= 0) return;
        setY(distance);
        ind.classList.add('ptr-visible');
        ind.classList.toggle('ptr-ready', distance >= THRESHOLD);
    }, { passive: true });

    document.addEventListener('touchend', function() {
        if (!pulling || distance <= 0) { reset(); return; }
        if (distance >= THRESHOLD) {
            busy = true;
            ind.classList.remove('ptr-ready');
            ind.classList.add('ptr-loading');
            ind.style.transform = 'translateX(-50%) translateY(20px)';
            if (typeof window._ptrRefresh === 'function') window._ptrRefresh();
            setTimeout(function() { busy = false; reset(); }, 2500);
        } else {
            reset();
        }
    }, { passive: true });
});

function openToolDrawer() {
  document.getElementById('toolDrawer').classList.add('open');
  document.getElementById('toolDrawerOverlay').classList.add('open');
}
function closeToolDrawer() {
  document.getElementById('toolDrawer').classList.remove('open');
  document.getElementById('toolDrawerOverlay').classList.remove('open');
}
