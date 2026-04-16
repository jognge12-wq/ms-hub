// ================================================================
// tools 共通スクリプト（inspection / payment / property / schedule）
// ハンバーガー → サイドドロワーの開閉
// ================================================================
function openToolDrawer() {
  document.getElementById('toolDrawer').classList.add('open');
  document.getElementById('toolDrawerOverlay').classList.add('open');
}
function closeToolDrawer() {
  document.getElementById('toolDrawer').classList.remove('open');
  document.getElementById('toolDrawerOverlay').classList.remove('open');
}
