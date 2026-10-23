/* ============================================================
   PoliMusic - utilidades del panel de administracion
   ============================================================ */
(function () {
  'use strict';

  /* Confirmar antes de eliminar */
  document.addEventListener('submit', function (e) {
    var f = e.target;
    if (f && f.getAttribute && f.getAttribute('data-confirm')) {
      if (!window.confirm(f.getAttribute('data-confirm'))) {
        e.preventDefault();
      }
    }
  });

  /* Cerrar banners tipo chip con la tecla Escape */
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      Array.prototype.slice.call(document.querySelectorAll('.alert[data-dismiss="true"]')).forEach(function (el) {
        el.style.display = 'none';
      });
    }
  });
})();
