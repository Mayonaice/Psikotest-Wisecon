window.DCT3Modal = (function(){
  const overlay = () => document.getElementById('notif-overlay');
  const modal = () => document.getElementById('notif-modal');
  const titleEl = () => document.getElementById('notif-title');
  const msgEl = () => document.getElementById('notif-message');
  const iconEl = () => document.getElementById('notif-icon');

  function show(title, message, type){
    if (titleEl()) titleEl().textContent = title || 'Notifikasi';
    if (msgEl()) msgEl().textContent = message || '';
    if (iconEl()) {
      var base = document.baseURI || '/';
      if (!base.endsWith('/')) base = base + '/';
      var map = {
        info: base + 'dct3/icons/info.svg',
        error: base + 'dct3/icons/error.svg',
        success: base + 'dct3/icons/success.svg'
      };
      iconEl().src = map[type] || map.info;
    }
    if (overlay()) overlay().classList.remove('hidden');
    if (modal()) modal().classList.remove('hidden');
  }
  function hide(){
    if (overlay()) overlay().classList.add('hidden');
    if (modal()) modal().classList.add('hidden');
  }
  document.addEventListener('click', function(e){
    if (e.target && e.target.id === 'notif-close') hide();
    if (e.target && e.target.id === 'notif-overlay') hide();
  });
  return { show, hide };
})();
