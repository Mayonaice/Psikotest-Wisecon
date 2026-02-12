(function(){
  const TYPE_IDS = {
    information: { overlay: 'dct-information-overlay', message: 'dct-information-message' },
    alert:       { overlay: 'dct-alert-overlay',       message: 'dct-alert-message' },
    failed:      { overlay: 'dct-failed-overlay',      message: 'dct-failed-message' },
    success:     { overlay: 'dct-success-overlay',     message: 'dct-success-message' },
    confirmation:{ overlay: 'dct-confirmation-overlay',message: 'dct-confirmation-message' }
  };

  const state = { information:{}, alert:{}, failed:{}, success:{}, confirmation:{} };

  function show(type, opts){
    const ids = TYPE_IDS[type];
    if(!ids) return;
    const overlay = document.getElementById(ids.overlay);
    const msgEl   = document.getElementById(ids.message);
    if(!overlay || !msgEl) return;
    const message = (opts && opts.message) || '';
    if (opts && opts.html) { msgEl.innerHTML = message; } else { msgEl.textContent = message; }
    overlay.style.display = 'flex';
    // Optional title override
    const titleEl = overlay.querySelector('.dct-title');
    if (titleEl && opts && typeof opts.title === 'string') {
      titleEl.textContent = opts.title;
    }

    // Optional button label overrides
    const okBtn = overlay.querySelector('[data-action="ok"]');
    if (okBtn && opts && typeof opts.okText === 'string') {
      okBtn.textContent = opts.okText;
    }
    const yesBtn = overlay.querySelector('[data-action="yes"]');
    if (yesBtn && opts && typeof opts.yesText === 'string') {
      yesBtn.textContent = opts.yesText;
    }
    const noBtn = overlay.querySelector('[data-action="no"]');
    if (noBtn && opts && typeof opts.noText === 'string') {
      noBtn.textContent = opts.noText;
    }

    state[type] = { onOk:opts && opts.onOk, onYes:opts && opts.onYes, onNo:opts && opts.onNo };
  }

  function hide(type){
    const ids = TYPE_IDS[type];
    if(!ids) return;
    const overlay = document.getElementById(ids.overlay);
    if(overlay) overlay.style.display = 'none';
    state[type] = {};
  }

  function attach(){
    Object.keys(TYPE_IDS).forEach(function(type){
      const ids = TYPE_IDS[type];
      const overlay = document.getElementById(ids.overlay);
      if(!overlay) return;
      const closeBtn = overlay.querySelector('.dct-close');
      if(closeBtn){ closeBtn.addEventListener('click', function(){ hide(type); }); }

      overlay.querySelectorAll('[data-action]').forEach(function(btn){
        btn.addEventListener('click', function(){
          const action = btn.getAttribute('data-action');
          const ctx = state[type] || {};
          hide(type);
          if(action === 'ok'   && typeof ctx.onOk  === 'function') ctx.onOk();
          if(action === 'yes'  && typeof ctx.onYes === 'function') ctx.onYes();
          if(action === 'no'   && typeof ctx.onNo  === 'function') ctx.onNo();
        });
      });
    });
  }

  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', attach);
  } else {
    attach();
  }

  window.DCT3Modals = {
    open: show,
    close: hide
  };
})();