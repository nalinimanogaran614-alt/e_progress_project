(function () {
  'use strict';

  function updateDate() {
    const el = document.getElementById('epDateText');
    if (!el) return;
    const now = new Date();
    const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    el.textContent = days[now.getDay()] + ' • ' + String(now.getDate()).padStart(2,'0') + ' ' + months[now.getMonth()] + ' ' + now.getFullYear();
  }

  function passwordEyes() {
    document.querySelectorAll('input[type="password"]').forEach(function (input) {
      if (input.dataset.eyeReady === '1') return;
      input.dataset.eyeReady = '1';
      const wrap = input.parentElement;
      if (!wrap || wrap.dataset.passwordWrap === '1') return;
      wrap.dataset.passwordWrap = '1';
      wrap.classList.add('ep-password-wrap');
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'ep-password-eye';
      btn.setAttribute('aria-label','Show password');
      btn.innerHTML = '<i class="bi bi-eye"></i>';
      btn.addEventListener('click', function () {
        const showing = input.type === 'text';
        input.type = showing ? 'password' : 'text';
        btn.innerHTML = showing ? '<i class="bi bi-eye"></i>' : '<i class="bi bi-eye-slash"></i>';
        btn.setAttribute('aria-label', showing ? 'Show password' : 'Hide password');
      });
      wrap.appendChild(btn);
    });
  }

  function keyboardNavigation() {
    document.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter') return;
      const target = e.target;
      if (!target || !['INPUT','SELECT','TEXTAREA'].includes(target.tagName)) return;
      if (target.type === 'submit' || target.dataset.enterSubmit === '1') return;
      const form = target.form;
      if (!form) return;
      const fields = Array.from(form.querySelectorAll('input:not([type="hidden"]):not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled])'))
        .filter(el => el.offsetParent !== null && el.type !== 'button');
      const idx = fields.indexOf(target);
      if (idx < 0) return;
      e.preventDefault();
      const next = fields[idx + 1];
      if (next) {
        next.focus();
        if (typeof next.select === 'function' && (next.tagName === 'INPUT' || next.tagName === 'TEXTAREA')) next.select();
      } else {
        const submit = form.querySelector('button[type="submit"],input[type="submit"]');
        if (submit) submit.click();
      }
    }, true);
  }

  function dateDefaults() {
    const now = new Date();
    const iso = now.toISOString().slice(0,10);
    document.querySelectorAll('[data-current-date]').forEach(el => { if (!el.value) el.value = iso; });
    document.querySelectorAll('input[type="date"]').forEach(el => { if (el.dataset.minToday === '1') el.min = iso; });
  }

  function enableDragScroll(el) {
    if (el.dataset.dragScroll === '1') return;
    el.dataset.dragScroll = '1';
    let down=false,startX=0,left=0;
    el.addEventListener('pointerdown',e=>{if(e.pointerType==='mouse'&&e.button!==0)return;down=true;startX=e.clientX;left=el.scrollLeft;el.classList.add('dragging');el.setPointerCapture?.(e.pointerId)});
    el.addEventListener('pointermove',e=>{if(!down)return;el.scrollLeft=left-(e.clientX-startX)});
    const stop=()=>{down=false;el.classList.remove('dragging')};
    el.addEventListener('pointerup',stop);el.addEventListener('pointercancel',stop);el.addEventListener('mouseleave',()=>{if(down)stop()});
  }

  updateDate();
  setInterval(updateDate,60000);
  passwordEyes();
  keyboardNavigation();
  dateDefaults();
  document.querySelectorAll('.faculty-table-scroll,.marks-table-wrap,.view-table-wrap,.hidden-scroll').forEach(enableDragScroll);
  new MutationObserver(passwordEyes).observe(document.body,{childList:true,subtree:true});
})();
