(function () {
  var zone = document.getElementById('dropzone');
  var input = document.getElementById('file');
  var label = document.getElementById('file-name');
  if (!zone || !input) { return; }
  function show() {
    if (input.files && input.files.length) {
      label.textContent = input.files[0].name;
      label.hidden = false;
    } else {
      label.hidden = true;
    }
  }
  input.addEventListener('change', show);
  ['dragenter', 'dragover'].forEach(function (ev) {
    zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.remove('over'); });
  });
  zone.addEventListener('drop', function (e) {
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
      try { input.files = e.dataTransfer.files; } catch (err) { /* tarayıcı desteklemiyorsa seçim yapılır */ }
      show();
    }
  });
})();
