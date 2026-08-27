/* =========================================================
   REGALO PARTY · puente con el backend de Google Apps Script
   ---------------------------------------------------------
   Cambia API_URL por la URL /exec de tu despliegue si algún
   día vuelves a desplegar el script.
   ========================================================= */
var API_URL = 'https://script.google.com/macros/s/AKfycbyIzpI5-PnW3lG1lYGwcZPnlzgs2lZ51tYNzO8ZrpLn2BbN-duhbKbKdR5gPy3xafBUkw/exec';

/* Lee ?evento=... de la URL (paneles de terceros) */
function paramURL(n){
  try { return new URLSearchParams(location.search).get(n) || ''; } catch(e){ return ''; }
}

/* --- JSONP: funciona siempre, aunque el navegador bloquee CORS --- */
function _jsonp(payload){
  return new Promise(function(resolve, reject){
    var name = '__rp_cb_' + Math.random().toString(36).slice(2);
    var s = document.createElement('script');
    var t = setTimeout(function(){ limpiar(); reject(new Error('timeout')); }, 30000);
    function limpiar(){
      clearTimeout(t);
      try { delete window[name]; } catch(e){ window[name] = undefined; }
      if (s.parentNode) s.parentNode.removeChild(s);
    }
    window[name] = function(r){ limpiar(); resolve(r); };
    var q = ['callback=' + name];
    Object.keys(payload).forEach(function(k){
      if (payload[k] === undefined || payload[k] === null) return;
      q.push(encodeURIComponent(k) + '=' + encodeURIComponent(payload[k]));
    });
    s.onerror = function(){ limpiar(); reject(new Error('red')); };
    s.src = API_URL + '?' + q.join('&');
    document.head.appendChild(s);
  });
}

/* --- POST (sin límite de tamaño). Si falla, caemos a JSONP --- */
function _post(payload){
  return fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
    redirect: 'follow'
  }).then(function(r){ return r.json(); });
}

/* API pública: api('guardar', {pin:..., d:{...}}) -> Promise */
function api(accion, params){
  var p = Object.assign({ api: accion }, params || {});
  if (p.d && typeof p.d === 'object') p.d = JSON.stringify(p.d);
  var esEscritura = ['guardar','aportar','manual','subirfoto','borrarfoto'].indexOf(accion) > -1;
  /* La foto viaja en base64: no cabe en una URL, asi que siempre por POST */
  if (accion === 'subirfoto') return _post(p);

  if (esEscritura) {
    return _post(p).catch(function(){ return _jsonp(p); });
  }
  return _jsonp(p).catch(function(){ return _post(p); });
}
