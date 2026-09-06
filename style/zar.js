/* Žiara: živé svetlo v pozadí úvodu. Jeden shader, žiadna knižnica, 0 cudzích domén.
 *
 * Prečo vlastný a nie unicorn.studio: ten sa načítava z cudzieho CDN, čo naša
 * CSP zakazuje, a na každej stránke tvrdíme, že sa nič neposiela tretím stranám.
 * Odmerané 6. 9. 2026 (ops/design/pohyb-dokaz.mjs): unicorn.studio beží na
 * siedmich WebGL plátnach plus prehrávané videá a mení 38 % plochy za snímku.
 * Toto je jedno plátno, jeden fragment shader, asi 5 kB, a mení sa tak pomaly,
 * že to čítanie neruší.
 *
 * Zodpovednosť voči návštevníkovi je tu dôležitejšia než efekt, preto:
 *   - `prefers-reduced-motion: reduce` vykreslí JEDEN snímok a slučka sa nespustí;
 *   - keď úvod odscrolluje z obrazu, slučka sa zastaví (IntersectionObserver);
 *   - keď je karta prehliadača skrytá, slučka sa zastaví (visibilitychange);
 *   - kreslíme najviac 30-krát za sekundu, nie 60, rozdiel nie je vidieť
 *     a spotreba je polovičná;
 *   - na úzkych obrazovkách kreslíme na polovičnom rozlíšení, lebo telefón
 *     má hustý displej a shader by zbytočne počítal štvornásobok bodov;
 *   - keď WebGL nie je (staré zariadenie, vypnutá akcelerácia, šetrič energie),
 *     plátno sa odstráni a ostane CSS prechod pod ním, ktorý vyzerá dobre sám.
 * Plátno je `aria-hidden` a `pointer-events:none`, takže do obsluhy nezasahuje.
 *
 * Použitie:
 *   <canvas class="zar-plocha" data-zar="stuhy" aria-hidden="true"></canvas>
 *   <script src="/style/zar.js" defer></script>
 * Varianty: stuhy (výrazná, domovská), vlny (pokojná, veľa textu),
 * prach (hĺbka, produkty), luc (takmer statická, právne texty).
 * Neznámy názov spadne na stuhy. Jedno plátno na stránku, len v úvode.
 */
(function () {
  'use strict';

  var plocha = document.querySelector('canvas.zar-plocha');
  if (!plocha) return;

  var tichy = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  var gl = null;
  try {
    gl = plocha.getContext('webgl2', { alpha: true, antialias: false, depth: false, stencil: false, powerPreference: 'low-power' })
      || plocha.getContext('webgl', { alpha: true, antialias: false, depth: false, stencil: false });
  } catch (e) { gl = null; }
  if (!gl) { plocha.remove(); return; }

  var VRCHOL = 'attribute vec2 p;void main(){gl_Position=vec4(p,0.,1.);}';

  /* Shader: dvakrát zohnutý šum (domain warping) namapovaný na našu žeravú
     paletu. Žiadny tvar, ktorý by sa dal pomenovať, len svetlo, ktoré sa
     preteká. Farby sú tie isté ako tokeny v paper.css. */
  /* Spolocny zaklad pre vsetky varianty: sum, fBm, rozptyl a nasa paleta.
     Rozptyl (dither) nie je ozdoba: bez neho su na tmavom prechode vidiet pruhy,
     lebo osem bitov na kanal nestaci na jemny prechod v tmavych tonoch. */
  var ZAKLAD = [
    'precision highp float;',
    'uniform vec2 rozmer;',
    'uniform float cas;',
    'float sum(vec2 v){ return fract(sin(dot(v, vec2(127.1, 311.7))) * 43758.5453123); }',
    'float hladky(vec2 v){',
    '  vec2 i = floor(v), f = fract(v);',
    '  vec2 u = f * f * (3.0 - 2.0 * f);',
    '  return mix(mix(sum(i), sum(i + vec2(1,0)), u.x),',
    '             mix(sum(i + vec2(0,1)), sum(i + vec2(1,1)), u.x), u.y);',
    '}',
    'float fbm(vec2 v){',
    '  float h = 0.0, a = 0.5;',
    '  for (int i = 0; i < 5; i++) { h += a * hladky(v); v = v * 2.03 + 17.3; a *= 0.5; }',
    '  return h;',
    '}',
    'float rozptyl(vec2 s){ return fract(sin(dot(s, vec2(12.9898, 78.233))) * 43758.5453) - 0.5; }',
    'const vec3 UHLIK  = vec3(0.42, 0.15, 0.07);',
    'const vec3 ZERAZ  = vec3(0.97, 0.40, 0.24);',
    'const vec3 JANTAR = vec3(1.00, 0.72, 0.40);',
    'const vec3 BIELA  = vec3(1.00, 0.94, 0.88);',
    'float utlm(vec2 uv){',
    '  float zhora = smoothstep(1.02, 0.02, uv.y);',
    '  float okraj = smoothstep(0.0, 0.26, uv.x) * smoothstep(1.0, 0.74, uv.x);',
    '  return zhora * mix(0.35, 1.0, okraj);',
    '}',
    'vec4 zloz(vec3 farba, float sila, float mierka){',
    '  float a = clamp(sila * mierka, 0.0, 0.92) + rozptyl(gl_FragCoord.xy) * 0.012;',
    '  return vec4(farba, clamp(a, 0.0, 1.0));',
    '}',
    ''
  ].join('\n');

  /* Styri varianty. Vyberaju sa cez data-zar na platne.
     Kazda stranka ma mat len jedno platno a len v uvode: pat platien by
     z peknej stranky spravilo vetrak. */
  var VARIANTY = {};

  // stuhy: tri vodorovne pasy svetla rozvlnene sumom, kazdy inou rychlostou.
  // Ostre jadro a mekky rozptyl okolo. Najvyraznejsi variant, pre domovsku stranku.
  VARIANTY.stuhy = [
    'float stuha(vec2 p, float posun, float rychlost, float hrubka, float vlna){',
    '  float y = posun',
    '          + vlna * (fbm(vec2(p.x * 1.15 + cas * rychlost, posun * 7.0)) - 0.5)',
    '          + vlna * 0.45 * (fbm(vec2(p.x * 2.9 - cas * rychlost * 0.6, posun * 3.0)) - 0.5);',
    '  float d = abs(p.y - y);',
    '  return pow(hrubka / (d + hrubka), 2.6);',
    '}',
    'void main(){',
    '  vec2 uv = gl_FragCoord.xy / rozmer.xy;',
    '  vec2 p = uv; p.x *= rozmer.x / rozmer.y;',
    '  float s1 = stuha(p, 0.62, 0.045, 0.055, 0.42);',
    '  float s2 = stuha(p, 0.48, 0.028, 0.090, 0.55);',
    '  float s3 = stuha(p, 0.74, 0.062, 0.035, 0.30);',
    '  float opar = smoothstep(0.42, 0.95, fbm(p * 1.7 + vec2(cas * 0.016, -cas * 0.011))) * 0.5;',
    '  float jas = s1 * 0.85 + s2 * 0.55 + s3 * 0.70 + opar * 0.35;',
    '  vec3 f = mix(UHLIK, ZERAZ, clamp(jas * 1.5, 0.0, 1.0));',
    '  f = mix(f, JANTAR, clamp(jas * 0.95 - 0.35, 0.0, 1.0));',
    '  f = mix(f, BIELA,  clamp(jas * 0.80 - 0.72, 0.0, 1.0));',
    '  gl_FragColor = zloz(f, clamp(jas, 0.0, 1.4) * utlm(uv), 0.72);',
    '}'
  ].join('\n');

  // vlny: hladky prelievany prechod cez dvojite zohnutie suradnic. Pokojnejsi
  // nez stuhy, bez kresby. Pre stranky, kde je vela textu.
  VARIANTY.vlny = [
    'void main(){',
    '  vec2 uv = gl_FragCoord.xy / rozmer.xy;',
    '  vec2 p = uv; p.x *= rozmer.x / rozmer.y;',
    '  float t = cas * 0.045;',
    '  vec2 q = vec2(fbm(p * 1.3 + vec2(0.0, t)), fbm(p * 1.3 + vec2(5.2, 1.3 - t)));',
    '  vec2 r = vec2(fbm(p * 1.3 + 2.4 * q + vec2(1.7, 9.2) + 0.18 * t),',
    '                fbm(p * 1.3 + 2.4 * q + vec2(8.3, 2.8) - 0.15 * t));',
    '  float jadro = clamp(length(q) * 1.15 - 0.34, 0.0, 1.0);',
    '  float teplo = clamp(r.x * 1.05 - 0.30, 0.0, 1.0);',
    '  vec3 f = mix(UHLIK, ZERAZ, jadro);',
    '  f = mix(f, JANTAR, teplo * 0.8);',
    '  gl_FragColor = zloz(f, (jadro * 0.9 + teplo * 0.5) * utlm(uv), 0.62);',
    '}'
  ].join('\n');

  // prach: pole svietiacich bodov, ktore sa unasaju a blikaju. Dava hlbku
  // bez kresby. Pre produktove stranky.
  VARIANTY.prach = [
    'float bod(vec2 p, float mriezka, float rychlost, float posun){',
    '  vec2 g = p * mriezka;',
    '  g.y += cas * rychlost + posun;',
    '  vec2 i = floor(g), f = fract(g);',
    '  float s = sum(i + posun);',
    '  vec2 stred = vec2(0.5) + 0.34 * vec2(sin(s * 17.0 + cas * 0.25), cos(s * 11.0 + cas * 0.19));',
    '  float d = length(f - stred);',
    '  float velkost = mix(0.020, 0.075, fract(s * 7.3));',
    '  float blik = 0.55 + 0.45 * sin(cas * (0.25 + fract(s * 3.1) * 0.4) + s * 30.0);',
    '  return pow(velkost / (d + velkost), 3.2) * blik * step(0.42, s);',
    '}',
    'void main(){',
    '  vec2 uv = gl_FragCoord.xy / rozmer.xy;',
    '  vec2 p = uv; p.x *= rozmer.x / rozmer.y;',
    '  float b = bod(p, 7.0, 0.020, 0.0) * 0.9',
    '          + bod(p, 12.0, 0.032, 3.7) * 0.6',
    '          + bod(p, 19.0, 0.048, 8.1) * 0.35;',
    '  float zaves = smoothstep(0.30, 0.95, fbm(p * 1.1 + vec2(cas * 0.012, 0.0))) * 0.42;',
    '  vec3 f = mix(UHLIK, JANTAR, clamp(b * 1.1, 0.0, 1.0));',
    '  f = mix(f, BIELA, clamp(b * 0.7 - 0.45, 0.0, 1.0));',
    '  gl_FragColor = zloz(f, (b * 0.85 + zaves * 0.5) * utlm(uv), 0.80);',
    '}'
  ].join('\n');

  // luc: jeden siroky sikmy pruh svetla, ktory velmi pomaly prechadza plochou.
  // Najpokojnejsi variant, takmer staticky. Pre pravne texty a dokumentaciu.
  VARIANTY.luc = [
    'void main(){',
    '  vec2 uv = gl_FragCoord.xy / rozmer.xy;',
    '  vec2 p = uv; p.x *= rozmer.x / rozmer.y;',
    '  float uhol = -0.55;',
    '  float os = p.x * cos(uhol) - p.y * sin(uhol);',
    '  float stred = 0.15 + 0.5 * sin(cas * 0.021);',
    '  float d = abs(os - stred);',
    '  float luc = pow(0.30 / (d + 0.30), 3.4);',
    '  float zrno = fbm(p * 2.4 + vec2(cas * 0.010, -cas * 0.008));',
    '  float jas = luc * (0.65 + 0.5 * zrno);',
    '  vec3 f = mix(UHLIK, ZERAZ, clamp(jas * 1.25, 0.0, 1.0));',
    '  f = mix(f, JANTAR, clamp(jas * 0.7 - 0.30, 0.0, 1.0));',
    '  gl_FragColor = zloz(f, jas * utlm(uv), 0.52);',
    '}'
  ].join('\n');

  var nazov = (plocha.getAttribute('data-zar') || 'stuhy').toLowerCase();
  var FRAGMENT = ZAKLAD + (VARIANTY[nazov] || VARIANTY.stuhy);

  function shader(typ, zdroj) {
    var s = gl.createShader(typ);
    gl.shaderSource(s, zdroj);
    gl.compileShader(s);
    if (!gl.getShaderParameter(s, gl.COMPILE_STATUS)) return null;
    return s;
  }

  var vs = shader(gl.VERTEX_SHADER, VRCHOL);
  var fs = shader(gl.FRAGMENT_SHADER, FRAGMENT);
  if (!vs || !fs) { plocha.remove(); return; }

  var program = gl.createProgram();
  gl.attachShader(program, vs);
  gl.attachShader(program, fs);
  gl.linkProgram(program);
  if (!gl.getProgramParameter(program, gl.LINK_STATUS)) { plocha.remove(); return; }
  gl.useProgram(program);

  var buf = gl.createBuffer();
  gl.bindBuffer(gl.ARRAY_BUFFER, buf);
  gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 3, -1, -1, 3]), gl.STATIC_DRAW);
  var loc = gl.getAttribLocation(program, 'p');
  gl.enableVertexAttribArray(loc);
  gl.vertexAttribPointer(loc, 2, gl.FLOAT, false, 0, 0);

  var uRozmer = gl.getUniformLocation(program, 'rozmer');
  var uCas = gl.getUniformLocation(program, 'cas');
  gl.enable(gl.BLEND);
  gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);

  function zmenRozmer() {
    var r = plocha.getBoundingClientRect();
    // Na telefóne kreslíme na polovičnom rozlíšení: shader je plynulý gradient,
    // takže rozdiel nie je vidieť, a počítania je štvrtina.
    var hustota = Math.min(window.devicePixelRatio || 1, r.width < 720 ? 1 : 1.5);
    var w = Math.max(1, Math.round(r.width * hustota));
    var h = Math.max(1, Math.round(r.height * hustota));
    if (plocha.width !== w || plocha.height !== h) {
      plocha.width = w; plocha.height = h;
      gl.viewport(0, 0, w, h);
    }
    gl.uniform2f(uRozmer, plocha.width, plocha.height);
  }

  var zaciatok = null;
  var bezi = false;
  var vidno = true;
  var poslednaKresba = 0;
  var SNIMOK = 1000 / 30;   // 30 za sekundu stačí, 60 by len hrialo batériu

  function kresli(teraz) {
    if (zaciatok === null) zaciatok = teraz;
    gl.uniform1f(uCas, (teraz - zaciatok) / 1000);
    gl.drawArrays(gl.TRIANGLES, 0, 3);
  }

  function slucka(teraz) {
    if (!bezi) return;
    if (teraz - poslednaKresba >= SNIMOK) { poslednaKresba = teraz; kresli(teraz); }
    window.requestAnimationFrame(slucka);
  }

  function spusti() {
    if (bezi || tichy || !vidno || document.hidden) return;
    bezi = true;
    window.requestAnimationFrame(slucka);
  }
  function zastav() { bezi = false; }

  zmenRozmer();
  // Prvy snimok vzdy, aj v tichom rezime. V tichom pridame posun, lebo
  // v case 0 je sum najplochejsi a obraz by bol takmer prazdny.
  kresli(performance.now() + (tichy ? 9000 : 0));

  if (!tichy) {
    if ('IntersectionObserver' in window) {
      new IntersectionObserver(function (z) {
        vidno = z[0].isIntersecting;
        if (vidno) spusti(); else zastav();
      }, { threshold: 0 }).observe(plocha);
    }
    document.addEventListener('visibilitychange', function () {
      if (document.hidden) zastav(); else spusti();
    });
    spusti();
  }

  var cakaNaRozmer = false;
  window.addEventListener('resize', function () {
    if (cakaNaRozmer) return;
    cakaNaRozmer = true;
    window.requestAnimationFrame(function () {
      cakaNaRozmer = false;
      zmenRozmer();
      if (!bezi) kresli(performance.now() + (tichy ? 9000 : 0));
    });
  }, { passive: true });
})();
