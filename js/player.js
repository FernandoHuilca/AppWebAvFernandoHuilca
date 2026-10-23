/* ============================================================
   PoliMusic - Reproductor (Vanilla JS)
   Depende de window.POLIMUSIC_CTX = { base, songs[] }
   ============================================================ */
(function () {
  'use strict';

  function $(id) { return document.getElementById(id); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fmt(sec) {
    if (!isFinite(sec)) sec = 0;
    sec = Math.max(0, Math.floor(sec));
    return Math.floor(sec / 60) + ':' + pad(sec % 60);
  }

  var ctx = window.POLIMUSIC_CTX || {};
  var base = ctx.base || '/';
  var songs = (window.POLIMUSIC_SONGS !== undefined ? window.POLIMUSIC_SONGS : ctx.songs) || [];
  var audio = $('audioPlayer');
  if (!audio) return;

  var cards = Array.prototype.slice.call(document.querySelectorAll('.song-card'));
  var elTitle = $('playerTitle'), elArtist = $('playerArtist'), elArt = $('playerArt');
  var btnPlay = $('btnPlay'), icoPlay = $('icoPlay'), icoPause = $('icoPause');
  var btnPrev = $('btnPrev'), btnNext = $('btnNext');
  var btnShuffle = $('btnShuffle'), btnRepeat = $('btnRepeat');
  var seekBar = $('seekBar'), timeNow = $('timeNow'), timeLeft = $('timeLeft'), volBar = $('volBar');
  var bar = $('playerBar');

  var current = -1;
  var shuffle = false;
  var repeat = false;
  var played = {};
  var toastTimer = null;

  function toast(msg) {
    var t = $('toast');
    if (!t) return;
    t.textContent = msg;
    t.classList.add('show');
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 3400);
  }

  function setPlaying(on) {
    if (icoPlay && icoPause) {
      icoPlay.style.display = on ? 'none' : '';
      icoPause.style.display = on ? '' : 'none';
    }
    if (btnPlay) btnPlay.setAttribute('aria-pressed', on ? 'true' : 'false');
    if (bar) bar.classList.toggle('is-playing', on);
  }

  function markCard() {
    cards.forEach(function (c) {
      var on = parseInt(c.getAttribute('data-index'), 10) === current;
      c.classList.toggle('is-playing', on);
    });
  }

  function hasUrl(i) { return i >= 0 && i < songs.length && !!songs[i].u; }

  function nextAvailable(from, dir) {
    if (songs.length === 0) return -1;
    var steps = 0, i = from;
    while (steps < songs.length) {
      i = (i + dir + songs.length) % songs.length;
      if (hasUrl(i)) return i;
      steps++;
    }
    return -1;
  }

  function updateMeta(s) {
    if (elTitle) elTitle.textContent = s.t || '(sin nombre)';
    if (elArtist) elArtist.textContent = (s.a && s.a.length) ? s.a : 'Artista desconocido';
    if (elArt) elArt.src = (s.c && s.c.length) ? s.c : (base + 'assets/placeholder.svg');
  }

  function bumpPlays(songId) {
    var counter = document.querySelector('span[data-plays="' + songId + '"]');
    if (counter) counter.textContent = String((parseInt(counter.textContent, 10) || 0) + 1);
    var url = base + 'song_plays.asp?id=' + encodeURIComponent(String(songId));
    if (typeof fetch === 'function') {
      fetch(url, { method: 'POST', credentials: 'same-origin' }).catch(function () {});
    }
  }

  function loadSong(i, autoplay) {
    if (i < 0 || i >= songs.length) return;
    current = i;
    markCard();
    updateMeta(songs[i]);

    if (!songs[i].u) {
      var other = nextAvailable(current, 1);
      if (other >= 0 && other !== current) {
        toast('No hay audio para "' + songs[i].t + '". Pasando a la siguiente disponible.');
        loadSong(other, autoplay);
      } else {
        audio.removeAttribute('src');
        try { audio.load(); } catch (e) {}
        setPlaying(false);
        toast('No hay archivos de audio disponibles en la carpeta SongFiles.');
      }
      return;
    }

    audio.src = songs[i].u;
    if (autoplay) {
      var pr = audio.play();
      if (pr && pr.catch) pr.catch(function () { setPlaying(false); });
    }
  }

  /* ------------------------------------------------ eventos del audio */
  audio.addEventListener('play', function () {
    setPlaying(true);
    if (current >= 0 && !played[current]) {
      played[current] = true;
      bumpPlays(songs[current].id);
    }
  });
  audio.addEventListener('pause', function () { setPlaying(false); });

  audio.addEventListener('ended', function () {
    var nx = nextAvailable(current, 1);
    if (nx < 0) { setPlaying(false); return; }
    if (repeat && (nx === current || songs.length === 1)) {
      audio.currentTime = 0;
      var rp = audio.play();
      if (rp && rp.catch) rp.catch(function () {});
    } else {
      loadSong(nx, true);
    }
  });

  audio.addEventListener('error', function () {
    if (current < 0) return;
    var s = songs[current];
    setPlaying(false);
    var nx = nextAvailable(current, 1);
    if (nx >= 0 && nx !== current) {
      toast('El audio de "' + s.t + '" no esta disponible. Pasando a la siguiente.');
      loadSong(nx, true);
    } else {
      toast('El archivo "' + s.t + '" no esta disponible en la carpeta de audios.');
    }
  });

  audio.addEventListener('timeupdate', function () {
    if (isFinite(audio.duration) && audio.duration > 0) {
      seekBar.max = Math.floor(audio.duration);
      if (!seekDragging) seekBar.value = Math.floor(audio.currentTime);
      if (timeLeft) timeLeft.textContent = fmt(audio.duration - audio.currentTime);
    }
    if (timeNow) timeNow.textContent = fmt(audio.currentTime);
  });
  audio.addEventListener('loadedmetadata', function () {
    if (isFinite(audio.duration)) seekBar.max = Math.floor(audio.duration);
  });

  /* ------------------------------------------------ controles */
  var seekDragging = false;

  function toggleCurrent() {
    if (current < 0) { loadSong(0, true); return; }
    if (!audio.src) { loadSong(current, true); return; }
    if (audio.paused) {
      var p = audio.play();
      if (p && p.catch) p.catch(function () {});
    } else {
      audio.pause();
    }
  }

  if (btnPlay) btnPlay.addEventListener('click', toggleCurrent);
  if (btnNext) btnNext.addEventListener('click', function () {
    var nx = nextAvailable(current, 1);
    if (nx >= 0) loadSong(nx, true);
  });
  if (btnPrev) btnPrev.addEventListener('click', function () {
    if (audio.currentTime > 3 && current >= 0) {
      audio.currentTime = 0;
    } else {
      var pv = nextAvailable(current, -1);
      if (pv >= 0) loadSong(pv, true);
    }
  });

  if (btnShuffle) btnShuffle.addEventListener('click', function () {
    shuffle = !shuffle;
    btnShuffle.classList.toggle('is-on', shuffle);
    btnShuffle.setAttribute('aria-pressed', shuffle ? 'true' : 'false');
    toast(shuffle ? 'Modo aleatorio activado.' : 'Modo aleatorio desactivado.');
  });
  if (btnRepeat) btnRepeat.addEventListener('click', function () {
    repeat = !repeat;
    btnRepeat.classList.toggle('is-on', repeat);
    btnRepeat.setAttribute('aria-pressed', repeat ? 'true' : 'false');
    toast(repeat ? 'Modo repetir activado.' : 'Modo repetir desactivado.');
  });

  seekBar.addEventListener('input', function () {
    seekDragging = true;
    if (timeNow) timeNow.textContent = fmt(parseInt(seekBar.value, 10));
  });
  seekBar.addEventListener('change', function () {
    seekDragging = false;
    if (isFinite(audio.duration)) audio.currentTime = parseInt(seekBar.value, 10);
  });
  if (volBar) {
    audio.volume = 1;
    volBar.addEventListener('input', function () {
      audio.volume = parseInt(volBar.value, 10) / 100;
    });
  }

  /* ------------------------------------------------ tarjetas */
  cards.forEach(function (c) {
    var idx = parseInt(c.getAttribute('data-index'), 10);
    c.addEventListener('click', function () {
      if (idx === current && audio.src) {
        toggleCurrent();
      } else {
        loadSong(idx, true);
      }
    });
    c.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        if (idx === current && audio.src) { toggleCurrent(); } else { loadSong(idx, true); }
      }
    });
  });

  /* ------------------------------------------------ buscador */
  var searchInput = $('songSearch');
  if (searchInput) {
    searchInput.addEventListener('input', function () {
      var q = searchInput.value.trim().toLowerCase();
      cards.forEach(function (c) {
        var text = (c.innerText || '').toLowerCase();
        c.style.display = (q === '' || text.indexOf(q) !== -1) ? '' : 'none';
      });
    });
  }

  markCard();
})();
