/* ALPHA BRAIN — サイト共通スクリプト
   全ページ共通：ナビの背景切替／モバイルメニュー／スクロール表示
   トップ専用の処理（4分割ヒーロー・CTAキャンバス）は要素の有無で判定して動く。 */

(function () {
    'use strict';

    /* ---- ナビ：スクロールしたら背景を敷く ---- */
    var nav = document.getElementById('nav');
    if (nav) {
        var onScroll = function () { nav.classList.toggle('scrolled', window.scrollY > 60); };
        window.addEventListener('scroll', onScroll, { passive: true });
        // 下層ページは先頭バナーが短いので、読み込み時点の位置で判定しておく
        onScroll();
    }

    /* ---- モバイルメニュー ---- */
    window.toggleMenu = function () {
        var m = document.getElementById('mobile-menu');
        if (m) m.style.display = m.style.display === 'block' ? 'none' : 'block';
    };
    window.closeMenu = function () {
        var m = document.getElementById('mobile-menu');
        if (m) m.style.display = 'none';
    };

    /* ---- 現在ページをナビでハイライト ---- */
    (function () {
        var page = (window.location.pathname.split('/').pop() || 'index.html');
        document.querySelectorAll('[data-page]').forEach(function (el) {
            if (el.getAttribute('data-page') === page) el.classList.add('active');
        });
    })();

    /* ---- スクロールで現れる ---- */
    (function () {
        var els = document.querySelectorAll('.reveal');
        if (!els.length) return;
        if (!('IntersectionObserver' in window)) {
            els.forEach(function (el) { el.classList.add('visible'); });
            return;
        }
        var obs = new IntersectionObserver(function (entries) {
            entries.forEach(function (e) {
                if (e.isIntersecting) { e.target.classList.add('visible'); obs.unobserve(e.target); }
            });
        }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });
        els.forEach(function (el) { obs.observe(el); });
    })();

    /* ---- トップ：モバイルのヒーロースライドショー（ズーム＋ホワイトアウト） ---- */
    (function () {
        var panels = document.querySelectorAll('.hero-panel');
        var flash = document.getElementById('hero-flash');
        if (!panels.length || !flash) return;
        var current = 0, timer = null;
        function isMobile() { return window.innerWidth <= 820; }
        function activate(idx) {
            panels.forEach(function (p) { p.classList.remove('active'); });
            void panels[idx].offsetWidth; // アニメーションを頭から流し直す
            panels[idx].classList.add('active');
        }
        function initSlide() {
            if (!isMobile()) {
                panels.forEach(function (p) { p.classList.remove('active'); });
                if (timer) { clearInterval(timer); timer = null; }
                return;
            }
            activate(0);
            current = 0;
            if (timer) clearInterval(timer);
            timer = setInterval(function () {
                if (!isMobile()) return;
                flash.style.transition = 'none';
                flash.style.opacity = '0';
                void flash.offsetWidth;
                flash.style.transition = 'opacity 0.35s ease';
                flash.style.opacity = '1';
                setTimeout(function () {
                    panels[current].classList.remove('active');
                    current = (current + 1) % panels.length;
                    activate(current);
                    flash.style.opacity = '0';
                }, 340);
            }, 5000);
        }
        initSlide();
        window.addEventListener('resize', initSlide);
    })();

    /* ---- CTA の背景キャンバス（金色のにじみ） ---- */
    (function () {
        var canvas = document.getElementById('cta-canvas');
        if (!canvas) return;
        var ctx = canvas.getContext('2d'), t = 0;
        function draw() {
            var W = canvas.width = canvas.offsetWidth;
            var H = canvas.height = canvas.offsetHeight;
            ctx.clearRect(0, 0, W, H);
            t += 0.005;
            for (var i = 0; i < 5; i++) {
                var x = W * (0.1 + i * 0.2 + Math.sin(t + i) * 0.05);
                var y = H * (0.3 + Math.cos(t * 0.7 + i * 1.3) * 0.25);
                var g = ctx.createRadialGradient(x, y, 0, x, y, 240);
                g.addColorStop(0, 'rgba(184,151,58,0.06)');
                g.addColorStop(1, 'transparent');
                ctx.fillStyle = g;
                ctx.beginPath(); ctx.arc(x, y, 240, 0, Math.PI * 2); ctx.fill();
            }
            requestAnimationFrame(draw);
        }
        draw();
    })();
})();
