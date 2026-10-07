'use strict';

// Motion enhances a fully rendered page; it never carries content or navigation.
const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)');
if (!reducedMotion.matches && 'IntersectionObserver' in window) {
  const observer = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.remove('awaiting-reveal');
        entry.target.classList.add('revealed');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.08 });
  document.documentElement.classList.add('motion-ready');
  document.querySelectorAll('.feature-card, .section-heading, .story-text, .workflow-band .screenshot-frame, .getting-started').forEach(element => {
    element.dataset.reveal = '';
    element.classList.add('awaiting-reveal');
    observer.observe(element);
  });
  reducedMotion.addEventListener('change', event => {
    if (event.matches) {
      observer.disconnect();
      document.documentElement.classList.remove('motion-ready');
    }
  });
}

const scene = document.querySelector('.hero-visual');
if (scene && window.matchMedia('(hover: hover) and (pointer: fine)').matches) {
  scene.addEventListener('pointermove', event => {
    if (reducedMotion.matches) return;
    const bounds = scene.getBoundingClientRect();
    scene.style.setProperty('--tilt-y', `${((event.clientX - bounds.left) / bounds.width - 0.5) * 5}deg`);
    scene.style.setProperty('--tilt-x', `${-((event.clientY - bounds.top) / bounds.height - 0.5) * 4}deg`);
  });
  const resetTilt = () => {
    scene.style.setProperty('--tilt-x', '0deg');
    scene.style.setProperty('--tilt-y', '0deg');
  };
  scene.addEventListener('pointerleave', resetTilt);
  reducedMotion.addEventListener('change', resetTilt);
}

// Screenshots stay readable links when JavaScript is unavailable.
const viewer = document.querySelector('.image-dialog');
if (viewer && typeof viewer.showModal === 'function') {
  const image = viewer.querySelector('img');
  const title = viewer.querySelector('#image-dialog-title');
  document.querySelectorAll('[data-screenshot]').forEach(link => {
    link.addEventListener('click', event => {
      event.preventDefault();
      image.src = link.href;
      image.alt = link.querySelector('img').alt;
      viewer.querySelector('.dialog-image').style.cssText = link.dataset.framing;
      title.textContent = link.dataset.screenshot;
      viewer.showModal();
    });
  });
  viewer.querySelector('.dialog-close').addEventListener('click', () => viewer.close());
  viewer.addEventListener('click', event => {
    if (event.target === viewer) {
      const box = viewer.getBoundingClientRect();
      if (event.clientX < box.left || event.clientX > box.right ||
          event.clientY < box.top || event.clientY > box.bottom) viewer.close();
    }
  });
}
