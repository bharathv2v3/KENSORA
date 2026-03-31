// =============================================
// KENSORA — Main JavaScript
// =============================================

document.addEventListener('DOMContentLoaded', () => {

  // ─── Navbar ───
  const navbar = document.querySelector('.navbar');
  const hamburger = document.querySelector('.navbar__hamburger');
  const mobileNav = document.querySelector('.navbar__mobile');

  const updateNavbar = () => {
    if (!navbar) return;
    if (window.scrollY > 60) {
      navbar.classList.add('scrolled');
      navbar.classList.remove('transparent');
    } else {
      navbar.classList.remove('scrolled');
      if (navbar.dataset.transparent === 'true') navbar.classList.add('transparent');
    }
  };

  window.addEventListener('scroll', updateNavbar, { passive: true });
  if (navbar && navbar.dataset.transparent === 'true') navbar.classList.add('transparent');
  updateNavbar();

  if (hamburger && mobileNav) {
    hamburger.addEventListener('click', () => {
      hamburger.classList.toggle('open');
      mobileNav.classList.toggle('open');
    });
    mobileNav.querySelectorAll('a').forEach(a => {
      a.addEventListener('click', () => {
        hamburger.classList.remove('open');
        mobileNav.classList.remove('open');
      });
    });
  }

  // Active nav link
  const currentPage = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.navbar__nav a, .navbar__mobile a').forEach(link => {
    const href = link.getAttribute('href');
    if (href && (href === currentPage || (currentPage === '' && href === 'index.html'))) {
      link.classList.add('active');
    }
  });

  // ─── Scroll Reveal ───
  const revealObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        revealObserver.unobserve(entry.target);
      }
    });
  }, { threshold: 0.12, rootMargin: '0px 0px -60px 0px' });

  document.querySelectorAll('.reveal').forEach(el => revealObserver.observe(el));

  // ─── Animated Counters ───
  const animateCounter = (el) => {
    const target = parseInt(el.dataset.target, 10);
    const suffix = el.dataset.suffix || '';
    const duration = 2000;
    const step = target / (duration / 16);
    let current = 0;
    const timer = setInterval(() => {
      current = Math.min(current + step, target);
      el.textContent = Math.floor(current).toLocaleString() + suffix;
      if (current >= target) clearInterval(timer);
    }, 16);
  };

  const counterObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        animateCounter(entry.target);
        counterObserver.unobserve(entry.target);
      }
    });
  }, { threshold: 0.5 });

  document.querySelectorAll('[data-target]').forEach(el => counterObserver.observe(el));

  // ─── Testimonials Carousel ───
  const carousel = document.querySelector('.testimonials-track');
  if (carousel) {
    let idx = 0;
    const slides = carousel.querySelectorAll('.testimonial-card');
    const dots = document.querySelectorAll('.testi-dot');
    const prevBtn = document.querySelector('.testi-prev');
    const nextBtn = document.querySelector('.testi-next');

    const goTo = (n) => {
      slides[idx].classList.remove('active');
      if (dots[idx]) dots[idx].classList.remove('active');
      idx = (n + slides.length) % slides.length;
      slides[idx].classList.add('active');
      if (dots[idx]) dots[idx].classList.add('active');
    };

    slides[0]?.classList.add('active');
    dots[0]?.classList.add('active');
    if (prevBtn) prevBtn.addEventListener('click', () => goTo(idx - 1));
    if (nextBtn) nextBtn.addEventListener('click', () => goTo(idx + 1));
    dots.forEach((dot, i) => dot.addEventListener('click', () => goTo(i)));
    setInterval(() => goTo(idx + 1), 5000);
  }

  // ─── Page Hero Ken Burns ───
  const heroBg = document.querySelector('.page-hero__bg');
  if (heroBg) setTimeout(() => heroBg.closest('.page-hero')?.classList.add('loaded'), 100);

  // ─── Cursor Glow (desktop only) ───
  if (window.matchMedia('(pointer:fine)').matches) {
    const glow = document.createElement('div');
    glow.className = 'cursor-glow';
    document.body.appendChild(glow);
    document.addEventListener('mousemove', e => {
      glow.style.left = e.clientX + 'px';
      glow.style.top  = e.clientY + 'px';
    });
  }

  // ─── Smooth Anchor Scroll ───
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', e => {
      const target = document.querySelector(anchor.getAttribute('href'));
      if (target) {
        e.preventDefault();
        const offset = document.querySelector('.navbar')?.offsetHeight || 80;
        window.scrollTo({ top: target.offsetTop - offset, behavior: 'smooth' });
      }
    });
  });

  // ─── Form multi-step ───
  const steps = document.querySelectorAll('.form-step');
  const stepDots = document.querySelectorAll('.step-dot');
  let currentStep = 0;

  const showStep = (n) => {
    steps.forEach((s, i) => {
      s.classList.toggle('active', i === n);
      if (stepDots[i]) stepDots[i].classList.toggle('active', i <= n);
    });
    currentStep = n;
  };

  document.querySelectorAll('.btn-next-step').forEach(btn => {
    btn.addEventListener('click', () => {
      if (currentStep < steps.length - 1) showStep(currentStep + 1);
    });
  });
  document.querySelectorAll('.btn-prev-step').forEach(btn => {
    btn.addEventListener('click', () => {
      if (currentStep > 0) showStep(currentStep - 1);
    });
  });
  if (steps.length) showStep(0);

  // ─── Contact form submit ───
  const contactForm = document.getElementById('consultationForm');
  if (contactForm) {
    contactForm.addEventListener('submit', e => {
      e.preventDefault();
      const successDiv = document.getElementById('formSuccess');
      if (successDiv) {
        contactForm.style.display = 'none';
        successDiv.style.display = 'block';
      }
    });
  }

});

// ─── Cursor Glow Styles (injected) ───
const glowStyles = document.createElement('style');
glowStyles.textContent = `.cursor-glow{pointer-events:none;position:fixed;width:300px;height:300px;border-radius:50%;background:radial-gradient(circle,rgba(181,154,106,0.06) 0%,transparent 70%);transform:translate(-50%,-50%);z-index:9999;transition:left 0.2s ease,top 0.2s ease;}`;
document.head.appendChild(glowStyles);
